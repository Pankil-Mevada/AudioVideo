#include <systemc>
#include <tlm>
#include <tlm_utils/simple_initiator_socket.h>
#include <tlm_utils/simple_target_socket.h>

#include <array>
#include <cstdint>
#include <cstring>
#include <iomanip>
#include <iostream>

using namespace sc_core;
using namespace tlm;

namespace map {
constexpr std::uint64_t CAN_BASE = 0x4000;
constexpr std::uint64_t RAM_BASE = 0x8000;
constexpr std::uint64_t REGION_SIZE = 0x1000;
} // namespace map

// A tiny memory-mapped CAN controller.
//   0x00 CONTROL: write bit 0 to enable
//   0x04 TX_ID:   CAN identifier
//   0x08 TX_DATA: four payload bytes
//   0x0C COMMAND: write bit 0 to transmit
//   0x10 STATUS:  bit 0 means transmit complete
class CanController : public sc_module {
public:
    tlm_utils::simple_target_socket<CanController> socket{"socket"};
    sc_out<bool> irq{"irq"};

    SC_CTOR(CanController) {
        socket.register_b_transport(this, &CanController::b_transport);
        SC_THREAD(transmit_worker);
    }

private:
    enum Register : std::uint64_t {
        CONTROL = 0x00,
        TX_ID = 0x04,
        TX_DATA = 0x08,
        COMMAND = 0x0C,
        STATUS = 0x10,
    };

    std::uint32_t control_ = 0;
    std::uint32_t tx_id_ = 0;
    std::uint32_t tx_data_ = 0;
    std::uint32_t status_ = 0;
    sc_event transmit_requested_;

    void b_transport(tlm_generic_payload& tx, sc_time& delay) {
        if (tx.get_data_length() != sizeof(std::uint32_t) ||
            tx.get_data_ptr() == nullptr) {
            tx.set_response_status(TLM_BURST_ERROR_RESPONSE);
            return;
        }

        std::uint32_t value = 0;
        const auto offset = tx.get_address();

        // Both the bus and device add latency to the transaction. The
        // initiator consumes the accumulated delay after b_transport returns.
        delay += sc_time(20, SC_NS);

        if (tx.is_write()) {
            std::memcpy(&value, tx.get_data_ptr(), sizeof(value));
            switch (offset) {
            case CONTROL: control_ = value; break;
            case TX_ID: tx_id_ = value; break;
            case TX_DATA: tx_data_ = value; break;
            case COMMAND:
                if ((value & 1U) != 0U && (control_ & 1U) != 0U) {
                    status_ = 0;
                    transmit_requested_.notify(delay);
                }
                break;
            default:
                tx.set_response_status(TLM_ADDRESS_ERROR_RESPONSE);
                return;
            }
        } else if (tx.is_read()) {
            switch (offset) {
            case CONTROL: value = control_; break;
            case TX_ID: value = tx_id_; break;
            case TX_DATA: value = tx_data_; break;
            case STATUS: value = status_; break;
            default:
                tx.set_response_status(TLM_ADDRESS_ERROR_RESPONSE);
                return;
            }
            std::memcpy(tx.get_data_ptr(), &value, sizeof(value));
        } else {
            tx.set_response_status(TLM_COMMAND_ERROR_RESPONSE);
            return;
        }

        tx.set_dmi_allowed(false);
        tx.set_response_status(TLM_OK_RESPONSE);
    }

    void transmit_worker() {
        irq.write(false);
        while (true) {
            wait(transmit_requested_);
            wait(100, SC_US); // modeled CAN transmission time

            std::cout << sc_time_stamp() << " CAN TX id=0x" << std::hex
                      << tx_id_ << " data=0x" << tx_data_ << std::dec << '\n';

            status_ = 1;
            irq.write(true);
            wait(1, SC_US);
            irq.write(false);
        }
    }
};

class Ram : public sc_module {
public:
    tlm_utils::simple_target_socket<Ram> socket{"socket"};

    SC_CTOR(Ram) {
        socket.register_b_transport(this, &Ram::b_transport);
        bytes_.fill(0);
    }

private:
    std::array<unsigned char, map::REGION_SIZE> bytes_{};

    void b_transport(tlm_generic_payload& tx, sc_time& delay) {
        const auto address = tx.get_address();
        const auto length = tx.get_data_length();

        if (tx.get_data_ptr() == nullptr || address + length > bytes_.size()) {
            tx.set_response_status(TLM_ADDRESS_ERROR_RESPONSE);
            return;
        }

        if (tx.is_write()) {
            std::memcpy(bytes_.data() + address, tx.get_data_ptr(), length);
        } else if (tx.is_read()) {
            std::memcpy(tx.get_data_ptr(), bytes_.data() + address, length);
        } else {
            tx.set_response_status(TLM_COMMAND_ERROR_RESPONSE);
            return;
        }

        delay += sc_time(10, SC_NS);
        tx.set_response_status(TLM_OK_RESPONSE);
    }
};

// Address-decoding bus: converts system addresses into target-local offsets.
class Bus : public sc_module {
public:
    tlm_utils::simple_target_socket<Bus> cpu_socket{"cpu_socket"};
    tlm_utils::simple_initiator_socket<Bus> can_socket{"can_socket"};
    tlm_utils::simple_initiator_socket<Bus> ram_socket{"ram_socket"};

    SC_CTOR(Bus) {
        cpu_socket.register_b_transport(this, &Bus::b_transport);
    }

private:
    void b_transport(tlm_generic_payload& tx, sc_time& delay) {
        const auto system_address = tx.get_address();
        delay += sc_time(5, SC_NS); // interconnect latency

        if (system_address >= map::CAN_BASE &&
            system_address < map::CAN_BASE + map::REGION_SIZE) {
            tx.set_address(system_address - map::CAN_BASE);
            can_socket->b_transport(tx, delay);
        } else if (system_address >= map::RAM_BASE &&
                   system_address < map::RAM_BASE + map::REGION_SIZE) {
            tx.set_address(system_address - map::RAM_BASE);
            ram_socket->b_transport(tx, delay);
        } else {
            tx.set_response_status(TLM_ADDRESS_ERROR_RESPONSE);
        }

        tx.set_address(system_address); // restore address for the initiator
    }
};

class Cpu : public sc_module {
public:
    tlm_utils::simple_initiator_socket<Cpu> socket{"socket"};
    sc_in<bool> irq{"irq"};

    SC_CTOR(Cpu) {
        SC_THREAD(software);
        SC_METHOD(interrupt_handler);
        sensitive << irq.pos();
        dont_initialize();
    }

private:
    void software() {
        wait(1, SC_US); // reset/boot time

        std::cout << sc_time_stamp() << " CPU: enable CAN\n";
        write32(map::CAN_BASE + 0x00, 1);
        write32(map::CAN_BASE + 0x04, 0x123);
        write32(map::CAN_BASE + 0x08, 0xDEADBEEF);

        std::cout << sc_time_stamp() << " CPU: request CAN transmission\n";
        write32(map::CAN_BASE + 0x0C, 1);

        wait(irq.posedge_event());
        const auto status = read32(map::CAN_BASE + 0x10);
        write32(map::RAM_BASE, status); // application records the result

        std::cout << sc_time_stamp() << " CPU: status=" << status
                  << ", saved to RAM; test passed\n";
        sc_stop();
    }

    void interrupt_handler() {
        std::cout << sc_time_stamp() << " CPU: CAN interrupt observed\n";
    }

    void write32(std::uint64_t address, std::uint32_t value) {
        transport(TLM_WRITE_COMMAND, address, value);
    }

    std::uint32_t read32(std::uint64_t address) {
        std::uint32_t value = 0;
        transport(TLM_READ_COMMAND, address, value);
        return value;
    }

    void transport(tlm_command command, std::uint64_t address,
                   std::uint32_t& value) {
        tlm_generic_payload tx;
        tx.set_command(command);
        tx.set_address(address);
        tx.set_data_ptr(reinterpret_cast<unsigned char*>(&value));
        tx.set_data_length(sizeof(value));
        tx.set_streaming_width(sizeof(value));
        tx.set_byte_enable_ptr(nullptr);
        tx.set_dmi_allowed(false);
        tx.set_response_status(TLM_INCOMPLETE_RESPONSE);

        sc_time delay = SC_ZERO_TIME;
        socket->b_transport(tx, delay);
        wait(delay); // consume the delay accumulated by bus and target

        if (tx.is_response_error()) {
            SC_REPORT_FATAL("CPU", tx.get_response_string().c_str());
        }
    }
};

int sc_main(int, char**) {
    Cpu cpu{"cpu"};
    Bus bus{"bus"};
    CanController can{"can"};
    Ram ram{"ram"};
    sc_signal<bool> can_irq{"can_irq"};

    cpu.socket.bind(bus.cpu_socket);
    bus.can_socket.bind(can.socket);
    bus.ram_socket.bind(ram.socket);
    can.irq(can_irq);
    cpu.irq(can_irq);

    sc_start();
    return 0;
}

