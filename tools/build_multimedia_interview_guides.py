from pathlib import Path
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "documents"
OUT.mkdir(exist_ok=True)

NAVY = "17365D"
PALE = "EAF1F8"
LIGHT = "F3F5F7"
GRID = "D9D9D9"
BLACK = RGBColor(0, 0, 0)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, color=GRID, size="6"):
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = borders.find(qn(f"w:{edge}"))
        if tag is None:
            tag = OxmlElement(f"w:{edge}")
            borders.append(tag)
        tag.set(qn("w:val"), "single")
        tag.set(qn("w:sz"), size)
        tag.set(qn("w:color"), color)


def repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    cant_split.set(qn("w:val"), "true")
    tr_pr.append(cant_split)


def configure_document(doc, short_title):
    section = doc.sections[0]
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    section.left_margin = Inches(0.78)
    section.right_margin = Inches(0.78)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Aptos"
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Aptos")
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Aptos")
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(5)
    normal.paragraph_format.line_spacing = 1.08

    for name, size, before, after in (
        ("Title", 25, 0, 12),
        ("Heading 1", 18, 14, 7),
        ("Heading 2", 14, 10, 5),
        ("Heading 3", 11.5, 7, 3),
    ):
        style = styles[name]
        style.font.name = "Aptos Display" if name != "Normal" else "Aptos"
        style._element.rPr.rFonts.set(qn("w:ascii"), style.font.name)
        style._element.rPr.rFonts.set(qn("w:hAnsi"), style.font.name)
        style.font.size = Pt(size)
        style.font.bold = name != "Title" or True
        style.font.color.rgb = BLACK
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True
        if name == "Title":
            p_pr = style._element.get_or_add_pPr()
            border = p_pr.find(qn("w:pBdr"))
            if border is not None:
                p_pr.remove(border)

    if "Code" not in styles:
        code_style = styles.add_style("Code", 1)
    else:
        code_style = styles["Code"]
    code_style.font.name = "Consolas"
    code_style._element.rPr.rFonts.set(qn("w:ascii"), "Consolas")
    code_style._element.rPr.rFonts.set(qn("w:hAnsi"), "Consolas")
    code_style.font.size = Pt(8)
    code_style.paragraph_format.left_indent = Inches(0.18)
    code_style.paragraph_format.right_indent = Inches(0.08)
    code_style.paragraph_format.space_before = Pt(3)
    code_style.paragraph_format.space_after = Pt(7)
    code_style.paragraph_format.line_spacing = 1.0

    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(short_title + "  |  ")
    r.font.size = Pt(8)
    r.font.color.rgb = RGBColor(85, 85, 85)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    p._p.append(fld)


def shade_code(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), LIGHT)
    p_pr.append(shd)


def add_code(doc, code):
    p = doc.add_paragraph(style="Code")
    p.paragraph_format.keep_together = True
    p.add_run(code.strip("\n"))
    shade_code(p)
    return p


def add_bullets(doc, items, level=0):
    for item in items:
        p = doc.add_paragraph(style="List Bullet" if level == 0 else "List Bullet 2")
        p.add_run(item)


def add_numbered(doc, items):
    for index, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.first_line_indent = Inches(-0.25)
        p.add_run(f"{index}.  ")
        p.add_run(item)


def add_table(doc, headers, rows, widths=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    hdr = table.rows[0]
    repeat_table_header(hdr)
    prevent_row_split(hdr)
    for i, text in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = str(text)
        set_cell_shading(cell, NAVY)
        set_cell_margins(cell)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for run in cell.paragraphs[0].runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
            run.font.size = Pt(9)
    for row_index, values in enumerate(rows):
        row = table.add_row()
        prevent_row_split(row)
        for i, value in enumerate(values):
            cell = row.cells[i]
            cell.text = str(value)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if row_index % 2:
                set_cell_shading(cell, PALE)
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(1)
                for run in p.runs:
                    run.font.size = Pt(8.7)
        if widths:
            for i, width in enumerate(widths):
                row.cells[i].width = Inches(width)
    if widths:
        for i, width in enumerate(widths):
            hdr.cells[i].width = Inches(width)
    doc.add_paragraph().paragraph_format.space_after = Pt(1)
    return table


def add_chapter(doc, number, title, objective=None):
    heading = doc.add_heading(f"{number} {title}", level=1)
    heading.paragraph_format.page_break_before = True
    if objective:
        p = doc.add_paragraph()
        r = p.add_run("Learning objective  ")
        r.bold = True
        p.add_run(objective)


def add_cover(doc, title, subtitle, audience):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(90)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    doc.add_paragraph(title, style="Title")
    s = doc.add_paragraph(subtitle)
    s.runs[0].font.size = Pt(14)
    s.runs[0].font.color.rgb = RGBColor(55, 55, 55)
    doc.add_paragraph()
    a = doc.add_paragraph()
    a.add_run("Prepared for  ").bold = True
    a.add_run(audience)
    d = doc.add_paragraph()
    d.add_run("Scope  ").bold = True
    d.add_run("Portable multimedia fundamentals, Linux pipelines, hardware acceleration, real-time design, implementation labs, and interview preparation")
    doc.add_page_break()


def add_manual_contents(doc, entries):
    doc.add_heading("Contents", level=1)
    for number, title in entries:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.add_run(f"{number}  ").bold = True
        p.add_run(title)


def build_handbook():
    doc = Document()
    configure_document(doc, "Multimedia Systems Handbook")
    add_cover(
        doc,
        "Video Audio and Multimedia Systems Engineering",
        "Hardware software protocols coding and real time interview preparation",
        "Engineers preparing for NVIDIA Qualcomm and related multimedia systems roles",
    )

    contents = [
        ("1", "Role map and preparation strategy"),
        ("2", "Computer architecture and systems programming foundation"),
        ("3", "Digital video fundamentals"),
        ("4", "Video compression and codec architecture"),
        ("5", "Digital audio and signal processing"),
        ("6", "Containers streaming and network protocols"),
        ("7", "Linux multimedia architecture"),
        ("8", "GStreamer and FFmpeg programming"),
        ("9", "Hardware acceleration on NVIDIA and Qualcomm platforms"),
        ("10", "Real time performance memory and concurrency"),
        ("11", "Debugging validation and quality measurement"),
        ("12", "Integrated design case study"),
        ("13", "Interview question bank with model answers"),
        ("14", "Official references and next steps"),
    ]
    add_manual_contents(doc, contents)
    doc.add_heading("How to use this handbook", level=2)
    doc.add_paragraph(
        "Study each chapter in sequence, compile the examples, and keep a laboratory notebook containing command output, timing measurements, failures, and design decisions. The interview target is not memorization. You should be able to trace one frame or audio block from hardware capture through memory, processing, compression, transport, decoding, synchronization, and presentation."
    )
    add_bullets(doc, [
        "First pass: explain each pipeline aloud without looking at notes.",
        "Second pass: implement the code and deliberately introduce failures.",
        "Third pass: measure latency, throughput, CPU use, copies, and queue depth.",
        "Final pass: answer each design question with requirements, trade-offs, failure modes, and verification evidence.",
    ])

    add_chapter(doc, 1, "Role map and preparation strategy", "Understand what multimedia systems interviews evaluate and how hardware and software responsibilities meet.")
    doc.add_paragraph(
        "Current NVIDIA multimedia and streaming roles consistently emphasize C and C++, multithreading, Linux or Windows systems programming, low-latency pipelines, codecs, GStreamer, RTP or WebRTC, and hardware acceleration. Qualcomm multimedia hiring materials similarly identify Linux or QNX, device drivers, automotive audio video camera graphics display, GStreamer, C and C++, and SoC architecture. Treat this as a skills map rather than a promise about one interview loop."
    )
    add_table(doc, ["Interview lane", "What you must explain", "What you must build or debug"], [
        ("C and C++", "Object lifetime, ownership, atomics, caches, virtual dispatch, templates", "Ring buffers, parsers, thread-safe queues, resource wrappers"),
        ("Operating systems", "Processes, threads, scheduling, virtual memory, mmap, ioctl, poll", "Capture loop, nonblocking I/O, timeout and shutdown handling"),
        ("Video", "YUV, chroma subsampling, stride, colorimetry, codec stages", "Frame conversion, validation, decode or encode pipeline"),
        ("Audio", "PCM, sampling, quantization, channel layout, DSP, latency", "WAV parser, FIR filter, resampler or mixer"),
        ("Streaming", "RTP, RTCP, RTSP, jitter, loss, congestion, synchronization", "Packetizer, jitter buffer, timestamp mapping"),
        ("Hardware", "DMA, IOMMU, caches, accelerators, interrupts, zero copy", "Trace buffers across CPU, VPU, GPU, DSP and display"),
        ("Performance", "Throughput, latency, bandwidth, queueing and backpressure", "Profile and remove copies, stalls and oversized queues"),
    ], [1.0, 2.75, 3.05])
    doc.add_heading("A strong interview answer pattern", level=2)
    add_numbered(doc, [
        "Clarify input format, resolution, frame rate, channels, latency target, power limit, and failure policy.",
        "Draw the data path and identify ownership at every boundary.",
        "Calculate bandwidth and compute budgets before selecting APIs.",
        "Define threading, queues, timestamps, backpressure, and shutdown behavior.",
        "Explain hardware acceleration and fallback behavior.",
        "Name observability signals and tests that prove correctness and performance.",
    ])

    add_chapter(doc, 2, "Computer architecture and systems programming foundation", "Connect C++ code to memory, scheduling, cache behavior, and hardware I/O.")
    doc.add_heading("Data movement is the central cost", level=2)
    doc.add_paragraph(
        "A 3840 by 2160 NV12 frame contains 1.5 bytes per pixel, or about 12.44 MB. At 60 frames per second, one read of every byte is roughly 746 MB/s. A pipeline that copies each frame four times moves nearly 3 GB/s before codec, inference, or display work. Real systems therefore use buffer pools, DMA, shared handles, cache-aware access, and asynchronous engines."
    )
    add_table(doc, ["Mechanism", "Meaning", "Multimedia implication"], [
        ("DMA", "A device transfers data without the CPU copying each byte", "Camera, codec and audio engines write or read buffers directly"),
        ("IOMMU", "Translates and protects device-visible addresses", "Shared buffers require valid mappings and lifetime control"),
        ("Cache coherence", "Defines how CPU and device views become consistent", "Incorrect synchronization can show stale or partially written frames"),
        ("Interrupt", "Device signals completion or an event", "High rates require batching or polling to control overhead"),
        ("Memory mapping", "Kernel or device memory appears in user address space", "V4L2 and device APIs avoid read-copy loops"),
        ("Fence", "Represents asynchronous completion and dependency", "Producer and consumer coordinate without blocking the whole CPU"),
    ], [1.15, 2.45, 3.2])
    doc.add_heading("RAII wrapper for an operating system file descriptor", level=2)
    add_code(doc, r'''
#include <unistd.h>
#include <utility>

class UniqueFd {
public:
    explicit UniqueFd(int fd = -1) noexcept : fd_(fd) {}
    ~UniqueFd() { reset(); }
    UniqueFd(const UniqueFd&) = delete;
    UniqueFd& operator=(const UniqueFd&) = delete;
    UniqueFd(UniqueFd&& other) noexcept : fd_(std::exchange(other.fd_, -1)) {}
    UniqueFd& operator=(UniqueFd&& other) noexcept {
        if (this != &other) reset(std::exchange(other.fd_, -1));
        return *this;
    }
    int get() const noexcept { return fd_; }
    explicit operator bool() const noexcept { return fd_ >= 0; }
    void reset(int next = -1) noexcept {
        if (fd_ >= 0) ::close(fd_);
        fd_ = next;
    }
private:
    int fd_;
};''')
    doc.add_paragraph(
        "Interview discussion: explain why copying is forbidden, why move operations transfer ownership, how exceptions interact with cleanup, and why blocking system calls still need cancellation and shutdown design."
    )
    doc.add_heading("Single producer single consumer ring buffer", level=2)
    add_code(doc, r'''
template<class T, std::size_t N>
class SpscRing {
    static_assert((N & (N - 1)) == 0, "N must be a power of two");
public:
    bool push(T value) {
        const auto h = head_.load(std::memory_order_relaxed);
        const auto next = (h + 1) & (N - 1);
        if (next == tail_.load(std::memory_order_acquire)) return false;
        data_[h] = std::move(value);
        head_.store(next, std::memory_order_release);
        return true;
    }
    bool pop(T& out) {
        const auto t = tail_.load(std::memory_order_relaxed);
        if (t == head_.load(std::memory_order_acquire)) return false;
        out = std::move(data_[t]);
        tail_.store((t + 1) & (N - 1), std::memory_order_release);
        return true;
    }
private:
    alignas(64) std::atomic<std::size_t> head_{0};
    alignas(64) std::atomic<std::size_t> tail_{0};
    std::array<T, N> data_{};
};''')
    doc.add_paragraph(
        "Acquire and release ordering publishes an initialized element before the consumer observes the new head. Separate cache lines reduce false sharing. The queue still needs an explicit full-buffer policy: block, drop newest, drop oldest, or signal overload. For live video, dropping an old non-reference frame may be preferable to accumulating latency."
    )

    add_chapter(doc, 3, "Digital video fundamentals", "Reason precisely about pixels, planes, color, memory layout, bandwidth, and frame timing.")
    doc.add_heading("From light to a frame buffer", level=2)
    doc.add_paragraph(
        "A camera sensor measures filtered light at photosites. An image signal processor performs operations such as demosaicing, black-level correction, denoising, white balance, tone mapping, color correction, and scaling. The output may be RGB or YCbCr and is stored in one or more memory planes. A video engineer must distinguish visible width from row stride, luma resolution from chroma resolution, coded dimensions from crop dimensions, and color values from colorimetry metadata."
    )
    add_table(doc, ["Format", "Layout", "Bytes per pixel", "Typical use"], [
        ("RGB24", "Packed R G B", "3", "CPU image tools and display interchange"),
        ("RGBA", "Packed four channels", "4", "Graphics and compositing"),
        ("I420", "Y plane then U then V, 4:2:0", "1.5", "Software codecs and interchange"),
        ("NV12", "Y plane then interleaved UV, 4:2:0", "1.5", "Hardware codecs and camera pipelines"),
        ("YUYV", "Packed Y U Y V, 4:2:2", "2", "Camera capture and uncompressed transport"),
        ("P010", "10-bit values in 16-bit containers, 4:2:0", "3", "HDR and 10-bit hardware pipelines"),
    ], [0.85, 2.75, 1.15, 2.0])
    doc.add_heading("Stride-safe luma average for NV12", level=2)
    add_code(doc, r'''
double average_luma(const uint8_t* y, int width, int height, int y_stride) {
    if (!y || width <= 0 || height <= 0 || y_stride < width)
        throw std::invalid_argument("invalid NV12 luma plane");
    uint64_t sum = 0;
    for (int row = 0; row < height; ++row) {
        const uint8_t* line = y + static_cast<size_t>(row) * y_stride;
        for (int col = 0; col < width; ++col) sum += line[col];
    }
    return static_cast<double>(sum) / (static_cast<uint64_t>(width) * height);
}''')
    doc.add_paragraph(
        "A common interview trap is using width as the row step. Hardware allocations often align rows to a larger stride. Reading padding as pixels corrupts analysis; assuming tightly packed chroma can cause out-of-bounds access. Always carry width, height, pixel format, per-plane stride, offset, allocation size, and color metadata together."
    )
    doc.add_heading("Color conversion reasoning", level=2)
    doc.add_paragraph(
        "A YCbCr-to-RGB conversion depends on matrix coefficients, range, bit depth, and transfer characteristics. BT.601, BT.709, and BT.2020 are not interchangeable. Limited-range 8-bit luma commonly uses nominal code values 16 through 235, while full range uses 0 through 255. A visually plausible image can still be technically wrong if the pipeline loses this metadata."
    )
    add_code(doc, r'''
static uint8_t clip8(int x) { return static_cast<uint8_t>(std::clamp(x, 0, 255)); }

// Integer approximation for limited-range BT.601 YCbCr.
Rgb yuv_to_rgb(uint8_t y, uint8_t u, uint8_t v) {
    int c = static_cast<int>(y) - 16;
    int d = static_cast<int>(u) - 128;
    int e = static_cast<int>(v) - 128;
    return {
        clip8((298*c + 409*e + 128) >> 8),
        clip8((298*c - 100*d - 208*e + 128) >> 8),
        clip8((298*c + 516*d + 128) >> 8)
    };
}''')
    doc.add_heading("Frame timing", level=2)
    doc.add_paragraph(
        "At 60 frames per second, the nominal frame interval is 16.667 ms. Presentation timestamps describe when content belongs on the playback timeline; decoding timestamps describe decode order when reordering exists. A monotonic clock should measure local elapsed time. Wall clocks can jump and should not drive timeout arithmetic."
    )

    add_chapter(doc, 4, "Video compression and codec architecture", "Explain encoder and decoder stages, reference pictures, rate control, latency, and failure propagation.")
    doc.add_paragraph(
        "Modern video codecs reduce spatial redundancy within a picture and temporal redundancy between pictures. The conceptual encoder performs partitioning, prediction, transform, quantization, entropy coding, and in-loop filtering. The decoder reverses entropy coding, inverse quantization and transform, adds prediction, applies filters, and stores reference pictures. Practical designs also manage buffers, timestamps, reordering, metadata, error recovery, and hardware queues."
    )
    add_table(doc, ["Concept", "Meaning", "Interview implication"], [
        ("I or intra picture", "Uses spatial prediction without earlier pictures", "Random access and recovery, but usually larger"),
        ("P picture", "Predicts from earlier references", "Lower bitrate with forward dependency"),
        ("B picture", "Can use multiple temporal references", "Better compression but reordering and added latency"),
        ("GOP", "Pattern of independently and predictively coded pictures", "Controls seekability, recovery, bitrate and latency"),
        ("QP", "Controls quantization strength", "Higher QP normally means fewer bits and more distortion"),
        ("VBV or HRD", "Decoder buffer model", "Rate control must prevent overflow and underflow"),
        ("IDR or refresh", "Breaks reference dependency in a defined way", "Used for channel changes, recovery and segment boundaries"),
    ], [1.0, 2.6, 3.15])
    doc.add_heading("Latency budget example", level=2)
    add_table(doc, ["Stage", "Budget", "Design lever"], [
        ("Capture", "8 ms", "Short exposure and bounded driver queue"),
        ("Preprocess", "3 ms", "GPU or DSP, fused kernels, zero copy"),
        ("Encode", "12 ms", "Low-latency preset, no B frames, bounded lookahead"),
        ("Network", "35 ms", "Regional path, pacing, congestion control"),
        ("Jitter buffer", "25 ms", "Adaptive target and loss recovery"),
        ("Decode", "8 ms", "Hardware decode and shallow queues"),
        ("Render", "16 ms", "Vsync policy and late-frame handling"),
    ], [1.35, 1.0, 4.4])
    doc.add_paragraph(
        "The total is not automatically the sum of maximum queue depths because stages can overlap, but buffered stages contribute directly to glass-to-glass delay. The correct answer includes measurement points at capture, enqueue, encode completion, packet send, packet receive, decode completion, and presentation."
    )
    doc.add_heading("Annex B start-code scanner exercise", level=2)
    add_code(doc, r'''
std::vector<std::span<const uint8_t>> split_annex_b(std::span<const uint8_t> bytes) {
    std::vector<size_t> starts;
    for (size_t i = 0; i + 3 < bytes.size(); ++i) {
        if (bytes[i] == 0 && bytes[i+1] == 0 && bytes[i+2] == 1) {
            starts.push_back(i); i += 2;
        } else if (i + 4 <= bytes.size() && bytes[i] == 0 && bytes[i+1] == 0 &&
                   bytes[i+2] == 0 && bytes[i+3] == 1) {
            starts.push_back(i); i += 3;
        }
    }
    std::vector<std::span<const uint8_t>> units;
    for (size_t n = 0; n < starts.size(); ++n) {
        size_t prefix = bytes[starts[n] + 2] == 1 ? 3 : 4;
        size_t begin = starts[n] + prefix;
        size_t end = n + 1 < starts.size() ? starts[n+1] : bytes.size();
        if (begin < end) units.emplace_back(bytes.subspan(begin, end - begin));
    }
    return units;
}''')
    doc.add_paragraph(
        "Production code must define behavior for empty units, trailing zeros, truncated start codes, untrusted input sizes, emulation-prevention bytes, and length-prefixed container samples. Fuzz the parser and never assume a complete access unit arrives in one network packet."
    )

    add_chapter(doc, 5, "Digital audio and signal processing", "Understand PCM, latency, buffering, filtering, resampling, mixing, and audio hardware interfaces.")
    doc.add_heading("Sampling and representation", level=2)
    doc.add_paragraph(
        "PCM represents sampled amplitude values. A 48 kHz stereo stream with 16-bit samples carries 48,000 times 2 times 2 bytes, or 192,000 bytes per second before container and transport overhead. Interleaved stereo stores L0 R0 L1 R1. Planar audio stores all left samples followed by all right samples. Endianness, signedness, valid bits, container width, and channel order must be explicit."
    )
    add_table(doc, ["Term", "Definition", "Typical failure"], [
        ("Sample rate", "Samples per second per channel", "Pitch or duration error when clocks disagree"),
        ("Bit depth", "Quantization precision", "Clipping, noise or wrong sign extension"),
        ("Period", "Frames processed per wakeup", "Too small wastes CPU; too large increases latency"),
        ("Buffer", "Multiple periods held by device or software", "Underrun, overrun or excessive delay"),
        ("XRUN", "Playback underrun or capture overrun", "Audible gap and timeline discontinuity"),
        ("Drift", "Clock rate mismatch between endpoints", "Gradual queue growth or starvation"),
    ], [1.0, 2.65, 3.1])
    doc.add_heading("Saturating stereo mixer", level=2)
    add_code(doc, r'''
int16_t saturate_i16(int32_t x) {
    return static_cast<int16_t>(std::clamp<int32_t>(x, -32768, 32767));
}

void mix_s16(const int16_t* a, const int16_t* b, int16_t* out,
             size_t sample_count, float gain_a, float gain_b) {
    for (size_t i = 0; i < sample_count; ++i) {
        int32_t mixed = static_cast<int32_t>(std::lround(a[i] * gain_a)) +
                        static_cast<int32_t>(std::lround(b[i] * gain_b));
        out[i] = saturate_i16(mixed);
    }
}''')
    doc.add_paragraph(
        "Questions to expect: why does direct addition clip, why is floating-point headroom useful, when is fixed point preferred, how should gains ramp to avoid clicks, and how does channel layout affect mixing?"
    )
    doc.add_heading("Finite impulse response filter", level=2)
    add_code(doc, r'''
class Fir {
public:
    explicit Fir(std::vector<float> taps)
        : taps_(std::move(taps)), history_(taps_.size(), 0.0f) {}

    float process(float x) {
        history_[write_] = x;
        float y = 0.0f;
        size_t index = write_;
        for (float tap : taps_) {
            y += tap * history_[index];
            index = index == 0 ? history_.size() - 1 : index - 1;
        }
        write_ = (write_ + 1) % history_.size();
        return y;
    }
private:
    std::vector<float> taps_, history_;
    size_t write_ = 0;
};''')
    doc.add_paragraph(
        "A causal FIR output is a weighted sum of current and earlier samples. Linear-phase symmetric taps preserve waveform shape but add group delay. Optimize long filters with SIMD, block processing, or frequency-domain convolution. Do not allocate, lock, log, or perform unbounded work in a real-time audio callback."
    )
    doc.add_heading("Audio video synchronization", level=2)
    doc.add_paragraph(
        "Choose a master clock, often audio playback because the hardware consumes samples continuously. Convert both streams to one timeline. If video is early, wait; if slightly late, present immediately; if very late, drop according to policy. If audio and video clocks drift, correct gradually with resampling or clock adjustment rather than periodically inserting a large discontinuity."
    )

    add_chapter(doc, 6, "Containers streaming and network protocols", "Separate elementary streams, containers, signaling, transport, timing, reliability, and congestion control.")
    add_table(doc, ["Layer", "Examples", "Responsibility"], [
        ("Codec", "H.264, HEVC, AV1, AAC, Opus", "Compress and reconstruct media"),
        ("Container", "MP4, Matroska, MPEG-TS", "Store tracks, timestamps, metadata and indexing"),
        ("Transport", "RTP over UDP, QUIC, TCP", "Carry packets between endpoints"),
        ("Control", "RTCP, RTSP, SDP, WebRTC signaling", "Describe, negotiate and monitor sessions"),
        ("Application", "Player, conference, camera service", "Policy, user experience and recovery"),
    ], [1.0, 2.45, 3.3])
    doc.add_heading("RTP essentials", level=2)
    doc.add_paragraph(
        "RTP provides sequence numbers, timestamps, payload type, source identity, and marker semantics. Sequence numbers expose loss and reordering. Timestamps represent the media sampling timeline, not wall-clock time. RTCP reports reception statistics and maps RTP time to a reference time. RTP itself does not guarantee delivery or congestion safety."
    )
    add_code(doc, r'''
// Correctly classify 16-bit sequence numbers across wraparound when the
// distance is known to be less than half the sequence space.
bool sequence_is_newer(uint16_t a, uint16_t b) {
    return static_cast<int16_t>(a - b) > 0;
}

uint32_t video_timestamp_90k(std::chrono::nanoseconds elapsed) {
    using i128 = __int128_t;
    return static_cast<uint32_t>((i128(elapsed.count()) * 90000) / 1000000000);
}''')
    doc.add_heading("Jitter buffer design", level=2)
    add_numbered(doc, [
        "Parse and validate packets without trusting payload length or extension fields.",
        "Extend wrapped sequence numbers and timestamps into a larger local domain.",
        "Insert packets into a bounded reordering structure.",
        "Estimate network jitter and choose a bounded playout target.",
        "Detect missing packets and apply wait, retransmission, concealment, or frame-drop policy.",
        "Emit complete access units with reconstructed timestamps and discontinuity markers.",
        "Feed RTCP and congestion-control statistics without blocking the media path.",
    ])
    doc.add_heading("Protocols to compare in interviews", level=2)
    add_table(doc, ["Protocol", "Best mental model", "Key trade-off"], [
        ("RTSP", "Session control for media servers", "Usually controls an RTP-based stream; not the media codec"),
        ("WebRTC", "Interactive real-time media system", "NAT traversal, encryption and congestion control add complexity"),
        ("HLS or DASH", "Segmented HTTP adaptive streaming", "Scalable delivery with more latency than interactive media"),
        ("SRT or RIST", "Reliable contribution transport", "Loss recovery improves robustness but adds delay"),
        ("MPEG-TS", "Packetized broadcast-friendly container", "Fast recovery and multiplexing with container overhead"),
    ], [1.0, 2.8, 2.95])

    add_chapter(doc, 7, "Linux multimedia architecture", "Trace capture and playback across user space, kernel drivers, DMA buffers, codecs, display, and audio devices.")
    doc.add_heading("V4L2 capture sequence", level=2)
    add_numbered(doc, [
        "Open the device and query capabilities.",
        "Enumerate and negotiate pixel format, dimensions, frame interval, and controls.",
        "Request a buffer pool using MMAP, USERPTR, or DMABUF according to driver support.",
        "Map or import buffers, then queue every available buffer.",
        "Start streaming and wait with poll or epoll.",
        "Dequeue a completed buffer, consume it, then queue it again promptly.",
        "Stop streaming and release mappings in a failure-safe shutdown path.",
    ])
    add_code(doc, r'''
v4l2_buffer buffer{};
buffer.type = V4L2_BUF_TYPE_VIDEO_CAPTURE;
buffer.memory = V4L2_MEMORY_MMAP;

if (::ioctl(fd, VIDIOC_DQBUF, &buffer) == -1) {
    if (errno == EAGAIN) return NoFrameYet;
    if (errno == EINTR) return Retry;
    throw_system_error("VIDIOC_DQBUF");
}

process_frame(mapped[buffer.index], buffer.bytesused,
              buffer.timestamp, buffer.sequence);

if (::ioctl(fd, VIDIOC_QBUF, &buffer) == -1)
    throw_system_error("VIDIOC_QBUF");''')
    doc.add_paragraph(
        "A complete implementation validates buffer indices and bytes used, handles multi-planar formats, propagates discontinuities, drains safely during shutdown, and never holds scarce capture buffers longer than necessary. DMABUF export or import can connect the camera, codec, GPU, and display without CPU copies when all components support the format and synchronization model."
    )
    doc.add_heading("ALSA PCM lifecycle", level=2)
    doc.add_paragraph(
        "An ALSA application opens a PCM device, configures access mode, sample format, channels, rate, period and buffer size, prepares the stream, and repeatedly reads or writes frames. It must recover from underrun, overrun, suspend, short transfers, and interruption. Latency depends on hardware, period size, buffer size, scheduling, and any software plugins in the path."
    )
    doc.add_heading("Driver interview model", level=2)
    add_bullets(doc, [
        "Probe acquires clocks, regulators, interrupts, memory ranges and DMA capabilities.",
        "User APIs negotiate formats and allocate or import buffers.",
        "The driver programs descriptors and hardware registers, then starts the engine.",
        "An interrupt or polling path records completion and returns the buffer.",
        "Power management preserves ordering across suspend, resume and runtime idle.",
        "Error paths unwind partial initialization in reverse order.",
    ])

    add_chapter(doc, 8, "GStreamer and FFmpeg programming", "Build, inspect, extend, and debug graph-based multimedia pipelines.")
    doc.add_heading("GStreamer object model", level=2)
    doc.add_paragraph(
        "Elements perform work; pads connect elements; caps describe compatible media; buffers carry data and timestamps; events carry control information; queries request information; buses deliver application messages; clocks and segments map stream time to running time. Dynamic pads appear only after a demuxer discovers a stream, so applications often connect them in a pad-added callback."
    )
    add_code(doc, r'''
int main(int argc, char** argv) {
    gst_init(&argc, &argv);
    GError* error = nullptr;
    GstElement* pipeline = gst_parse_launch(
        "filesrc location=input.mp4 ! qtdemux name=d "
        "d.video_0 ! queue ! decodebin ! videoconvert ! autovideosink",
        &error);
    if (!pipeline) throw std::runtime_error(error ? error->message : "pipeline error");
    gst_element_set_state(pipeline, GST_STATE_PLAYING);
    GstBus* bus = gst_element_get_bus(pipeline);
    GstMessage* msg = gst_bus_timed_pop_filtered(
        bus, GST_CLOCK_TIME_NONE,
        static_cast<GstMessageType>(GST_MESSAGE_ERROR | GST_MESSAGE_EOS));
    if (msg) gst_message_unref(msg);
    gst_object_unref(bus);
    gst_element_set_state(pipeline, GST_STATE_NULL);
    gst_object_unref(pipeline);
}''')
    doc.add_paragraph(
        "For a production application, inspect error details and debug strings, manage dynamic pads deliberately, handle state-change failures, support cancellation, and avoid assuming a fixed pad name. Use queues only with explicit size and leak policy; every queue can add latency."
    )
    doc.add_heading("Custom transform plugin mental model", level=2)
    add_bullets(doc, [
        "Declare sink and source pad templates with accurate caps.",
        "Negotiate input and output formats before processing buffers.",
        "Map buffers with correct read or write flags and respect memory type.",
        "Preserve or transform PTS, DTS, duration, flags and metadata intentionally.",
        "Return flow errors rather than silently dropping corrupted state.",
        "Separate control-plane property changes from streaming-thread data access.",
    ])
    doc.add_heading("FFmpeg architecture", level=2)
    doc.add_paragraph(
        "libavformat demultiplexes packets, libavcodec decodes or encodes frames, libswscale converts video, libswresample converts audio, and filters build processing graphs. Packets contain compressed data; frames contain decoded samples. The send-packet or receive-frame API is asynchronous: one input may yield zero, one, or several outputs, and end-of-stream requires flushing."
    )
    add_code(doc, r'''
while (av_read_frame(format, packet) >= 0) {
    if (packet->stream_index == video_index) {
        check(avcodec_send_packet(decoder, packet));
        while (true) {
            int rc = avcodec_receive_frame(decoder, frame);
            if (rc == AVERROR(EAGAIN) || rc == AVERROR_EOF) break;
            check(rc);
            consume(frame); // retain or copy only with correct reference ownership
            av_frame_unref(frame);
        }
    }
    av_packet_unref(packet);
}
check(avcodec_send_packet(decoder, nullptr)); // flush delayed frames''')

    add_chapter(doc, 9, "Hardware acceleration on NVIDIA and Qualcomm platforms", "Map portable pipeline concepts onto GPU, VPU, DSP, codec, camera, and display engines.")
    doc.add_heading("NVIDIA data path", level=2)
    doc.add_paragraph(
        "A representative NVIDIA path is capture or demux, NVDEC hardware decode, CUDA or a dedicated post-processing engine, TensorRT inference when required, compositor, NVENC hardware encode, and RTP or container output. NVDEC and NVENC are dedicated codec engines rather than ordinary CUDA kernels. DeepStream builds accelerated analytics pipelines on GStreamer and carries structured metadata alongside buffers. Exact capabilities depend on GPU, platform, driver, and SDK version, so query capabilities instead of hard-coding assumptions."
    )
    add_table(doc, ["Component", "Responsibility", "Interview focus"], [
        ("NVDEC", "Hardware video decode", "Parser callbacks, surfaces, mapping, decode and display order"),
        ("NVENC", "Hardware video encode", "Preset, rate control, input surfaces, async output, latency"),
        ("CUDA", "Programmable GPU processing", "Transfers, pitch, kernels, streams, occupancy and synchronization"),
        ("DeepStream", "GStreamer-based accelerated analytics", "Batching, plugins, metadata, memory types and latency"),
        ("NvMedia or platform APIs", "Automotive and embedded media primitives", "Safety, surfaces, synchronization and platform constraints"),
    ], [1.25, 2.4, 3.15])
    doc.add_heading("CUDA pitch-safe grayscale kernel", level=2)
    add_code(doc, r'''
__global__ void threshold_luma(const uint8_t* y, size_t pitch,
                               uint8_t* out, size_t out_pitch,
                               int width, int height, uint8_t threshold) {
    int x = blockIdx.x * blockDim.x + threadIdx.x;
    int row = blockIdx.y * blockDim.y + threadIdx.y;
    if (x >= width || row >= height) return;
    const uint8_t* src_line = y + static_cast<size_t>(row) * pitch;
    uint8_t* dst_line = out + static_cast<size_t>(row) * out_pitch;
    dst_line[x] = src_line[x] >= threshold ? 255 : 0;
}

dim3 block(32, 8);
dim3 grid((width + block.x - 1) / block.x,
          (height + block.y - 1) / block.y);
threshold_luma<<<grid, block, 0, stream>>>(y, pitch, out, out_pitch,
                                          width, height, 128);''')
    doc.add_paragraph(
        "Explain coalesced access, pitch, asynchronous streams, launch error checking, completion events, pinned host memory, and why unnecessary host round trips destroy pipeline performance."
    )
    doc.add_heading("Qualcomm data path", level=2)
    doc.add_paragraph(
        "A representative Qualcomm path can combine camera and ISP, hardware video codec, GPU, Hexagon DSP or AI accelerator, display, and GStreamer-based multimedia components. Qualcomm Linux documentation describes VPU-backed video and GStreamer integration, while the Intelligent Multimedia SDK presents camera, video, audio, and AI pipelines. The portable interview skill is to place each operation on the most suitable engine while preserving buffer compatibility, synchronization, power efficiency, and observability."
    )
    add_table(doc, ["Engine", "Good fit", "Questions to ask"], [
        ("CPU", "Control, parsing, irregular algorithms", "Can work be vectorized or moved off the critical thread?"),
        ("GPU", "Highly parallel pixel and graphics work", "Are transfers, format conversion or synchronization dominant?"),
        ("DSP", "Low-power streaming signal processing", "What fixed-point, alignment, RPC and cache constraints apply?"),
        ("VPU or codec", "Video encode and decode", "Which formats, profiles, surfaces, rates and queue depths are supported?"),
        ("ISP", "Sensor-domain image processing", "Which operations must happen before RGB or YUV output?"),
        ("NPU or AI engine", "Neural inference", "What tensor layout, quantization and preprocessing path is required?"),
    ], [1.1, 2.45, 3.3])
    doc.add_heading("Portable capability negotiation", level=2)
    add_code(doc, r'''
struct VideoCaps {
    Codec codec;
    PixelFormat format;
    int width, height;
    Rational frame_rate;
    int bit_depth;
    MemoryDomain memory;
};

Expected<PipelinePlan, Error> choose_pipeline(const Request& request,
                                               const DeviceCaps& device) {
    auto decoder = device.find_decoder(request.input_codec, request.bit_depth);
    if (!decoder) return unexpected(Error::UnsupportedCodec);
    auto surface = intersect(decoder->outputs, device.processor_inputs);
    if (!surface) return unexpected(Error::NoZeroCopyFormat);
    return PipelinePlan{*decoder, *surface, choose_queue_depth(request.latency)};
}''')

    add_chapter(doc, 10, "Real time performance memory and concurrency", "Design bounded, measurable pipelines that remain stable under overload.")
    doc.add_heading("Throughput and latency are different", level=2)
    doc.add_paragraph(
        "A pipeline can sustain 60 frames per second yet still have one second of latency if queues hold 60 frames. Throughput asks how much work completes per unit time. Latency asks how long one item takes end to end. Tail latency matters because occasional stalls cause visible freezes, audio gaps, and queue growth."
    )
    add_table(doc, ["Problem", "Symptom", "Evidence", "Fix direction"], [
        ("Queue buildup", "Latency increases over time", "Queue depth and item age", "Backpressure, dropping, faster stage"),
        ("Memory copies", "High bandwidth and CPU use", "Memcpy samples, bus counters", "Shared surfaces, DMABUF, fusion"),
        ("Lock contention", "Threads sleep despite available compute", "Mutex wait profile", "Ownership partitioning, SPSC queues"),
        ("False sharing", "Poor scaling across cores", "Cache invalidation counters", "Alignment and data layout"),
        ("Priority inversion", "Real-time thread misses deadline", "Scheduling trace", "Priority inheritance and bounded locks"),
        ("Thermal throttling", "Performance degrades after warmup", "Frequency and thermal telemetry", "Power-aware load distribution"),
    ], [1.15, 1.85, 1.7, 2.05])
    doc.add_heading("Bounded pipeline contract", level=2)
    add_bullets(doc, [
        "Every queue has a capacity and documented overflow policy.",
        "Every blocking operation has a timeout or cancellation mechanism.",
        "Every buffer has one clear owner or reference-counted lifetime.",
        "Every asynchronous handoff carries a timestamp and completion signal.",
        "Every stage exports latency, processing time, queue depth, drops and error counters.",
        "Shutdown stops producers, drains or discards by policy, joins threads, and releases hardware in dependency order.",
    ])
    doc.add_heading("Performance timer", level=2)
    add_code(doc, r'''
class ScopedTimer {
public:
    ScopedTimer(Histogram& h) : histogram_(h), start_(Clock::now()) {}
    ~ScopedTimer() {
        auto elapsed = std::chrono::duration_cast<std::chrono::microseconds>(
            Clock::now() - start_);
        histogram_.record(elapsed.count());
    }
private:
    using Clock = std::chrono::steady_clock;
    Histogram& histogram_;
    Clock::time_point start_;
};''')
    doc.add_paragraph(
        "Use a monotonic clock, record distributions rather than averages alone, avoid high-overhead logging on the hot path, and correlate measurements using frame or packet identifiers. Warm up caches and hardware before benchmarking."
    )

    add_chapter(doc, 11, "Debugging validation and quality measurement", "Turn symptoms into evidence across formats, timing, memory, transport, and hardware boundaries.")
    add_table(doc, ["Symptom", "Likely categories", "First checks"], [
        ("Green or purple video", "Plane order, stride, chroma, range", "Dump raw frame, verify format and per-plane metadata"),
        ("Diagonal tearing", "Incorrect row stride", "Compare width, bytes per pixel and allocated pitch"),
        ("Periodic freeze", "Queue, keyframe loss, blocking I/O", "Timeline queue depth, packet loss and decoder errors"),
        ("Audio clicks", "XRUN, discontinuity, gain step", "PCM state, timestamps, callback duration"),
        ("A/V drift", "Clock mismatch or timestamp conversion", "Plot clock difference and queue occupancy"),
        ("High CPU", "Copies, conversion, polling, contention", "Profile samples, syscalls and memory bandwidth"),
        ("Hardware works once", "Lifetime, cache, fence, reset ordering", "Trace buffer ownership and completion events"),
    ], [1.25, 2.7, 2.75])
    doc.add_heading("Quality metrics", level=2)
    add_bullets(doc, [
        "Video fidelity: PSNR, SSIM and perceptual metrics, interpreted with content and encode settings.",
        "Audio fidelity: SNR, THD plus noise, frequency response, clipping and perceptual listening tests.",
        "System performance: end-to-end latency, p50 p95 p99 stage time, throughput, drops, jitter, CPU, GPU, memory bandwidth and power.",
        "Protocol health: packet loss, reordering, retransmissions, round-trip time, jitter-buffer delay and bitrate adaptation.",
        "Robustness: malformed inputs, stream changes, device removal, timeout, out-of-memory, thermal throttling and repeated start-stop cycles.",
    ])
    doc.add_heading("Test frame metadata", level=2)
    add_code(doc, r'''
struct FrameTrace {
    uint64_t id;
    int64_t capture_ns;
    int64_t preprocess_begin_ns;
    int64_t encode_done_ns;
    int64_t network_send_ns;
    int64_t network_receive_ns;
    int64_t decode_done_ns;
    int64_t present_ns;
};

// Report stage deltas and end-to-end latency; never compare unrelated
// machines without a defined clock synchronization strategy.''')

    add_chapter(doc, 12, "Integrated design case study", "Design a low-latency automotive camera and audio streaming system from requirements through validation.")
    doc.add_heading("Requirements", level=2)
    add_bullets(doc, [
        "Four 1080p30 camera inputs and one stereo 48 kHz audio input.",
        "Optional object detection and overlay on each camera stream.",
        "Hardware H.265 encode with a 100 ms local glass-to-glass target.",
        "Encrypted remote streaming with graceful degradation on variable networks.",
        "Local recording with recoverable segments and accurate A/V synchronization.",
        "Bounded memory, restartable components, health telemetry and no unbounded queue growth.",
    ])
    doc.add_heading("Proposed architecture", level=2)
    add_code(doc, r'''
Camera/ISP -> capture buffers -> preprocess/inference -> overlay -> encoder
       |              |                 |                  |
       +--------- shared surfaces and explicit fences -----+
                                      |
Audio device -> period queue -> resample/mix -> audio encoder
                                      |
                         timestamp and sync service
                                      |
                 packetizer -> network / segment recorder''')
    doc.add_heading("Design decisions", level=2)
    add_table(doc, ["Decision", "Choice", "Reason and verification"], [
        ("Memory", "Shared hardware surfaces", "Remove full-frame CPU copies; trace imports, fences and cache operations"),
        ("Queues", "Two to four buffers per bounded stage", "Enough overlap without hiding excessive latency; graph item age"),
        ("Clock", "Monotonic media timeline with audio master", "Measure drift and gradually correct resampling ratio"),
        ("Overload", "Drop oldest eligible video before audio", "Protect interactivity and continuous audio; count every drop"),
        ("Network", "Adaptive bitrate and bounded jitter buffer", "Respond to loss and RTT while capping playout delay"),
        ("Recording", "Short finalized segments", "Limit corruption after power failure; verify recovery under fault injection"),
    ], [1.05, 2.15, 3.85])
    doc.add_heading("Failure injection plan", level=2)
    add_bullets(doc, [
        "Disconnect and reconnect one camera while other streams continue.",
        "Delay the inference stage until its queue overflows.",
        "Inject RTP loss, duplication, reordering, and burst jitter.",
        "Force audio XRUN and verify timeline recovery without a permanent offset.",
        "Change resolution or codec parameters midstream where supported.",
        "Terminate recording during a segment write and verify previous segments remain playable.",
        "Run at thermal limits and prove the degradation policy preserves core functionality.",
    ])

    add_chapter(doc, 13, "Interview question bank with model answers", "Practice concise explanations that include mechanism, trade-off, failure mode, and evidence.")
    qa = [
        ("Why is NV12 1.5 bytes per pixel?", "The Y plane has one byte per pixel. In 4:2:0, each 2 by 2 luma block shares one U and one V sample, adding two bytes for four pixels, or 0.5 byte per pixel."),
        ("What is the difference between stride and width?", "Width is the number of visible pixels. Stride is the byte distance between corresponding positions in adjacent rows and can include alignment padding."),
        ("Why can B frames increase latency?", "They can depend on future pictures, so encoding, decoding, and presentation may require reordering and additional buffered frames."),
        ("Why use a monotonic clock?", "It advances consistently for elapsed-time measurement and is not affected by wall-clock adjustments."),
        ("What is zero copy?", "It is a pipeline design where components share or import the same underlying storage rather than copying payload bytes. It still requires ownership, format, cache, and synchronization handling."),
        ("How do you prevent pipeline latency growth?", "Bound every queue, measure item age, apply backpressure or a documented drop policy, and ensure long-term service rate is at least arrival rate."),
        ("What does an IOMMU add?", "Device address translation and isolation. It enables controlled device access to mapped memory but adds mapping and invalidation responsibilities."),
        ("Why can audio not allocate in its callback?", "General allocators can lock or take unbounded time, violating a hard callback deadline and causing an underrun."),
        ("What is the difference between a packet and a frame in FFmpeg?", "A packet normally holds compressed elementary-stream data, while a frame holds decoded video samples or audio samples."),
        ("How does RTP detect loss?", "The receiver observes gaps in the 16-bit sequence-number progression, accounting for wraparound and reordering."),
        ("What does RTCP provide?", "Reception statistics, sender timing information, source descriptions, and feedback used for monitoring and adaptation."),
        ("How do you debug a green video frame?", "Confirm negotiated format, plane addresses, strides, offsets, chroma order, range and buffer synchronization, then dump a known frame for independent inspection."),
        ("What is backpressure?", "A downstream stage communicates limited capacity so upstream stages slow, block, adapt, or drop instead of creating unbounded queues."),
        ("What is temporal locality?", "Recently accessed data is likely to be accessed again; cache-aware algorithms keep working sets small and reuse loaded data."),
        ("How do hardware decoder surfaces affect throughput?", "Too few surfaces stall decode when frames remain referenced or displayed; too many consume memory and can increase queueing latency."),
        ("When would you choose a DSP?", "For sustained low-power signal or pixel processing with predictable kernels, especially when avoiding CPU wakeups and memory movement matters."),
        ("How do you handle device removal?", "Stop new work, cancel or timeout waits, mark buffers invalid, drain callbacks, release resources in dependency order, and expose a recoverable state transition."),
        ("What is caps negotiation in GStreamer?", "Connected pads agree on media type and fields such as format, dimensions, rate, and memory features so each element receives supported data."),
        ("What is a keyframe not sufficient to describe?", "Different codecs define several random-access picture types and reference rules; recovery behavior depends on the exact bitstream semantics and decoder state."),
        ("How do you measure glass-to-glass latency?", "Use a traceable visual or electronic event visible to both capture and display measurement, or propagate synchronized timestamps through all stages and validate clock alignment."),
    ]
    for question, answer in qa:
        p = doc.add_paragraph()
        p.add_run(question + "  ").bold = True
        p.add_run(answer)

    doc.add_heading("Coding prompts", level=2)
    add_bullets(doc, [
        "Implement a bounded blocking queue with shutdown and timed pop.",
        "Parse a WAV header without assuming packed host structs or native endianness.",
        "Convert an NV12 luma plane to a histogram while respecting stride.",
        "Reassemble fragmented network units from an untrusted byte stream.",
        "Extend a wrapping 16-bit sequence number into a monotonic 64-bit counter.",
        "Design an LRU cache and discuss whether it belongs on a real-time path.",
        "Find and correct a race in a buffer reference-counting scheme.",
        "Calculate memory bandwidth and queue memory for a multi-camera pipeline.",
    ])

    add_chapter(doc, 14, "Official references and next steps", "Continue with primary documentation and implement the companion projects.")
    refs = [
        ("NVIDIA Video Codec SDK", "https://docs.nvidia.com/video-technologies/video-codec-sdk/13.1/index.html"),
        ("NVIDIA NVDEC Programming Guide", "https://docs.nvidia.com/video-technologies/video-codec-sdk/13.1/nvdec-video-decoder-api-prog-guide/index.html"),
        ("NVIDIA NVENC Programming Guide", "https://docs.nvidia.com/video-technologies/video-codec-sdk/13.1/nvenc-video-encoder-api-prog-guide/index.html"),
        ("NVIDIA DeepStream Developer Guide", "https://docs.nvidia.com/deepstream"),
        ("Qualcomm Intelligent Multimedia SDK", "https://docs.qualcomm.com/bundle/publicresource/topics/80-70020-50/qimsdk_landing_page.html"),
        ("Qualcomm Linux software overview", "https://docs.qualcomm.com/bundle/publicresource/topics/80-70014-115/overview.html"),
        ("GStreamer Application Development Manual", "https://gstreamer.freedesktop.org/documentation/application-development/index.html"),
        ("Linux V4L2 userspace API", "https://docs.kernel.org/userspace-api/media/v4l/v4l2.html"),
        ("Linux V4L2 streaming I/O", "https://cdn.kernel.org/doc/html/latest/userspace-api/media/v4l/mmap.html"),
        ("ALSA project documentation", "https://www.alsa-project.org/wiki/Documentation"),
        ("FFmpeg libraries documentation", "https://ffmpeg.org/documentation.html"),
        ("IETF RTP specification RFC 3550", "https://www.rfc-editor.org/rfc/rfc3550"),
        ("WebRTC specifications overview", "https://www.w3.org/TR/webrtc/"),
    ]
    add_table(doc, ["Reference", "URL"], refs, [2.8, 4.0])
    doc.add_paragraph(
        "SDK versions and product capabilities change. Before using vendor-specific code, match the documentation to the installed driver, SDK, SoC or GPU, operating system, and target hardware. Portable concepts in the earlier chapters remain the foundation for capability negotiation and debugging."
    )

    path = OUT / "Multimedia_Systems_Engineering_Handbook.docx"
    doc.save(path)
    return path


def build_workbook():
    doc = Document()
    configure_document(doc, "Multimedia Project Workbook")
    add_cover(
        doc,
        "Multimedia Systems Project and Interview Workbook",
        "Eight week plan mini projects review checklists and mock interviews",
        "C and C++ engineers preparing for hardware accelerated video and audio roles",
    )
    add_manual_contents(doc, [
        ("1", "Eight week study plan"),
        ("2", "Development environment and engineering notebook"),
        ("3", "Mini project one raw video laboratory"),
        ("4", "Mini project two Linux camera capture"),
        ("5", "Mini project three real time audio processor"),
        ("6", "Mini project four RTP packetizer and jitter buffer"),
        ("7", "Mini project five GStreamer analytics plugin"),
        ("8", "Mini project six hardware accelerated transcoder"),
        ("9", "Capstone synchronized low latency streamer"),
        ("10", "Mock interview rounds"),
        ("11", "Self assessment and final checklist"),
    ])
    doc.add_paragraph(
        "This workbook turns theory into portfolio evidence. Every project requires a design note, compilable implementation, tests, measurements, failure injection, and a short spoken explanation. Vendor hardware is optional for the early projects; write portable CPU versions first, then add NVIDIA or Qualcomm acceleration where available."
    )

    add_chapter(doc, 1, "Eight week study plan", "Build interview readiness through measured implementation rather than passive reading.")
    add_table(doc, ["Week", "Study focus", "Build outcome", "Interview outcome"], [
        ("1", "C++ ownership, threads, atomics, Linux I/O", "RAII wrappers and bounded queues", "Explain memory model and failure-safe shutdown"),
        ("2", "Pixels, YUV, stride, colorimetry", "Raw frame inspector and converter", "Calculate bandwidth and debug color defects"),
        ("3", "Codecs, timestamps, containers", "Bitstream and metadata inspector", "Explain GOP, QP, reordering and latency"),
        ("4", "PCM, DSP, ALSA", "Real-time audio effect chain", "Explain periods, XRUN, drift and filtering"),
        ("5", "RTP, RTCP, jitter and loss", "Packetizer and jitter simulator", "Design low-latency network behavior"),
        ("6", "V4L2, GStreamer and FFmpeg", "Capture and graph-based processing", "Debug negotiation and device buffers"),
        ("7", "NVDEC, NVENC, CUDA, VPU and DSP", "Accelerated processing path", "Explain zero copy and capability negotiation"),
        ("8", "Integration, profiling and mock loops", "Capstone demo and report", "Deliver system-design and coding interviews"),
    ], [0.55, 2.1, 2.25, 2.25])
    doc.add_heading("Daily cadence", level=2)
    add_table(doc, ["Block", "Time", "Activity"], [
        ("Concept", "45 min", "Read one focused topic and draw its data path"),
        ("Implementation", "90 min", "Write or modify code without copying a finished solution"),
        ("Debug", "30 min", "Inject one failure and capture evidence"),
        ("Interview", "30 min", "Answer five questions aloud and one coding prompt"),
        ("Review", "15 min", "Record what changed, measured results, and next experiment"),
    ], [1.0, 0.9, 5.0])

    add_chapter(doc, 2, "Development environment and engineering notebook", "Create reproducible experiments and evidence that can be discussed in interviews.")
    add_bullets(doc, [
        "Linux development environment with GCC or Clang, CMake, GDB, sanitizers and perf.",
        "FFmpeg and ffprobe for generating, converting and inspecting media.",
        "GStreamer tools and development headers for pipeline experiments.",
        "V4L2 utilities for camera enumeration and controls where a camera is available.",
        "ALSA utilities and headers for PCM tests where audio hardware is available.",
        "Optional NVIDIA GPU or Jetson with matching driver and SDK, or Qualcomm development platform with matching SDK.",
    ])
    doc.add_heading("Repository structure", level=2)
    add_code(doc, r'''
multimedia-labs/
  CMakeLists.txt
  common/           # RAII, logging, timestamps, buffer metadata
  raw-video/        # format parser, conversions, tests
  v4l2-capture/     # camera capture and DMABUF experiments
  audio-dsp/        # WAV, ALSA, FIR and mixer
  rtp-lab/          # packetizer, loss simulator, jitter buffer
  gst-plugin/       # custom transform and pipeline app
  hw-transcoder/    # portable interface plus vendor back ends
  capstone/         # integrated A/V system
  testdata/         # small generated fixtures, never huge media files
  docs/             # design, results, traces and interview notes''')
    doc.add_heading("Evidence template for every experiment", level=2)
    add_numbered(doc, [
        "State the requirement and measurable success criterion.",
        "Record hardware, OS, compiler, library and driver versions.",
        "Draw the data path and buffer ownership.",
        "Record the exact command and input fixture.",
        "Capture latency distribution, throughput, CPU, memory and drops.",
        "Describe one failure, root cause, fix and regression test.",
        "Summarize trade-offs and the next experiment.",
    ])

    add_chapter(doc, 3, "Mini project one raw video laboratory", "Master formats, planes, stride, color conversion and frame validation.")
    doc.add_heading("Build", level=2)
    add_bullets(doc, [
        "Command-line tool that reads I420, NV12, YUYV and RGB frames with explicit dimensions and strides.",
        "Print per-plane offsets, sizes, minimum buffer size, luma range and histogram.",
        "Convert NV12 to RGB and write a portable image output.",
        "Generate synthetic color bars and odd-width negative tests.",
        "Compare scalar and SIMD or multithreaded implementations without changing results.",
    ])
    doc.add_heading("Acceptance tests", level=2)
    add_table(doc, ["Test", "Expected evidence"], [
        ("Tightly packed 640 by 480 NV12", "Correct 460800-byte minimum and visually correct output"),
        ("Stride greater than width", "Padding ignored and no diagonal corruption"),
        ("Truncated input", "Clear error before out-of-bounds access"),
        ("BT.601 versus BT.709", "Documented visible and numeric difference"),
        ("Sanitizers", "No address, undefined behavior or leak reports"),
        ("Performance", "MPixels per second and memory bandwidth estimate"),
    ], [2.6, 4.3])
    doc.add_heading("Interview defense", level=2)
    add_bullets(doc, [
        "Explain why chroma planes have different dimensions.",
        "Explain why stride cannot be derived from visible width.",
        "Calculate frame size and bandwidth for 1080p60 and 4K60.",
        "Describe how you would move the converter to CUDA, GPU shaders, or a DSP.",
    ])

    add_chapter(doc, 4, "Mini project two Linux camera capture", "Implement a robust V4L2 streaming application with bounded buffer ownership.")
    doc.add_heading("Build", level=2)
    add_bullets(doc, [
        "Enumerate formats, sizes, intervals and controls.",
        "Negotiate one capture format and allocate four MMAP buffers.",
        "Use poll with timeout, dequeue, timestamp, validate, process and requeue.",
        "Write periodic frames and metadata to disk without blocking capture.",
        "Add signal-safe shutdown coordination and repeated start-stop testing.",
        "Optional: export DMABUF and import it into a downstream hardware component.",
    ])
    doc.add_heading("Failure injection", level=2)
    add_bullets(doc, [
        "Disconnect the camera or stop delivery while poll is waiting.",
        "Delay the consumer until every capture buffer is occupied.",
        "Request an unsupported format and verify capability-driven fallback.",
        "Send interruption signals during open, stream-on, dequeue and shutdown.",
        "Corrupt a saved frame length and prove the offline reader rejects it.",
    ])
    doc.add_heading("Design review questions", level=2)
    add_bullets(doc, [
        "Who owns a dequeued buffer, and when can the driver reuse it?",
        "What happens when processing takes longer than one frame interval?",
        "How would multi-planar capture change the metadata and mapping logic?",
        "Where do cache synchronization and fences appear in a DMABUF path?",
    ])

    add_chapter(doc, 5, "Mini project three real time audio processor", "Implement deterministic audio buffering, DSP and recovery.")
    doc.add_heading("Build", level=2)
    add_bullets(doc, [
        "Parse and write PCM WAV files using byte-wise little-endian helpers.",
        "Implement gain ramp, mixer, FIR low-pass filter and level meter.",
        "Process fixed-size blocks with no allocation inside the processing function.",
        "Add an ALSA capture-to-playback mode if Linux audio hardware is available.",
        "Measure algorithmic delay, callback time and XRUN count.",
    ])
    doc.add_heading("Required edge cases", level=2)
    add_table(doc, ["Case", "Expected behavior"], [
        ("Unknown WAV chunk", "Skip safely using declared chunk length and alignment"),
        ("24-bit PCM", "Correct sign extension and conversion without host-struct assumptions"),
        ("Gain change", "Smooth ramp without discontinuity click"),
        ("Over-range mix", "Saturate or use documented floating-point headroom"),
        ("XRUN", "Recover, mark discontinuity, and preserve a coherent timeline"),
        ("Clock mismatch", "Demonstrate queue drift and gradual correction"),
    ], [2.1, 4.8])

    add_chapter(doc, 6, "Mini project four RTP packetizer and jitter buffer", "Learn packetization, wraparound, loss, reordering and bounded playout latency.")
    doc.add_heading("Build", level=2)
    add_bullets(doc, [
        "Packetize synthetic video access units into an RTP-like stream with MTU limits.",
        "Parse packet headers defensively and extend sequence numbers across wraparound.",
        "Simulate loss, duplication, delay, burst jitter and reordering with a deterministic seed.",
        "Reassemble frames in a bounded jitter buffer with a playout deadline.",
        "Report received, late, duplicate, missing, recovered and dropped packet counts.",
        "Add RTCP-style round-trip and jitter statistics as an extension.",
    ])
    doc.add_heading("Experiment matrix", level=2)
    add_table(doc, ["Loss", "Jitter", "Buffer target", "Observe"], [
        ("0 percent", "5 ms", "20 ms", "Baseline latency and no drops"),
        ("1 percent", "10 ms", "30 ms", "Incomplete frames and concealment policy"),
        ("5 percent burst", "20 ms", "50 ms", "Recovery versus latency trade-off"),
        ("0 percent", "80 ms", "30 ms", "Late drops despite no packet loss"),
        ("Reordering", "20 ms", "40 ms", "Sequence extension and duplicate handling"),
    ], [1.15, 1.15, 1.45, 3.15])
    doc.add_heading("Interview defense", level=2)
    add_bullets(doc, [
        "Explain why RTP timestamps are not packet arrival times.",
        "Explain half-range sequence-number comparison and its assumptions.",
        "Compare retransmission, forward error correction, concealment and dropping.",
        "Explain how a jitter buffer adapts without creating uncontrolled latency.",
    ])

    add_chapter(doc, 7, "Mini project five GStreamer analytics plugin", "Create a negotiated transform and debug real pipeline behavior.")
    doc.add_heading("Build", level=2)
    add_bullets(doc, [
        "Create a transform element that accepts a documented raw video format.",
        "Compute luma statistics and optionally draw a threshold overlay.",
        "Preserve timestamps, duration, discontinuity flags and relevant metadata.",
        "Expose thread-safe properties for threshold and enable state.",
        "Run file, camera and synthetic-source pipelines.",
        "Add a queue-overflow experiment and record latency versus drop behavior.",
    ])
    doc.add_heading("Pipeline tests", level=2)
    add_code(doc, r'''
# Functional test
gst-launch-1.0 videotestsrc num-buffers=300 ! \
  video/x-raw,format=NV12,width=1280,height=720,framerate=30/1 ! \
  yourfilter threshold=128 ! fakesink sync=true

# Negotiation inspection
GST_DEBUG=2 gst-launch-1.0 videotestsrc ! yourfilter ! fakesink

# Deliberate downstream slowdown and bounded queue
... ! queue max-size-buffers=3 leaky=downstream ! identity sleep-time=50000 ! ...''')
    doc.add_heading("Review checklist", level=2)
    add_bullets(doc, [
        "Caps templates agree with actual memory access.",
        "Buffer mapping is released on every success and error path.",
        "Properties cannot race with the streaming thread.",
        "The plugin returns correct flow results and does not swallow errors.",
        "No unexpected allocations or copies occur in steady state.",
    ])

    add_chapter(doc, 8, "Mini project six hardware accelerated transcoder", "Design one portable pipeline with software and vendor-specific back ends.")
    doc.add_heading("Build", level=2)
    add_bullets(doc, [
        "Demux input, decode, scale or filter, encode and mux output.",
        "Implement a software baseline before enabling hardware acceleration.",
        "Query supported codec, profile, level, resolution, bit depth and surface formats.",
        "Keep frames in device-compatible memory when possible.",
        "Use asynchronous submission with bounded in-flight work.",
        "Measure throughput, p95 frame time, CPU, GPU or accelerator use, copies and power where available.",
    ])
    doc.add_heading("NVIDIA extension", level=2)
    add_bullets(doc, [
        "Use the Video Codec SDK or an appropriate GStreamer or FFmpeg hardware path.",
        "Inspect NVDEC output surfaces and avoid host copies before CUDA or NVENC.",
        "Compare low-latency and quality-oriented encoder configurations.",
        "Prove driver and SDK compatibility checks and report unsupported capabilities clearly.",
    ])
    doc.add_heading("Qualcomm extension", level=2)
    add_bullets(doc, [
        "Use the platform-supported GStreamer or Intelligent Multimedia SDK path.",
        "Inspect the memory type and format accepted by camera, video, GPU, DSP and display components.",
        "Measure whether format conversion or host access breaks zero-copy flow.",
        "Compare CPU, GPU, DSP or accelerator placement for one preprocessing operation.",
    ])
    doc.add_heading("Results table to complete", level=2)
    add_table(doc, ["Backend", "FPS", "p95 latency", "CPU", "Copies per frame", "Power"], [
        ("Software", "", "", "", "", ""),
        ("NVIDIA", "", "", "", "", ""),
        ("Qualcomm", "", "", "", "", ""),
    ], [1.35, 0.85, 1.25, 0.75, 1.35, 0.85])

    add_chapter(doc, 9, "Capstone synchronized low latency streamer", "Integrate capture, processing, codec, network, audio, timing, telemetry and recovery.")
    doc.add_heading("Mandatory features", level=2)
    add_bullets(doc, [
        "One live or synthetic video source and one audio source.",
        "Bounded processing queues with explicit overload policies.",
        "Video and audio encoding, timestamp mapping, packetization and receiver playback.",
        "A jitter and loss simulator between sender and receiver.",
        "A/V synchronization and drift telemetry.",
        "Optional hardware acceleration behind capability-driven interfaces.",
        "Clean cancellation, repeated start-stop, and automatic recovery from at least three injected failures.",
    ])
    doc.add_heading("Definition of done", level=2)
    add_table(doc, ["Area", "Evidence required"], [
        ("Correctness", "Known fixtures, unit tests, sanitizer run and playable output"),
        ("Latency", "Timestamp trace showing p50 p95 p99 and queue contribution"),
        ("Throughput", "Sustained target rate without unbounded queue growth"),
        ("Robustness", "Failure-injection log and recovery behavior"),
        ("Quality", "Video and audio checks plus documented encode settings"),
        ("Architecture", "Data path, ownership, thread and state diagrams"),
        ("Interview demo", "Five-minute walkthrough and ten-minute design defense"),
    ], [1.25, 5.65])
    doc.add_heading("Five minute portfolio explanation", level=2)
    add_numbered(doc, [
        "State the user-visible requirement and numerical targets.",
        "Show the architecture and follow one video frame and audio block.",
        "Explain the hardest ownership, timing or synchronization problem.",
        "Show measurements before and after one optimization.",
        "Demonstrate one injected failure and recovery.",
        "Close with limitations and the next engineering step.",
    ])

    add_chapter(doc, 10, "Mock interview rounds", "Practice coding, fundamentals, debugging and system design under realistic time constraints.")
    rounds = [
        ("Round 1", "C and C++ coding", "60 min", "Bounded queue, parser, RAII, memory safety and complexity"),
        ("Round 2", "Multimedia fundamentals", "60 min", "YUV, codecs, PCM, timestamps, containers and protocols"),
        ("Round 3", "Linux and hardware", "60 min", "V4L2, ALSA, DMA, IOMMU, cache, interrupts and driver flow"),
        ("Round 4", "Debugging", "45 min", "Analyze corrupted video, XRUN, A/V drift and growing latency"),
        ("Round 5", "System design", "60 min", "Design a low-latency multi-camera or conferencing pipeline"),
        ("Round 6", "Project deep dive", "45 min", "Defend measurements, trade-offs, failures and personal decisions"),
    ]
    add_table(doc, ["Round", "Focus", "Time", "Evaluation"], rounds, [0.8, 1.6, 0.8, 3.7])
    doc.add_heading("System design prompt", level=2)
    doc.add_paragraph(
        "Design a 4K60 low-latency remote camera system with microphone audio, hardware video codec, adaptive network transport, local recording and optional AI overlay. The system must operate for hours without latency growth and recover from packet loss and temporary device failure."
    )
    add_bullets(doc, [
        "Clarify resolution, quality, end-to-end latency, network, power, platform and safety constraints.",
        "Calculate raw bandwidth, encoded bitrate range, buffer memory and compute budget.",
        "Choose formats, engines, queues, threads, clocks and overload policies.",
        "Describe capability negotiation and software fallback.",
        "Define security boundaries, malformed-input handling and resource limits.",
        "Design telemetry, test fixtures, fault injection and acceptance criteria.",
    ])
    doc.add_heading("Debugging prompt", level=2)
    doc.add_paragraph(
        "The stream begins at 80 ms latency and reaches 900 ms after fifteen minutes. CPU is below 50 percent and no packet loss is reported. Explain your investigation."
    )
    add_numbered(doc, [
        "Plot queue depth and oldest-item age at every stage.",
        "Compare arrival and service rates over long windows rather than average CPU use.",
        "Check timestamp domains, clock drift, rate conversion and pacing.",
        "Look for a slow leak in retained references or decoder surfaces.",
        "Check whether a sink synchronizes to timestamps that advance at the wrong rate.",
        "Apply a bounded queue policy only after identifying which stage cannot keep pace.",
    ])

    add_chapter(doc, 11, "Self assessment and final checklist", "Decide whether you can explain, implement, measure and defend the complete pipeline.")
    assessment = [
        ("C++", "I can explain ownership, atomics and cache behavior and implement bounded concurrent structures."),
        ("Video", "I can calculate layouts and bandwidth and debug stride, color and timestamp failures."),
        ("Audio", "I can implement block DSP and explain XRUN, periods, latency, drift and resampling."),
        ("Codecs", "I can trace encode and decode stages and connect GOP and rate control to latency and quality."),
        ("Protocols", "I can explain RTP, RTCP, RTSP and WebRTC layering and design a jitter buffer."),
        ("Linux", "I can trace V4L2 and ALSA streaming and discuss mmap, ioctl, poll and DMABUF."),
        ("Frameworks", "I can build and debug GStreamer and FFmpeg pipelines and ownership."),
        ("Hardware", "I can place work across CPU, GPU, DSP, VPU, ISP and NPU with capability checks."),
        ("Performance", "I can distinguish throughput from latency and produce percentile measurements."),
        ("Debugging", "I form hypotheses, collect evidence, isolate stages and add regression tests."),
    ]
    add_table(doc, ["Area", "Ready when this statement is true", "Score 0 to 3"], [(a, b, "") for a, b in assessment], [1.15, 4.95, 0.85])
    doc.add_heading("Scoring", level=2)
    add_bullets(doc, [
        "0: I cannot explain it yet.",
        "1: I can explain it with notes.",
        "2: I can implement or debug it independently.",
        "3: I can teach it, discuss trade-offs, and support claims with measurements.",
    ])
    doc.add_heading("Final interview checklist", level=2)
    add_bullets(doc, [
        "Prepare two project stories with a difficult bug, your evidence, your decision and measured result.",
        "Practice writing compilable C++ without an IDE and narrate assumptions before coding.",
        "Review raw bandwidth, audio bitrate and latency calculations until they are automatic.",
        "Practice drawing one complete data path in under five minutes.",
        "Never claim zero copy without explaining storage sharing, synchronization and lifetime.",
        "Never optimize before naming the metric and showing a baseline.",
        "Ask the interviewer for requirements before selecting a codec, protocol or hardware engine.",
    ])

    path = OUT / "Multimedia_Mini_Projects_and_Interview_Workbook.docx"
    doc.save(path)
    return path


if __name__ == "__main__":
    print(build_handbook())
    print(build_workbook())
