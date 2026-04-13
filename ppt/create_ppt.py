from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

prs = Presentation()
prs.slide_width  = Inches(13.33)
prs.slide_height = Inches(7.5)

NAVY   = RGBColor(0x1a, 0x3a, 0x5c)
GOLD   = RGBColor(0xf5, 0x9e, 0x0b)
BLUE   = RGBColor(0x25, 0x63, 0xeb)
GREEN  = RGBColor(0x10, 0xb9, 0x81)
RED    = RGBColor(0xef, 0x44, 0x44)
WHITE  = RGBColor(0xff, 0xff, 0xff)
LGRAY  = RGBColor(0xf0, 0xf4, 0xf8)
MGRAY  = RGBColor(0x64, 0x74, 0x8b)

blank_layout = prs.slide_layouts[6]  # completely blank

def add_rect(slide, left, top, w, h, fill=NAVY, line=None):
    shape = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background() if line is None else None
    return shape

def add_text_box(slide, text, left, top, w, h, size=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox

def slide_bg(slide, color=LGRAY):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

# ─────────────────────────────────────────────
# SLIDE 1 — TITLE
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, NAVY)
add_rect(sl, 0, 0, 13.33, 7.5, NAVY)
# Gold accent bar
add_rect(sl, 0, 5.8, 13.33, 0.12, GOLD)
add_text_box(sl, "RPA COURSE  ·  ASSIGNMENT 04", 1, 0.6, 11, 0.5, size=11, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
add_text_box(sl, "E-Commerce Order Processing", 1, 1.3, 11, 0.9, size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text_box(sl, "RPA Bot — Troubleshooting & Debugging Case Study", 1, 2.2, 11, 0.8, size=22, bold=False, color=RGBColor(0x93,0xc5,0xfd), align=PP_ALIGN.CENTER)
add_text_box(sl, "A comprehensive study on diagnosing, debugging, and optimizing a\nUiPath RPA bot for end-to-end e-commerce order processing", 1, 3.1, 11, 0.9, size=13, color=RGBColor(0x94,0xa3,0xb8), align=PP_ALIGN.CENTER)
add_text_box(sl, "Student: Omais Siddiqui     |     Platform: UiPath + Orchestrator     |     April 2026", 1, 6.0, 11, 0.4, size=11, color=RGBColor(0x64,0x74,0x8b), align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────
# SLIDE 2 — AGENDA
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, LGRAY)
add_rect(sl, 0, 0, 13.33, 1.1, NAVY)
add_text_box(sl, "Presentation Agenda", 0.5, 0.2, 12, 0.7, size=26, bold=True, color=WHITE)

items = [
    ("01", "RPA Error Log",          "5 realistic exceptions across workflow stages",        BLUE),
    ("02", "Debugging Strategy",     "Logging, monitoring, retry logic, decision flow",      GREEN),
    ("03", "Optimization",           "Performance, reliability, scheduling, before/after",   GOLD),
    ("04", "Monitoring Dashboard",   "KPIs, queue health, alerts, bot instance status",      RED),
    ("05", "AI-Assisted Development","Vibe coding with Claude Code — how it was built",      RGBColor(0x93,0x33,0xea)),
]
for i, (num, title, sub, color) in enumerate(items):
    x = 0.4 + (i % 3) * 4.3
    y = 1.4 + (i // 3) * 2.8
    add_rect(sl, x, y, 3.9, 2.4, WHITE)
    add_rect(sl, x, y, 3.9, 0.08, color)
    add_text_box(sl, num, x+0.15, y+0.15, 0.7, 0.55, size=28, bold=True, color=color)
    add_text_box(sl, title, x+0.15, y+0.75, 3.6, 0.45, size=14, bold=True, color=NAVY)
    add_text_box(sl, sub,   x+0.15, y+1.2,  3.6, 0.9,  size=10, color=MGRAY)

# ─────────────────────────────────────────────
# SLIDE 3 — ERRORS TABLE
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, LGRAY)
add_rect(sl, 0, 0, 13.33, 1.1, NAVY)
add_text_box(sl, "Section 1 — RPA Error Log: 5 Exceptions", 0.5, 0.2, 12, 0.7, size=24, bold=True, color=WHITE)

errors = [
    ("E-01", "SelectorNotFoundException", "System",   "Order Placement",    "HIGH",  "Stale CSS selector after frontend deploy"),
    ("E-02", "InvalidOrderDataException", "Business", "Data Extraction",    "HIGH",  "Missing ZIP code in ERP international feed"),
    ("E-03", "OracleDBTimeoutException",  "System",   "Inventory Check",    "HIGH",  "Missing DB index causing full table scans"),
    ("E-04", "DuplicateOrderException",   "Business", "Order Submission",   "MEDIUM","No idempotency key on retry logic"),
    ("E-05", "PDFParsingFailureException","App",      "Invoice Generation", "LOW",   "Scanned PDF — OCR not enabled"),
]
hdrs = ["#", "Error Name", "Type", "Stage", "Severity", "Root Cause"]
col_w = [0.45, 3.0, 1.1, 1.6, 1.0, 5.4]
col_x = [0.25]
for w in col_w[:-1]: col_x.append(col_x[-1]+w)

# Header row
add_rect(sl, 0.25, 1.2, 12.85, 0.42, NAVY)
for j, (h, x, w) in enumerate(zip(hdrs, col_x, col_w)):
    add_text_box(sl, h, x+0.05, 1.22, w-0.1, 0.38, size=10, bold=True, color=WHITE)

row_colors = [WHITE, LGRAY, WHITE, LGRAY, WHITE]
sev_colors = {"HIGH": RED, "MEDIUM": GOLD, "LOW": GREEN}
type_colors = {"System": BLUE, "Business": GOLD, "App": RGBColor(0x93,0x33,0xea)}
for i, (eid, name, etype, stage, sev, cause) in enumerate(errors):
    y = 1.65 + i*0.95
    add_rect(sl, 0.25, y, 12.85, 0.9, row_colors[i])
    vals = [eid, name, etype, stage, sev, cause]
    for j, (val, x, w) in enumerate(zip(vals, col_x, col_w)):
        c = WHITE if row_colors[i]==NAVY else NAVY
        if j == 2: c = type_colors.get(val, BLUE)
        if j == 4: c = sev_colors.get(val, MGRAY)
        bold = j in (0,1,4)
        add_text_box(sl, val, x+0.05, y+0.05, w-0.1, 0.8, size=9 if j==5 else 10, bold=bold, color=c)

# ─────────────────────────────────────────────
# SLIDE 4 — DEBUGGING STRATEGY
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, LGRAY)
add_rect(sl, 0, 0, 13.33, 1.1, NAVY)
add_text_box(sl, "Section 2 — Debugging Strategy & Framework", 0.5, 0.2, 12, 0.7, size=24, bold=True, color=WHITE)

debug_cards = [
    ("📋 Logging",        BLUE,  "• Structured levels: INFO, WARN, ERROR, FATAL\n• Include: OrderID, BotInstanceID, UTC timestamp\n• Screenshot on every exception\n• Ship to Orchestrator → Elasticsearch → Kibana"),
    ("📊 Monitoring",     GREEN, "• Real-time KPI tracking (60s refresh)\n• PagerDuty: alert if success rate <90%\n• SLA violation detection (4-hour threshold)\n• Anomaly detection → auto Jira tickets"),
    ("🔁 Reproduction",   GOLD,  "• Isolate failing transaction from queue\n• Snapshot environment state\n• Replay in UiPath Debug mode\n• Reproduce 3× before applying any fix"),
    ("🔄 Retry Logic",    RED,   "• RetryScope: 3 retries, 5s backoff\n• Idempotency key check before each order POST\n• System exceptions → retry then escalate\n• Business exceptions → exception queue"),
]
for i, (title, color, body) in enumerate(debug_cards):
    x = 0.3 + (i%2)*6.4
    y = 1.25 + (i//2)*2.8
    add_rect(sl, x, y, 6.1, 2.55, WHITE)
    add_rect(sl, x, y, 6.1, 0.08, color)
    add_text_box(sl, title, x+0.15, y+0.15, 5.8, 0.45, size=13, bold=True, color=NAVY)
    add_text_box(sl, body,  x+0.15, y+0.65, 5.8, 1.8,  size=10, color=RGBColor(0x1e,0x29,0x3b))

# ─────────────────────────────────────────────
# SLIDE 5 — DEBUGGING FLOW
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, LGRAY)
add_rect(sl, 0, 0, 13.33, 1.1, NAVY)
add_text_box(sl, "Debugging Decision Flow — 10-Step Process", 0.5, 0.2, 12, 0.7, size=24, bold=True, color=WHITE)

steps = [
    ("Error\nDetected",   RED),
    ("Log\nCaptured",     BLUE),
    ("Alert\nTriggered",  GREEN),
    ("Classify\nType",    RGBColor(0x93,0x33,0xea)),
    ("Reproduce\nin DEV", GOLD),
    ("Apply\nFix",        RGBColor(0x13,0x4e,0x4a)),
    ("Test 50\nOrders",   BLUE),
    ("Deploy\nto PROD",   GREEN),
    ("Monitor\n24hr",     RGBColor(0x14,0x53,0x2d)),
    ("Close\nIncident",   NAVY),
]
for i, (label, color) in enumerate(steps):
    x = 0.3 + i * 1.27
    add_rect(sl, x, 1.4, 1.1, 0.9, color)
    add_text_box(sl, label, x+0.05, 1.42, 1.0, 0.86, size=8, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if i < len(steps)-1:
        add_text_box(sl, "→", x+1.1, 1.6, 0.18, 0.5, size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

# Exception path boxes
paths = [
    ("System Exception", BLUE,  "Auto-retry 3× → If persistent: pause bot → Notify DevOps → Hotfix selector/infrastructure",        0.3),
    ("Business Exception",GOLD, "Move to Exception Queue → Human review within 2 hours → Reprocess with corrected data",              4.5),
    ("App Exception",    RGBColor(0x93,0x33,0xea), "Enable OCR fallback → Reroute to alternate parser → Log for vendor notification", 8.7),
]
for title, color, body, x in paths:
    add_rect(sl, x, 2.65, 4.0, 1.5, WHITE)
    add_rect(sl, x, 2.65, 4.0, 0.07, color)
    add_text_box(sl, title, x+0.12, 2.72, 3.76, 0.38, size=11, bold=True, color=color)
    add_text_box(sl, body,  x+0.12, 3.12, 3.76, 0.9,  size=9,  color=NAVY)

add_text_box(sl, "Each exception type has its own handling path — reducing Mean Time to Resolution (MTTR)", 0.3, 4.5, 12.7, 0.4, size=11, color=MGRAY, align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────
# SLIDE 6 — OPTIMIZATION
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, LGRAY)
add_rect(sl, 0, 0, 13.33, 1.1, NAVY)
add_text_box(sl, "Section 3 — Bot Optimization Strategies", 0.5, 0.2, 12, 0.7, size=24, bold=True, color=WHITE)

opt_cards = [
    ("⚡ Performance",  BLUE,  "• 5 parallel bot lanes via Work Queues\n• DB composite index: 30s → <200ms\n• Lazy UI loading: -40% wait time\n• Batch inventory API calls (50 per batch)"),
    ("🛡️ Reliability",  GREEN, "• Idempotency keys eliminate duplicates\n• Dynamic anchor-based selectors\n• JSON schema validation on ERP feed\n• Circuit breaker on DB error spikes"),
    ("🔧 Maintainability",GOLD,"• Centralized selector library in Orchestrator\n• Modular .xaml workflows per stage\n• Config-driven (no code deploy for params)\n• OCR pipeline for scanned PDFs"),
    ("🗓️ Scheduling",   RGBColor(0x93,0x33,0xea), "• Peak (8AM–6PM): 5 bot instances\n• Off-peak (6PM–8AM): 2 instances\n• Maintenance: 2AM–4AM daily window\n• VIP queue: 15-minute SLA guarantee"),
]
for i, (title, color, body) in enumerate(opt_cards):
    x = 0.3 + (i%2)*6.4
    y = 1.25 + (i//2)*2.8
    add_rect(sl, x, y, 6.1, 2.55, WHITE)
    add_rect(sl, x, y, 6.1, 0.08, color)
    add_text_box(sl, title, x+0.15, y+0.15, 5.8, 0.45, size=13, bold=True, color=NAVY)
    add_text_box(sl, body,  x+0.15, y+0.65, 5.8, 1.8,  size=10, color=RGBColor(0x1e,0x29,0x3b))

# ─────────────────────────────────────────────
# SLIDE 7 — BEFORE / AFTER
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, LGRAY)
add_rect(sl, 0, 0, 13.33, 1.1, NAVY)
add_text_box(sl, "Before vs. After Optimization — Key Metrics", 0.5, 0.2, 12, 0.7, size=24, bold=True, color=WHITE)

# Before column
add_rect(sl, 0.3, 1.2, 5.8, 5.9, RGBColor(0xff,0xf5,0xf5))
add_rect(sl, 0.3, 1.2, 5.8, 0.5, RED)
add_text_box(sl, "❌  BEFORE Optimization", 0.45, 1.25, 5.5, 0.4, size=13, bold=True, color=WHITE)
before_items = [
    "Success rate: 71%",
    "Processing time: 4.2 min/order",
    "Duplicate orders: ~12 incidents/week",
    "DB timeouts: 80+ per peak hour",
    "Manual overhead: 3.5 hours/day",
    "Selector failures after every frontend deploy",
    "Throughput: 120 orders/hour",
    "Config changes require code deployment",
]
for i, item in enumerate(before_items):
    add_text_box(sl, "✗  " + item, 0.45, 1.85+i*0.55, 5.5, 0.5, size=11, color=RGBColor(0x7f,0x1d,0x1d))

# After column
add_rect(sl, 7.2, 1.2, 5.8, 5.9, RGBColor(0xf0,0xfd,0xf4))
add_rect(sl, 7.2, 1.2, 5.8, 0.5, GREEN)
add_text_box(sl, "✅  AFTER Optimization", 7.35, 1.25, 5.5, 0.4, size=13, bold=True, color=WHITE)
after_items = [
    "Success rate: 97.4% (+26%)",
    "Processing time: 1.8 min/order (57% faster)",
    "Duplicate orders: 0 incidents",
    "DB timeouts: eliminated (<200ms)",
    "Manual overhead: 12 min/day (96% less)",
    "Dynamic selectors survive all deployments",
    "Throughput: 600 orders/hour (5× increase)",
    "Zero-downtime config updates via Orchestrator",
]
for i, item in enumerate(after_items):
    add_text_box(sl, "✓  " + item, 7.35, 1.85+i*0.55, 5.5, 0.5, size=11, color=RGBColor(0x06,0x5f,0x46))

# VS divider
add_text_box(sl, "VS", 6.3, 3.7, 0.7, 0.7, size=20, bold=True, color=NAVY, align=PP_ALIGN.CENTER)

# ─────────────────────────────────────────────
# SLIDE 8 — DASHBOARD KPIS
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, LGRAY)
add_rect(sl, 0, 0, 13.33, 1.1, NAVY)
add_text_box(sl, "Section 4 — Monitoring Dashboard KPIs", 0.5, 0.2, 12, 0.7, size=24, bold=True, color=WHITE)

kpis = [
    ("Total Orders Today", "4,827",  "▲ 12% vs yesterday",        BLUE),
    ("Success Rate",       "97.4%",  "▲ Above 95% SLA target",     GREEN),
    ("Failed Orders",      "125",    "▼ 43% vs pre-optimization",  RED),
    ("Retries Executed",   "218",    "▼ 60% improvement",          RGBColor(0xf9,0x73,0x16)),
    ("Avg Process Time",   "1.8 min","▼ 57% faster per order",     RGBColor(0x14,0xb8,0xa6)),
    ("Bot Uptime (30d)",   "99.2%",  "▲ Target: 99% — Exceeded",   RGBColor(0x93,0x33,0xea)),
]
for i, (label, value, trend, color) in enumerate(kpis):
    x = 0.3 + (i%3)*4.25
    y = 1.3  + (i//3)*2.65
    add_rect(sl, x, y, 3.95, 2.35, color)
    add_text_box(sl, label, x+0.15, y+0.18, 3.65, 0.4,  size=10, bold=True, color=WHITE)
    add_text_box(sl, value, x+0.15, y+0.6,  3.65, 1.0,  size=30, bold=True, color=WHITE)
    add_text_box(sl, trend, x+0.15, y+1.75, 3.65, 0.42, size=10, color=RGBColor(0xa7,0xf3,0xd0))

# ─────────────────────────────────────────────
# SLIDE 9 — AI USAGE
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, LGRAY)
add_rect(sl, 0, 0, 13.33, 1.1, NAVY)
add_text_box(sl, "Section 5 — AI-Assisted Development: Vibe Coding with Claude Code", 0.5, 0.2, 12.3, 0.7, size=22, bold=True, color=WHITE)

add_rect(sl, 0.3, 1.2, 12.73, 1.1, RGBColor(0xef,0xf6,0xff))
add_text_box(sl, "What is Vibe Coding?", 0.5, 1.25, 5, 0.4, size=12, bold=True, color=NAVY)
add_text_box(sl,
    "Vibe coding = describing what you want in natural language → AI generates the implementation.\n"
    "The developer acts as architect & reviewer; Claude Code handles implementation details.",
    0.5, 1.6, 12.3, 0.55, size=10, color=RGBColor(0x1e,0x29,0x3b))

ai_items = [
    ("💬 Prompt Engineering",    RGBColor(0x25,0x63,0xeb), "Full project requirements described in natural language → Claude structured architecture, layout, and component design from plain English prompts."),
    ("🎨 UI/UX Generation",      GREEN,                    "Design specs like 'enterprise-style, navy/gold color scheme, card layout' became actual CSS gradients, grid systems, and styled components — no manual CSS."),
    ("📋 Domain Synthesis",      GOLD,                     "RPA error scenarios, root causes, and business impacts generated by feeding Claude the domain context — technically accurate UiPath + e-commerce content."),
    ("🔄 Iterative Refinement",  RED,                      "Each section refined through prompt → generate → review loop. The debugging flow evolved from a list to a visual step-by-step arrow flow via a single prompt."),
    ("⚙️ Full Stack Generation", RGBColor(0x93,0x33,0xea), "Claude Code generated HTML/CSS/JS, Python scripts for PPT/Word, folder structure, GitHub setup commands, and this video script — all from conversation."),
    ("📊 Consistent Mock Data",  RGBColor(0x14,0xb8,0xa6), "Dashboard KPIs internally consistent and realistic for mid-scale e-commerce. Before/after metrics logically aligned with described optimizations."),
]
for i, (title, color, body) in enumerate(ai_items):
    x = 0.3 + (i%2)*6.45
    y = 2.5  + (i//2)*1.55
    add_rect(sl, x, y, 6.15, 1.4, WHITE)
    add_rect(sl, x, y, 6.15, 0.07, color)
    add_text_box(sl, title, x+0.12, y+0.12, 5.9, 0.35, size=11, bold=True, color=NAVY)
    add_text_box(sl, body,  x+0.12, y+0.5,  5.9, 0.8,  size=9,  color=RGBColor(0x1e,0x29,0x3b))

# ─────────────────────────────────────────────
# SLIDE 10 — KEY METRICS SUMMARY
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, NAVY)
add_text_box(sl, "Key Results Summary", 0.5, 0.4, 12, 0.8, size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_rect(sl, 3, 1.1, 7.33, 0.07, GOLD)

metrics = [
    ("↑26%",   "Success Rate Gain",     "71% → 97.4%",    BLUE),
    ("5×",     "Throughput Increase",   "120 → 600 orders/hr", GREEN),
    ("57%",    "Faster Processing",     "4.2min → 1.8min",     RGBColor(0xf9,0x73,0x16)),
    ("96%",    "Manual Work Reduced",   "3.5hr → 12min/day",   RGBColor(0x93,0x33,0xea)),
    ("0",      "Duplicate Orders",      "Per month post-fix",  RED),
    ("99.2%",  "Bot Uptime (30-day)",   "Exceeds 99% target",  RGBColor(0x14,0xb8,0xa6)),
]
for i, (val, label, sub, color) in enumerate(metrics):
    x = 0.5 + (i%3)*4.1
    y = 1.5  + (i//3)*2.5
    add_rect(sl, x, y, 3.7, 2.15, RGBColor(0x1e,0x4a,0x76))
    add_rect(sl, x, y, 3.7, 0.09, color)
    add_text_box(sl, val,   x+0.15, y+0.2,  3.4, 0.8, size=36, bold=True, color=color)
    add_text_box(sl, label, x+0.15, y+1.05, 3.4, 0.45, size=12, bold=True, color=WHITE)
    add_text_box(sl, sub,   x+0.15, y+1.5,  3.4, 0.45, size=10, color=RGBColor(0x93,0xc5,0xfd))

# ─────────────────────────────────────────────
# SLIDE 11 — CLOSING
# ─────────────────────────────────────────────
sl = prs.slides.add_slide(blank_layout)
slide_bg(sl, NAVY)
add_rect(sl, 0, 2.8, 13.33, 0.12, GOLD)
add_text_box(sl, "Thank You", 1, 0.8, 11, 1.2, size=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
add_text_box(sl, "RPA Course · Assignment 04 · E-Commerce Order Processing Bot Case Study", 1, 2.1, 11, 0.55, size=14, color=RGBColor(0x93,0xc5,0xfd), align=PP_ALIGN.CENTER)
add_text_box(sl, "Omais Siddiqui  |  April 2026  |  Built with Claude Code AI-Assisted Development", 1, 3.2, 11, 0.5, size=12, color=RGBColor(0x64,0x74,0x8b), align=PP_ALIGN.CENTER)
add_text_box(sl, "All project files available in:  /Desktop/Assignment_04_RPA/", 1, 3.85, 11, 0.45, size=11, color=GOLD, align=PP_ALIGN.CENTER)

out = "/Users/omaissaeedsiddiqui/Desktop/Assignment_04_RPA/ppt/Assignment_04_RPA_Presentation.pptx"
prs.save(out)
print(f"Saved: {out}")
