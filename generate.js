const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, Footer, Header, TabStopType, TabStopPosition,
  PageBreak, PageNumber, PageNumberElement
} = require('docx');
const fs = require('fs');

// Color palette
const C = {
  navy:    "1F3864",
  blue:    "2E5FA3",
  lightBlue: "D9E2F3",
  midBlue: "BDD0EA",
  accent:  "4472C4",
  teal:    "1F7391",
  gray1:   "F2F5FA",
  gray2:   "E8EDF5",
  gray3:   "D0D8E8",
  white:   "FFFFFF",
  text:    "1A1A2E",
  muted:   "555577",
  code:    "F0F4FF",
  codeBorder: "B0BCD8",
};

const border = (color="CCCCCC") => ({ style: BorderStyle.SINGLE, size: 1, color });
const borders = (color="CCCCCC") => ({ top: border(color), bottom: border(color), left: border(color), right: border(color) });
const noBorder = () => ({ style: BorderStyle.NONE, size: 0, color: "FFFFFF" });
const noBorders = () => ({ top: noBorder(), bottom: noBorder(), left: noBorder(), right: noBorder() });

// Content width: 9360 DXA (US Letter - 1" margins each side)
const CONTENT_WIDTH = 9360;

function hr() {
  return new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.midBlue, space: 1 } },
    spacing: { before: 120, after: 120 }
  });
}

function spacer(before=80, after=80) {
  return new Paragraph({ children: [], spacing: { before, after } });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial", size: 36, bold: true, color: C.navy })],
    spacing: { before: 400, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.blue, space: 4 } }
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial", size: 28, bold: true, color: C.blue })],
    spacing: { before: 300, after: 120 }
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial", size: 24, bold: true, color: C.teal })],
    spacing: { before: 240, after: 100 }
  });
}

function h4(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_4,
    children: [new TextRun({ text, font: "Arial", size: 22, bold: true, color: C.navy })],
    spacing: { before: 200, after: 80 }
  });
}

function para(text, opts={}) {
  return new Paragraph({
    children: [new TextRun({ text, font: "Arial", size: 22, color: C.text, ...opts })],
    spacing: { before: 60, after: 80 }
  });
}

function inlineCode(text) {
  return new TextRun({ text, font: "Courier New", size: 20, color: C.teal, bold: true });
}

function codeBlock(lines) {
  const content = Array.isArray(lines) ? lines.join('\n') : lines;
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [new TableRow({
      children: [new TableCell({
        borders: borders(C.codeBorder),
        shading: { fill: C.code, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 200, right: 200 },
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        children: content.split('\n').map(line =>
          new Paragraph({
            children: [new TextRun({ text: line || ' ', font: "Courier New", size: 18, color: "1D3461" })],
            spacing: { before: 0, after: 0 }
          })
        )
      })]
    })]
  });
}

function headerRow(cells, widths) {
  return new TableRow({
    tableHeader: true,
    children: cells.map((c, i) => new TableCell({
      borders: borders(C.blue),
      shading: { fill: C.navy, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 140, right: 140 },
      width: { size: widths[i], type: WidthType.DXA },
      children: [new Paragraph({
        children: [new TextRun({ text: c, font: "Arial", size: 20, bold: true, color: C.white })],
        spacing: { before: 0, after: 0 }
      })]
    }))
  });
}

function dataRow(cells, widths, shade=false) {
  return new TableRow({
    children: cells.map((c, i) => new TableCell({
      borders: borders(C.gray3),
      shading: { fill: shade ? C.gray1 : C.white, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 140, right: 140 },
      width: { size: widths[i], type: WidthType.DXA },
      children: [new Paragraph({
        children: [new TextRun({ text: c, font: "Arial", size: 20, color: C.text })],
        spacing: { before: 0, after: 0 }
      })]
    }))
  });
}

function bullet(text, level=0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    children: [new TextRun({ text, font: "Arial", size: 22, color: C.text })],
    spacing: { before: 40, after: 40 }
  });
}

function infoBox(title, text) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH],
    rows: [
      new TableRow({ children: [new TableCell({
        borders: borders(C.blue),
        shading: { fill: C.lightBlue, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 20, left: 160, right: 160 },
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: title, font: "Arial", size: 22, bold: true, color: C.navy })], spacing: { before: 0, after: 0 } })]
      })] }),
      new TableRow({ children: [new TableCell({
        borders: borders(C.blue),
        shading: { fill: C.gray1, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 100, left: 160, right: 160 },
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text, font: "Arial", size: 21, color: C.text })], spacing: { before: 0, after: 0 } })]
      })] })
    ]
  });
}

// ─── TITLE PAGE ────────────────────────────────────────────────────────────────
function buildTitlePage() {
  return [
    spacer(2000, 0),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [CONTENT_WIDTH],
      rows: [
        new TableRow({ children: [new TableCell({
          borders: noBorders(),
          shading: { fill: C.navy, type: ShadingType.CLEAR },
          margins: { top: 400, bottom: 120, left: 400, right: 400 },
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "Document Agents", font: "Arial", size: 64, bold: true, color: C.white })],
              spacing: { before: 0, after: 60 }
            }),
          ]
        })] }),
        new TableRow({ children: [new TableCell({
          borders: noBorders(),
          shading: { fill: C.blue, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 240, left: 400, right: 400 },
          width: { size: CONTENT_WIDTH, type: WidthType.DXA },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: "Comprehensive Architecture & Implementation Guide", font: "Arial", size: 28, color: C.lightBlue, italics: true })],
              spacing: { before: 0, after: 0 }
            })
          ]
        })] }),
      ]
    }),
    spacer(400, 0),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [CONTENT_WIDTH],
      rows: [new TableRow({ children: [new TableCell({
        borders: borders(C.midBlue),
        shading: { fill: C.gray1, type: ShadingType.CLEAR },
        margins: { top: 200, bottom: 200, left: 360, right: 360 },
        width: { size: CONTENT_WIDTH, type: WidthType.DXA },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Repository: document-agents  |  Organization: Peer AI", font: "Arial", size: 22, color: C.muted })], spacing: { before: 0, after: 60 } }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Domain: AI-powered clinical/regulatory document processing", font: "Arial", size: 22, color: C.muted })], spacing: { before: 0, after: 60 } }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Primary LLM: OpenAI GPT-5.4 via LiteLLM  |  Agent Framework: Google ADK", font: "Arial", size: 22, color: C.navy, bold: true })], spacing: { before: 0, after: 0 } }),
        ]
      })] })]
    }),
    spacer(1000, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── OVERVIEW SECTION ─────────────────────────────────────────────────────────
function buildOverview() {
  const w = [3120, 3120, 3120];
  return [
    h1("1. High-Level Overview"),
    para("This repository contains a suite of AI-powered agents that automate the creation and review of clinical/regulatory documents. The system processes source documents (protocols, SAPs, TFLs, templates, style guides) uploaded to S3 and produces structured JSON templates ('roadmaps') that a downstream content execution platform renders into final documents."),
    spacer(),
    h2("System Architecture"),
    para("The system uses a Gateway Router as the unified entry point. An AWS SQS message containing an operation_id triggers the Gateway, which resolves the document type from the backend API and routes to the correct agent."),
    spacer(),
    infoBox("Architecture Flow", "SQS Message → Gateway Router → [CSR Agent | Generic Roadmap Agent] — Data Review Agent runs standalone, triggered by a separate SQS message."),
    spacer(120, 0),
    h2("Agent Summary"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: w,
      rows: [
        headerRow(["Agent", "Input", "Purpose"], w),
        dataRow(["Data Review Agent", "Uploaded PDFs (protocol, SAP, TFLs)", "Analyzes and catalogs uploaded documents for downstream agents"], w, false),
        dataRow(["CSR Agent", "Sponsor template + protocol + SAP + TFLs + style guides", "Generates CSR content roadmaps as structured JSON"], w, true),
        dataRow(["Generic Roadmap Agent", "Template/example PDFs + source documents", "Generates roadmaps for ANY clinical document type"], w, false),
      ]
    }),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── TECH STACK ────────────────────────────────────────────────────────────────
function buildTechStack() {
  const w1 = [2600, 2800, 3960];
  const w2 = [3120, 6240];
  return [
    h1("2. Technology Stack"),
    h2("Core Technologies"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: w1,
      rows: [
        headerRow(["Component", "Technology", "Purpose"], w1),
        dataRow(["Language", "Python 3.11+", "All agents"], w1, false),
        dataRow(["Agent Framework", "Google ADK (google.adk)", "Agent orchestration, tool calling, session state"], w1, true),
        dataRow(["LLM Gateway", "LiteLLM", "Unified API for multiple LLM providers"], w1, false),
        dataRow(["Primary LLM", "OpenAI GPT-5.4", "All reasoning, content generation, classification"], w1, true),
        dataRow(["Vision Model", "OpenAI GPT-5.4 (vision mode)", "PDF page analysis, heading detection"], w1, false),
        dataRow(["Data Validation", "Pydantic v2", "Schema validation for all DTOs and template output"], w1, true),
        dataRow(["PDF Processing", "PyMuPDF (fitz)", "Text extraction, page rendering, table detection"], w1, false),
        dataRow(["Dependency Mgmt", "Poetry", "Per-agent dependency isolation"], w1, true),
        dataRow(["Containerization", "Docker + Docker Compose", "Production deployment"], w1, false),
        dataRow(["Object Storage", "AWS S3 / MinIO (local)", "PDF storage and retrieval"], w1, true),
        dataRow(["Message Queue", "AWS SQS", "Trigger agent execution"], w1, false),
        dataRow(["Authentication", "AWS Cognito (OAuth 2.0)", "Service-to-service auth"], w1, true),
        dataRow(["Secrets", "AWS Secrets Manager", "API keys, client secrets"], w1, false),
        dataRow(["Observability", "Phoenix (Arize)", "Distributed tracing, LLM call monitoring"], w1, true),
        dataRow(["HTTP Client", "httpx (async)", "Backend API communication"], w1, false),
      ]
    }),
    spacer(200, 0),
    h2("Key Python Dependencies"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: w2,
      rows: [
        headerRow(["Package", "Purpose"], w2),
        dataRow(["google-adk", "Agent Development Kit"], w2, false),
        dataRow(["litellm", "LLM API abstraction"], w2, true),
        dataRow(["pydantic", "Data validation"], w2, false),
        dataRow(["pymupdf (fitz)", "PDF processing"], w2, true),
        dataRow(["httpx", "Async HTTP client"], w2, false),
        dataRow(["phoenix-otel", "Observability / tracing"], w2, true),
        dataRow(["boto3", "AWS SDK"], w2, false),
        dataRow(["peerai-db", "Internal DB package"], w2, true),
        dataRow(["peerai-logging", "Internal logging package"], w2, false),
      ]
    }),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── ADK CONCEPTS ─────────────────────────────────────────────────────────────
function buildADK() {
  return [
    h1("3. Agent Framework Concepts (Google ADK)"),
    para("All agents in this repository are built on Google's Agent Development Kit (ADK), which provides several key abstractions for building AI-powered pipelines."),
    spacer(),

    h2("3.1 Agent Types"),

    h3("Agent — Tool-Calling Agent"),
    para("The standard ADK agent. It receives an instruction (system prompt), a model, and a list of tools. The LLM can decide to call tools, and ADK handles the tool-call dispatch loop automatically."),
    bullet("Used by: CSR Agent's individual agents (Sponsor Template Agent, Study Agent, Decision Agent), Data Review Agent"),
    spacer(80, 0),
    codeBlock([
      "agent = Agent(",
      "    name='sponsor_template_agent',",
      "    model=LiteLlm(model='openai/responses/gpt-5.4'),",
      "    instruction='You are a sponsor template analyst...',",
      "    tools=[FunctionTool(func=search_sponsor_template_tool)],",
      "    output_key='template_instructions',",
      ")"
    ]),

    spacer(160, 0),
    h3("LlmAgent — Instruction-Only Agent"),
    para("A lighter agent that makes LLM calls without multi-turn tool loops. Used when the agent needs a single LLM generation and writes the result to state."),
    bullet("Used by: Generic Roadmap Agent's Structure Agent, Content Planner, Prompt Writer, Validation Fixer"),
    spacer(80, 0),
    codeBlock([
      "agent = LlmAgent(",
      "    name='structure_agent',",
      "    model=LiteLlm(model='openai/responses/gpt-5.4', reasoning_effort='high'),",
      "    instruction=build_structure_decide_instruction(...),",
      "    tools=[],",
      "    output_key='section_plan',",
      ")"
    ]),

    spacer(160, 0),
    h3("SequentialAgent — Pipeline Orchestrator"),
    para("Executes sub-agents one after another in a fixed order. Each agent can read the outputs of all previous agents via session state."),
    bullet("Used by: CSR Agent's main pipeline, Generic Roadmap Agent's per-section pipeline"),

    spacer(160, 0),
    h3("ParallelAgent — Concurrent Execution"),
    para("Executes sub-agents concurrently. All sub-agents must write to different state keys to avoid conflicts. Provides ~50% time savings for parallel document extraction."),
    bullet("Used by: CSR Agent (protocol + SAP in parallel), Generic Roadmap Agent (tier-based parallel content generation)"),

    spacer(160, 0),
    h3("BaseAgent — Custom Agent"),
    para("The most flexible option. You subclass BaseAgent and implement _run_async_impl() as an async generator that yields Event objects. This allows full programmatic control over the execution flow."),
    bullet("Used by: Generic Roadmap Agent's RoadmapOrchestrator, ContentGenerationOrchestrator, PageAnalyzer, SectionAssembler, PromptSanitizer"),

    spacer(160, 0),
    h2("3.2 Key ADK Concepts"),
    h3("Session State"),
    para("A shared dictionary (ctx.session.state) that persists across all agents in a pipeline. Agents write results using output_key and read other agents' outputs from state."),

    spacer(80, 0),
    h3("FunctionTool"),
    para("Wraps a Python function (sync or async) as a tool the LLM can call. ADK handles serialization, dispatch, and result injection into the conversation."),

    spacer(80, 0),
    h3("ToolContext"),
    para("Injected into tool functions automatically. Provides access to session state, allowing tools to cache results and read shared data without explicit parameter passing."),

    spacer(80, 0),
    h3("LiteLlm Model Wrapper"),
    para("ADK's bridge to LiteLLM, supporting reasoning effort, retry policies, and provider-specific options."),
    codeBlock([
      "model = LiteLlm(",
      "    model='openai/responses/gpt-5.4',",
      "    reasoning_effort='high',",
      "    timeout=1200,",
      "    num_retries=3,",
      "    retry_strategy='exponential_backoff_retry',",
      ")"
    ]),

    spacer(80, 0),
    h3("Lazy Initialization Pattern"),
    para("Agents are NOT created at import time. Instead, Python's __getattr__ module hook delays creation until first access, ensuring catalogs and file indexes are initialized before agents are built."),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── GATEWAY ──────────────────────────────────────────────────────────────────
function buildGateway() {
  return [
    h1("4. Gateway Router"),
    para("The Gateway is the unified entry point for the document_agents Docker container. It routes incoming SQS messages to the correct agent based on document type."),
    spacer(80, 0),

    h2("Agent Registry"),
    codeBlock([
      "AGENT_REGISTRY = {",
      "    'CSR':      'csr_agent.src.server',",
      "    'PROTOCOL': 'generic_roadmap_agent.src.server',",
      "}"
    ]),
    spacer(160, 0),

    h2("Execution Flow"),
    bullet("Step 1 — Load Environment: Load .env (Docker: /app/.env from SSM, Local: gateway/.env), validate agent registry"),
    bullet("Step 2 — Resolve Routing: Extract operation_id from SQS_MESSAGE env var, call backend API with 3 retries + exponential backoff to resolve document type"),
    bullet("Step 3 — Activate Agent: Map document_type to agent module, activate agent's virtualenv (prepend site-packages), load agent-specific .env"),
    bullet("Step 4 — Dispatch: importlib.import_module(agent_module_path) then await agent_module.main_async()"),
    bullet("Step 5 — Error Handling: Phase 1 (gateway-owned) vs Phase 2 (agent-owned) errors handled separately"),
    spacer(160, 0),

    h2("Key Features"),
    para("Per-agent virtualenvs ensure isolated Python dependencies. The retry mechanism uses exponential backoff: 1s → 2s → 4s for document type resolution. A health check at startup validates all registered agents have venvs and source directories."),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── DATA REVIEW AGENT ────────────────────────────────────────────────────────
function buildDataReview() {
  const w1 = [3120, 2080, 4160];
  const w2 = [3500, 1500, 4360];
  return [
    h1("5. Data Review Agent"),
    para("The Data Review Agent is the first step in the document processing pipeline. When a user uploads source files, this agent downloads and analyzes each file, classifies file roles, extracts Tables of Contents, splits multi-section TFL files, creates structured data source records, and performs a data sufficiency review."),
    spacer(80, 0),

    h2("Model Configuration"),
    codeBlock([
      "MODEL = 'openai/gpt-5.4'",
      "REASONING_EFFORT = 'low'            # Default (fast, cheap)",
      "REVIEW_AGENT_REASONING_EFFORT = 'medium'  # Data review assessment"
    ]),
    spacer(160, 0),

    h2("Pipeline Steps"),
    bullet("Step 1 — Get Operation Metadata: Fetch operation from API, get document name/type, initialize thinking trace manager"),
    bullet("Step 2 — Check Reusable Data Sources: Query existing data sources by source_file_id, check checksums, skip processing for reusable files"),
    bullet("Step 3 — Download Files from S3: Fetch source file metadata from API, download each file, build sanitized local filenames"),
    bullet("Step 4 — Classify File Roles: FileRoleClassifier extracts first 2 pages as text + images, sends to LLM for role classification (Protocol, SAP, TFL, Template, etc.)"),
    bullet("Step 5 — Process Each File: TOC extraction, TFL splitting (if applicable), data source creation"),
    bullet("Step 6 — Data Review: Extract text/images from pages 1-2, batch files (batch size=5), send to LLM with review prompt"),
    bullet("Step 7 — Update Statuses: Update template with data_source_details, update document/operation status"),
    spacer(160, 0),

    h2("Key Services"),
    h3("TOCProcessorService"),
    para("Extracts Table of Contents from PDF documents by scanning pages 1-60, using LLM to parse TOC entries, performing QC verification (sampling entries and verifying content at claimed page ranges), and applying fallback content-analysis extraction when needed."),

    h3("FileRoleClassifier"),
    para("Classifies uploaded files into roles using LLM. Extracts first 1-2 pages as text + images (multimodal) and returns classification: Protocol, SAP, TFL, Template, Example CSR, Style Guide, etc."),

    h3("LLMClient"),
    para("Custom wrapper around LiteLLM with sync/async call support, retry logic with exponential backoff on transient errors (502, 503, 5xx, timeouts, rate limits), and Phoenix callback management."),
    spacer(160, 0),

    h2("Configuration Constants"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: w1,
      rows: [
        headerRow(["Constant", "Value", "Purpose"], w1),
        dataRow(["TOC_START_PAGE", "1", "Start scanning for TOC"], w1, false),
        dataRow(["TOC_END_PAGE", "60", "End scanning for TOC"], w1, true),
        dataRow(["TOC_QC_SAMPLE_SIZE", "10", "Number of TOC entries to sample for QC"], w1, false),
        dataRow(["TOC_QC_THRESHOLD", "0.70", "Minimum pass rate for TOC QC"], w1, true),
        dataRow(["DATA_REVIEW_BATCH_SIZE", "5", "Files per LLM review batch"], w1, false),
        dataRow(["DATA_REVIEW_START_PAGE", "1", "Start page for data review"], w1, true),
        dataRow(["DATA_REVIEW_END_PAGE", "2", "End page for data review"], w1, false),
        dataRow(["CONTENT_ANALYSIS_BATCH_SIZE", "10", "Pages per content analysis batch"], w1, true),
        dataRow(["MAX_DATA_SOURCE_TABLE_NAME_LENGTH", "25", "Max chars for DB table name"], w1, false),
        dataRow(["PIPELINE_VERSION", "1.0.0", "Version tag for cache invalidation"], w1, true),
      ]
    }),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── CSR AGENT ────────────────────────────────────────────────────────────────
function buildCSR() {
  const w3 = [2200, 1400, 5760];
  const w4 = [2600, 6760];
  return [
    h1("6. CSR Agent"),
    para("The CSR Agent automates Clinical Study Report (CSR) authoring by processing sponsor templates, study documents, and generating structured writing prompts for each section of the CSR."),
    spacer(80, 0),

    infoBox("Trigger", "The CSR Agent is triggered by an SQS message containing an operation_id. The Gateway Router resolves document type as 'CSR' and dispatches to csr_agent.src.server."),
    spacer(160, 0),

    h2("Source Document Categories"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [2400, 1200, 5760],
      rows: [
        headerRow(["Document", "Required", "Description"], [2400, 1200, 5760]),
        dataRow(["Sponsor Template PDF", "Yes", "The CSR shell template with instructional text, red-text guidelines, table structures, and section headings"], [2400, 1200, 5760], false),
        dataRow(["Protocol PDF(s)", "No", "Clinical study protocol — study design, objectives, endpoints, inclusion/exclusion criteria, visit schedules"], [2400, 1200, 5760], true),
        dataRow(["SAP PDF(s)", "No", "Statistical Analysis Plan — analysis populations, statistical methods, handling of missing data, multiplicity adjustments"], [2400, 1200, 5760], false),
        dataRow(["TFL PDF(s)", "No", "Tables, Figures, and Listings — actual clinical data outputs. Required for data sections 10-13"], [2400, 1200, 5760], true),
        dataRow(["Style Guide PDF", "No", "Sponsor-specific formatting and style conventions (abbreviation standards, table formatting rules, terminology preferences)"], [2400, 1200, 5760], false),
      ]
    }),
    spacer(200, 0),

    h2("3-Tier Document Classification System"),
    h3("Tier 1 — DataSourceType"),
    para("Documents with type == 'PROTOCOL' are classified as Non-TFL. All others (TABLES/LISTINGS/FIGURES/CIOMS) are TFLs and downloaded to the tfl/ subdirectory."),
    h3("Tier 2 — source_context_role"),
    para("Distinguishes between TEMPLATE (sponsor template PDF), REFERENCE (style guide), EXAMPLE (example document), SOURCE (study document), or falls back to display_name/key heuristics."),
    h3("Tier 3 — SOURCE Sub-classification"),
    para("Classifies SOURCE documents as SAP (if 'sap' or 'statistical analysis plan' appears in key/name) or Protocol (if 'protocol' appears). Falls back to GPT-5.4 LLM classification."),
    spacer(160, 0),

    h2("Two-Phase Section Routing"),
    bullet("Phase 1 — Shell Sections (1-9): If template does NOT have content_details for all sections 10-13. Processes sections 1-4 (static) and 5-9 (extracted from template TOC)."),
    bullet("Phase 2 — Data Sections (10-13): If template ALREADY has content_details for sections 10-13 from a prior run. Requires TFL inputs gate (tfl_dir_path with files + tfl_catalog.json)."),
    spacer(160, 0),

    h2("Sequential Agent Pipeline (Per Section)"),
    para("Each section passes through a 10-step sequential pipeline:"),
    bullet("Step 1 — Sponsor Template Agent: Extracts template requirements using search_sponsor_template_tool; writes to state['template_instructions']"),
    bullet("Step 2 — Decision Agent: Detects red text (instructional) + table structures; sets flags needs_static_instructional_agent, needs_table_agent"),
    bullet("Step 3 — Parallel Study Data Extraction: Protocol Agent + SAP Agent run concurrently (~50% time savings vs sequential)"),
    bullet("Step 4 — Study Agent: Merges protocol + SAP outputs, applies SAP-driven scope filtering; writes to state['study_context']"),
    bullet("Step 5 — Static Instructional Agent: Conditional — only if needs_static_instructional_agent; extracts verbatim red-colored instructional text"),
    bullet("Step 6 — Table Prompt Generator Agent: Conditional — only if needs_table_agent; generates table-specific writing prompts"),
    bullet("Step 7 — Prompt Generator Agent: Combines template + study + style + static content to generate comprehensive writing prompts"),
    bullet("Step 8 — Prompt Refinement Agent: Cross-references with sponsor template requirements, refines and polishes the prompt"),
    bullet("Step 9 — JSON Output Agent: Validates generated prompts against schema, saves to JSON files"),
    spacer(160, 0),

    h2("Environment Variables"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: w4,
      rows: [
        headerRow(["Variable", "Purpose"], w4),
        dataRow(["SQS_MESSAGE", "JSON with operation_id (primary trigger)"], w4, false),
        dataRow(["OPERATION_ID", "Fallback operation ID for local testing"], w4, true),
        dataRow(["S3_BUCKET / MINIO_BUCKET_NAME", "S3/MinIO bucket for source documents"], w4, false),
        dataRow(["OPENAI_API_KEY / OPENAI_SECRET_KEY", "OpenAI API key (direct or AWS Secrets Manager name)"], w4, true),
        dataRow(["START_SECTION", "First section to process (default: 1)"], w4, false),
        dataRow(["SECTION_NUMBER", "Last section to process (default: 13)"], w4, true),
        dataRow(["SKIP_SECTIONS", "Comma-separated section numbers to skip"], w4, false),
        dataRow(["CSR_AGENT_MODEL", "LLM model identifier (default: openai/gpt-5.4)"], w4, true),
        dataRow(["GPT_5_REASONING_EFFORT", "Reasoning effort level (default: high)"], w4, false),
        dataRow(["CSR_AGENT_MAX_LLM_CALLS", "Max LLM calls per pipeline run (default: 3000)"], w4, true),
        dataRow(["CSR_AGENT_EVAL_PIPELINE", "Enable eval pipeline mode (default: False)"], w4, false),
      ]
    }),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── GENERIC ROADMAP AGENT ────────────────────────────────────────────────────
function buildGeneric() {
  const wDoc = [2600, 1560, 5200];
  const wLimits = [3500, 1500, 4360];
  return [
    h1("7. Generic Roadmap Agent"),
    para("The Generic Roadmap Agent is a document-type-agnostic pipeline that generates structured template definitions ('roadmaps') for ANY clinical document type. It is the most architecturally sophisticated agent in the system."),
    spacer(80, 0),

    h2("Supported Document Types"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: wDoc,
      rows: [
        headerRow(["Document Type", "Key", "Notes"], wDoc),
        dataRow(["Plain Language Summary", "PLAIN_LANGUAGE_SUMMARY", "Grade 6-8 reading level"], wDoc, false),
        dataRow(["Informed Consent Form", "ICF", "Grade 6-8 reading level"], wDoc, true),
        dataRow(["Investigator's Brochure", "IB", "Abbreviation + citation post-processing"], wDoc, false),
        dataRow(["Briefing Book", "BRIEFING_BOOK", "Mixed tense"], wDoc, true),
        dataRow(["Nonclinical Overview (Module 2.4)", "IND_2_4", "Mixed tense"], wDoc, false),
        dataRow(["Clinical Study Report", "CSR", "Past tense"], wDoc, true),
        dataRow(["Clinical Trial Protocol", "PROTOCOL", "Future/present tense"], wDoc, false),
        dataRow(["Toxicology Written Summary (CTD 2.6.6)", "IND_2_6_6", "Past tense"], wDoc, true),
      ]
    }),
    spacer(200, 0),

    h2("4-Phase Pipeline"),
    h3("Phase 1 — Setup (Pure Python, No LLM)"),
    bullet("Build CatalogRegistry from data_source_details + PDF paths"),
    bullet("Resolve DocumentContext for document_type (display name, tense guidance, reading level)"),
    bullet("Load Skills: sponsor + doc-type specific conventions from skills/ directory"),
    bullet("Cache full TEMPLATE/EXAMPLE text for reuse across phases"),
    spacer(120, 0),

    h3("Phase 2 — Structure (LLM)"),
    para("Phase 2a — Route Selection (deterministic): Analyzes data source characteristics and chooses one of four routes:"),
    bullet("simple: <150 pages, ≤3 data sources"),
    bullet("standard: moderate complexity"),
    bullet("appendix_heavy: >5 data sources, >50 avg pages"),
    bullet("tfl_heavy: >20 source docs, <10 avg pages", 1),
    spacer(80, 0),
    para("Phase 2b — Pre-Reader (parallel async): Detects headings from TEMPLATE/EXAMPLE PDFs and performs route-dependent source distillation. All tasks run in parallel via asyncio.gather."),
    para("Phase 2c — Structure Decide: StructureAgent (LlmAgent, reasoning_effort='high') produces HardenedSectionPlan JSON with section keys, headings, tiers, and applicability flags."),
    para("Phase 2d — Structure Review: StructureReviewer agent (reasoning_effort='medium') compares plan against extracted headings and flags corrections."),
    para("Phase 2e — Deterministic Cleanup: Deduplication (exact key → fuzzy heading → level-0) and applicability enforcement."),
    spacer(120, 0),

    h3("Phase 3 — Content Generation (N-tier Parallel + Batched)"),
    para("Deterministic short-circuit handles standard sections without LLM calls: TOC, List of Tables, List of Figures, List of Abbreviations, References, and tables-only sections all get appropriate placeholders."),
    para("For all other sections, N-Tier Execution runs independent (Tier 1) sections in parallel batches of up to 5, then waits for prior tiers before running dependent sections (Tier 2+)."),
    para("Per-section pipeline (SequentialAgent): SourceAllocator → PageAnalyzer → ContentPlanner → PlanValidator → PromptWriter → SectionAssembler → PromptSanitizer"),
    spacer(120, 0),

    h3("Phase 4 — Assembly + Validation"),
    bullet("Deterministic merge of all section outputs from state"),
    bullet("Post-process configuration: ABBREVIATION and CITATION configs for applicable document types"),
    bullet("Schema validation against Pydantic template schema"),
    bullet("Two-phase fix loop (up to 3 retries): deterministic fixer first, then LLM ValidationFixer fallback"),
    bullet("PATCH assembled roadmap to template service API, update operation status"),
    spacer(160, 0),

    h2("Skills System"),
    para("Skills inject sponsor-specific and document-type-specific conventions into agent prompts at runtime. They are Markdown files with YAML frontmatter stored in the src/skills/ directory, organized by universal/, doc_types/, and sponsors/."),
    para("Skills are filtered by sponsor_id + document_type, matched per pipeline stage, and applied via trigger conditions (section heading regex or always-on). When multiple skills share the same topic, most specific wins (sponsor > doc_type > universal)."),
    spacer(160, 0),

    h2("Resource Limits & Timeouts"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: wLimits,
      rows: [
        headerRow(["Setting", "Default", "Purpose"], wLimits),
        dataRow(["MAX_PARALLEL_SECTIONS", "5", "Concurrent sections per batch"], wLimits, false),
        dataRow(["MAX_VALIDATION_RETRIES", "3", "Phase 4 fix attempts"], wLimits, true),
        dataRow(["MAX_AGENT_TOOL_ROUNDS", "50", "Max tool calls per agent"], wLimits, false),
        dataRow(["PROMPT_CHAR_LIMIT", "1,088,000", "~272K tokens — fits most model context windows"], wLimits, true),
        dataRow(["VISION_MAX_PAGES_PER_REQUEST", "3", "Pages per vision API call"], wLimits, false),
        dataRow(["MAX_TOC_SEARCH_PAGES", "60", "Max pages to scan for TOC"], wLimits, true),
        dataRow(["LLM_REQUEST_TIMEOUT", "900s (15 min)", "Per-request LLM timeout"], wLimits, false),
        dataRow(["LLM_REVIEWER_TIMEOUT", "300s (5 min)", "Structure reviewer timeout"], wLimits, true),
        dataRow(["MAX_LLM_CALLS", "3000", "Total LLM calls per pipeline run"], wLimits, false),
      ]
    }),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── MODEL CONFIG + OBSERVABILITY ─────────────────────────────────────────────
function buildModelAndObs() {
  const wModel = [2800, 2000, 2000, 2560];
  return [
    h1("8. Shared Infrastructure"),
    para("Common code shared between the CSR Agent and Generic Roadmap Agent is located in document_agents/shared/ and includes HTTP clients, API service clients, shared Pydantic DTOs, build utilities, and token management."),
    spacer(80, 0),
    bullet("service_helpers.py — HTTP client init, operation metadata fetch"),
    bullet("agents/ — Shared abbreviation and citation agent logic"),
    bullet("build_json_template/ — Data source utilities, post-processing, curly brace escaping"),
    bullet("dto/ — Shared Pydantic schemas for data sources, documents, operations, templates"),
    bullet("http_client/ — AsyncHTTPClient with automatic Cognito token refresh"),
    bullet("services/peerai/ — Backend API service clients (data source, document, operation, template)"),
    bullet("utils/ — API utilities, ECS log stream resolution, logging, token management"),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] }),

    h1("9. Model Configuration & LLM Integration"),
    h2("Model Summary"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: wModel,
      rows: [
        headerRow(["Agent", "Model", "Reasoning Effort", "API Format"], wModel),
        dataRow(["Data Review Agent", "openai/gpt-5.4", "low / medium", "LiteLLM Responses API"], wModel, false),
        dataRow(["CSR Agent", "openai/responses/gpt-5.4", "high", "LiteLLM Responses API (ADK)"], wModel, true),
        dataRow(["Generic Roadmap — Default", "openai/responses/gpt-5.4", "low", "LiteLLM Responses API (ADK)"], wModel, false),
        dataRow(["Generic Roadmap — Structure", "openai/responses/gpt-5.4", "high", "LiteLLM Responses API (ADK)"], wModel, true),
        dataRow(["Generic Roadmap — Reviewer", "openai/responses/gpt-5.4", "medium", "LiteLLM Responses API (ADK)"], wModel, false),
        dataRow(["Generic Roadmap — Allocator", "openai/responses/gpt-5.4", "high", "LiteLLM Responses API (ADK)"], wModel, true),
        dataRow(["Generic Roadmap — Vision", "openai/gpt-5.4", "low", "LiteLLM (standard)"], wModel, false),
      ]
    }),
    spacer(200, 0),
    h2("Why Responses API?"),
    para("GPT-5.4 requires OpenAI's Responses API (/v1/responses) when combining reasoning_effort with function tools. LiteLLM dispatches via this API when the model name contains a 'responses/' segment:"),
    codeBlock([
      "'openai/responses/gpt-5.4'  # → /v1/responses endpoint (tools + reasoning)",
      "'openai/gpt-5.4'            # → /v1/chat/completions endpoint (no tools+reasoning)"
    ]),
    spacer(160, 0),
    h2("Retry Strategy"),
    bullet("Retryable errors: 5xx, timeouts, connection errors, rate limits, bad gateway"),
    bullet("NOT retryable: 400, 401, 403, 404, 422, auth errors, model not found"),
    bullet("Strategy: Exponential backoff with jitter"),
    bullet("Max retries: 3 (configurable per agent)"),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] }),

    h1("10. Observability & Tracing"),
    para("All agents integrate with Phoenix (Arize) for distributed tracing of agent execution, LLM calls, tool calls, and error traces."),
    spacer(80, 0),
    h2("Phoenix Configuration"),
    bullet("CSR Agent: Enabled via PHOENIX_ENABLED=true, runs Phoenix server on port 6006"),
    bullet("Data Review Agent: Enabled via ENABLE_OBSERVABILITY=True, uses OBSERVABILITY_ENDPOINT"),
    bullet("Generic Roadmap Agent: Enabled via GENERIC_ROADMAP_AGENT_TRACING_ENABLED=true"),
    spacer(160, 0),
    h2("What's Traced"),
    bullet("Agent execution traces (per-phase, per-section)"),
    bullet("LLM calls (prompt, response, token usage)"),
    bullet("Tool calls (search_document, etc.)"),
    bullet("Section processing times"),
    bullet("Error traces and failure reasons"),
    spacer(160, 0),
    h2("Thinking Trace Manager"),
    para("Both the CSR Agent and Data Review Agent include a ThinkingTraceManager that sends real-time progress updates to the backend API: phase start/complete/fail events, per-section progress tracking, batch progress (started, completed, failed counts), and ECS log stream attachment for debugging."),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── DEPLOYMENT + EVAL ────────────────────────────────────────────────────────
function buildDeploymentAndEval() {
  return [
    h1("11. Deployment Architecture"),
    h2("Docker Setup"),
    codeBlock([
      "Container Structure:",
      "  /opt/venvs/csr_agent/            (CSR virtualenv)",
      "  /opt/venvs/generic_roadmap_agent/ (Roadmap virtualenv)",
      "  /app/gateway/server.py           (Entry point)",
      "  /app/csr_agent/                  (CSR source)",
      "  /app/generic_roadmap_agent/      (Roadmap source)",
      "  /app/shared/                     (Shared code)",
      "  /app/.env                        (SSM parameters)",
      "",
      "Ports:",
      "  8000 — Document Agents",
      "  6006 — Phoenix (observability)"
    ]),
    spacer(160, 0),

    h2("Production Flow"),
    bullet("Step 1: SQS message arrives with operation_id"),
    bullet("Step 2: ECS task starts Docker container"),
    bullet("Step 3: start.sh fetches SSM parameters and writes /app/.env"),
    bullet("Step 4: gateway/server.py loads env, validates agents, resolves document type from API, activates correct agent venv, dispatches to agent's main_async()"),
    bullet("Step 5: Agent downloads PDFs from S3, runs pipeline (LLM calls), PATCHes results to API, updates operation status"),
    bullet("Step 6: Container exits"),
    spacer(160, 0),

    h2("Key Environment Variables"),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [3500, 5860],
      rows: [
        headerRow(["Variable", "Purpose"], [3500, 5860]),
        dataRow(["SQS_MESSAGE", "JSON with operation_id (trigger)"], [3500, 5860], false),
        dataRow(["OPENAI_API_KEY", "OpenAI API key (local development)"], [3500, 5860], true),
        dataRow(["OPENAI_SECRET_KEY", "AWS Secrets Manager key name (production)"], [3500, 5860], false),
        dataRow(["S3_BUCKET", "Source document bucket"], [3500, 5860], true),
        dataRow(["API_SERVICE_URL", "Backend API endpoint"], [3500, 5860], false),
        dataRow(["COGNITO_TOKEN_URL", "OAuth token endpoint"], [3500, 5860], true),
        dataRow(["COGNITO_APP_CLIENT_ID", "OAuth client ID"], [3500, 5860], false),
        dataRow(["LOCAL", "'True' for local development (MinIO, local Cognito)"], [3500, 5860], true),
      ]
    }),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] }),

    h1("12. Evaluation & CI/CD"),
    h2("Data Review Agent Evaluations"),
    para("Ground truth is stored in Phoenix under dataset folders. Each study requires datasource_details.jsonl and data_review_result.jsonl. CI runs src.tests.run_ci_cd_eval and compares agent output against Phoenix ground truth. Judge prompts evaluate accuracy of TOC extraction, file classification, and data source creation."),
    spacer(160, 0),

    h2("CSR Agent Evaluations"),
    para("The EVAL_SQS environment variable provides S3 keys and folder IDs. Setting CSR_AGENT_EVAL_PIPELINE=True enables eval mode, which downloads golden eval data from S3 for comparison. Template structure, section rules, and prompt quality are evaluated."),
    spacer(160, 0),

    h2("Generic Roadmap Agent"),
    para("Uses the same Phoenix evaluation framework. Validates template schema compliance and compares section structure against ground truth."),
    spacer(80, 0),
    new Paragraph({ children: [new PageBreak()] })
  ];
}

// ─── SUMMARY ──────────────────────────────────────────────────────────────────
function buildSummary() {
  return [
    h1("Summary"),
    para("This repository implements a sophisticated multi-agent system for clinical document automation, built on Google ADK with GPT-5.4 via LiteLLM."),
    spacer(120, 0),
    new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [2400, 6960],
      rows: [
        headerRow(["Agent", "Role"], [2400, 6960]),
        dataRow(["Data Review Agent", "The 'intake' system that catalogs and classifies uploaded clinical documents"], [2400, 6960], false),
        dataRow(["CSR Agent", "A 10-step sequential pipeline with parallel optimization for Clinical Study Report authoring"], [2400, 6960], true),
        dataRow(["Generic Roadmap Agent", "A 4-phase, N-tier parallel pipeline for generating roadmaps for any clinical document type"], [2400, 6960], false),
        dataRow(["Gateway Router", "Intelligent dispatcher that routes SQS messages to the correct agent based on document type"], [2400, 6960], true),
      ]
    }),
    spacer(200, 0),
    para("All agents share common infrastructure for API communication (AsyncHTTPClient with Cognito auth), observability (Phoenix/Arize tracing), error handling (exponential backoff retries), and template schema validation (Pydantic v2). The system processes clinical study documents from S3 and produces structured JSON templates that a downstream platform renders into final regulatory documents.")
  ];
}

// ─── ASSEMBLE DOCUMENT ────────────────────────────────────────────────────────
const allContent = [
  ...buildTitlePage(),
  ...buildOverview(),
  ...buildTechStack(),
  ...buildADK(),
  ...buildGateway(),
  ...buildDataReview(),
  ...buildCSR(),
  ...buildGeneric(),
  ...buildModelAndObs(),
  ...buildDeploymentAndEval(),
  ...buildSummary(),
];

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22, color: C.text } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Arial", size: 36, bold: true, color: C.navy },
        paragraph: { spacing: { before: 400, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Arial", size: 28, bold: true, color: C.blue },
        paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Arial", size: 24, bold: true, color: C.teal },
        paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 2 } },
      { id: "Heading4", name: "Heading 4", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Arial", size: 22, bold: true, color: C.navy },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 3 } },
    ]
  },
  numbering: {
    config: [
      { reference: "bullets", levels: [
        { level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
          style: { run: { font: "Arial", size: 22, color: C.blue }, paragraph: { indent: { left: 720, hanging: 360 }, spacing: { before: 40, after: 40 } } } },
        { level: 1, format: LevelFormat.BULLET, text: "\u25E6", alignment: AlignmentType.LEFT,
          style: { run: { font: "Arial", size: 22, color: C.muted }, paragraph: { indent: { left: 1080, hanging: 360 }, spacing: { before: 40, after: 40 } } } },
      ]},
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [
            new TextRun({ text: "Document Agents — Architecture & Implementation Guide", font: "Arial", size: 18, color: C.muted }),
          ],
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.midBlue, space: 1 } },
          spacing: { before: 0, after: 100 }
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          children: [
            new TextRun({ text: "Peer AI  |  Confidential", font: "Arial", size: 18, color: C.muted }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: C.muted, tab: true }),
          ],
          tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.midBlue, space: 1 } },
          spacing: { before: 100, after: 0 }
        })]
      })
    },
    children: allContent,
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('document-agents-guide.docx', buffer);
  console.log('Done!');
}).catch(err => {
  console.error(err);
  process.exit(1);
});