import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "D:/Engg/diss/distributed-search-engine/ppt_build/output";
const FINAL = "D:/Engg/diss/distributed-search-engine/Distributed_Search_Engine_FA1_Technical_Presentation.pptx";
const W = 1280;
const H = 720;

const c = {
  navy: "#102A43",
  ink: "#17212B",
  muted: "#5D6B78",
  pale: "#F5F8FB",
  line: "#D8E0E8",
  teal: "#0E7C7B",
  blue: "#3267B1",
  violet: "#6C5CE7",
  amber: "#F2A93B",
  red: "#C2413A",
  green: "#2E7D32",
  white: "#FFFFFF",
};

const sources =
  "[Sources]\n" +
  "D:/Engg/diss/distributed_search_engine_project_context.md\n" +
  "D:/Engg/diss/distributed-search-engine/README.md\n" +
  "D:/Engg/diss/distributed-search-engine/app.py\n" +
  "D:/Engg/diss/distributed-search-engine/crawler/crawler.py\n" +
  "D:/Engg/diss/distributed-search-engine/parser/html_parser.py\n" +
  "D:/Engg/diss/distributed-search-engine/indexing/partitioner.py\n" +
  "D:/Engg/diss/distributed-search-engine/indexing/index_builder.py\n" +
  "Verified local command: python app.py stats";

async function writeBlob(path, blob) {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

function text(slide, value, pos, style = {}) {
  const box = slide.shapes.add({
    geometry: "textbox",
    position: pos,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  box.text = value;
  box.text.style = {
    fontSize: style.fontSize ?? 20,
    bold: style.bold ?? false,
    color: style.color ?? c.ink,
    alignment: style.alignment ?? "left",
  };
  return box;
}

function footer(slide, n) {
  text(slide, "Distributed Search Engine | FA-1", { left: 76, top: 684, width: 360, height: 22 }, { fontSize: 13, color: c.muted });
  text(slide, String(n).padStart(2, "0"), { left: 1156, top: 684, width: 48, height: 22 }, { fontSize: 14, bold: true, color: c.muted, alignment: "right" });
}

function title(slide, heading, subheading, n) {
  text(slide, heading, { left: 76, top: 46, width: 1060, height: 56 }, { fontSize: 38, bold: true, color: c.navy });
  if (subheading) text(slide, subheading, { left: 76, top: 104, width: 1000, height: 34 }, { fontSize: 18, color: c.muted });
  slide.shapes.add({ geometry: "rect", position: { left: 76, top: 146, width: 112, height: 5 }, fill: c.teal, line: { style: "solid", fill: "none", width: 0 } });
  footer(slide, n);
}

function surface(slide, pos, fill = c.white, line = c.line) {
  return slide.shapes.add({
    geometry: "roundRect",
    position: pos,
    fill,
    line: { style: "solid", fill: line, width: 1 },
    borderRadius: 8,
  });
}

function node(slide, label, pos, fill, fontSize = 17) {
  const s = slide.shapes.add({
    geometry: "roundRect",
    position: pos,
    fill,
    line: { style: "solid", fill: "none", width: 0 },
    borderRadius: 8,
  });
  s.text = label;
  s.text.style = { fontSize, bold: true, color: c.white, alignment: "center" };
  return s;
}

function arrow(slide, pos, fill = c.navy) {
  slide.shapes.add({
    geometry: "rightArrow",
    position: {
      ...pos,
      width: Math.max(pos.width, 34),
      height: Math.max(pos.height, 26),
    },
    fill,
    line: { style: "solid", fill: "none", width: 0 },
  });
}

function bullet(slide, value, left, top, width, color = c.teal, size = 19) {
  slide.shapes.add({ geometry: "ellipse", position: { left, top: top + 8, width: 8, height: 8 }, fill: color, line: { style: "solid", fill: "none", width: 0 } });
  text(slide, value, { left: left + 22, top, width, height: 34 }, { fontSize: size, color: c.ink });
}

function notes(slide, value) {
  slide.speakerNotes.textFrame.setText(value + "\n\n" + sources);
  slide.speakerNotes.setVisible(true);
}

function twoColumnBullets(slide, leftTitle, leftBullets, rightTitle, rightBullets, n) {
  title(slide, "Implementation is divided into two connected layers", "The final system combines offline indexing with online distributed query processing.", n);
  text(slide, leftTitle, { left: 118, top: 216, width: 440, height: 34 }, { fontSize: 25, bold: true, color: c.teal });
  leftBullets.forEach((b, i) => bullet(slide, b, 128, 286 + i * 58, 420, c.teal));
  text(slide, rightTitle, { left: 704, top: 216, width: 440, height: 34 }, { fontSize: 25, bold: true, color: c.blue });
  rightBullets.forEach((b, i) => bullet(slide, b, 714, 286 + i * 58, 420, c.blue));
}

function cover(p) {
  const slide = p.slides.add();
  slide.background.fill = c.pale;
  slide.shapes.add({ geometry: "rect", position: { left: 0, top: 0, width: 22, height: H }, fill: c.teal, line: { style: "solid", fill: "none", width: 0 } });
  text(slide, "Fault-Tolerant Distributed Search Engine", { left: 88, top: 128, width: 850, height: 128 }, { fontSize: 52, bold: true, color: c.navy });
  text(slide, "Partitioned indexing and parallel query processing", { left: 92, top: 282, width: 780, height: 40 }, { fontSize: 24, color: c.muted });
  text(slide, "FA-1 Technical Presentation | Distributed Systems", { left: 92, top: 352, width: 650, height: 30 }, { fontSize: 20, bold: true, color: c.teal });
  surface(slide, { left: 850, top: 186, width: 286, height: 210 }, c.white);
  text(slide, "Scope", { left: 894, top: 224, width: 200, height: 30 }, { fontSize: 25, bold: true, color: c.navy, alignment: "center" });
  bullet(slide, "offline indexing layer", 900, 286, 210, c.teal, 18);
  bullet(slide, "online query layer", 900, 338, 210, c.blue, 18);
  footer(slide, 1);
  notes(slide, "Open by stating that the presentation covers the complete distributed search engine, from offline data preparation to online query processing.");
}

function problem(p) {
  const slide = p.slides.add();
  slide.background.fill = c.white;
  title(slide, "Why search is a distributed-systems problem", "A centralized search system limits scalability and availability.", 2);
  text(slide, "Centralized search", { left: 122, top: 210, width: 300, height: 32 }, { fontSize: 25, bold: true, color: c.navy });
  node(slide, "all documents", { left: 118, top: 292, width: 150, height: 54 }, c.muted, 17);
  arrow(slide, { left: 292, top: 305, width: 56, height: 28 }, c.red);
  node(slide, "single index\nsingle query processor", { left: 378, top: 276, width: 210, height: 86 }, c.navy, 18);
  node(slide, "bottleneck / single failure point", { left: 180, top: 456, width: 300, height: 56 }, c.red, 17);
  text(slide, "one machine handles indexing and queries", { left: 156, top: 536, width: 360, height: 34 }, { fontSize: 18, color: c.muted, alignment: "center" });
  text(slide, "Distributed search", { left: 742, top: 210, width: 300, height: 32 }, { fontSize: 25, bold: true, color: c.navy });
  node(slide, "Coordinator", { left: 710, top: 338, width: 170, height: 58 }, c.navy);
  text(slide, "fan-out query", { left: 912, top: 312, width: 120, height: 26 }, { fontSize: 16, bold: true, color: c.navy, alignment: "center" });
  arrow(slide, { left: 914, top: 352, width: 66, height: 30 }, c.navy);
  node(slide, "Node 1", { left: 1024, top: 276, width: 138, height: 50 }, c.teal);
  node(slide, "Node 2", { left: 1024, top: 354, width: 138, height: 50 }, c.blue);
  node(slide, "Node 3", { left: 1024, top: 432, width: 138, height: 50 }, c.violet);
  text(slide, "parallel local search", { left: 998, top: 508, width: 190, height: 26 }, { fontSize: 18, bold: true, color: c.teal, alignment: "center" });
  text(slide, "Data is partitioned and query work is sent to multiple nodes in parallel.", { left: 720, top: 552, width: 500, height: 50 }, { fontSize: 20, color: c.ink, alignment: "center" });
  notes(slide, "Explain the project motivation using distributed-systems terms: scalability, parallel processing, and fault tolerance.");
}

function architecture(p) {
  const slide = p.slides.add();
  slide.background.fill = c.pale;
  title(slide, "Architecture has an offline layer and an online layer", "The offline layer prepares shard/index files; the online layer serves distributed queries using those files.", 3);
  const steps = [
    ["Crawl", c.teal],
    ["Parse", c.blue],
    ["Shard", c.violet],
    ["Index", c.amber],
    ["Search nodes", c.blue],
    ["Coordinator", c.navy],
    ["Web UI", c.green],
  ];
  steps.forEach((s, i) => {
    const left = 74 + i * 166;
    node(slide, s[0], { left, top: 300, width: 122, height: 58 }, s[1], 16);
    if (i < steps.length - 1) arrow(slide, { left: left + 124, top: 316, width: 40, height: 26 }, i < 3 ? c.teal : c.blue);
  });
  text(slide, "Offline layer", { left: 160, top: 228, width: 300, height: 32 }, { fontSize: 25, bold: true, color: c.teal, alignment: "center" });
  text(slide, "Online layer", { left: 818, top: 228, width: 300, height: 32 }, { fontSize: 25, bold: true, color: c.blue, alignment: "center" });
  surface(slide, { left: 96, top: 440, width: 500, height: 78 }, "#E8F5F4", "#B8DEDC");
  text(slide, "Offline output: raw pages, documents.jsonl, shard files, index files, global statistics", { left: 128, top: 462, width: 436, height: 34 }, { fontSize: 19, color: c.ink, alignment: "center" });
  surface(slide, { left: 690, top: 440, width: 500, height: 78 }, "#EEF3FA", "#BDD0EA");
  text(slide, "Runtime output: user query, parallel node responses, merged ranked result list", { left: 722, top: 462, width: 436, height: 34 }, { fontSize: 19, color: c.ink, alignment: "center" });
  notes(slide, "Use this as the linked-system overview. The deck should not look like separate files; it should show a continuous architecture.");
}

function split(p) {
  const slide = p.slides.add();
  slide.background.fill = c.white;
  twoColumnBullets(
    slide,
    "Offline data/indexing layer",
    ["controlled crawling from allowed domains", "HTML parsing into document records", "document partitioning into shards", "local inverted indexes with global stats"],
    "Online query/runtime layer",
    ["gRPC search-node services", "coordinator fan-out and timeout", "score-based result aggregation", "web/API result display with node status"],
    4,
  );
  surface(slide, { left: 312, top: 548, width: 660, height: 54 }, "#FFF8EA", "#F4D49A");
  text(slide, "Common contract: node_id, shard path, index path, QueryRequest, SearchResult", { left: 346, top: 564, width: 592, height: 24 }, { fontSize: 19, bold: true, color: c.ink, alignment: "center" });
  notes(slide, "Explain this as a system decomposition rather than a person-wise split. The important point is the contract between offline artifacts and runtime services.");
}

function offlinePipeline(p) {
  const slide = p.slides.add();
  slide.background.fill = c.pale;
  title(slide, "Offline pipeline builds the data and index foundation", "The FA-1 implementation prepares the files loaded by search nodes.", 5);
  const rows = [
    ["Crawler", "Visits allowed domains, respects robots/limits, saves raw HTML offline"],
    ["Parser", "Extracts title, URL, source domain, visible text, and token count"],
    ["Partitioner", "Splits documents into node-owned shard files using range partitioning"],
    ["Index builder", "Creates token postings, document metadata, and global IDF statistics"],
  ];
  rows.forEach((r, i) => {
    const top = 210 + i * 86;
    text(slide, r[0], { left: 128, top, width: 210, height: 30 }, { fontSize: 24, bold: true, color: [c.teal, c.blue, c.violet, c.amber][i] });
    text(slide, r[1], { left: 372, top: top + 2, width: 690, height: 32 }, { fontSize: 20, color: c.ink });
    slide.shapes.add({ geometry: "line", position: { left: 128, top: top + 52, width: 900, height: 0 }, line: { style: "solid", fill: c.line, width: 1 }, fill: "none" });
  });
  notes(slide, "Describe this as the offline indexing layer. Mention that generated artifacts are stored and can be inspected before the runtime starts.");
}

function partitionedIndex(p) {
  const slide = p.slides.add();
  slide.background.fill = c.white;
  title(slide, "Partitioned indexing makes each node independent", "Each search node can load only its shard and its local index.", 6);
  node(slide, "documents", { left: 150, top: 286, width: 150, height: 54 }, c.navy);
  arrow(slide, { left: 322, top: 297, width: 64, height: 34 }, c.navy);
  node(slide, "shard 1", { left: 402, top: 234, width: 130, height: 48 }, c.teal);
  node(slide, "shard 2", { left: 402, top: 310, width: 130, height: 48 }, c.blue);
  node(slide, "shard 3", { left: 402, top: 386, width: 130, height: 48 }, c.violet);
  text(slide, "Range partitioning", { left: 142, top: 456, width: 390, height: 30 }, { fontSize: 21, bold: true, color: c.navy, alignment: "center" });
  surface(slide, { left: 710, top: 232, width: 400, height: 270 }, "#F8FAFC", c.line);
  text(slide, "Local inverted index", { left: 778, top: 272, width: 270, height: 30 }, { fontSize: 24, bold: true, color: c.teal, alignment: "center" });
  text(slide, "term -> [(doc_id, tf, positions)]\n\nquery terms are matched locally\n\nscores use global_stats.json", { left: 756, top: 328, width: 310, height: 136 }, { fontSize: 19, color: c.ink, alignment: "center" });
  notes(slide, "Explain why independent shards are important for search-node processes. Each node can load one partition instead of the full corpus.");
}

function codeEvidence(p) {
  const slide = p.slides.add();
  slide.background.fill = c.pale;
  title(slide, "Core implementation follows a staged pipeline", "The code keeps crawling, parsing, sharding, and indexing as separate steps.", 7);
  surface(slide, { left: 92, top: 206, width: 512, height: 384 }, "#F8FAFC", c.line);
  text(slide, "Pipeline entry point", { left: 126, top: 236, width: 300, height: 30 }, { fontSize: 24, bold: true, color: c.navy });
  text(slide,
    "def run_pipeline(args):\n" +
    "    crawl(config)\n" +
    "    parse_html(raw_dir, documents.jsonl)\n" +
    "    partition_documents(documents, shards, n)\n" +
    "    build_indexes(shards, indexes)",
    { left: 126, top: 302, width: 420, height: 158 },
    { fontSize: 20, color: c.ink },
  );
  text(slide, "Why this matters", { left: 126, top: 500, width: 240, height: 28 }, { fontSize: 22, bold: true, color: c.teal });
  text(slide, "Each stage has a clear input and output, so failures are easier to locate and individual stages can be rerun.", { left: 126, top: 534, width: 420, height: 46 }, { fontSize: 18, color: c.ink });

  surface(slide, { left: 690, top: 206, width: 438, height: 384 }, c.white, c.line);
  text(slide, "Index structure", { left: 728, top: 236, width: 260, height: 30 }, { fontSize: 24, bold: true, color: c.navy });
  text(slide,
    "index = {\n" +
    "  \"node_id\": \"node_1\",\n" +
    "  \"documents\": {...},\n" +
    "  \"terms\": {\n" +
    "    \"query\": [\n" +
    "      {\"doc_id\": \"...\", \"tf\": 2}\n" +
    "    ]\n" +
    "  }\n" +
    "}",
    { left: 728, top: 298, width: 330, height: 206 },
    { fontSize: 18, color: c.ink },
  );
  text(slide, "This is the local state loaded by every search-node process.", { left: 728, top: 536, width: 340, height: 28 }, { fontSize: 18, color: c.muted, alignment: "center" });
  notes(slide, "Show a small code-level view without overwhelming the presentation. The snippet is simplified from app.py and the generated index structure.");
}

function outputEvidence(p) {
  const slide = p.slides.add();
  slide.background.fill = c.white;
  title(slide, "Current output proves the offline layer is working", "These commands can be shown directly during evaluation.", 8);
  surface(slide, { left: 90, top: 204, width: 520, height: 372 }, "#102A43", "#102A43");
  text(slide, "python app.py stats", { left: 124, top: 236, width: 300, height: 28 }, { fontSize: 22, bold: true, color: c.white });
  text(slide,
    "Raw HTML pages: 1000\n" +
    "Metadata records: 1000\n" +
    "Parsed documents: 1000\n" +
    "Global indexed documents: 1000\n" +
    "node_1: 334 documents, 35803 terms\n" +
    "node_2: 334 documents, 13726 terms\n" +
    "node_3: 332 documents, 21854 terms",
    { left: 124, top: 292, width: 428, height: 210 },
    { fontSize: 18, color: c.white },
  );
  surface(slide, { left: 680, top: 204, width: 520, height: 372 }, "#F8FAFC", c.line);
  text(slide, "python app.py search \"parallel query processing\"", { left: 714, top: 236, width: 430, height: 28 }, { fontSize: 21, bold: true, color: c.navy });
  text(slide,
    "1. The Kubernetes API | Kubernetes\n" +
    "   Score: 0.9916 | Node: node_1\n\n" +
    "2. PostgreSQL 18.6, 17.11, 16.15...\n" +
    "   Score: 0.9632 | Node: node_2\n\n" +
    "3. Labels and Selectors | Kubernetes\n" +
    "   Score: 0.9484 | Node: node_1",
    { left: 714, top: 294, width: 410, height: 188 },
    { fontSize: 18, color: c.ink },
  );
  text(slide, "The result list shows which partition produced each match.", { left: 714, top: 518, width: 400, height: 36 }, { fontSize: 17, color: c.muted, alignment: "center" });
  notes(slide, "Use this as concrete project evidence. The stats and search output were captured from the current local implementation.");
}

function queryRuntime(p) {
  const slide = p.slides.add();
  slide.background.fill = c.pale;
  title(slide, "Distributed query runtime serves the prepared indexes", "This layer turns local shard indexes into one searchable service.", 9);
  const rows = [
    ["Search node service", "Loads one shard and local index; tokenizes query; returns local top-K results"],
    ["Coordinator service", "Receives query; calls all nodes in parallel; handles timeout; merges results"],
    ["API / Web UI", "Accepts user input; displays ranked results, source URLs, scores, and node status"],
    ["Config layer", "Stores node IDs, host addresses, ports, shard assignments, and timeouts"],
  ];
  rows.forEach((r, i) => {
    const top = 202 + i * 90;
    surface(slide, { left: 114, top, width: 1050, height: 64 }, i % 2 === 0 ? c.white : "#F8FAFC", c.line);
    text(slide, r[0], { left: 150, top: top + 16, width: 270, height: 28 }, { fontSize: 22, bold: true, color: [c.blue, c.navy, c.green, c.violet][i] });
    text(slide, r[1], { left: 464, top: top + 18, width: 650, height: 26 }, { fontSize: 19, color: c.ink });
  });
  notes(slide, "This slide explains the online runtime components: search-node services, coordinator, interface, and configuration.");
}

function grpcContract(p) {
  const slide = p.slides.add();
  slide.background.fill = c.white;
  title(slide, "gRPC defines the communication contract", "The coordinator and search nodes communicate through request and response messages.", 10);
  surface(slide, { left: 112, top: 210, width: 456, height: 350 }, "#F8FAFC", c.line);
  text(slide, "QueryRequest", { left: 152, top: 246, width: 260, height: 30 }, { fontSize: 24, bold: true, color: c.blue });
  text(slide, "query_text\n\ntop_k\n\nrequest_id\n\ntimeout_ms", { left: 170, top: 306, width: 260, height: 160 }, { fontSize: 21, color: c.ink });
  surface(slide, { left: 712, top: 210, width: 456, height: 350 }, "#F8FAFC", c.line);
  text(slide, "SearchResponse", { left: 752, top: 246, width: 280, height: 30 }, { fontSize: 24, bold: true, color: c.teal });
  text(slide, "node_id\n\nstatus\n\nresults[]\n\nlatency_ms", { left: 770, top: 306, width: 260, height: 160 }, { fontSize: 21, color: c.ink });
  arrow(slide, { left: 598, top: 368, width: 78, height: 36 }, c.muted);
  notes(slide, "Map this to Unit 2 RPC and message-oriented communication: the proto schema becomes the formal interface between processes.");
}

function coordinator(p) {
  const slide = p.slides.add();
  slide.background.fill = c.pale;
  title(slide, "Coordinator controls fan-out, timeout, and merge", "The coordinator is the central runtime component in the online layer.", 11);
  const steps = [
    ["1", "receive query and top-K limit"],
    ["2", "send QueryRequest to all nodes"],
    ["3", "wait with per-node timeout"],
    ["4", "merge local top-K results"],
    ["5", "return ranked response with node status"],
  ];
  steps.forEach((s, i) => {
    const top = 194 + i * 78;
    node(slide, s[0], { left: 156, top, width: 48, height: 48 }, i < 4 ? c.blue : c.amber, 18);
    text(slide, s[1], { left: 240, top: top + 8, width: 470, height: 28 }, { fontSize: 22, color: c.ink });
  });
  surface(slide, { left: 772, top: 246, width: 330, height: 244 }, c.white, c.line);
  text(slide, "Merge rule", { left: 832, top: 284, width: 210, height: 28 }, { fontSize: 24, bold: true, color: c.navy, alignment: "center" });
  text(slide, "combine node results\nsort by score\nremove duplicate URLs\napply top-K cut-off\nshow node status", { left: 814, top: 332, width: 250, height: 132 }, { fontSize: 19, color: c.ink, alignment: "center" });
  notes(slide, "Explain the coordinator algorithm in viva terms. This is the part where parallel query processing is managed.");
}

function searchNode(p) {
  const slide = p.slides.add();
  slide.background.fill = c.white;
  title(slide, "Each search node performs local retrieval", "Nodes do not need the full corpus; they search only their own partition.", 12);
  text(slide, "Node startup", { left: 128, top: 206, width: 260, height: 30 }, { fontSize: 24, bold: true, color: c.blue });
  bullet(slide, "load assigned shard file", 134, 276, 390, c.blue);
  bullet(slide, "load local inverted index", 134, 334, 390, c.blue);
  bullet(slide, "register node_id and address", 134, 392, 390, c.blue);
  text(slide, "Query handling", { left: 704, top: 206, width: 260, height: 30 }, { fontSize: 24, bold: true, color: c.teal });
  bullet(slide, "tokenize incoming query", 710, 276, 390, c.teal);
  bullet(slide, "lookup postings in local index", 710, 334, 390, c.teal);
  bullet(slide, "return local top-K results", 710, 392, 390, c.teal);
  surface(slide, { left: 278, top: 520, width: 724, height: 82 }, "#FFF8EA", "#F4D49A");
  text(slide, "This design supports scalability because adding a node means assigning another partition.", { left: 318, top: 542, width: 644, height: 44 }, { fontSize: 18, color: c.ink, alignment: "center" });
  notes(slide, "Discuss search-node responsibility and scalability. The node has local state and a small RPC interface.");
}

function uiAndFaults(p) {
  const slide = p.slides.add();
  slide.background.fill = c.pale;
  title(slide, "UI and fault handling make the system demonstrable", "The faculty demo should show normal results and unavailable-node behavior.", 13);
  text(slide, "Web/API display", { left: 126, top: 214, width: 320, height: 30 }, { fontSize: 24, bold: true, color: c.green });
  bullet(slide, "query input box", 134, 284, 360, c.green);
  bullet(slide, "ranked result list", 134, 342, 360, c.green);
  bullet(slide, "source URL and score", 134, 400, 360, c.green);
  bullet(slide, "node status/latency", 134, 458, 360, c.green);
  text(slide, "Failure behavior", { left: 704, top: 214, width: 320, height: 30 }, { fontSize: 24, bold: true, color: c.red });
  bullet(slide, "set deadline per node call", 712, 284, 390, c.red);
  bullet(slide, "skip failed node response", 712, 342, 390, c.red);
  bullet(slide, "return partial results", 712, 400, 390, c.red);
  bullet(slide, "show which node failed", 712, 458, 390, c.red);
  notes(slide, "This slide explains the user-facing and fault-tolerant pieces of the online layer. It links to Unit 2 fault-tolerance self-study.");
}

function syllabus(p) {
  const slide = p.slides.add();
  slide.background.fill = c.white;
  title(slide, "FA-1 concepts are visible across the system", "The project maps to Unit 1 and Unit 2 through concrete implementation choices.", 14);
  text(slide, "Unit 1", { left: 132, top: 216, width: 190, height: 30 }, { fontSize: 25, bold: true, color: c.teal });
  bullet(slide, "architecture: coordinator + search nodes", 140, 286, 470, c.teal);
  bullet(slide, "design issue: partitioning + ranking", 140, 344, 470, c.teal);
  bullet(slide, "scalability through sharding", 140, 402, 470, c.teal);
  bullet(slide, "middleware boundary between services", 140, 460, 470, c.teal);
  text(slide, "Unit 2", { left: 710, top: 216, width: 190, height: 30 }, { fontSize: 25, bold: true, color: c.blue });
  bullet(slide, "RPC using gRPC service calls", 718, 286, 430, c.blue);
  bullet(slide, "request/response message schema", 718, 344, 430, c.blue);
  bullet(slide, "node identifiers and addresses", 718, 402, 430, c.blue);
  bullet(slide, "timeouts for partial results", 718, 460, 430, c.blue);
  notes(slide, "Use this as the direct mapping to the FA-1 syllabus expected by faculty.");
}

function demo(p) {
  const slide = p.slides.add();
  slide.background.fill = c.pale;
  title(slide, "Evaluation demo sequence", "The demo should prove that the offline and online layers connect into one system.", 15);
  const rows = [
    ["1", "Run offline pipeline", "generate documents, shards, indexes"],
    ["2", "Start search-node services", "each node loads one shard and index"],
    ["3", "Start coordinator", "register node addresses and send fan-out query"],
    ["4", "Search from UI/API", "display merged ranked results"],
    ["5", "Stop one node", "show timeout and partial result behavior"],
  ];
  rows.forEach((r, i) => {
    const top = 194 + i * 78;
    node(slide, r[0], { left: 122, top, width: 48, height: 48 }, i < 4 ? c.teal : c.red, 18);
    text(slide, r[1], { left: 212, top: top + 2, width: 330, height: 28 }, { fontSize: 22, bold: true, color: c.navy });
    text(slide, r[2], { left: 570, top: top + 6, width: 520, height: 28 }, { fontSize: 19, color: c.muted });
  });
  notes(slide, "Close with the exact demonstration plan. It covers offline artifacts and online runtime behavior.");
}

async function main() {
  await fs.mkdir(OUT, { recursive: true });
  const p = Presentation.create({ slideSize: { width: W, height: H } });
  cover(p);
  problem(p);
  architecture(p);
  split(p);
  offlinePipeline(p);
  partitionedIndex(p);
  codeEvidence(p);
  outputEvidence(p);
  queryRuntime(p);
  grpcContract(p);
  coordinator(p);
  searchNode(p);
  uiAndFaults(p);
  syllabus(p);
  demo(p);

  const inspect = await p.inspect({ kind: "slide,textbox,shape,notes", maxChars: 20000 });
  await fs.writeFile(`${OUT}/inspect.ndjson`, inspect.ndjson);
  for (const [i, slide] of p.slides.items.entries()) {
    const stem = `slide-${String(i + 1).padStart(2, "0")}`;
    await writeBlob(`${OUT}/${stem}.png`, await p.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${OUT}/${stem}.layout.json`, await layout.text());
  }
  await writeBlob(`${OUT}/deck-montage.webp`, await p.export({ format: "webp", montage: true, scale: 1 }));
  const pptx = await PresentationFile.exportPptx(p);
  await pptx.save(FINAL);
  console.log(FINAL);
}

main().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});
