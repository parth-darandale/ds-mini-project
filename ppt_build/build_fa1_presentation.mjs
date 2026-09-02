import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "D:/Engg/diss/distributed-search-engine/ppt_build/output";
const FINAL = "D:/Engg/diss/distributed-search-engine/Distributed_Search_Engine_FA1_Presentation.pptx";

const W = 1280;
const H = 720;
const page = { left: 76, top: 54, width: 1128, height: 604 };

const colors = {
  navy: "#102A43",
  ink: "#17212B",
  muted: "#596A7A",
  pale: "#F5F8FB",
  line: "#D7E1EA",
  teal: "#0E7C7B",
  blue: "#3267B1",
  violet: "#6C5CE7",
  amber: "#F2A93B",
  red: "#C2413A",
  green: "#2E7D32",
  white: "#FFFFFF",
};

const sourceBlock =
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

function addText(slide, text, position, style = {}) {
  const box = slide.shapes.add({
    geometry: "textbox",
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  box.text = text;
  box.text.style = {
    fontSize: style.fontSize ?? 20,
    bold: style.bold ?? false,
    color: style.color ?? colors.ink,
    alignment: style.alignment ?? "left",
  };
  return box;
}

function addFooter(slide, number) {
  addText(slide, "Distributed Search Engine | FA-1", { left: 76, top: 684, width: 360, height: 22 }, { fontSize: 13, color: colors.muted });
  addText(slide, String(number).padStart(2, "0"), { left: 1156, top: 684, width: 48, height: 22 }, { fontSize: 14, bold: true, color: colors.muted, alignment: "right" });
}

function addTitle(slide, title, subtitle, number) {
  addText(slide, title, { left: page.left, top: 46, width: 1040, height: 56 }, { fontSize: 38, bold: true, color: colors.navy });
  if (subtitle) addText(slide, subtitle, { left: page.left, top: 104, width: 980, height: 34 }, { fontSize: 18, color: colors.muted });
  slide.shapes.add({
    geometry: "rect",
    position: { left: page.left, top: 145, width: 112, height: 5 },
    fill: colors.teal,
    line: { style: "solid", fill: "none", width: 0 },
  });
  addFooter(slide, number);
}

function addSurface(slide, position, fill = colors.white, line = colors.line) {
  return slide.shapes.add({
    geometry: "roundRect",
    position,
    fill,
    line: { style: "solid", fill: line, width: 1 },
    borderRadius: 10,
  });
}

function addNode(slide, text, position, fill, textColor = colors.white, fontSize = 18) {
  const shape = slide.shapes.add({
    geometry: "roundRect",
    position,
    fill,
    line: { style: "solid", fill: "none", width: 0 },
    borderRadius: 10,
  });
  shape.text = text;
  shape.text.style = { fontSize, bold: true, color: textColor, alignment: "center" };
  return shape;
}

function addArrow(slide, position, fill = colors.muted) {
  slide.shapes.add({
    geometry: "rightArrow",
    position,
    fill,
    line: { style: "solid", fill: "none", width: 0 },
  });
}

function addLine(slide, x1, y1, x2, y2, fill = colors.muted, width = 3) {
  slide.shapes.add({
    geometry: "line",
    position: {
      left: Math.min(x1, x2),
      top: Math.min(y1, y2),
      width: Math.abs(x2 - x1),
      height: Math.abs(y2 - y1),
    },
    line: { style: "solid", fill, width },
    fill: "none",
  });
}

function connect(slide, from, to, options = {}) {
  slide.shapes.connect(from, to, {
    kind: options.kind ?? "straight",
    fromSide: options.fromSide,
    toSide: options.toSide,
    line: { style: "solid", fill: options.fill ?? colors.muted, width: options.width ?? 3 },
    head: options.head ?? { type: "none" },
  });
}

function addBullet(slide, text, left, top, width, color = colors.teal) {
  slide.shapes.add({
    geometry: "ellipse",
    position: { left, top: top + 8, width: 9, height: 9 },
    fill: color,
    line: { style: "solid", fill: "none", width: 0 },
  });
  addText(slide, text, { left: left + 22, top, width, height: 42 }, { fontSize: 19, color: colors.ink });
}

function addNotes(slide, text) {
  slide.speakerNotes.textFrame.setText(text + "\n\n" + sourceBlock);
  slide.speakerNotes.setVisible(true);
}

function cover(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  slide.shapes.add({ geometry: "rect", position: { left: 0, top: 0, width: 22, height: H }, fill: colors.teal, line: { style: "solid", fill: "none", width: 0 } });
  addText(slide, "Fault-Tolerant Distributed Search Engine", { left: 86, top: 118, width: 840, height: 132 }, { fontSize: 52, bold: true, color: colors.navy });
  addText(slide, "Partitioned indexing and parallel query processing", { left: 90, top: 274, width: 760, height: 40 }, { fontSize: 24, color: colors.muted });
  addText(slide, "FA-1 Technical Presentation | Distributed Systems", { left: 90, top: 352, width: 640, height: 32 }, { fontSize: 20, bold: true, color: colors.teal });
  addSurface(slide, { left: 854, top: 164, width: 286, height: 298 }, colors.white, colors.line);
  addText(slide, "Presentation focus", { left: 894, top: 204, width: 210, height: 34 }, { fontSize: 24, bold: true, color: colors.navy, alignment: "center" });
  addBullet(slide, "system architecture", 900, 278, 190, colors.teal);
  addBullet(slide, "Member 1 implementation", 900, 330, 200, colors.blue);
  addBullet(slide, "FA-1 concepts", 900, 382, 190, colors.violet);
  addFooter(slide, 1);
  addNotes(slide, "Open as a technical student presentation. Keep the focus on design, implementation decisions, and FA-1 distributed systems concepts.");
}

function problem(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Why search is a distributed-systems problem", "Large document collections create storage, processing, and availability challenges.", 2);
  addText(slide, "Centralized approach", { left: 110, top: 208, width: 310, height: 32 }, { fontSize: 25, bold: true, color: colors.navy });
  addNode(slide, "Single server\nindex + queries", { left: 162, top: 292, width: 200, height: 98 }, colors.navy, colors.white, 20);
  addLine(slide, 262, 390, 262, 470, colors.muted, 3);
  addNode(slide, "bottleneck", { left: 164, top: 470, width: 196, height: 50 }, colors.red, colors.white, 18);
  addText(slide, "Limited CPU and storage; one failure can stop search.", { left: 112, top: 552, width: 340, height: 58 }, { fontSize: 20, color: colors.ink, alignment: "center" });

  addText(slide, "Distributed approach", { left: 724, top: 208, width: 310, height: 32 }, { fontSize: 25, bold: true, color: colors.navy });
  const coord = addNode(slide, "Coordinator", { left: 850, top: 282, width: 174, height: 58 }, colors.navy);
  const n1 = addNode(slide, "Node 1", { left: 690, top: 424, width: 132, height: 56 }, colors.teal);
  const n2 = addNode(slide, "Node 2", { left: 878, top: 424, width: 132, height: 56 }, colors.blue);
  const n3 = addNode(slide, "Node 3", { left: 1066, top: 424, width: 132, height: 56 }, colors.violet);
  connect(slide, coord, n1, { fromSide: "bottom", toSide: "top" });
  connect(slide, coord, n2, { fromSide: "bottom", toSide: "top" });
  connect(slide, coord, n3, { fromSide: "bottom", toSide: "top" });
  addText(slide, "Documents are partitioned; query work happens in parallel.", { left: 700, top: 552, width: 470, height: 58 }, { fontSize: 20, color: colors.ink, alignment: "center" });
  addNotes(slide, "Explain the shift from one machine doing all storage and query work to independent nodes coordinated through communication.");
}

function architecture(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  addTitle(slide, "Architecture separates indexing from query serving", "FA-1 establishes the offline foundation; FA-2 can extend the online distributed runtime.", 3);
  addText(slide, "Offline indexing path", { left: 92, top: 204, width: 360, height: 32 }, { fontSize: 25, bold: true, color: colors.teal });
  const offline = [["Crawler", 92, colors.teal], ["Parser", 260, colors.blue], ["Partitioner", 428, colors.violet], ["Index builder", 596, colors.amber]];
  for (let i = 0; i < offline.length; i++) {
    addNode(slide, offline[i][0], { left: offline[i][1], top: 292, width: 132, height: 64 }, offline[i][2], colors.white, 17);
    if (i < offline.length - 1) addArrow(slide, { left: offline[i][1] + 140, top: 314, width: 36, height: 22 });
  }
  addSurface(slide, { left: 104, top: 420, width: 598, height: 90 }, "#E8F5F4", "#B8DEDC");
  addText(slide, "Output: document shards + local inverted indexes + global statistics", { left: 134, top: 448, width: 540, height: 34 }, { fontSize: 21, color: colors.ink, alignment: "center" });

  addText(slide, "Online query path", { left: 810, top: 204, width: 330, height: 32 }, { fontSize: 25, bold: true, color: colors.blue });
  const client = addNode(slide, "Client", { left: 820, top: 286, width: 126, height: 52 }, colors.green);
  const coordinator = addNode(slide, "Coordinator", { left: 988, top: 286, width: 150, height: 52 }, colors.navy);
  addArrow(slide, { left: 952, top: 302, width: 32, height: 22 });
  const searchNodes = addNode(slide, "Search\nnodes", { left: 906, top: 414, width: 146, height: 68 }, colors.blue);
  connect(slide, coordinator, searchNodes, { fromSide: "bottom", toSide: "top" });
  addSurface(slide, { left: 780, top: 534, width: 384, height: 58 }, colors.white, colors.line);
  addText(slide, "gRPC calls and result aggregation are Member 2's runtime layer.", { left: 806, top: 552, width: 330, height: 26 }, { fontSize: 18, color: colors.ink, alignment: "center" });
  addNotes(slide, "Show the boundary between Member 1 and Member 2 without making the deck look like separate files. The architecture is one linked system.");
}

function memberSplit(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Team split keeps one continuous system contract", "Both members understand the whole design, but each part has a clear responsibility.", 4);
  addText(slide, "Member 1: data and indexing layer", { left: 116, top: 204, width: 430, height: 32 }, { fontSize: 25, bold: true, color: colors.teal });
  addBullet(slide, "controlled crawling and offline storage", 126, 272, 430, colors.teal);
  addBullet(slide, "HTML parsing into document records", 126, 328, 430, colors.teal);
  addBullet(slide, "partitioned shards for future nodes", 126, 384, 430, colors.teal);
  addBullet(slide, "local inverted indexes and global stats", 126, 440, 430, colors.teal);
  addText(slide, "Member 2: distributed query layer", { left: 710, top: 204, width: 430, height: 32 }, { fontSize: 25, bold: true, color: colors.blue });
  addBullet(slide, "gRPC services for search nodes", 720, 272, 430, colors.blue);
  addBullet(slide, "coordinator fan-out and merge", 720, 328, 430, colors.blue);
  addBullet(slide, "parallel query processing demo", 720, 384, 430, colors.blue);
  addBullet(slide, "timeouts and partial-result handling", 720, 440, 430, colors.blue);
  addSurface(slide, { left: 290, top: 548, width: 700, height: 54 }, "#FFF8EA", "#F4D49A");
  addText(slide, "Shared contract: node_id, shard file, local index file, result schema", { left: 326, top: 564, width: 628, height: 26 }, { fontSize: 20, bold: true, color: colors.ink, alignment: "center" });
  addNotes(slide, "Use this slide to answer how two team members divided work while still building one connected distributed search engine.");
}

function pipeline(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  addTitle(slide, "Member 1 pipeline creates the searchable corpus", "The output of each stage becomes the input for the next stage.", 5);
  const steps = [["Seed URLs", "configured allowlist", colors.navy], ["Crawler", "robots, BFS, delay", colors.teal], ["Raw HTML", "offline page files", colors.green], ["Parser", "visible text extraction", colors.blue], ["Documents", "JSONL records", colors.violet], ["Indexes", "postings + IDF", colors.amber]];
  for (let i = 0; i < steps.length; i++) {
    const left = 82 + i * 188;
    addNode(slide, steps[i][0], { left, top: 248, width: 128, height: 58 }, steps[i][2], colors.white, 16);
    addText(slide, steps[i][1], { left: left - 10, top: 324, width: 148, height: 44 }, { fontSize: 17, color: colors.ink, alignment: "center" });
    if (i < steps.length - 1) addArrow(slide, { left: left + 136, top: 266, width: 36, height: 22 });
  }
  addSurface(slide, { left: 160, top: 470, width: 960, height: 80 }, colors.white, colors.line);
  addText(slide, "Design decision: crawling, parsing, partitioning, and indexing are separate modules so each stage can be tested or replaced independently.", { left: 204, top: 494, width: 872, height: 34 }, { fontSize: 21, color: colors.ink, alignment: "center" });
  addNotes(slide, "Walk through the modules from configured seed URLs to the generated indexes. Emphasize separation of concerns.");
}

function crawlerDesign(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Crawler design is controlled, polite, and reproducible", "For FA-1, the crawler stores web pages offline before processing them.", 6);
  addSurface(slide, { left: 112, top: 210, width: 420, height: 350 }, "#F8FAFC", colors.line);
  addText(slide, "Control rules", { left: 150, top: 244, width: 250, height: 32 }, { fontSize: 25, bold: true, color: colors.navy });
  addBullet(slide, "only allowed domains are visited", 154, 308, 310, colors.teal);
  addBullet(slide, "robots.txt is checked", 154, 364, 310, colors.teal);
  addBullet(slide, "request delay prevents aggressive crawling", 154, 420, 330, colors.teal);
  addBullet(slide, "HTML is saved before parsing", 154, 476, 330, colors.teal);
  addText(slide, "Crawler sequence", { left: 716, top: 210, width: 280, height: 32 }, { fontSize: 25, bold: true, color: colors.navy });
  const y = [284, 362, 440, 518];
  const labels = ["queue URL", "fetch page", "extract links", "save metadata"];
  for (let i = 0; i < labels.length; i++) {
    addNode(slide, labels[i], { left: 704, top: y[i], width: 210, height: 48 }, i % 2 === 0 ? colors.blue : colors.teal, colors.white, 17);
    if (i < labels.length - 1) addLine(slide, 809, y[i] + 48, 809, y[i + 1], colors.muted, 3);
  }
  addNotes(slide, "Explain that the crawler is intentionally limited for an academic demo: it is controlled by configuration and produces repeatable offline input.");
}

function documentModel(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  addTitle(slide, "Parser converts HTML pages into logical documents", "Indexing should work on clean text and stable metadata, not raw HTML.", 7);
  addText(slide, "Raw HTML", { left: 108, top: 212, width: 220, height: 32 }, { fontSize: 25, bold: true, color: colors.navy });
  addSurface(slide, { left: 110, top: 268, width: 326, height: 260 }, colors.white, colors.line);
  addText(slide, "<html>\n  <title>...</title>\n  <nav>...</nav>\n  <main>\n    Visible content\n  </main>\n</html>", { left: 146, top: 306, width: 250, height: 178 }, { fontSize: 20, color: colors.ink });
  addArrow(slide, { left: 484, top: 380, width: 70, height: 36 }, colors.muted);
  addText(slide, "Document record", { left: 646, top: 212, width: 300, height: 32 }, { fontSize: 25, bold: true, color: colors.navy });
  addSurface(slide, { left: 648, top: 268, width: 452, height: 260 }, colors.white, colors.line);
  addText(slide, "{\n  doc_id,\n  url,\n  source_domain,\n  title,\n  text,\n  token_count\n}", { left: 700, top: 302, width: 300, height: 190 }, { fontSize: 22, color: colors.ink });
  addSurface(slide, { left: 238, top: 570, width: 804, height: 50 }, "#E8F5F4", "#B8DEDC");
  addText(slide, "This step makes search independent from website-specific HTML structure.", { left: 270, top: 584, width: 740, height: 24 }, { fontSize: 20, color: colors.ink, alignment: "center" });
  addNotes(slide, "Describe the parser as the boundary between messy web pages and clean searchable documents.");
}

function indexStructure(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Partitioned inverted indexes make local node search possible", "Each search node can answer using only its shard and index.", 8);
  addText(slide, "Partitioning", { left: 102, top: 200, width: 250, height: 32 }, { fontSize: 25, bold: true, color: colors.navy });
  addNode(slide, "documents", { left: 132, top: 278, width: 150, height: 54 }, colors.navy);
  addArrow(slide, { left: 308, top: 294, width: 50, height: 24 });
  addNode(slide, "shard 1", { left: 394, top: 232, width: 130, height: 48 }, colors.teal, colors.white, 16);
  addNode(slide, "shard 2", { left: 394, top: 306, width: 130, height: 48 }, colors.blue, colors.white, 16);
  addNode(slide, "shard 3", { left: 394, top: 380, width: 130, height: 48 }, colors.violet, colors.white, 16);
  addText(slide, "range-based split by document order", { left: 120, top: 468, width: 410, height: 30 }, { fontSize: 19, color: colors.muted, alignment: "center" });
  addText(slide, "Index per shard", { left: 704, top: 200, width: 250, height: 32 }, { fontSize: 25, bold: true, color: colors.navy });
  addSurface(slide, { left: 684, top: 262, width: 394, height: 224 }, "#F8FAFC", colors.line);
  addText(slide, "term -> posting list", { left: 738, top: 296, width: 270, height: 30 }, { fontSize: 24, bold: true, color: colors.teal, alignment: "center" });
  addText(slide, "distributed -> [(D17, tf=3), (D82, tf=2)]\nindexing -> [(D04, tf=5), (D91, tf=1)]\nquery -> [(D11, tf=2), (D54, tf=4)]", { left: 722, top: 352, width: 320, height: 96 }, { fontSize: 18, color: colors.ink });
  addText(slide, "global_stats.json stores corpus-level values used during ranking.", { left: 678, top: 524, width: 430, height: 44 }, { fontSize: 20, color: colors.ink, alignment: "center" });
  addNotes(slide, "Explain that each node holds a partition and builds a local inverted index. Global statistics keep scores comparable when results are merged.");
}

function queryFlow(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  addTitle(slide, "Parallel query processing is the next layer over these indexes", "The same query is sent to all search nodes, then the coordinator merges results.", 9);
  const userQuery = addNode(slide, "User query", { left: 98, top: 322, width: 150, height: 58 }, colors.green);
  addArrow(slide, { left: 270, top: 338, width: 54, height: 26 });
  const coordinator = addNode(slide, "Coordinator", { left: 354, top: 316, width: 164, height: 70 }, colors.navy);
  const n1 = addNode(slide, "Node 1\nlocal top-K", { left: 704, top: 222, width: 170, height: 66 }, colors.teal, colors.white, 17);
  const n2 = addNode(slide, "Node 2\nlocal top-K", { left: 704, top: 320, width: 170, height: 66 }, colors.blue, colors.white, 17);
  const n3 = addNode(slide, "Node 3\nlocal top-K", { left: 704, top: 418, width: 170, height: 66 }, colors.violet, colors.white, 17);
  connect(slide, coordinator, n1, { fromSide: "right", toSide: "left" });
  connect(slide, coordinator, n2, { fromSide: "right", toSide: "left" });
  connect(slide, coordinator, n3, { fromSide: "right", toSide: "left" });
  addArrow(slide, { left: 908, top: 338, width: 54, height: 26 });
  addNode(slide, "Merged\nranked list", { left: 998, top: 316, width: 170, height: 70 }, colors.amber, colors.white, 17);
  addSurface(slide, { left: 250, top: 548, width: 780, height: 54 }, colors.white, colors.line);
  addText(slide, "FA-1 prepares the shard/index files; Member 2 exposes them through RPC services.", { left: 282, top: 564, width: 716, height: 24 }, { fontSize: 20, color: colors.ink, alignment: "center" });
  addNotes(slide, "Use this slide to connect Member 1 outputs to Unit 2 communication. The query path is the runtime layer built on top of the prepared indexes.");
}

function faultTolerance(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Fault tolerance is handled through node independence", "A failed shard reduces coverage but should not stop the complete system.", 10);
  const coordinator = addNode(slide, "Coordinator", { left: 536, top: 214, width: 170, height: 58 }, colors.navy);
  const n1 = addNode(slide, "Node 1\navailable", { left: 250, top: 380, width: 160, height: 64 }, colors.teal, colors.white, 17);
  const n2 = addNode(slide, "Node 2\ntimeout", { left: 560, top: 380, width: 160, height: 64 }, colors.red, colors.white, 17);
  const n3 = addNode(slide, "Node 3\navailable", { left: 870, top: 380, width: 160, height: 64 }, colors.violet, colors.white, 17);
  connect(slide, coordinator, n1, { fromSide: "bottom", toSide: "top" });
  connect(slide, coordinator, n2, { fromSide: "bottom", toSide: "top" });
  connect(slide, coordinator, n3, { fromSide: "bottom", toSide: "top" });
  addSurface(slide, { left: 178, top: 522, width: 924, height: 66 }, "#FFF8EA", "#F4D49A");
  addText(slide, "Coordinator returns partial results and reports the unavailable node instead of failing the whole query.", { left: 220, top: 542, width: 840, height: 26 }, { fontSize: 21, color: colors.ink, alignment: "center" });
  addNotes(slide, "Explain the intended fault-tolerance behavior for the project. Member 1's independent shards make this possible because each node can work separately.");
}

function syllabus(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  addTitle(slide, "FA-1 syllabus concepts are directly represented", "The implementation maps to Unit 1 and Unit 2 rather than only being a coding exercise.", 11);
  addText(slide, "Unit 1: Distributed Systems", { left: 124, top: 216, width: 430, height: 32 }, { fontSize: 25, bold: true, color: colors.teal });
  addBullet(slide, "architecture: coordinator and search nodes", 132, 286, 430, colors.teal);
  addBullet(slide, "design issue: partitioning and ranking", 132, 346, 430, colors.teal);
  addBullet(slide, "scalability: add nodes and shards", 132, 406, 430, colors.teal);
  addBullet(slide, "middleware boundary between services", 132, 466, 430, colors.teal);
  addText(slide, "Unit 2: Communication", { left: 708, top: 216, width: 430, height: 32 }, { fontSize: 25, bold: true, color: colors.blue });
  addBullet(slide, "RPC: gRPC coordinator-to-node calls", 716, 286, 430, colors.blue);
  addBullet(slide, "message flow: fan-out and gather", 716, 346, 430, colors.blue);
  addBullet(slide, "identifiers: node IDs and addresses", 716, 406, 430, colors.blue);
  addBullet(slide, "fault tolerance: timeout and partial result", 716, 466, 430, colors.blue);
  addNotes(slide, "This slide is for faculty alignment with FA-1 guidelines. Keep answers precise and tied to visible implementation pieces.");
}

function demoPlan(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Demonstration plan for evaluation", "The demo should prove the implementation step by step.", 12);
  const rows = [
    ["1", "Show configuration", "allowed domains, per-site limits, crawler settings"],
    ["2", "Run pipeline command", "crawl -> parse -> shard -> index"],
    ["3", "Inspect generated files", "raw pages, JSONL documents, shard files, indexes"],
    ["4", "Run local search", "verify index lookup before distributed runtime"],
    ["5", "Connect Member 2", "start nodes, send RPC query, aggregate results"],
  ];
  for (let i = 0; i < rows.length; i++) {
    const top = 198 + i * 80;
    addNode(slide, rows[i][0], { left: 120, top, width: 54, height: 54 }, i < 4 ? colors.teal : colors.amber, colors.white, 18);
    addText(slide, rows[i][1], { left: 214, top: top - 2, width: 330, height: 30 }, { fontSize: 23, bold: true, color: colors.navy });
    addText(slide, rows[i][2], { left: 214, top: top + 36, width: 760, height: 26 }, { fontSize: 18, color: colors.muted });
  }
  addSurface(slide, { left: 820, top: 504, width: 290, height: 84 }, "#E8F5F4", "#B8DEDC");
  addText(slide, "Expected viva answer", { left: 858, top: 518, width: 214, height: 24 }, { fontSize: 19, bold: true, color: colors.navy, alignment: "center" });
  addText(slide, "Distribution comes from partitioned indexes and parallel node processing.", { left: 850, top: 552, width: 230, height: 30 }, { fontSize: 17, color: colors.ink, alignment: "center" });
  addNotes(slide, "Close with an evaluation-ready flow. The student should be able to run commands and explain each generated artifact.");
}

async function main() {
  await fs.mkdir(OUT, { recursive: true });
  const presentation = Presentation.create({ slideSize: { width: W, height: H } });
  cover(presentation);
  problem(presentation);
  architecture(presentation);
  memberSplit(presentation);
  pipeline(presentation);
  crawlerDesign(presentation);
  documentModel(presentation);
  indexStructure(presentation);
  queryFlow(presentation);
  faultTolerance(presentation);
  syllabus(presentation);
  demoPlan(presentation);

  const inspect = await presentation.inspect({ kind: "slide,textbox,shape,chart,notes", maxChars: 18000 });
  await fs.writeFile(`${OUT}/inspect.ndjson`, inspect.ndjson);

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    await writeBlob(`${OUT}/${stem}.png`, await presentation.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`${OUT}/${stem}.layout.json`, await layout.text());
  }

  await writeBlob(`${OUT}/deck-montage.webp`, await presentation.export({ format: "webp", montage: true, scale: 1 }));
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(FINAL);
  console.log(FINAL);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
