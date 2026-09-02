import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const OUT = "D:/Engg/diss/distributed-search-engine/ppt_build/output";
const FINAL = "D:/Engg/diss/distributed-search-engine/Distributed_Search_Engine_FA1_Presentation.pptx";

const W = 1280;
const H = 720;
const page = { left: 72, top: 58, width: 1136, height: 602 };

const colors = {
  navy: "#102A43",
  ink: "#17212B",
  muted: "#5D6B78",
  pale: "#F4F7FA",
  line: "#D8E0E8",
  teal: "#0E7C7B",
  cyan: "#36A3C7",
  green: "#2E7D32",
  amber: "#F2A93B",
  blue: "#3267B1",
  violet: "#6C5CE7",
  white: "#FFFFFF",
};

const websites = [
  "Python Docs",
  "gRPC",
  "Kubernetes",
  "MDN",
  "Docker",
  "PostgreSQL",
  "Hadoop",
  "NGINX",
  "GeeksforGeeks",
  "MongoDB",
];

const sourceBlock =
  "[Sources]\n" +
  "D:/Engg/diss/distributed_search_engine_project_context.md\n" +
  "D:/Engg/diss/distributed-search-engine/README.md\n" +
  "Verified local command: python app.py stats\n" +
  "Generated corpus files: data/metadata.jsonl, data/documents.jsonl, data/shards, indexes/global_stats.json";

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
  addText(slide, "FA-1 Distributed Systems Mini Project", { left: 72, top: 686, width: 360, height: 22 }, { fontSize: 13, color: colors.muted });
  addText(slide, String(number).padStart(2, "0"), { left: 1160, top: 684, width: 48, height: 24 }, { fontSize: 14, bold: true, color: colors.muted, alignment: "right" });
}

function addTitle(slide, title, subtitle, number) {
  addText(slide, title, { left: page.left, top: 48, width: 980, height: 54 }, { fontSize: 37, bold: true, color: colors.navy });
  if (subtitle) {
    addText(slide, subtitle, { left: page.left, top: 102, width: 980, height: 36 }, { fontSize: 18, color: colors.muted });
  }
  slide.shapes.add({
    geometry: "rect",
    position: { left: 72, top: 142, width: 112, height: 5 },
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
    borderRadius: 16,
    shadow: "shadow-sm",
  });
}

function addMetric(slide, value, label, position, accent = colors.teal) {
  addSurface(slide, position, colors.white, colors.line);
  slide.shapes.add({
    geometry: "rect",
    position: { left: position.left, top: position.top, width: 8, height: position.height },
    fill: accent,
    line: { style: "solid", fill: "none", width: 0 },
  });
  addText(slide, value, { left: position.left + 26, top: position.top + 22, width: position.width - 44, height: 54 }, { fontSize: 40, bold: true, color: colors.navy });
  addText(slide, label, { left: position.left + 28, top: position.top + 82, width: position.width - 50, height: 48 }, { fontSize: 17, color: colors.muted });
}

function addProcessNode(slide, label, position, fill, textColor = colors.white) {
  const shape = slide.shapes.add({
    geometry: "roundRect",
    position,
    fill,
    line: { style: "solid", fill: "none", width: 0 },
    borderRadius: 14,
    shadow: "shadow-sm",
  });
  shape.text = label;
  shape.text.style = { fontSize: 18, bold: true, color: textColor, alignment: "center" };
  return shape;
}

function addArrow(slide, position, fill = colors.teal) {
  const arrow = slide.shapes.add({
    geometry: "rightArrow",
    position,
    fill,
    line: { style: "solid", fill: "none", width: 0 },
  });
  return arrow;
}

function addNotes(slide, text) {
  slide.speakerNotes.textFrame.setText(text + "\n\n" + sourceBlock);
  slide.speakerNotes.setVisible(true);
}

function cover(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  slide.shapes.add({ geometry: "rect", position: { left: 0, top: 0, width: 1280, height: 720 }, fill: colors.pale, line: { style: "solid", fill: "none", width: 0 } });
  slide.shapes.add({ geometry: "rect", position: { left: 0, top: 0, width: 26, height: 720 }, fill: colors.teal, line: { style: "solid", fill: "none", width: 0 } });
  addText(slide, "Fault-Tolerant Distributed Search Engine", { left: 86, top: 112, width: 705, height: 150 }, { fontSize: 52, bold: true, color: colors.navy });
  addText(slide, "Partitioned Indexing and Parallel Query Processing", { left: 90, top: 280, width: 680, height: 42 }, { fontSize: 24, color: colors.muted });
  addText(slide, "FA-1 Mini Project | Distributed Systems", { left: 90, top: 352, width: 520, height: 32 }, { fontSize: 18, bold: true, color: colors.teal });
  addMetric(slide, "1,000", "offline documents prepared for search", { left: 828, top: 126, width: 300, height: 144 }, colors.teal);
  addMetric(slide, "10", "technology websites crawled", { left: 828, top: 304, width: 300, height: 144 }, colors.blue);
  addMetric(slide, "3", "distributed search-node shards", { left: 828, top: 482, width: 300, height: 144 }, colors.amber);
  addFooter(slide, 1);
  addNotes(slide, "Opening slide. Establish that the project is a miniature distributed search engine and that Member 1 has completed the offline corpus and indexing pipeline.");
}

function problem(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Centralized search becomes a bottleneck", "The project turns search into a real distributed-systems problem.", 2);
  addText(slide, "Traditional centralized search", { left: 94, top: 210, width: 360, height: 32 }, { fontSize: 24, bold: true, color: colors.navy });
  const center = addProcessNode(slide, "Single machine\nindex + query", { left: 150, top: 298, width: 230, height: 110 }, colors.navy);
  const boxes = [
    ["Storage limit", 92, 482, colors.amber],
    ["Query overload", 320, 482, colors.amber],
    ["Single failure point", 548, 482, colors.amber],
  ];
  for (const [label, left, top, color] of boxes) addProcessNode(slide, label, { left, top, width: 184, height: 58 }, color);
  addText(slide, "Distributed search response", { left: 742, top: 210, width: 360, height: 32 }, { fontSize: 24, bold: true, color: colors.navy });
  const n1 = addProcessNode(slide, "Node 1", { left: 740, top: 294, width: 130, height: 64 }, colors.teal);
  const n2 = addProcessNode(slide, "Node 2", { left: 908, top: 294, width: 130, height: 64 }, colors.blue);
  const n3 = addProcessNode(slide, "Node 3", { left: 1076, top: 294, width: 130, height: 64 }, colors.violet);
  addProcessNode(slide, "Coordinator", { left: 888, top: 432, width: 170, height: 68 }, colors.navy);
  addText(slide, "Partition data, search in parallel, aggregate ranked results.", { left: 740, top: 546, width: 466, height: 54 }, { fontSize: 22, color: colors.ink });
  addNotes(slide, "Explain the motivation: one central index cannot scale well for storage, query load, and fault tolerance. The proposed architecture partitions the index across independent nodes.");
}

function architecture(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  addTitle(slide, "FA-1 architecture has offline and online halves", "Our completed part prepares everything the distributed query layer needs.", 3);
  const nodes = [
    ["Crawler", 82, 238, colors.teal],
    ["Parser", 252, 238, colors.blue],
    ["Partitioner", 422, 238, colors.violet],
    ["Indexes", 592, 238, colors.amber],
    ["Search Nodes", 804, 238, colors.teal],
    ["Coordinator", 974, 238, colors.navy],
  ];
  for (let i = 0; i < nodes.length; i++) {
    const [label, left, top, color] = nodes[i];
    addProcessNode(slide, label, { left, top, width: 132, height: 70 }, color);
    if (i < nodes.length - 1) addArrow(slide, { left: left + 140, top: top + 22, width: 40, height: 26 }, colors.muted);
  }
  addText(slide, "Member 1: offline corpus and indexes", { left: 86, top: 370, width: 650, height: 30 }, { fontSize: 24, bold: true, color: colors.navy });
  addText(slide, "Member 2: online query layer", { left: 806, top: 370, width: 410, height: 34 }, { fontSize: 23, bold: true, color: colors.navy });
  addSurface(slide, { left: 82, top: 420, width: 650, height: 104 }, "#E8F5F4", "#B8DEDC");
  addText(slide, "Raw HTML -> logical documents -> three shards -> local inverted indexes -> global IDF statistics", { left: 112, top: 454, width: 590, height: 48 }, { fontSize: 22, color: colors.ink });
  addSurface(slide, { left: 806, top: 420, width: 410, height: 104 }, "#EEF3FA", "#BDD0EA");
  addText(slide, "gRPC search nodes, coordinator, local top-K results, global aggregation, failure display", { left: 836, top: 448, width: 350, height: 64 }, { fontSize: 20, color: colors.ink });
  addNotes(slide, "Position Member 1 as the completed data foundation and Member 2 as the online distributed search path.");
}

function pipeline(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Member 1 turns web pages into searchable shards", "The crawler stores real HTML offline before any indexing is performed.", 4);
  const y = 240;
  const items = [
    ["Public URLs", colors.navy],
    ["Raw HTML", colors.teal],
    ["Documents", colors.blue],
    ["Shards", colors.violet],
    ["Indexes", colors.amber],
  ];
  for (let i = 0; i < items.length; i++) {
    const left = 92 + i * 222;
    addProcessNode(slide, items[i][0], { left, top: y, width: 150, height: 78 }, items[i][1]);
    if (i < items.length - 1) addArrow(slide, { left: left + 158, top: y + 27, width: 50, height: 28 }, colors.muted);
  }
  const desc = [
    ["Controlled crawl", "BFS, allowlist, robots.txt, delay"],
    ["Offline storage", "data/raw/000001.html ..."],
    ["HTML parsing", "title, URL, visible content"],
    ["Range split", "shard_1, shard_2, shard_3"],
    ["TF-IDF prep", "local postings + global IDF"],
  ];
  for (let i = 0; i < desc.length; i++) {
    const left = 76 + i * 222;
    addText(slide, desc[i][0], { left, top: 360, width: 180, height: 28 }, { fontSize: 20, bold: true, color: colors.navy, alignment: "center" });
    addText(slide, desc[i][1], { left, top: 395, width: 180, height: 58 }, { fontSize: 17, color: colors.muted, alignment: "center" });
  }
  addSurface(slide, { left: 170, top: 540, width: 940, height: 70 }, "#FFF8EA", "#F4D49A");
  addText(slide, "Key principle: the crawler does not search. It produces a clean offline corpus for downstream parsing, sharding, and indexing.", { left: 202, top: 560, width: 876, height: 34 }, { fontSize: 22, color: colors.ink, alignment: "center" });
  addNotes(slide, "Walk through the offline pipeline and stress that raw web pages are preserved separately from parsed logical documents.");
}

function corpus(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  addTitle(slide, "The FA-1 corpus is balanced across 10 sources", "Each configured website contributed exactly 100 offline documents.", 5);
  addMetric(slide, "1,000", "raw HTML pages stored offline", { left: 82, top: 198, width: 250, height: 132 }, colors.teal);
  addMetric(slide, "1,000", "parsed logical documents", { left: 82, top: 360, width: 250, height: 132 }, colors.blue);
  addMetric(slide, "10 x 100", "websites by documents", { left: 82, top: 522, width: 250, height: 132 }, colors.amber);
  slide.charts.add("bar", {
    position: { left: 386, top: 202, width: 780, height: 382 },
    categories: websites,
    series: [{ name: "Documents", values: websites.map(() => 100), fill: colors.teal }],
    hasLegend: false,
    barOptions: { direction: "column", grouping: "clustered", gapWidth: 70 },
    yAxis: { min: 0, max: 120, majorUnit: 20, majorGridlines: { style: "solid", fill: colors.line, width: 1 } },
    dataLabels: { showValue: true, position: "outEnd", textStyle: { fontSize: 12, fill: colors.ink } },
    title: "Documents per source",
    titleTextStyle: { fontSize: 18, bold: true, fill: colors.navy },
  });
  addNotes(slide, "Use this slide as evidence that the corpus is real, offline, and balanced across 10 technical websites.");
}

function sharding(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Range-based sharding prepares independent search nodes", "The 1,000 documents are split into three node-owned partitions.", 6);
  const shardData = [
    ["Node 1", "334 docs", "35,803 terms", colors.teal],
    ["Node 2", "334 docs", "13,726 terms", colors.blue],
    ["Node 3", "332 docs", "21,854 terms", colors.violet],
  ];
  for (let i = 0; i < shardData.length; i++) {
    const left = 110 + i * 370;
    addSurface(slide, { left, top: 214, width: 300, height: 214 }, colors.white, colors.line);
    slide.shapes.add({ geometry: "rect", position: { left, top: 214, width: 300, height: 12 }, fill: shardData[i][3], line: { style: "solid", fill: "none", width: 0 } });
    addText(slide, shardData[i][0], { left: left + 28, top: 248, width: 244, height: 36 }, { fontSize: 28, bold: true, color: colors.navy, alignment: "center" });
    addText(slide, shardData[i][1], { left: left + 40, top: 306, width: 220, height: 44 }, { fontSize: 34, bold: true, color: shardData[i][3], alignment: "center" });
    addText(slide, shardData[i][2], { left: left + 44, top: 368, width: 212, height: 30 }, { fontSize: 20, color: colors.muted, alignment: "center" });
  }
  slide.charts.add("doughnut", {
    position: { left: 388, top: 468, width: 500, height: 160 },
    categories: ["Node 1", "Node 2", "Node 3"],
    series: [{ name: "Documents", values: [334, 334, 332], points: [{ idx: 0, fill: colors.teal }, { idx: 1, fill: colors.blue }, { idx: 2, fill: colors.violet }] }],
    hasLegend: true,
    legend: { position: "right", textStyle: { fontSize: 14, fill: colors.ink } },
    dataLabels: { showValue: true, position: "center", textStyle: { fontSize: 12, fill: colors.white, bold: true } },
    doughnutOptions: { holeSize: 58 },
  });
  addNotes(slide, "Explain that each future search-node process loads only its own shard and local index. This is the direct link to partitioned indexing.");
}

function indexStructure(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  addTitle(slide, "The inverted index stores where every term appears", "Posting lists make document lookup fast inside each shard.", 7);
  addSurface(slide, { left: 90, top: 208, width: 400, height: 352 }, colors.white, colors.line);
  addText(slide, "Term dictionary", { left: 126, top: 238, width: 300, height: 32 }, { fontSize: 24, bold: true, color: colors.navy });
  const terms = [["distributed", colors.teal], ["indexing", colors.blue], ["parallel", colors.violet], ["query", colors.amber]];
  for (let i = 0; i < terms.length; i++) {
    addProcessNode(slide, terms[i][0], { left: 126, top: 298 + i * 54, width: 190, height: 38 }, terms[i][1]);
    addArrow(slide, { left: 326, top: 306 + i * 54, width: 56, height: 22 }, colors.muted);
  }
  addSurface(slide, { left: 572, top: 208, width: 610, height: 352 }, colors.white, colors.line);
  addText(slide, "Posting list example", { left: 608, top: 238, width: 320, height: 32 }, { fontSize: 24, bold: true, color: colors.navy });
  const postings = [
    "D17: frequency=3, positions=[4, 18, 32]",
    "D82: frequency=2, positions=[9, 47]",
    "D214: frequency=5, positions=[3, 11, 42, 67, 88]",
  ];
  for (let i = 0; i < postings.length; i++) {
    addText(slide, postings[i], { left: 620, top: 306 + i * 64, width: 520, height: 30 }, { fontSize: 22, color: colors.ink });
  }
  addText(slide, "Global IDF keeps scores comparable across nodes.", { left: 620, top: 520, width: 510, height: 72 }, { fontSize: 23, bold: true, color: colors.teal });
  addNotes(slide, "Explain inverted index fundamentals: term to posting list. Posting list records frequency and positions, while global IDF supports fair ranking across shards.");
}

function handoff(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "Member 2 can now attach the distributed query layer", "The handoff files define a clear contract between both parts.", 8);
  const leftItems = [
    "data/shards/shard_1.jsonl",
    "data/shards/shard_2.jsonl",
    "data/shards/shard_3.jsonl",
    "indexes/node_1/index.json",
    "indexes/node_2/index.json",
    "indexes/node_3/index.json",
    "indexes/global_stats.json",
  ];
  addSurface(slide, { left: 86, top: 202, width: 514, height: 388 }, "#F8FAFC", colors.line);
  addText(slide, "Files produced by Member 1", { left: 120, top: 232, width: 420, height: 30 }, { fontSize: 24, bold: true, color: colors.navy });
  for (let i = 0; i < leftItems.length; i++) {
    addText(slide, leftItems[i], { left: 124, top: 286 + i * 38, width: 420, height: 24 }, { fontSize: 18, color: colors.ink });
  }
  addSurface(slide, { left: 690, top: 202, width: 490, height: 388 }, "#E8F5F4", "#B8DEDC");
  addText(slide, "Distributed runtime contract", { left: 728, top: 232, width: 380, height: 30 }, { fontSize: 24, bold: true, color: colors.navy });
  const runtime = [
    "Each node loads one shard and one local index.",
    "Coordinator sends query to all nodes using gRPC.",
    "Nodes return local top-K results.",
    "Coordinator merges and ranks final results.",
    "If one node fails, other shards still respond.",
  ];
  for (let i = 0; i < runtime.length; i++) {
    addText(slide, runtime[i], { left: 730, top: 292 + i * 50, width: 392, height: 30 }, { fontSize: 20, color: colors.ink });
  }
  addNotes(slide, "Use this slide to show team coordination. Member 1 has a concrete output contract; Member 2 uses those files for gRPC services and the web interface.");
}

function syllabus(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.pale;
  addTitle(slide, "FA-1 concepts are visible in the implementation", "The project maps directly to Unit 1 and Unit 2 expectations.", 9);
  addSurface(slide, { left: 92, top: 204, width: 500, height: 384 }, colors.white, colors.line);
  addText(slide, "Unit 1: Distributed Systems", { left: 128, top: 236, width: 400, height: 34 }, { fontSize: 25, bold: true, color: colors.teal });
  const u1 = [
    "Multiple independent search-node partitions",
    "Coordinator + search-node architecture",
    "Scalability through document sharding",
    "Design issue: partitioning and global ranking",
  ];
  for (let i = 0; i < u1.length; i++) addText(slide, u1[i], { left: 132, top: 306 + i * 58, width: 390, height: 30 }, { fontSize: 20, color: colors.ink });
  addSurface(slide, { left: 688, top: 204, width: 500, height: 384 }, colors.white, colors.line);
  addText(slide, "Unit 2: Communication", { left: 724, top: 236, width: 400, height: 34 }, { fontSize: 25, bold: true, color: colors.blue });
  const u2 = [
    "gRPC planned for coordinator-to-node RPC",
    "Node IDs and addresses for each service",
    "Parallel request/response message flow",
    "Timeouts allow partial results on node failure",
  ];
  for (let i = 0; i < u2.length; i++) addText(slide, u2[i], { left: 728, top: 306 + i * 58, width: 390, height: 30 }, { fontSize: 20, color: colors.ink });
  addNotes(slide, "Frame the implementation using the exact FA-1 syllabus language: architecture, design issues, middleware, RPC, identifiers, and fault tolerance.");
}

function demoPlan(presentation) {
  const slide = presentation.slides.add();
  slide.background.fill = colors.white;
  addTitle(slide, "The faculty demo can show evidence, not just code", "A short command sequence proves the offline pipeline is complete.", 10);
  const steps = [
    ["1", "Show corpus stats", "python app.py stats"],
    ["2", "Run a local index search", "python app.py search \"parallel query processing\""],
    ["3", "Open generated files", "data/raw, documents.jsonl, data/shards, indexes"],
    ["4", "Connect to Member 2", "Launch nodes and coordinator for gRPC demo"],
  ];
  for (let i = 0; i < steps.length; i++) {
    const top = 206 + i * 96;
    addProcessNode(slide, steps[i][0], { left: 102, top, width: 58, height: 58 }, i < 3 ? colors.teal : colors.amber);
    addText(slide, steps[i][1], { left: 190, top: top - 2, width: 390, height: 32 }, { fontSize: 24, bold: true, color: colors.navy });
    addText(slide, steps[i][2], { left: 190, top: top + 38, width: 780, height: 26 }, { fontSize: 18, color: colors.muted });
  }
  addSurface(slide, { left: 856, top: 204, width: 270, height: 314 }, "#102A43", "#102A43");
  addText(slide, "Evaluation message", { left: 890, top: 244, width: 204, height: 30 }, { fontSize: 22, bold: true, color: colors.white, alignment: "center" });
  addText(slide, "The system is distributed because indexes are partitioned across independent nodes, not because distribution was added artificially.", { left: 888, top: 306, width: 210, height: 146 }, { fontSize: 21, color: colors.white, alignment: "center" });
  addNotes(slide, "Close by connecting evidence to evaluation. The offline part is complete and ready for the distributed runtime demonstration.");
}

async function main() {
  await fs.mkdir(OUT, { recursive: true });
  const presentation = Presentation.create({ slideSize: { width: W, height: H } });
  cover(presentation);
  problem(presentation);
  architecture(presentation);
  pipeline(presentation);
  corpus(presentation);
  sharding(presentation);
  indexStructure(presentation);
  handoff(presentation);
  syllabus(presentation);
  demoPlan(presentation);

  const inspect = await presentation.inspect({ kind: "slide,textbox,shape,chart,notes", maxChars: 16000 });
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
