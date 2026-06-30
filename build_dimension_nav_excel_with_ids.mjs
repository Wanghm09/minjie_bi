import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const rootDir = "D:/skilltohtml";
const sourcePath = path.join(rootDir, "维度一级导航明细文档.md");
const attachmentPath = "C:/Users/whmyd/.codex/attachments/b9031cca-dbe5-44ed-b382-b209cbf30490/pasted-text.txt";
const outputDir = path.join(rootDir, "outputs", "dimension-nav-detail");
const outputPath = path.join(outputDir, "维度一级导航明细_含ID.xlsx");
const previewPath = path.join(outputDir, "维度一级导航明细_含ID_preview.png");

const markdown = await fs.readFile(sourcePath, "utf8");
const dimensionDocs = JSON.parse(await fs.readFile(attachmentPath, "utf8"));

const categoryOrder = [
  "地域门店",
  "时段日期",
  "账单消费",
  "排号订餐",
  "外送客服",
  "会员标签",
  "菜品",
  "供应链",
  "营销活动",
  "财务",
  "人事",
];

const explicitLabelMap = {
  "省份/城市": ["省份", "城市"],
  "大区经理/门店": ["大区经理", "门店名称"],
  "菜品分类/名称": ["菜品一级分类", "菜品二级分类", "总部菜品名称"],
  "券模板ID编码": ["券模版id编码"],
  "券模板ID名称": ["券模版id名称"],
  "人事_员工姓名": ["人事_员工姓名"],
};

const byLabel = new Map();
for (const item of dimensionDocs) {
  const label = String(item.label || "").trim();
  if (!byLabel.has(label)) byLabel.set(label, []);
  byLabel.get(label).push(item);
}

const detailRows = [];
const matchedIds = new Set();
const sectionRegex = /### 3\.\d+ ([^\n]+)\n\n([\s\S]*?)(?=\n### 3\.|\n## 4\.|$)/g;
let sectionMatch;
while ((sectionMatch = sectionRegex.exec(markdown))) {
  const category = sectionMatch[1].trim();
  const tableText = sectionMatch[2];
  for (const line of tableText.split(/\r?\n/)) {
    if (!line.startsWith("|")) continue;
    if (line.includes("---") || line.includes("序号")) continue;
    const cells = line.split("|").slice(1, -1).map((cell) => cell.trim());
    if (cells.length !== 4) continue;
    const [index, displayName, fieldName, defaultStatus] = cells;
    const sourceLabels = explicitLabelMap[displayName] || [displayName];
    const matched = sourceLabels.flatMap((label) => byLabel.get(label) || []);
    matched.forEach((item) => matchedIds.add(item.dimensionId));
    detailRows.push({
      category,
      index: Number(index),
      displayName,
      fieldName,
      dimensionIds: matched.map((item) => item.dimensionId).join("\n") || "未匹配",
      sourceKeys: matched.map((item) => item.key).join("\n") || "-",
      sourceLabels: matched.map((item) => item.label).join("\n") || "-",
      defaultStatus,
      docDefaultVisible: matched.length
        ? matched.map((item) => (item.defaultVisible ? "true" : "false")).join("\n")
        : "-",
      matchStatus: matched.length ? "已匹配" : "未匹配",
    });
  }
}

const summaryRows = categoryOrder.map((category, index) => {
  const rows = detailRows.filter((row) => row.category === category);
  const defaultRows = rows.filter((row) => row.defaultStatus === "默认展示");
  const idCount = rows.reduce((sum, row) => sum + (row.dimensionIds === "未匹配" ? 0 : row.dimensionIds.split("\n").length), 0);
  return {
    index: index + 1,
    category,
    total: rows.length,
    idCount,
    defaultCount: defaultRows.length,
    hiddenCount: rows.length - defaultRows.length,
    defaultNames: defaultRows.map((row) => row.displayName).join("、") || "-",
  };
});

const attachmentRows = dimensionDocs.map((item, index) => {
  const matchedRow = detailRows.find((row) => row.dimensionIds.split("\n").includes(item.dimensionId));
  return [
    index + 1,
    item.dimensionId,
    item.key,
    item.label,
    item.sortNo,
    item.enabled ? "true" : "false",
    item.valueDisplayMode,
    item.defaultVisible ? "true" : "false",
    matchedRow?.category || "未纳入当前一级导航",
    matchedRow?.displayName || "-",
  ];
});

const workbook = Workbook.create();
const overview = workbook.worksheets.add("一级导航总览");
const detail = workbook.worksheets.add("维度明细含ID");
const source = workbook.worksheets.add("附件ID清单");

for (const sheet of [overview, detail, source]) {
  sheet.showGridLines = false;
}

overview.getRange("A1:G1").merge();
overview.getRange("A1").values = [["维度一级导航ID总览"]];
overview.getRange("A2:G2").merge();
overview.getRange("A2").values = [["按一级导航统计维度数量、已匹配 dimensionId 数量和默认展示维度。"]];
overview.getRange("A4:G4").values = [["序号", "一级导航", "维度总数", "匹配ID数", "默认展示数", "默认隐藏数", "默认展示维度"]];
overview.getRangeByIndexes(4, 0, summaryRows.length, 7).values = summaryRows.map((row) => [
  row.index,
  row.category,
  row.total,
  row.idCount,
  row.defaultCount,
  row.hiddenCount,
  row.defaultNames,
]);

detail.getRange("A1:I1").merge();
detail.getRange("A1").values = [["维度一级导航明细含ID"]];
detail.getRange("A2:I2").merge();
detail.getRange("A2").values = [["dimensionId 来自用户提供的维度文档；组合维度会在同一单元格内换行展示多个 ID。"]];
detail.getRange("A4:I4").values = [["一级导航", "序号", "页面展示名称", "维度字段", "dimensionId", "附件key", "附件label", "默认状态", "匹配状态"]];
detail.getRangeByIndexes(4, 0, detailRows.length, 9).values = detailRows.map((row) => [
  row.category,
  row.index,
  row.displayName,
  row.fieldName,
  row.dimensionIds,
  row.sourceKeys,
  row.sourceLabels,
  row.defaultStatus,
  row.matchStatus,
]);

source.getRange("A1:J1").merge();
source.getRange("A1").values = [["附件ID清单"]];
source.getRange("A2:J2").merge();
source.getRange("A2").values = [["保留用户提供文档中的全部 dimensionId，并标注是否纳入当前维度一级导航明细。"]];
source.getRange("A4:J4").values = [["序号", "dimensionId", "key", "label", "sortNo", "enabled", "valueDisplayMode", "defaultVisible", "匹配一级导航", "匹配页面展示名称"]];
source.getRangeByIndexes(4, 0, attachmentRows.length, 10).values = attachmentRows;

function styleSheet(sheet, titleRange, subtitleRange, headerRange, usedRange, widths) {
  titleRange.format = {
    fill: "#EAF3FF",
    font: { bold: true, color: "#0F3D82", size: 16 },
  };
  subtitleRange.format = {
    fill: "#F7FBFF",
    font: { color: "#4E5969", size: 10 },
  };
  headerRange.format = {
    fill: "#1769E0",
    font: { bold: true, color: "#FFFFFF" },
  };
  usedRange.format = {
    borders: {
      insideHorizontal: { style: "thin", color: "#DCE6F2" },
      insideVertical: { style: "thin", color: "#E8EEF7" },
      top: { style: "thin", color: "#BFD2EA" },
      bottom: { style: "thin", color: "#BFD2EA" },
      left: { style: "thin", color: "#BFD2EA" },
      right: { style: "thin", color: "#BFD2EA" },
    },
    font: { size: 10, color: "#1D2129" },
    wrapText: true,
  };
  for (let i = 0; i < widths.length; i += 1) {
    sheet.getRangeByIndexes(0, i, 1, 1).format.columnWidth = widths[i];
  }
  sheet.getRange("A1").format.rowHeight = 28;
  sheet.getRange("A2").format.rowHeight = 24;
  sheet.freezePanes.freezeRows(4);
}

styleSheet(
  overview,
  overview.getRange("A1:G1"),
  overview.getRange("A2:G2"),
  overview.getRange("A4:G4"),
  overview.getRangeByIndexes(3, 0, summaryRows.length + 1, 7),
  [8, 14, 12, 12, 14, 14, 70],
);

styleSheet(
  detail,
  detail.getRange("A1:I1"),
  detail.getRange("A2:I2"),
  detail.getRange("A4:I4"),
  detail.getRangeByIndexes(3, 0, detailRows.length + 1, 9),
  [14, 8, 24, 24, 34, 24, 24, 12, 12],
);

styleSheet(
  source,
  source.getRange("A1:J1"),
  source.getRange("A2:J2"),
  source.getRange("A4:J4"),
  source.getRangeByIndexes(3, 0, attachmentRows.length + 1, 10),
  [8, 34, 26, 22, 10, 10, 20, 14, 20, 24],
);

overview.tables.add(`A4:G${summaryRows.length + 4}`, true, "DimensionNavIdOverview");
detail.tables.add(`A4:I${detailRows.length + 4}`, true, "DimensionNavDetailWithIds");
source.tables.add(`A4:J${attachmentRows.length + 4}`, true, "DimensionIdSourceList");

const check = await workbook.inspect({
  kind: "table",
  range: "维度明细含ID!A1:I25",
  include: "values",
  tableMaxRows: 25,
  tableMaxCols: 9,
  maxChars: 3000,
});
console.log(check.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({
  sheetName: "维度明细含ID",
  range: "A1:I25",
  scale: 1,
  format: "png",
});
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);

console.log(outputPath);
