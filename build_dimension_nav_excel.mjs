import fs from "node:fs/promises";
import path from "node:path";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const rootDir = "D:/skilltohtml";
const sourcePath = path.join(rootDir, "维度一级导航明细文档.md");
const outputDir = path.join(rootDir, "outputs", "dimension-nav-detail");
const outputPath = path.join(outputDir, "维度一级导航明细.xlsx");
const previewPath = path.join(outputDir, "维度一级导航明细_preview.png");

const markdown = await fs.readFile(sourcePath, "utf8");

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

const detailRows = [];
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
    detailRows.push({
      category,
      index: Number(index),
      displayName,
      fieldName,
      defaultStatus,
    });
  }
}

const summaryRows = categoryOrder.map((category, index) => {
  const rows = detailRows.filter((row) => row.category === category);
  const defaultRows = rows.filter((row) => row.defaultStatus === "默认展示");
  return {
    index: index + 1,
    category,
    total: rows.length,
    defaultCount: defaultRows.length,
    hiddenCount: rows.length - defaultRows.length,
    defaultNames: defaultRows.map((row) => row.displayName).join("、") || "-",
  };
});

const workbook = Workbook.create();

const overview = workbook.worksheets.add("一级导航总览");
const detail = workbook.worksheets.add("维度明细");
const defaultSummary = workbook.worksheets.add("默认展示汇总");

for (const sheet of [overview, detail, defaultSummary]) {
  sheet.showGridLines = false;
}

overview.getRange("A1:F1").merge();
overview.getRange("A1").values = [["维度一级导航总览"]];
overview.getRange("A2:F2").merge();
overview.getRange("A2").values = [["按维度选择区左侧一级快速导航整理，统计每个一级分类下的维度数量和默认展示数量。"]];
overview.getRange("A4:F4").values = [["序号", "一级导航", "维度总数", "默认展示数", "默认隐藏数", "默认展示维度"]];
overview.getRangeByIndexes(4, 0, summaryRows.length, 6).values = summaryRows.map((row) => [
  row.index,
  row.category,
  row.total,
  row.defaultCount,
  row.hiddenCount,
  row.defaultNames,
]);

detail.getRange("A1:E1").merge();
detail.getRange("A1").values = [["维度一级导航明细"]];
detail.getRange("A2:E2").merge();
detail.getRange("A2").values = [["每一行对应一个维度，可按一级导航、默认状态筛选。"]];
detail.getRange("A4:E4").values = [["一级导航", "序号", "页面展示名称", "维度字段", "默认状态"]];
detail.getRangeByIndexes(4, 0, detailRows.length, 5).values = detailRows.map((row) => [
  row.category,
  row.index,
  row.displayName,
  row.fieldName,
  row.defaultStatus,
]);

defaultSummary.getRange("A1:C1").merge();
defaultSummary.getRange("A1").values = [["默认展示维度汇总"]];
defaultSummary.getRange("A2:C2").merge();
defaultSummary.getRange("A2").values = [["页面初始化时默认展示的维度，按一级导航归类。"]];
defaultSummary.getRange("A4:C4").values = [["一级导航", "默认展示数量", "默认展示维度"]];
defaultSummary.getRangeByIndexes(4, 0, summaryRows.length, 3).values = summaryRows.map((row) => [
  row.category,
  row.defaultCount,
  row.defaultNames,
]);

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
  overview.getRange("A1:F1"),
  overview.getRange("A2:F2"),
  overview.getRange("A4:F4"),
  overview.getRangeByIndexes(3, 0, summaryRows.length + 1, 6),
  [8, 14, 12, 14, 14, 70],
);
styleSheet(
  detail,
  detail.getRange("A1:E1"),
  detail.getRange("A2:E2"),
  detail.getRange("A4:E4"),
  detail.getRangeByIndexes(3, 0, detailRows.length + 1, 5),
  [14, 8, 24, 24, 12],
);
styleSheet(
  defaultSummary,
  defaultSummary.getRange("A1:C1"),
  defaultSummary.getRange("A2:C2"),
  defaultSummary.getRange("A4:C4"),
  defaultSummary.getRangeByIndexes(3, 0, summaryRows.length + 1, 3),
  [14, 14, 90],
);

overview.tables.add(`A4:F${summaryRows.length + 4}`, true, "DimensionNavOverview");
detail.tables.add(`A4:E${detailRows.length + 4}`, true, "DimensionNavDetail");
defaultSummary.tables.add(`A4:C${summaryRows.length + 4}`, true, "DimensionDefaultSummary");

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

const preview = await workbook.render({
  sheetName: "维度明细",
  range: "A1:E25",
  scale: 1,
  format: "png",
});

await fs.mkdir(outputDir, { recursive: true });
await fs.writeFile(previewPath, new Uint8Array(await preview.arrayBuffer()));

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);

console.log(outputPath);
