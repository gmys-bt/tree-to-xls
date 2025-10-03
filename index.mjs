import * as XLSX from 'xlsx';
import { tree, columnMapping } from './js_tree_with_constants.mjs';

const flatten = (treeStructure, level = 0) => {
  const rows = [];

  for (const node of treeStructure) {
    const row = [];

    // put values on respective col indexes
    columnMapping.forEach((k, i) => {
      row[i] = node[k];
    });

    // just two levels, 0 and 1 are necessary for outlines
    rows.push({ row, level: level > 1 ? 2 : level });

    if (node.children) {
      const parsedChildren = flatten(node.children, level + 1);
      rows.push(...parsedChildren.rows);
    }
  }

  return { rows };
};

const result = flatten(tree);

const dataToExport = [columnMapping, ...result.rows.map((row) => row.row)];
const ws = XLSX.utils.aoa_to_sheet(dataToExport);

// console.log(result.rows);

if (!ws['outline']) ws['!outline'] = {};
ws['!outline'].above = true;

// Set outline levels (row-level, 0-based, skip header)
result.rows.forEach((r, i) => {
  // i+1 because worksheetData has header at 0
  const rowNum = i + 1; // 1-based, +1 for header
  if (!ws['!rows']) ws['!rows'] = [];
  ws['!rows'][rowNum] = ws['!rows'][rowNum] || {};
  ws['!rows'][rowNum].level = r.level;
});

const filePath = 'tree-outline.xlsx';

// Create workbook and write file
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
XLSX.writeFile(wb, filePath);

console.log('file written to:', filePath);
