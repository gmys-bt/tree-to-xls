import * as XLSX from 'xlsx';

const tree = [
  {
    groupName: 'group 1',
    children: [
      {
        'Task Name': 'nuggets',
        address: 'some address',
        city: 'chicago',
        children: [
          {
            name: 'nested record',
          },
        ],
      },
      {
        'Task Name': 'fries',
        address: 'some address',
        city: 'new york',
      },
      {
        'Task Name': 'ice cream',
        address: 'some address',
        city: 'boston',
      },
    ],
  },
  {
    groupName: 'group 2',
    children: [
      {
        'Task Name': 'nuggets',
        address: 'some address',
        city: 'chicago',
        children: [
          {
            name: 'nested record',
          },
        ],
      },
      {
        'Task Name': 'fries',
        address: 'some address',
        city: 'new york',
      },
      {
        'Task Name': 'ice cream',
        address: 'some address',
        city: 'boston',
      },
    ],
  },
];

const flatten = (treeStructure, level = 0) => {
  const rows = [];
  const headers = [];

  for (const node of treeStructure) {
    const row = [];
    // at what cell do we start inserting data
    const baseLevel = level > 0 ? level - 1 : level;

    // get cell names based on props
    const keys = Object.keys(node).filter((k) => k !== 'children');

    // deduplicated headers per level
    headers[level] = keys.map((k) => `${level - 1}.${k}`);

    keys.forEach((k, i) => {
      row[baseLevel + i] = node[k];
    });

    // just two levels, 0 and 1 are necessary for outlines
    rows.push({ row, level: level > 1 ? 2 : level });

    if (node.children) {
      const parsedChildren = flatten(node.children, level + keys.length);
      rows.push(...parsedChildren.rows);
      headers.push(...parsedChildren.headers);
    }
  }

  return { rows, headers: headers.filter(Boolean).flat() };
};

const result = flatten(tree);

const deduplicatedHeaders = [
  ...new Set(result.headers.slice(1, result.headers.length)),
].map((h) => h.split('.')[1]);

console.log(result.rows);

const dataToExport = [
  deduplicatedHeaders,
  ...result.rows.map((row) => row.row),
];
const ws = XLSX.utils.aoa_to_sheet(dataToExport);

console.log(dataToExport);

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

// Create workbook and write file
const wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
XLSX.writeFile(wb, 'tree-outline.xlsx');
