import * as XLSX from 'xlsx';

const columnMapping = [
  'Task Name',
  'WBS Code (PPSA)',
  'Role Name (Pricing Sheet Original Reference - Do n',
  'Assigned To',
  'Budgeted Hours',
  'Effort (hrs)',
  'Worked Task Hours',
  'Remaining Task Hours',
  'Duration',
  'Start Date',
  'End Date',
  'Predecessors',
  '% Allocation',
  'Task Type',
  'Comments',
  'Open for Time',
  'Worked Minutes (PPSA)',
  'Remaining Minutes (PPSA)',
];

const tree = [
  {
    name: 'P0 (Phase 0)',
    'WBS Code (PPSA)': '1.1.1',
    children: [
      {
        name: 'P0-Delivery / Project Management',
        'WBS Code (PPSA)': '1.1.1',
        children: [
          {
            'Task Name': 'Phase 0 Day Before',
            'WBS Code (PPSA)': '1.1.1',
            'Role Name (Pricing Sheet Original Reference - Do n': '',
            'Assigned To': '',
            'Budgeted Hours': 618,
            'Effort (hrs)': 618,
            'Worked Task Hours': 0,
            'Remaining Task Hours': 0,
            Duration: '68d',
            'Start Date': '08/24/25',
            'End Date': '11/26/25',
            Predecessors: '',
            '% Allocation': '',
            'Task Type': '',
            Comments: '',
            'Open for Time': 'True',
            'Worked Minutes (PPSA)': '',
            'Remaining Minutes (PPSA)': 37080,
          },
        ],
      },
    ],
  },
];

const flatten = (treeStructure, level = 0) => {
  const rows = [];

  for (const node of treeStructure) {
    const row = [];

    // get cell names based on props
    const keys = Object.keys(node).filter((k) => k !== 'children');

    // put values on respective col indexes
    keys.forEach((k, i) => {
      row[i] = node[k];
    });

    // just two levels, 0 and 1 are necessary for outlines
    rows.push({ row, level: level > 1 ? 2 : level });

    if (node.children) {
      const parsedChildren = flatten(node.children, level + keys.length);
      rows.push(...parsedChildren.rows);
    }
  }

  return { rows };
};

const result = flatten(tree);

const dataToExport = [columnMapping, ...result.rows.map((row) => row.row)];
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
