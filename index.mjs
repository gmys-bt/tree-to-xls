import * as XLSX from 'xlsx';

const TASK_NAME = 'Task Name';
const WBS_CODE_PPSA = 'WBS Code (PPSA)';
const ROLE_NAME_PRICING_SHEET =
  'Role Name (Pricing Sheet Original Reference - Do n';
const ASSIGNED_TO = 'Assigned To';
const BUDGETED_HOURS = 'Budgeted Hours';
const EFFORT_HRS = 'Effort (hrs)';
const WORKED_TASK_HOURS = 'Worked Task Hours';
const REMAINING_TASK_HOURS = 'Remaining Task Hours';
const DURATION = 'Duration';
const START_DATE = 'Start Date';
const END_DATE = 'End Date';
const PREDECESSORS = 'Predecessors';
const PERCENT_ALLOCATION = '% Allocation';
const TASK_TYPE = 'Task Type';
const COMMENTS = 'Comments';
const OPEN_FOR_TIME = 'Open for Time';
const WORKED_MINUTES_PPSA = 'Worked Minutes (PPSA)';
const REMAINING_MINUTES_PPSA = 'Remaining Minutes (PPSA)';

// An array, to guarantee order
const columnMapping = [
  TASK_NAME,
  WBS_CODE_PPSA,
  ROLE_NAME_PRICING_SHEET,
  ASSIGNED_TO,
  BUDGETED_HOURS,
  EFFORT_HRS,
  WORKED_TASK_HOURS,
  REMAINING_TASK_HOURS,
  DURATION,
  START_DATE,
  END_DATE,
  PREDECESSORS,
  PERCENT_ALLOCATION,
  TASK_TYPE,
  COMMENTS,
  OPEN_FOR_TIME,
  WORKED_MINUTES_PPSA,
  REMAINING_MINUTES_PPSA,
];

const tree = [
  {
    [TASK_NAME]: 'P0 (Phase 0)',
    [WBS_CODE_PPSA]: '1.1.1',
    children: [
      {
        [TASK_NAME]: 'P0-Delivery / Project Management',
        [WBS_CODE_PPSA]: '1.1.1',
        children: [
          {
            [TASK_NAME]: 'Phase 0 Day Before',
            [WBS_CODE_PPSA]: '1.1.1',
            [ROLE_NAME_PRICING_SHEET]: '',
            [ASSIGNED_TO]: '',
            [BUDGETED_HOURS]: 618,
            [EFFORT_HRS]: 618,
            [WORKED_TASK_HOURS]: 0,
            [REMAINING_TASK_HOURS]: 0,
            [DURATION]: '68d',
            [START_DATE]: '08/24/25',
            [END_DATE]: '11/26/25',
            [PREDECESSORS]: '',
            [PERCENT_ALLOCATION]: '',
            [TASK_TYPE]: '',
            [COMMENTS]: '',
            [OPEN_FOR_TIME]: 'True',
            [WORKED_MINUTES_PPSA]: '',
            [REMAINING_MINUTES_PPSA]: 37080,
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

console.log(result.rows);

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
