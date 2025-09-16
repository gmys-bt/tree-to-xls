import * as XLSX from 'xlsx';

const tree = [
  {
    name: 'group 1',
    children: [
      {
        name: 'nuggets',
        address: 'some address',
        city: 'chicago',
        children: [
          {
            name: 'nested record',
          },
        ],
      },
      {
        name: 'fries',
        address: 'some address',
        city: 'new york',
      },
      {
        name: 'ice cream',
        address: 'some address',
        city: 'boston',
      },
    ],
  },
];

function flattenTree(nodes, level = 0, maxDepth = getMaxDepth(tree)) {
  const result = [];

  for (const node of nodes) {
    const row = {};

    // columns (using array indices for column positions)
    for (let i = 0; i <= maxDepth; i++) {
      row[i] = i === level ? node.name : '';
    }

    row[maxDepth + 1] = node.address || '';
    row[maxDepth + 2] = node.city || '';

    // outline level for this row (this is crucial for expand / collapse)
    row._outlineLevel = level;

    result.push(row);

    // add children (recursive)
    if (node.children && node.children.length > 0) {
      result.push(...flattenTree(node.children, level + 1, maxDepth));
    }
  }

  return result;
}

function getMaxDepth(nodes, currentDepth = 0) {
  let maxDepth = currentDepth;

  for (const node of nodes) {
    if (node.children && node.children.length > 0) {
      const childDepth = getMaxDepth(node.children, currentDepth + 1);
      maxDepth = Math.max(maxDepth, childDepth);
    }
  }

  return maxDepth;
}

const treeToGroupedXLSX = (treeData) => {
  const flatData = flattenTree(treeData);
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(flatData, { skipHeader: true });

  // set row outline levels for grouping
  if (!worksheet['!rows']) worksheet['!rows'] = [];

  flatData.forEach((row, index) => {
    if (!worksheet['!rows'][index]) worksheet['!rows'][index] = {};
    worksheet['!rows'][index].level = row._outlineLevel;
  });

  // remove the outline level from the actual data before creating the sheet - so that it wont be rendered in the xlsx itself
  const cleanData = flatData.map((row) => {
    const { _outlineLevel, ...cleanRow } = row;
    return cleanRow;
  });

  // recreate worksheet with clean data, without outline levels
  const cleanWorksheet = XLSX.utils.json_to_sheet(cleanData, {
    skipHeader: true,
  });

  // apply the row grouping to the clean worksheet
  if (!cleanWorksheet['!rows']) cleanWorksheet['!rows'] = [];
  flatData.forEach((row, index) => {
    if (!cleanWorksheet['!rows'][index]) cleanWorksheet['!rows'][index] = {};
    cleanWorksheet['!rows'][index].level = row._outlineLevel;
  });

  // Add the clean worksheet to workbook
  XLSX.utils.book_append_sheet(workbook, cleanWorksheet, 'Tree Data');

  // console.log('Data structure:');
  // console.table(flatData);

  return workbook;
};

XLSX.writeFile(treeToGroupedXLSX(tree), 'tree_output.xlsx');
