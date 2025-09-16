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

    // level columns (using array indices for column positions)
    for (let i = 0; i <= maxDepth; i++) {
      row[i] = i === level ? node.name : '';
    }

    row[maxDepth + 1] = node.address || '';
    row[maxDepth + 2] = node.city || '';

    result.push(row);

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
// Flatten the tree structure
const flatData = flattenTree(tree);

const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(flatData);

XLSX.utils.book_append_sheet(workbook, worksheet, 'Tree Data');
XLSX.writeFile(workbook, 'tree_output.xlsx');

console.log('Data structure:');
console.table(flatData);
