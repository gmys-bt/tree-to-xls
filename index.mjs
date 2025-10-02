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

const flatten = (treeStructure, level = 0) => {
  const rows = [];
  const headers = [];

  for (const node of treeStructure) {
    const row = [];
    const baseLevel = level > 0 ? level - 1 : level;

    const keys = Object.keys(node).filter((k) => k !== 'children');

    headers[level] = keys;

    keys.forEach((k, i) => {
      row[baseLevel + i] = node[k];
    });

    rows.push(row);

    if (node.children) {
      const parsedChildren = flatten(node.children, level + keys.length);
      rows.push(...parsedChildren.rows);
      headers.push(...parsedChildren.headers);
    }
  }

  return { rows, headers: headers.filter(Boolean).flat() };
};

const result = flatten(tree);

console.log(result.headers.slice(1, result.headers.length));
console.log(result.rows);
