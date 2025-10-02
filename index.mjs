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
    rows.push({ row, level: level > 1 ? 1 : level });

    if (node.children) {
      const parsedChildren = flatten(node.children, level + keys.length);
      rows.push(...parsedChildren.rows);
      headers.push(...parsedChildren.headers);
    }
  }

  return { rows, headers: headers.filter(Boolean).flat() };
};

const result = flatten(tree);

console.log(
  [...new Set(result.headers.slice(1, result.headers.length))].map(
    (h) => h.split('.')[1],
  ),
);
console.log(result.rows);
