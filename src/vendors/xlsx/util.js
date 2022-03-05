
// 深度扁平化routes
export function flatten(routes) {
  return routes.flatMap(route => route.children ? [ route, ...flatten(route.children) ] : [ route ]);
}

// 转化style对象为style字符串
export function toStyleString(style) {
  return [...style].map(key => `${key}: ${style[key]}`).join(';')
}

// 修复矩阵的宽度
export function fixMatrix(data, colLen) {
  for (const row of data) {
    for (let j = 0; j < colLen; j++) {
      if (!row[j]) {
        row[j] = '';
      }
    }
  }
  return data;
}

// 首字母大写
export function captain(str) {
  return `${str.charAt(0).toUpperCase()}${str.slice(1)}`;
}

// 连字符转驼峰
export function camelCase(str) {
  return str.split('-').map((part, index) => {
    if (index !== 0) {
      return captain(part);
    } else {
      return part;
    }
  }).join('');
}
