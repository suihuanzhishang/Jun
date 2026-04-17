/**
 * 大模型模拟执行测试脚本（Node.js 环境）
 * 模拟 WPS JSA 的 Application / Sheet / Range 接口，
 * 验证 compareAndInsertDateRange 的核心逻辑。
 */

// ── 模拟数据集（对应截图中的 Sheet1 A/B 列）─────────────────────────────────
// 第1行为标题，第2-8行为数据
var mockData = [
  ["姓名", "日期"],          // row 1 - header
  ["胡骏",  "2026-06-18"],   // row 2 → 在范围内
  ["SB俊", "2026-06-18"],   // row 3 → 在范围内
  ["H俊",  "2026-06-19"],   // row 4 → 在范围内
  ["回头俊","2026-06-20"],   // row 5 → 超出范围
  ["好俊",  "2026-06-18"],   // row 6 → 在范围内
  ["Hello君","2026-06-19"],  // row 7 → 在范围内
  ["物料均","2026-06-20"],   // row 8 → 超出范围
];

// ── 构造模拟 Sheet 对象 ────────────────────────────────────────────────────
var mockSheet = {
  // UsedRange.Rows.Count 返回已用行数
  UsedRange: { Rows: { Count: mockData.length } },

  // 模拟 Range(addr).Value 的读/写
  _store: {},

  Range: function(addr) {
    var self = this;
    return {
      // 读取时解析地址范围，返回 WPS JSA 风格的 1-based 二维数组
      get Value() {
        return self._readRange(addr);
      },
      // 写入时接收标准 0-based 二维数组
      set Value(arr) {
        self._writeRange(addr, arr);
      }
    };
  },

  // 解析 "B2:B8" 形式的地址，提取列字母和行号范围
  _parseAddr: function(addr) {
    var m = addr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i);
    if (!m) throw new Error("地址格式错误: " + addr);
    return { col: m[1].toUpperCase(), r1: parseInt(m[2]), r2: parseInt(m[4]) };
  },

  // 列字母转0-based索引（仅支持单字母列 A-Z）
  _colIndex: function(letter) { return letter.toUpperCase().charCodeAt(0) - 65; },

  // 读取指定列范围，返回 WPS JSA 风格的 1-based 二维数组（模拟真实环境）
  // 即：result[1][1] 是第一个单元格的值，result[0] 为 undefined
  _readRange: function(addr) {
    var p = this._parseAddr(addr);
    var colIdx = this._colIndex(p.col);
    var rowSpan = p.r2 - p.r1 + 1;

    if (rowSpan === 1) {
      // 单行时直接返回标量值（WPS JSA 行为）
      var rowData = mockData[p.r1 - 1];
      return rowData ? rowData[colIdx] : null;
    }

    // 多行：构造 1-based 二维数组（下标 1..rowSpan）
    // 用稀疏方式：result[i] 表示第 i 行（i 从 1 开始），result[0] 留为 undefined
    var result = new Array(rowSpan + 1); // 长度 rowSpan+1，index 0 空置
    for (var r = p.r1; r <= p.r2; r++) {
      var rowData = mockData[r - 1];
      var val = rowData ? rowData[colIdx] : null;
      var idx = r - p.r1 + 1;          // 1-based 行索引
      result[idx] = {};                 // 每行也是 1-based 对象
      result[idx][1] = val;             // 列索引从 1 开始
    }
    return result;
  },

  // 写入指定列范围，接收标准 0-based 二维数组（JS 惯例）
  _writeRange: function(addr, arr) {
    var p = this._parseAddr(addr);
    var colIdx = this._colIndex(p.col);
    for (var i = 0; i < arr.length; i++) {
      var r = p.r1 + i;
      while (mockData.length < r) mockData.push([]);
      if (!mockData[r - 1]) mockData[r - 1] = [];
      mockData[r - 1][colIdx] = arr[i][0];
    }
  }
};

// ── 模拟 Application 对象 ──────────────────────────────────────────────────
var Application = {
  ActiveWorkbook: {
    Sheets: function(name) {
      if (name === "Sheet1") return mockSheet;
      throw new Error("找不到工作表: " + name);
    }
  }
};

// ── 粘贴主函数（与 compareAndInsertDateRange.js 完全一致）──────────────────────
function compareAndInsertDateRange(sheetName, compareColumn, startDateStr, endDateStr, insertColumn) {
  var sheet = Application.ActiveWorkbook.Sheets(sheetName);

  function parseDate(str) {
    var match = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (!match) throw new Error("日期格式无法识别：" + str);
    var year  = parseInt(match[1], 10);
    var month = parseInt(match[2], 10);
    var day   = parseInt(match[3], 10);
    return new Date(Date.UTC(year, month - 1, day));
  }

  var startDate = parseDate(startDateStr);
  var endDate   = parseDate(endDateStr);

  function formatDateLabel(d) {
    var m = d.getUTCMonth() + 1;
    var day = d.getUTCDate();
    return m + "月" + day + "日";
  }

  var dateRangeLabel = formatDateLabel(startDate) + "-" + formatDateLabel(endDate);

  var lastRow = sheet.UsedRange.Rows.Count;
  if (lastRow < 2) return;

  var rowCount     = lastRow - 1;
  var compareStart = compareColumn + "2";
  var compareEnd   = compareColumn + lastRow;
  var rawValues    = sheet.Range(compareStart + ":" + compareEnd).Value;

  // 规范化为 0-based 一维数组（修复 WPS JSA 1-based 二维数组导致的崩溃）
  var compareArr = new Array(rowCount);
  if (rowCount === 1) {
    compareArr[0] = rawValues;
  } else {
    for (var r = 0; r < rowCount; r++) {
      var rowData = rawValues[r + 1];
      if (rowData === null || rowData === undefined) {
        compareArr[r] = null;
      } else if (rowData instanceof Date || typeof rowData !== "object") {
        compareArr[r] = rowData;
      } else {
        compareArr[r] = rowData[1] !== undefined ? rowData[1] : rowData[0];
      }
    }
  }

  var insertValues = new Array(rowCount);
  for (var i = 0; i < rowCount; i++) {
    insertValues[i] = [null];
  }

  var startTime = startDate.getTime();
  var endTime   = endDate.getTime();

  for (var i = 0; i < rowCount; i++) {
    var cellValue = compareArr[i];
    if (cellValue === null || cellValue === undefined || cellValue === "") continue;

    var cellDate;
    if (typeof cellValue === "number") {
      cellDate = new Date((cellValue - 25569) * 86400000);
    } else if (typeof cellValue === "string") {
      try { cellDate = parseDate(cellValue); } catch (e) { continue; }
    } else if (cellValue instanceof Date) {
      cellDate = cellValue;
    } else {
      continue;
    }

    var cellTime = Date.UTC(cellDate.getUTCFullYear(), cellDate.getUTCMonth(), cellDate.getUTCDate());

    if (cellTime >= startTime && cellTime <= endTime) {
      insertValues[i][0] = dateRangeLabel;
    }
  }

  sheet.Range(insertColumn + "2" + ":" + insertColumn + lastRow).Value = insertValues;
}

// ── 执行测试 ──────────────────────────────────────────────────────────────
console.log("=== 测试1：斜杠格式日期 2026/6/18 - 2026/6/19 ===");
compareAndInsertDateRange("Sheet1", "B", "2026/6/18", "2026/6/19", "C");

// 打印结果
var header = ["姓名", "日期", "日期段"];
console.log(header.join("\t\t"));
for (var r = 1; r < mockData.length; r++) {
  var row = mockData[r];
  console.log((row[0] || "").padEnd(8) + "\t" + (row[1] || "").padEnd(12) + "\t" + (row[2] || "（空）"));
}

// ── 预期结果断言 ──────────────────────────────────────────────────────────
var expected = [
  "6月18日-6月19日",  // row2 2026-06-18 ✓
  "6月18日-6月19日",  // row3 2026-06-18 ✓
  "6月18日-6月19日",  // row4 2026-06-19 ✓
  null,               // row5 2026-06-20 ✗
  "6月18日-6月19日",  // row6 2026-06-18 ✓
  "6月18日-6月19日",  // row7 2026-06-19 ✓
  null,               // row8 2026-06-20 ✗
];

var pass = true;
for (var i = 0; i < expected.length; i++) {
  var actual = mockData[i + 1][2] || null;
  if (actual !== expected[i]) {
    console.error("FAIL 第" + (i + 2) + "行: 期望=" + expected[i] + " 实际=" + actual);
    pass = false;
  }
}
if (pass) console.log("\n全部断言通过 ✓");

// ── 测试2：连字符格式日期 ─────────────────────────────────────────────────
console.log("\n=== 测试2：连字符格式日期 2026-06-20 - 2026-06-20（仅20日）===");
// 重置 C 列
for (var r = 1; r < mockData.length; r++) mockData[r][2] = undefined;
compareAndInsertDateRange("Sheet1", "B", "2026-06-20", "2026-06-20", "C");
for (var r = 1; r < mockData.length; r++) {
  var row = mockData[r];
  console.log((row[0] || "").padEnd(8) + "\t" + (row[1] || "").padEnd(12) + "\t" + (row[2] || "（空）"));
}

var expected2 = [null, null, null, "6月20日-6月20日", null, null, "6月20日-6月20日"];
var pass2 = true;
for (var i = 0; i < expected2.length; i++) {
  var actual = mockData[i + 1][2] || null;
  if (actual !== expected2[i]) {
    console.error("FAIL 第" + (i + 2) + "行: 期望=" + expected2[i] + " 实际=" + actual);
    pass2 = false;
  }
}
if (pass2) console.log("\n全部断言通过 ✓");
