/**
 * compareAndInsertDateRange
 * WPS JS宏（JSA）函数：比较指定列中的日期是否在给定范围内，
 * 并将格式化后的日期段字符串写入目标列。
 *
 * @param {string} sheetName      - 源工作表名称，例如 "Sheet1"
 * @param {string} compareColumn  - 存放待比较日期的列字母，例如 "B"
 * @param {string} startDateStr   - 范围开始日期字符串，例如 "2026/6/18" 或 "2026-06-18"
 * @param {string} endDateStr     - 范围结束日期字符串，例如 "2026/6/19" 或 "2026-06-19"
 * @param {string} insertColumn   - 写入日期段结果的列字母，例如 "C"
 */
function compareAndInsertDateRange(sheetName, compareColumn, startDateStr, endDateStr, insertColumn) {

  // ── 1. 获取目标工作表 ──────────────────────────────────────────────────────
  // 通过工作表名称从当前工作簿中取得对应的 Sheet 对象
  var sheet = Application.ActiveWorkbook.Sheets(sheetName);

  // ── 2. 解析日期字符串为 Date 对象 ─────────────────────────────────────────
  /**
   * parseDate：将 "2026/6/18" 或 "2026-06-18" 格式的字符串解析为 Date 对象。
   * 使用正则表达式提取年、月、日，然后用 Date.UTC 构造，
   * 避免因本地时区偏移导致日期错位。
   * @param {string} str - 日期字符串
   * @returns {Date}
   */
  function parseDate(str) {
    // 用正则匹配 "YYYY/M/D" 或 "YYYY-MM-DD"（分隔符可为 "/" 或 "-"）
    var match = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (!match) {
      // 若格式不匹配则抛出错误，提示调用者检查输入
      throw new Error("日期格式无法识别，请使用 YYYY/M/D 或 YYYY-MM-DD：" + str);
    }
    var year  = parseInt(match[1], 10); // 年份（十进制整数）
    var month = parseInt(match[2], 10); // 月份（1-12）
    var day   = parseInt(match[3], 10); // 日（1-31）
    // Date.UTC 返回 UTC 毫秒时间戳，月份参数需 -1（0-11）
    return new Date(Date.UTC(year, month - 1, day));
  }

  var startDate = parseDate(startDateStr); // 范围开始日期的 Date 对象
  var endDate   = parseDate(endDateStr);   // 范围结束日期的 Date 对象

  // ── 3. 构造格式化后的日期段字符串（仅计算一次，供循环复用）────────────────
  /**
   * formatDateLabel：将 Date 对象格式化为 "M月D日" 中文样式。
   * @param {Date} d - Date 对象（UTC）
   * @returns {string} 如 "6月18日"
   */
  function formatDateLabel(d) {
    var m = d.getUTCMonth() + 1; // getUTCMonth() 返回 0-11，需 +1
    var day = d.getUTCDate();    // getUTCDate() 返回日
    return m + "月" + day + "日";
  }

  // 最终写入单元格的日期段标签，例如 "6月18日-6月19日"
  var dateRangeLabel = formatDateLabel(startDate) + "-" + formatDateLabel(endDate);

  // ── 4. 确定数据行范围 ──────────────────────────────────────────────────────
  // UsedRange 返回工作表中已使用区域；Rows.Count 给出已用行总数
  var lastRow = sheet.UsedRange.Rows.Count;

  // 若工作表只有标题行（第1行）或为空，则无需处理
  if (lastRow < 2) {
    return;
  }

  // ── 5. 批量读取比较列数据（性能优化：一次性读入内存二维数组）────────────────
  // 构造比较列的整列范围地址，例如 "B2:B100001"
  var compareStart  = compareColumn + "2";                    // 数据起始单元格，如 "B2"
  var compareEnd    = compareColumn + lastRow;                 // 数据结束单元格，如 "B100001"
  // Range.Value 返回二维数组（行×列），一次 API 调用取回所有值，性能远优于逐行读取
  var compareCells  = sheet.Range(compareStart + ":" + compareEnd);
  var compareValues = compareCells.Value; // 二维数组，compareValues[i][0] 对应第 i+2 行

  // ── 6. 准备结果数组（与读入数组等长，初始化为 null）─────────────────────────
  var rowCount = lastRow - 1; // 数据行数（不含标题行）
  // 构造与 compareValues 相同维度的二维数组，用于批量写回
  var insertValues = new Array(rowCount);
  for (var i = 0; i < rowCount; i++) {
    insertValues[i] = [null]; // 每行单列，默认空值（不覆盖原有内容的替代方案）
  }

  // ── 7. 遍历比较列，逐行判断日期是否在范围内 ──────────────────────────────
  for (var i = 0; i < rowCount; i++) {
    var cellValue = compareValues[i][0]; // 取出当前行的单元格值

    // 跳过空单元格，避免无效转换
    if (cellValue === null || cellValue === undefined || cellValue === "") {
      continue;
    }

    var cellDate; // 用于存放当前行日期的 Date 对象

    if (typeof cellValue === "number") {
      // WPS/Excel 日期在底层以序列数（自1900-01-00起的天数）存储。
      // 将序列数转为毫秒时间戳：(序列数 - 25569) 得到自 Unix 纪元的天数，再乘以每天毫秒数。
      // 25569 = Excel 序列数对应 1970-01-01 的天偏移量（含 Excel 1900 闰年 Bug 修正）
      cellDate = new Date((cellValue - 25569) * 86400000);
    } else if (typeof cellValue === "string") {
      // 单元格存为文本时，尝试用 parseDate 解析
      try {
        cellDate = parseDate(cellValue);
      } catch (e) {
        // 无法解析时跳过该行
        continue;
      }
    } else if (cellValue instanceof Date) {
      // 若 JSA 环境直接返回 Date 对象，则直接使用
      cellDate = cellValue;
    } else {
      // 其他未知类型，跳过
      continue;
    }

    // 取当天零点的 UTC 时间戳，消除时分秒对比较结果的干扰
    var cellTime = Date.UTC(
      cellDate.getUTCFullYear(),
      cellDate.getUTCMonth(),
      cellDate.getUTCDate()
    );
    var startTime = startDate.getTime(); // 开始日期的时间戳（毫秒）
    var endTime   = endDate.getTime();   // 结束日期的时间戳（毫秒）

    // 判断当前行日期是否 >= 开始日期 且 <= 结束日期
    if (cellTime >= startTime && cellTime <= endTime) {
      insertValues[i][0] = dateRangeLabel; // 写入日期段标签
    }
  }

  // ── 8. 批量写回结果（一次 API 调用写入所有行，避免逐行写入的性能损耗）──────
  var insertStart = insertColumn + "2";          // 写入起始单元格，如 "C2"
  var insertEnd   = insertColumn + lastRow;       // 写入结束单元格，如 "C100001"
  // 将二维数组直接赋值给 Range.Value，一次性写入所有结果
  sheet.Range(insertStart + ":" + insertEnd).Value = insertValues;
}

// ── 使用示例（可在 WPS 宏编辑器中直接调用）────────────────────────────────────
// compareAndInsertDateRange("Sheet1", "B", "2026/6/18", "2026/6/19", "C");
// compareAndInsertDateRange("Sheet1", "B", "2026-06-18", "2026-06-19", "C");
