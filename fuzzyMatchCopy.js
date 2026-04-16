/**
 * 优化版：支持多次调用追加数据，并修复列偏移问题
 * 通配符支持：
 *   *        匹配任意数量字符（含零个）
 *   ?        匹配单个任意字符
 *   [abc]    匹配括号内任意一个字符
 *   [a-z]    匹配指定范围内的任意一个字符
 *   [^abc]   反向匹配：匹配不在括号内的字符
 *
 * @param {String}  sourceSheetName    源数据工作表名，如 "Sheet1"
 * @param {String}  sourceDataColIndex 源数据中用于匹配的列，如 "A" 或 "B"
 * @param {String}  matchSheetName     匹配规则工作表名，如 "Sheet2"
 * @param {String}  matchDataColIndex  匹配规则所在列，如 "E"
 * @param {String}  resultSheetName    结果写入的工作表名，如 "Sheet3"
 * @param {Boolean} isAppend           可选，true=追加到末尾，false/undefined=清空重写，默认 false
 */
function fuzzyMatchCopy(
    sourceSheetName,
    sourceDataColIndex,
    matchSheetName,
    matchDataColIndex,
    resultSheetName,
    isAppend
) {

    // ─────────────────────────────────────────────────────────────────────────
    // 工具函数：将 glob 风格的通配符模式转为正则表达式
    // 支持：* ? [abc] [a-z] [^abc]
    // ─────────────────────────────────────────────────────────────────────────
    function createRegexFromWildcard(pattern) {

        // 若模式为 null / undefined，返回 null 表示跳过
        if (pattern == null) return null;

        // 统一转为字符串，防止数值类型的规则出错
        let str = String(pattern);

        // 用于拼接最终正则字符串的缓冲区
        let regexStr = "";

        // 当前扫描位置
        let i = 0;

        // 逐字符解析，将 glob 语法转换为对应的正则片段
        while (i < str.length) {
            let ch = str[i]; // 取当前字符

            if (ch === "*") {
                // '*' → 匹配任意数量（含零个）的任意字符
                regexStr += ".*";
                i++;

            } else if (ch === "?") {
                // '?' → 精确匹配一个任意字符
                regexStr += ".";
                i++;

            } else if (ch === "[") {
                // '[' 开头：进入字符类解析，直到找到匹配的 ']'
                let j = i + 1; // j 从 '[' 之后开始扫描

                // 处理可选的 '^'（反向匹配）
                if (j < str.length && str[j] === "^") {
                    j++; // 跳过 '^'，稍后一并写入
                }

                // 处理紧跟在 '[' 或 '[^' 之后的 ']'（此时 ']' 是字面字符，不结束字符类）
                if (j < str.length && str[j] === "]") {
                    j++; // 跳过首位的字面 ']'
                }

                // 继续向后扫描，找到真正结束字符类的 ']'
                while (j < str.length && str[j] !== "]") {
                    j++;
                }

                if (j < str.length) {
                    // 找到了闭合的 ']'，提取完整的字符类原文（含两端括号）
                    let charClass = str.slice(i, j + 1); // 例如 "[abc]" "[A-Z]" "[^0-9]"

                    // 字符类内部无需额外转义（-, ^, ] 在 glob 语义下已正确书写），
                    // 直接拼入正则即可
                    regexStr += charClass;

                    i = j + 1; // 跳过整个字符类片段
                } else {
                    // 没有找到闭合的 ']'，把 '[' 当普通字面字符处理（转义后写入）
                    regexStr += "\\[";
                    i++;
                }

            } else {
                // 其余字符：转义正则元字符后作为字面字符写入
                // 需要转义的元字符集：. + ? ^ $ { } ( ) | [ ] \
                // 注意：'*' '?' '[' 已在上方分支处理，此处不会出现
                regexStr += ch.replace(/[.+^${}()|[\]\\]/g, "\\$&");
                i++;
            }
        }

        // 用 ^ 和 $ 锚定，确保整串匹配（而不是包含匹配）；不区分大小写（i 标志）
        return new RegExp("^" + regexStr + "$", "i");
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 第 1 步：读取源数据工作表
    // ─────────────────────────────────────────────────────────────────────────

    let app = Application; // 获取宿主应用对象（WPS / Excel 等）

    // 按名称取源数据工作表
    let srcSheet = app.Sheets.Item(sourceSheetName);

    // 以 A 列末行确定数据总行数，比 UsedRange 更稳健
    let srcLastRow = srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row;

    // 若没有数据行（仅标题或完全空表），直接退出
    if (srcLastRow < 2) return;

    // 获取已用区域，用于确认数据的最大列数
    let srcUsedRange = srcSheet.UsedRange;

    // 最大列数，用于 Resize，保证读取完整的行数据
    let srcMaxCol = srcUsedRange.Columns.Count;

    // 从 A2 开始读取整张表的数据区域（跳过第 1 行标题）
    // 始终从 A 列开始，使得 srcData[row][col] 的索引与列号一一对应
    let srcRange = srcSheet.Range("A2").Resize(srcLastRow - 1, srcMaxCol);

    // 一次性取出二维数组，避免在循环中多次访问单元格（性能关键）
    let srcData = srcRange.Value2;

    // 计算匹配列在 srcData 二维数组中的列下标（0-based）
    // 例如：sourceDataColIndex="B" → Column 属性为 2 → 数组索引为 2-1=1
    let matchColIndex = srcSheet.Range(sourceDataColIndex + "1").Column - 1;

    // ─────────────────────────────────────────────────────────────────────────
    // 第 2 步：读取匹配规则工作表
    // ─────────────────────────────────────────────────────────────────────────

    // 按名称取匹配规则工作表
    let matchSheet = app.Sheets.Item(matchSheetName);

    // 以规则列确定最后一行（即规则数据的末行）
    let matchLastRow = matchSheet.Cells(matchSheet.Rows.Count, matchDataColIndex).End(xlUp).Row;

    // 若没有规则行，直接退出
    if (matchLastRow < 2) return;

    // 读取从第 2 行到末行的规则列数据
    let matchRange = matchSheet.Range(
        matchDataColIndex + "2:" + matchDataColIndex + matchLastRow
    );

    // 取出规则值（二维数组）
    let matchData = matchRange.Value2;

    // 当规则只有一个单元格时，Value2 返回标量而非数组，需手动包装
    if (!Array.isArray(matchData)) matchData = [[matchData]];

    // ─────────────────────────────────────────────────────────────────────────
    // 第 3 步：核心匹配循环
    // ─────────────────────────────────────────────────────────────────────────

    // 存放所有命中的完整行数据，最终批量写入结果表
    let resultRows = [];

    // 遍历每条匹配规则
    for (let i = 0; i < matchData.length; i++) {

        // 取当前规则字符串（matchData 是二维数组，列下标固定为 0）
        let criteria = matchData[i][0];

        // 空规则直接跳过，避免把所有行都匹配上
        if (!criteria) continue;

        // 将规则字符串编译为正则表达式（含通配符解析）
        let regex = createRegexFromWildcard(criteria);

        // 遍历源数据的每一行
        for (let j = 0; j < srcData.length; j++) {

            // 从当前行中取出匹配列的单元格值
            let cellValue = srcData[j][matchColIndex];

            // 跳过空单元格
            if (cellValue == null) continue;

            // 将单元格值转为字符串后与正则匹配
            if (regex.test(String(cellValue))) {
                // 命中：将整行数据推入结果集
                resultRows.push(srcData[j]);
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 第 4 步：准备结果工作表
    // ─────────────────────────────────────────────────────────────────────────

    let resSheet = null;

    try {
        // 尝试按名称获取已存在的结果表
        resSheet = app.Sheets.Item(resultSheetName);
    } catch (e) {
        // 不存在则在末尾新建并命名
        resSheet = app.Sheets.Add(null, app.Sheets.Item(app.Sheets.Count));
        resSheet.Name = resultSheetName;
    }

    // 确定写入起始行
    let startRow = 2; // 默认从第 2 行开始（第 1 行留给标题）

    if (isAppend) {
        // 追加模式：找到 A 列当前最后一行，从其下一行写入
        let lastRow = resSheet.Cells(resSheet.Rows.Count, 1).End(xlUp).Row;
        startRow = lastRow + 1;

        // 若结果表完全空（lastRow=1 代表第 1 行无数据），至少从第 2 行开始
        if (startRow < 2) startRow = 2;

    } else {
        // 覆盖模式：清除第 2 行及以下的旧数据，保留第 1 行标题
        let resLastRow = resSheet.UsedRange.Row + resSheet.UsedRange.Rows.Count - 1;
        if (resLastRow >= 2) {
            // ClearContents 只清除值，不影响格式
            resSheet.Range("2:" + resLastRow).ClearContents();
        }
        startRow = 2; // 重置写入起始行
    }

    // ─────────────────────────────────────────────────────────────────────────
    // 第 5 步：批量写入结果
    // ─────────────────────────────────────────────────────────────────────────

    if (resultRows.length > 0) {
        // 根据结果行数和列数定位写入区域
        let resTargetRange = resSheet
            .Range("A" + startRow)
            .Resize(resultRows.length, resultRows[0].length);

        // 一次性赋值整个二维数组（比逐行写入快得多）
        resTargetRange.Value2 = resultRows;

        // 输出执行日志
        console.log(
            "已写入 " + resultRows.length +
            " 行数据到 " + resultSheetName +
            " (追加模式: " + isAppend + ")"
        );
    } else {
        // 本次无命中，记录日志
        console.log("本次无匹配数据。");
    }
}
