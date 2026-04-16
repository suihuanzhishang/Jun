/**
 * 大模型模拟执行校验
 * 脱离 WPS/Excel 宿主环境，纯 Node.js 运行，
 * 验证 createRegexFromWildcard 通配符解析逻辑以及 fuzzyMatchCopy 核心匹配流程。
 *
 * 运行方式：node fuzzyMatchCopy.test.js
 */

// ─────────────────────────────────────────────────────────────────────────────
// 将被测函数从 fuzzyMatchCopy.js 中提取并内联，以便在 Node 环境独立运行
// ─────────────────────────────────────────────────────────────────────────────

/**
 * 将 glob 风格的通配符模式转为正则表达式
 * 支持：* ? [abc] [a-z] [^abc]
 */
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

                // 字符类内部无需额外转义，直接拼入正则
                regexStr += charClass;

                i = j + 1; // 跳过整个字符类片段
            } else {
                // 没有找到闭合的 ']'，把 '[' 当普通字面字符处理（转义后写入）
                regexStr += "\\[";
                i++;
            }

        } else {
            // 其余字符：转义正则元字符后作为字面字符写入
            regexStr += ch.replace(/[.+^${}()|[\]\\]/g, "\\$&");
            i++;
        }
    }

    // 用 ^ 和 $ 锚定，确保整串匹配；不区分大小写（i 标志）
    return new RegExp("^" + regexStr + "$", "i");
}

// ─────────────────────────────────────────────────────────────────────────────
// 轻量断言工具
// ─────────────────────────────────────────────────────────────────────────────

// 记录测试统计
let passed = 0; // 通过计数
let failed = 0; // 失败计数

/**
 * 断言函数
 * @param {Boolean} condition   期望为 true 的表达式
 * @param {String}  description 本条测试的说明
 */
function assert(condition, description) {
    if (condition) {
        // 测试通过：打印绿色 PASS
        console.log("  ✓ PASS  " + description);
        passed++;
    } else {
        // 测试失败：打印红色 FAIL
        console.error("  ✗ FAIL  " + description);
        failed++;
    }
}

/**
 * 断言：pattern 应匹配 value
 */
function assertMatch(pattern, value) {
    let re = createRegexFromWildcard(pattern); // 将通配符模式转为正则
    assert(
        re !== null && re.test(String(value)),  // 非 null 且测试通过
        "「" + pattern + "」 应匹配 「" + value + "」"
    );
}

/**
 * 断言：pattern 不应匹配 value
 */
function assertNoMatch(pattern, value) {
    let re = createRegexFromWildcard(pattern); // 将通配符模式转为正则
    assert(
        re !== null && !re.test(String(value)), // 非 null 且测试失败（即不匹配）
        "「" + pattern + "」 不应匹配 「" + value + "」"
    );
}

// ─────────────────────────────────────────────────────────────────────────────
// 测试套件
// ─────────────────────────────────────────────────────────────────────────────

// ── 套件 1：星号 (*) ──────────────────────────────────────────────────────────
console.log("\n=== 套件 1：星号 (*) — 匹配任意数量字符（含零个）===");

// '*' 在中间：file*.txt
assertMatch   ("file*.txt", "file1.txt");       // 匹配一个数字
assertMatch   ("file*.txt", "file_test.txt");   // 匹配多个字符
assertMatch   ("file*.txt", "file.txt");        // 匹配零个字符
assertNoMatch ("file*.txt", "file.csv");        // 后缀不同，不匹配

// '*' 在开头：*.txt
assertMatch   ("*.txt", "readme.txt");          // 任意前缀
assertMatch   ("*.txt", ".txt");                // 零字符前缀
assertNoMatch ("*.txt", "readme.csv");          // 后缀不同，不匹配

// '*' 在结尾：file*
assertMatch   ("file*", "file");                // 零字符后缀
assertMatch   ("file*", "file_anything");       // 任意后缀
assertNoMatch ("file*", "doc_anything");        // 前缀不同，不匹配

// 双星号：a**b（两个 .* 拼接，功能同单星号）
assertMatch   ("a**b", "ab");                   // 两个 * 均匹配零字符
assertMatch   ("a**b", "a123b");                // 两个 * 共同匹配中间部分

// ── 套件 2：问号 (?) ──────────────────────────────────────────────────────────
console.log("\n=== 套件 2：问号 (?) — 匹配单个字符 ===");

// 'file?.txt'
assertMatch   ("file?.txt", "file1.txt");       // 匹配一个数字
assertMatch   ("file?.txt", "fileA.txt");       // 匹配一个字母
assertNoMatch ("file?.txt", "file12.txt");      // 两个字符，不匹配
assertNoMatch ("file?.txt", "file.txt");        // 零个字符，不匹配

// 多个 '?'：a??b
assertMatch   ("a??b", "a12b");                 // 恰好两个字符
assertMatch   ("a??b", "aXYb");                 // 恰好两个字符
assertNoMatch ("a??b", "a1b");                  // 只有一个字符，不匹配
assertNoMatch ("a??b", "a123b");                // 三个字符，不匹配

// ── 套件 3：方括号 [abc] ──────────────────────────────────────────────────────
console.log("\n=== 套件 3：方括号 [abc] — 匹配括号内任意一个字符 ===");

// 'file[123].txt'
assertMatch   ("file[123].txt", "file1.txt");   // 匹配 '1'
assertMatch   ("file[123].txt", "file2.txt");   // 匹配 '2'
assertMatch   ("file[123].txt", "file3.txt");   // 匹配 '3'
assertNoMatch ("file[123].txt", "file4.txt");   // '4' 不在集合中
assertNoMatch ("file[123].txt", "file12.txt");  // 两个字符，不匹配

// 字符集含特殊字符 [a.b]（'.' 是字面字符）
assertMatch   ("a[a.b]c", "a.c");               // 匹配字面点号
assertMatch   ("a[a.b]c", "aac");               // 匹配 'a'
assertNoMatch ("a[a.b]c", "axc");               // 'x' 不在集合中

// ── 套件 4：范围 [a-z] / [0-9] / [A-Z] ──────────────────────────────────────
console.log("\n=== 套件 4：范围表示 [a-z] [0-9] [A-Z] ===");

// 'file[A-Z].txt'
assertMatch   ("file[A-Z].txt", "fileA.txt");   // 'A' 在范围内
assertMatch   ("file[A-Z].txt", "fileZ.txt");   // 'Z' 在范围内
// 注意：正则标志含 'i'（不区分大小写），故小写字母也能命中大写范围
assertMatch   ("file[A-Z].txt", "filea.txt");   // 不区分大小写，命中
assertNoMatch ("file[A-Z].txt", "file1.txt");   // '1' 不在字母范围

// '[0-9]+'：此处规则只有字符类，无量词
// 单独字符类只匹配一个字符
assertMatch   ("data[0-9]", "data5");           // '5' 在 0-9 范围内
assertNoMatch ("data[0-9]", "dataA");           // 'A' 不在数字范围

// ── 套件 5：反向匹配 [^...] ───────────────────────────────────────────────────
console.log("\n=== 套件 5：反向匹配 [^abc] — 匹配不在括号内的字符 ===");

// 'file[^123].txt'
assertMatch   ("file[^123].txt", "fileA.txt");  // 'A' 不在 {1,2,3} 中
assertMatch   ("file[^123].txt", "file4.txt");  // '4' 不在 {1,2,3} 中
assertNoMatch ("file[^123].txt", "file1.txt");  // '1' 在 {1,2,3} 中，不匹配
assertNoMatch ("file[^123].txt", "file2.txt");  // '2' 在 {1,2,3} 中，不匹配
assertNoMatch ("file[^123].txt", "file3.txt");  // '3' 在 {1,2,3} 中，不匹配

// '[^0-9]'：匹配非数字的单个字符
assertMatch   ("data[^0-9]", "dataX");          // 'X' 是非数字
assertNoMatch ("data[^0-9]", "data5");          // '5' 是数字，不匹配

// ── 套件 6：组合通配符 ────────────────────────────────────────────────────────
console.log("\n=== 套件 6：组合通配符 ===");

// '*[0-9].log'：任意前缀 + 一个数字 + .log
assertMatch   ("*[0-9].log", "app1.log");       // 前缀 'app'，数字 '1'
assertMatch   ("*[0-9].log", "server9.log");    // 前缀 'server'，数字 '9'
assertNoMatch ("*[0-9].log", "appA.log");       // 'A' 不是数字

// '?[a-z]*'：一个任意字符 + 一个小写字母 + 任意后缀
assertMatch   ("?[a-z]*", "Xhello");            // 'X' 匹配 '?'，'h' 匹配 [a-z]，'ello' 匹配 *
assertMatch   ("?[a-z]*", "1a");                // '1' 匹配 '?'，'a' 匹配 [a-z]，'' 匹配 *
assertNoMatch ("?[a-z]*", "1");                 // 缺少 [a-z] 部分

// '*[^aeiou]*'：中间含一个非元音字符
assertMatch   ("*[^aeiou]*", "hello");          // 'h' 匹配 [^aeiou]

// ── 套件 7：边界和特殊情况 ────────────────────────────────────────────────────
console.log("\n=== 套件 7：边界和特殊情况 ===");

// null / undefined 返回 null
assert(createRegexFromWildcard(null) === null, "null 模式 → 返回 null");
assert(createRegexFromWildcard(undefined) === null, "undefined 模式 → 返回 null");

// 纯字面字符串（无通配符）
assertMatch   ("hello.txt", "hello.txt");       // 精确匹配（全小写）
// 正则含 'i' 标志，所有通配符匹配均不区分大小写
assertMatch   ("hello.txt", "Hello.Txt");       // 大小写不敏感，应命中
assertMatch   ("hello.txt", "HELLO.TXT");       // 全大写，同样命中

// 含正则元字符的字面字符串（需被正确转义）
assertMatch   ("a.b", "a.b");                   // '.' 应被转义为字面点
assertNoMatch ("a.b", "axb");                   // 'x' 不等于字面 '.'

// 数值类型的规则（自动 toString）
assertMatch   (123, "123");                     // 数值 123 转字符串后精确匹配

// 未闭合的 '['（当普通字面字符处理）
assertMatch   ("a[b", "a[b");                   // '[' 被视为字面字符
assertNoMatch ("a[b", "axb");                   // 不匹配其他形式

// ─────────────────────────────────────────────────────────────────────────────
// 核心业务逻辑模拟（不依赖 WPS/Excel 对象）
// ─────────────────────────────────────────────────────────────────────────────
console.log("\n=== 核心业务逻辑模拟 ===");

/**
 * 在 Node 环境模拟 fuzzyMatchCopy 的核心匹配逻辑（第 3 步）
 * 直接传入内存数组，绕开 WPS API 调用
 *
 * @param {Array}  srcData       源数据二维数组（行 × 列，已去掉标题行）
 * @param {Number} matchColIndex 匹配列在 srcData 中的 0-based 索引
 * @param {Array}  matchData     规则二维数组（每行第 0 列为规则字符串）
 * @returns {Array} 命中行的数组
 */
function simulateCoreMatch(srcData, matchColIndex, matchData) {
    let resultRows = []; // 存放命中行

    // 遍历每条规则
    for (let i = 0; i < matchData.length; i++) {
        let criteria = matchData[i][0]; // 取规则字符串
        if (!criteria) continue;        // 跳过空规则

        let regex = createRegexFromWildcard(criteria); // 编译正则

        // 遍历每条源数据
        for (let j = 0; j < srcData.length; j++) {
            let cellValue = srcData[j][matchColIndex]; // 取匹配列的值
            if (cellValue == null) continue;           // 跳过空单元格

            if (regex.test(String(cellValue))) {
                resultRows.push(srcData[j]); // 命中，记录整行
            }
        }
    }
    return resultRows;
}

// ── 场景 A：源数据匹配 B 列（索引 1），规则含 * 和 [0-9] ────────────────────

// 模拟 Sheet1 数据（去除标题行后）
// 列：[A:ID, B:FileName, C:Size]
let srcDataA = [
    [1,  "file1.txt",       1024],  // 行 0
    [2,  "file_test.txt",   2048],  // 行 1
    [3,  "file.txt",        512],   // 行 2
    [4,  "document.pdf",    4096],  // 行 3
    [5,  "fileA.txt",       128],   // 行 4
    [6,  "report2025.xlsx", 8192],  // 行 5
    [7,  "data5.log",       256],   // 行 6
    [8,  "dataX.log",       300],   // 行 7
    [9,  "readme.md",       64],    // 行 8
    [10, "file12.txt",      999],   // 行 9（两字符后缀，不匹配 file?.txt）
];

// 规则列（匹配 B 列，索引 1）
let matchDataA = [
    ["file*.txt"],      // * 通配符
    ["*[0-9].log"],     // 组合：* + [0-9] + .log
];

let resultA = simulateCoreMatch(srcDataA, 1, matchDataA);

// 期望命中：
//   file*.txt  → file1.txt(0), file_test.txt(1), file.txt(2), fileA.txt(4), file12.txt(9)
//   *[0-9].log → data5.log(6)
// 合计 6 行（注意：fileA.txt 也匹配 file*.txt）
assert(resultA.length === 6, "场景 A：应命中 6 行，实际=" + resultA.length);

// 验证具体命中的文件名（取 B 列，索引 1）
let namesA = resultA.map(function(r) { return r[1]; });
assert(namesA.indexOf("file1.txt")    >= 0, "场景 A：命中 file1.txt");
assert(namesA.indexOf("file_test.txt") >= 0,"场景 A：命中 file_test.txt");
assert(namesA.indexOf("file.txt")     >= 0, "场景 A：命中 file.txt");
assert(namesA.indexOf("fileA.txt")    >= 0, "场景 A：命中 fileA.txt");
assert(namesA.indexOf("file12.txt")   >= 0, "场景 A：命中 file12.txt");
assert(namesA.indexOf("data5.log")    >= 0, "场景 A：命中 data5.log");
assert(namesA.indexOf("document.pdf") < 0,  "场景 A：未命中 document.pdf");
assert(namesA.indexOf("dataX.log")    < 0,  "场景 A：未命中 dataX.log");

// ── 场景 B：规则含 ? 和 [^...] ───────────────────────────────────────────────

let srcDataB = [
    ["fileA.txt"],   // 行 0：? 匹配 'A'
    ["file1.txt"],   // 行 1：? 匹配 '1'，但 [^123] 不匹配 '1'
    ["file2.txt"],   // 行 2：[^123] 不匹配 '2'
    ["file4.txt"],   // 行 3：[^123] 匹配 '4'
    ["file12.txt"],  // 行 4：? 不匹配（两个字符），[^123] 也不匹配
];

// 规则 1：file?.txt  只匹配单字符后缀
// 规则 2：file[^123].txt  匹配非 1/2/3 后缀
let matchDataB = [
    ["file?.txt"],
    ["file[^123].txt"],
];

let resultB = simulateCoreMatch(srcDataB, 0, matchDataB);

// file?.txt     → fileA.txt(0), file1.txt(1), file2.txt(2), file4.txt(3)
// file[^123].txt → fileA.txt(0), file4.txt(3)
// 去重后合计（此模拟不去重，两条规则都命中 fileA/file4，会各记一次）：
// fileA(0)+file1(1)+file2(2)+file4(3) + fileA(0)+file4(3) = 6 条
assert(resultB.length === 6, "场景 B：应命中 6 条（含重复），实际=" + resultB.length);

let namesB = resultB.map(function(r) { return r[0]; });
assert(namesB.filter(function(n){return n==="fileA.txt";}).length === 2, "场景 B：fileA.txt 命中 2 次");
assert(namesB.filter(function(n){return n==="file4.txt";}).length === 2, "场景 B：file4.txt 命中 2 次");
assert(namesB.indexOf("file12.txt") < 0, "场景 B：file12.txt 未命中");

// ── 场景 C：空规则跳过，null 值跳过 ─────────────────────────────────────────

let srcDataC = [
    [null,    "skip_null"],  // A 列为 null，若匹配 A 列应跳过
    ["hello", "keep"],
    ["",      "skip_empty_string"],
];

// 规则匹配 A 列（索引 0）
let matchDataC = [
    [null],    // 空规则，应被跳过
    ["hello"], // 字面匹配
];

let resultC = simulateCoreMatch(srcDataC, 0, matchDataC);

// null 规则跳过；null 单元格跳过；"" 不匹配 "hello"
// 只有 srcDataC[1] 命中
assert(resultC.length === 1, "场景 C：应命中 1 行，实际=" + resultC.length);
assert(resultC[0][1] === "keep", "场景 C：命中行的第 2 列为 'keep'");

// ─────────────────────────────────────────────────────────────────────────────
// 汇总结果
// ─────────────────────────────────────────────────────────────────────────────
console.log("\n══════════════════════════════════════════");
console.log("测试完成：通过 " + passed + " / " + (passed + failed) + " 条");
if (failed > 0) {
    console.error("失败 " + failed + " 条，请检查上方 FAIL 项目。");
    process.exit(1); // 以非零退出码标记失败，便于 CI 集成
} else {
    console.log("全部通过 ✓");
}
