/**
 * Excel 函数目录数据 & 搜索
 *
 * 内置常用 Excel 函数元数据，供 formula.search 工具使用。
 */

interface ExcelFunction {
  name: string;
  category: string;
  description: string;
  syntax: string;
  example: string;
}

/** 常用 Excel 函数数据库 */
export const EXCEL_FUNCTIONS: ExcelFunction[] = [
  // 数学与统计
  { name: "SUM", category: "数学", description: "对一组数值求和", syntax: "SUM(number1, [number2], ...)", example: "=SUM(A1:A10)" },
  { name: "AVERAGE", category: "统计", description: "计算一组数值的算术平均值", syntax: "AVERAGE(number1, [number2], ...)", example: "=AVERAGE(B1:B20)" },
  { name: "COUNT", category: "统计", description: "计算包含数字的单元格数量", syntax: "COUNT(value1, [value2], ...)", example: "=COUNT(A1:A100)" },
  { name: "COUNTA", category: "统计", description: "计算非空单元格数量", syntax: "COUNTA(value1, [value2], ...)", example: "=COUNTA(A1:A100)" },
  { name: "COUNTIF", category: "统计", description: "统计满足条件的单元格数量", syntax: "COUNTIF(range, criteria)", example: '=COUNTIF(A1:A10,">50")' },
  { name: "COUNTIFS", category: "统计", description: "统计满足多个条件的单元格数量", syntax: "COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)", example: '=COUNTIFS(A1:A10,">50",B1:B10,"<100")' },
  { name: "SUMIF", category: "数学", description: "对满足条件的单元格求和", syntax: "SUMIF(range, criteria, [sum_range])", example: '=SUMIF(A1:A10,">50",B1:B10)' },
  { name: "SUMIFS", category: "数学", description: "对满足多个条件的单元格求和", syntax: "SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)", example: '=SUMIFS(C1:C10,A1:A10,">50",B1:B10,"<100")' },
  { name: "MAX", category: "统计", description: "返回一组数值中的最大值", syntax: "MAX(number1, [number2], ...)", example: "=MAX(A1:A100)" },
  { name: "MIN", category: "统计", description: "返回一组数值中的最小值", syntax: "MIN(number1, [number2], ...)", example: "=MIN(A1:A100)" },
  { name: "MEDIAN", category: "统计", description: "返回一组数值的中位数", syntax: "MEDIAN(number1, [number2], ...)", example: "=MEDIAN(A1:A100)" },
  { name: "STDEV", category: "统计", description: "基于样本估算标准偏差", syntax: "STDEV(number1, [number2], ...)", example: "=STDEV(A1:A100)" },
  { name: "ROUND", category: "数学", description: "将数字四舍五入到指定位数", syntax: "ROUND(number, num_digits)", example: "=ROUND(3.14159, 2)" },
  { name: "ROUNDUP", category: "数学", description: "将数字向上（远离零）舍入", syntax: "ROUNDUP(number, num_digits)", example: "=ROUNDUP(3.141, 2)" },
  { name: "ROUNDDOWN", category: "数学", description: "将数字向下（朝向零）舍入", syntax: "ROUNDDOWN(number, num_digits)", example: "=ROUNDDOWN(3.149, 2)" },
  { name: "INT", category: "数学", description: "将数字向下取整为最接近的整数", syntax: "INT(number)", example: "=INT(8.9)" },
  { name: "MOD", category: "数学", description: "返回两数相除的余数", syntax: "MOD(number, divisor)", example: "=MOD(10, 3)" },
  { name: "ABS", category: "数学", description: "返回数字的绝对值", syntax: "ABS(number)", example: "=ABS(-5)" },
  { name: "POWER", category: "数学", description: "返回数字的乘幂结果", syntax: "POWER(number, power)", example: "=POWER(2, 10)" },
  { name: "SQRT", category: "数学", description: "返回正平方根", syntax: "SQRT(number)", example: "=SQRT(16)" },
  { name: "PRODUCT", category: "数学", description: "计算所有参数的乘积", syntax: "PRODUCT(number1, [number2], ...)", example: "=PRODUCT(A1:A5)" },
  { name: "RAND", category: "数学", description: "返回0到1之间的随机数", syntax: "RAND()", example: "=RAND()" },
  { name: "RANDBETWEEN", category: "数学", description: "返回指定范围内的随机整数", syntax: "RANDBETWEEN(bottom, top)", example: "=RANDBETWEEN(1, 100)" },
  // 查找与引用
  { name: "VLOOKUP", category: "查找", description: "在表格首列查找值并返回同行指定列的值", syntax: "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])", example: "=VLOOKUP(D2, A1:B100, 2, FALSE)" },
  { name: "HLOOKUP", category: "查找", description: "在表格首行查找值并返回同列指定行的值", syntax: "HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])", example: "=HLOOKUP(A1, B1:E10, 3, FALSE)" },
  { name: "XLOOKUP", category: "查找", description: "在范围或数组中搜索并返回对应项", syntax: "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])", example: "=XLOOKUP(E2, A1:A100, B1:B100)" },
  { name: "INDEX", category: "查找", description: "返回表格或范围中的值或引用", syntax: "INDEX(array, row_num, [column_num])", example: "=INDEX(A1:C10, 3, 2)" },
  { name: "MATCH", category: "查找", description: "在范围中搜索指定项并返回其相对位置", syntax: "MATCH(lookup_value, lookup_array, [match_type])", example: '=MATCH("apple", A1:A100, 0)' },
  { name: "OFFSET", category: "查找", description: "从给定引用偏移指定行列返回引用", syntax: "OFFSET(reference, rows, cols, [height], [width])", example: "=OFFSET(A1, 5, 2)" },
  { name: "INDIRECT", category: "查找", description: "返回文本字符串指定的引用", syntax: "INDIRECT(ref_text, [a1])", example: '=INDIRECT("A" & B1)' },
  { name: "CHOOSE", category: "查找", description: "从值列表中选择一个值", syntax: "CHOOSE(index_num, value1, [value2], ...)", example: "=CHOOSE(2, A1, B1, C1)" },
  // 文本
  { name: "LEFT", category: "文本", description: "返回文本字符串最左边的字符", syntax: "LEFT(text, [num_chars])", example: '=LEFT(A1, 3)' },
  { name: "RIGHT", category: "文本", description: "返回文本字符串最右边的字符", syntax: "RIGHT(text, [num_chars])", example: '=RIGHT(A1, 4)' },
  { name: "MID", category: "文本", description: "返回文本字符串中从指定位置开始的特定数目字符", syntax: "MID(text, start_num, num_chars)", example: '=MID(A1, 2, 5)' },
  { name: "LEN", category: "文本", description: "返回文本字符串的字符数", syntax: "LEN(text)", example: "=LEN(A1)" },
  { name: "FIND", category: "文本", description: "在文本中查找子串的位置（区分大小写）", syntax: "FIND(find_text, within_text, [start_num])", example: '=FIND("@", A1)' },
  { name: "SEARCH", category: "文本", description: "在文本中查找子串的位置（不区分大小写）", syntax: "SEARCH(find_text, within_text, [start_num])", example: '=SEARCH("hello", A1)' },
  { name: "SUBSTITUTE", category: "文本", description: "将文本中的旧文本替换为新文本", syntax: "SUBSTITUTE(text, old_text, new_text, [instance_num])", example: '=SUBSTITUTE(A1, "old", "new")' },
  { name: "REPLACE", category: "文本", description: "替换文本中指定位置的字符", syntax: "REPLACE(old_text, start_num, num_chars, new_text)", example: '=REPLACE(A1, 1, 3, "NEW")' },
  { name: "TRIM", category: "文本", description: "去除文本前后空格", syntax: "TRIM(text)", example: "=TRIM(A1)" },
  { name: "UPPER", category: "文本", description: "将文本转换为大写", syntax: "UPPER(text)", example: "=UPPER(A1)" },
  { name: "LOWER", category: "文本", description: "将文本转换为小写", syntax: "LOWER(text)", example: "=LOWER(A1)" },
  { name: "PROPER", category: "文本", description: "将文本每个单词首字母大写", syntax: "PROPER(text)", example: "=PROPER(A1)" },
  { name: "CONCATENATE", category: "文本", description: "将多个文本字符串合并为一个", syntax: "CONCATENATE(text1, [text2], ...)", example: '=CONCATENATE(A1, " ", B1)' },
  { name: "TEXT", category: "文本", description: "将数值转换为指定格式的文本", syntax: "TEXT(value, format_text)", example: '=TEXT(A1, "0.00")' },
  { name: "VALUE", category: "文本", description: "将表示数字的文本转换为数字", syntax: "VALUE(text)", example: '=VALUE("123")' },
  { name: "TEXTJOIN", category: "文本", description: "用分隔符连接文本", syntax: "TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)", example: '=TEXTJOIN(",", TRUE, A1:A10)' },
  // 逻辑
  { name: "IF", category: "逻辑", description: "根据条件返回不同值", syntax: "IF(logical_test, value_if_true, [value_if_false])", example: '=IF(A1>50, "Pass", "Fail")' },
  { name: "IFS", category: "逻辑", description: "检查多个条件并返回第一个为真的对应值", syntax: "IFS(logical_test1, value1, [logical_test2, value2], ...)", example: '=IFS(A1>=90,"A",A1>=80,"B",A1>=70,"C")' },
  { name: "AND", category: "逻辑", description: "所有参数为 TRUE 时返回 TRUE", syntax: "AND(logical1, [logical2], ...)", example: "=AND(A1>0, B1>0)" },
  { name: "OR", category: "逻辑", description: "任一参数为 TRUE 时返回 TRUE", syntax: "OR(logical1, [logical2], ...)", example: "=OR(A1>0, B1>0)" },
  { name: "NOT", category: "逻辑", description: "反转参数的逻辑值", syntax: "NOT(logical)", example: "=NOT(A1>0)" },
  { name: "IFERROR", category: "逻辑", description: "如果公式出错则返回指定值", syntax: "IFERROR(value, value_if_error)", example: '=IFERROR(VLOOKUP(A1,B:C,2,0), "Not found")' },
  { name: "SWITCH", category: "逻辑", description: "根据表达式值返回对应结果", syntax: "SWITCH(expression, value1, result1, [value2, result2], ..., [default])", example: '=SWITCH(A1, 1, "One", 2, "Two", "Other")' },
  // 日期与时间
  { name: "TODAY", category: "日期", description: "返回当前日期", syntax: "TODAY()", example: "=TODAY()" },
  { name: "NOW", category: "日期", description: "返回当前日期和时间", syntax: "NOW()", example: "=NOW()" },
  { name: "DATE", category: "日期", description: "从年月日创建日期", syntax: "DATE(year, month, day)", example: "=DATE(2024, 1, 15)" },
  { name: "YEAR", category: "日期", description: "提取日期中的年份", syntax: "YEAR(date)", example: "=YEAR(A1)" },
  { name: "MONTH", category: "日期", description: "提取日期中的月份", syntax: "MONTH(date)", example: "=MONTH(A1)" },
  { name: "DAY", category: "日期", description: "提取日期中的日", syntax: "DAY(date)", example: "=DAY(A1)" },
  { name: "DATEDIF", category: "日期", description: "计算两个日期之间的差", syntax: "DATEDIF(start_date, end_date, unit)", example: '=DATEDIF(A1, B1, "Y")' },
  { name: "EOMONTH", category: "日期", description: "返回指定月数之前或之后的月末日期", syntax: "EOMONTH(start_date, months)", example: "=EOMONTH(A1, 0)" },
  { name: "EDATE", category: "日期", description: "返回指定月数之前或之后的日期", syntax: "EDATE(start_date, months)", example: "=EDATE(A1, 3)" },
  { name: "WEEKDAY", category: "日期", description: "返回日期对应的星期几", syntax: "WEEKDAY(serial_number, [return_type])", example: "=WEEKDAY(A1, 2)" },
  { name: "NETWORKDAYS", category: "日期", description: "计算两个日期之间的工作日数", syntax: "NETWORKDAYS(start_date, end_date, [holidays])", example: "=NETWORKDAYS(A1, B1)" },
  // 信息
  { name: "ISBLANK", category: "信息", description: "检查单元格是否为空", syntax: "ISBLANK(value)", example: "=ISBLANK(A1)" },
  { name: "ISERROR", category: "信息", description: "检查值是否为错误类型", syntax: "ISERROR(value)", example: "=ISERROR(A1/B1)" },
  { name: "ISNUMBER", category: "信息", description: "检查值是否为数字", syntax: "ISNUMBER(value)", example: "=ISNUMBER(A1)" },
  { name: "ISTEXT", category: "信息", description: "检查值是否为文本", syntax: "ISTEXT(value)", example: "=ISTEXT(A1)" },
  { name: "TYPE", category: "信息", description: "返回值的类型编号", syntax: "TYPE(value)", example: "=TYPE(A1)" },
  // 财务
  { name: "PMT", category: "财务", description: "计算贷款的每期付款额", syntax: "PMT(rate, nper, pv, [fv], [type])", example: "=PMT(0.05/12, 360, 100000)" },
  { name: "PV", category: "财务", description: "计算投资的现值", syntax: "PV(rate, nper, pmt, [fv], [type])", example: "=PV(0.05/12, 360, -1000)" },
  { name: "FV", category: "财务", description: "计算投资的终值", syntax: "FV(rate, nper, pmt, [pv], [type])", example: "=FV(0.05/12, 120, -100)" },
  { name: "RATE", category: "财务", description: "计算年金的每期利率", syntax: "RATE(nper, pmt, pv, [fv], [type], [guess])", example: "=RATE(360, -1000, 100000)" },
  { name: "NPV", category: "财务", description: "计算净现值", syntax: "NPV(rate, value1, [value2], ...)", example: "=NPV(0.1, B1:B5)" },
  { name: "IRR", category: "财务", description: "计算内部收益率", syntax: "IRR(values, [guess])", example: "=IRR(B1:B5)" },
  // 数组与动态数组
  { name: "FILTER", category: "数组", description: "根据条件筛选数据", syntax: "FILTER(array, include, [if_empty])", example: '=FILTER(A1:C100, B1:B100>50)' },
  { name: "SORT", category: "数组", description: "对范围或数组排序", syntax: "SORT(array, [sort_index], [sort_order], [by_col])", example: "=SORT(A1:C100, 2, -1)" },
  { name: "SORTBY", category: "数组", description: "按另一个范围排序", syntax: "SORTBY(array, by_array1, [sort_order1], ...)", example: "=SORTBY(A1:C100, C1:C100, -1)" },
  { name: "UNIQUE", category: "数组", description: "返回范围中的唯一值", syntax: "UNIQUE(array, [by_col], [exactly_once])", example: "=UNIQUE(A1:A100)" },
  { name: "SEQUENCE", category: "数组", description: "生成连续数字序列", syntax: "SEQUENCE(rows, [columns], [start], [step])", example: "=SEQUENCE(10, 1, 1, 1)" },
  { name: "TRANSPOSE", category: "数组", description: "转置行列方向", syntax: "TRANSPOSE(array)", example: "=TRANSPOSE(A1:D5)" },
];

/** 搜索 Excel 函数 */
export function searchExcelFunctions(query: string, category?: string): ExcelFunction[] {
  let results = EXCEL_FUNCTIONS;

  // 按类别筛选
  if (category) {
    const catLower = category.toLowerCase();
    results = results.filter(
      (f) => f.category.toLowerCase().includes(catLower)
    );
  }

  // 按关键词搜索
  if (query) {
    results = results.filter(
      (f) =>
        f.name.toLowerCase().includes(query) ||
        f.description.toLowerCase().includes(query) ||
        f.category.toLowerCase().includes(query)
    );
  }

  // 精确匹配排在前面
  results.sort((a, b) => {
    const aExact = a.name.toLowerCase() === query ? 0 : 1;
    const bExact = b.name.toLowerCase() === query ? 0 : 1;
    return aExact - bExact;
  });

  return results.slice(0, 20); // 最多返回 20 个结果
}
