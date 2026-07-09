# WPS / Office Excel 公式差异知识库

## 使用原则

这份内置知识库只记录写公式时最容易踩坑的差异，不做函数大全。

| 原则 | 建议 |
|---|---|
| 能写通用公式时 | 优先使用两边共有的基础函数和动态数组函数 |
| 涉及正则、联网、图片、股票数据时 | 默认按 WPS / Excel 分开写 |
| 涉及 `MAP/LAMBDA` 变量参与区域引用时 | WPS 中优先给变量加括号，如 `(x):A8` |
| 涉及 `GROUPBY/PIVOTBY` 时 | 已勾选支持动态数组函数时默认可用，重点确认参数表现和输出形状 |
| 要发给别人使用时 | 避免整列引用、避免合并单元格、避免依赖单边函数 |

## 公式输入细节

### 函数名

建议统一使用英文函数名，不依赖中文函数名或本地化函数名。

```excel
=FILTER(A2:D100,B2:B100="河北")
```

### 参数分隔符

公式参数分隔符可能受系统区域设置影响。知识库公式统一使用英文逗号。

```excel
=XLOOKUP(E2,A2:A100,C2:C100,"")
```

如果粘贴后软件提示公式错误，先检查分隔符是否应改成分号。

### 半角符号

公式中必须使用英文半角符号，例如 `""`、`() `、`:`、`,`、`*`。不要混入中文全角引号、括号、冒号、逗号。

### 动态数组公式只输入一次

动态数组公式会自动溢出，不需要向下拖满。结果区域必须为空，否则会出现 `#SPILL!` 或无法正常展开。

```excel
=FILTER(A2:D100,B2:B100="河北")
```

### `#` 和 `@`

| 符号 | 含义 | 注意 |
|---|---|---|
| `A1#` | 引用 A1 溢出的整个动态数组区域 | 只在确认有溢出结果时使用 |
| `@` | 隐式交叉 / 取当前行或首项 | Excel 与 WPS 自动插入和显示行为可能不同 |

建议：能不用 `@` 就不用。如果需要取第一项，优先写 `TAKE(x,1)` 或 `INDEX(x,1)`。

## 写入公式时的通用避坑

### 不要直接整列引用动态数组

容易卡顿：

```excel
=FILTER(A:D,B:B="河北")
```

更稳：

```excel
=FILTER(A2:D1000,B2:B1000="河北")
```

### 避免合并单元格参与计算

合并单元格在视觉上是一块，底层只有左上角有值，其余行是空。

常见处理：

```excel
=SCAN(0,A2:A100,LAMBDA(x,y,IF(y<>"",y,x)))
```

作用：把合并单元格的首行值向下补齐。

### 拼接数组时要处理尺寸不一致

`HSTACK`、`VSTACK` 遇到行列数不一致时，可能用 `#N/A` 补位。

```excel
=IFNA(HSTACK(A2:A10,"",B2:C10),"")
```

### 空值可能被当成 0

在算术运算中，空单元格可能按 `0` 参与计算。

```excel
=FILTER(A2:C100,A2:A100<>"")
```

```excel
=TOCOL(A2:A100,1)
```

## 必须区分 WPS / Excel 的公式

### 正则函数

| 场景 | WPS | Office Excel |
|---|---|---|
| 正则提取 | `REGEXP` | `REGEXEXTRACT` |
| 正则判断 | `REGEXP` 或其他模式 | `REGEXTEST` |
| 正则替换 | `REGEXP` 替换模式 | `REGEXREPLACE` |

WPS 示例：

```excel
=REGEXP(A2,"\d+|[A-z]+")
```

Excel 示例：

```excel
=REGEXEXTRACT(A2,"\d+|[A-z]+")
```

注意：不要把 WPS 的 `REGEXP` 直接复制到 Office Excel，也不要把 Excel 的 `REGEXEXTRACT`、`REGEXTEST`、`REGEXREPLACE` 直接复制到 WPS。

### `MAP/LAMBDA` 中变量参与区域引用

Excel 中通常可以写：

```excel
=MAP(A2:A8,LAMBDA(x,TAKE(TOCOL(x:A8,1),1)))
```

WPS 中更稳写法：

```excel
=MAP(A2:A8,LAMBDA(x,TAKE(TOCOL((x):A8,1),1)))
```

原因：Excel 更容易保留 `x` 的引用身份；WPS 中 `x` 可能先被当作值处理，导致 `x:A8` 解析失败。凡是 `LAMBDA` 变量要当作区域起点或终点，都优先写成 `(x)`，例如 `(x):A8`、`A2:(x)`。

### `GROUPBY/PIVOTBY`

在已勾选支持动态数组函数的环境中，`GROUPBY/PIVOTBY` 默认可用。重点关注参数设置、标题/汇总行、筛选数组和聚合结果形状。

通用思路：

```excel
=GROUPBY(分组字段,汇总值,SUM,0,0)
```

容易踩坑：

| 坑 | 说明 |
|---|---|
| 默认标题/汇总行 | 参数设置不同或省略时，结果可能多出标题行、总计行 |
| 聚合函数返回数组 | 如 `TOROW`、自定义 `LAMBDA` 返回多值，可能形成嵌套数组 |
| WPS 显示嵌套数组 | 可能只显示每组数组的第一个元素 |

WPS 常用展开法：

```excel
=REDUCE(,DROP(GROUPBY(LEFT(A2:A11,4),A2:A11,TOROW,,0),,1),VSTACK)
```

建议：简单求和、计数优先使用通用写法；聚合后还要横向展开时，优先准备 `REDUCE + VSTACK` 修正写法。

### Web、图片、股票类函数

这些不建议写成 WPS / Excel 通用公式。

| 函数 | 问题 |
|---|---|
| `STOCKHISTORY` | Excel 365 股票历史数据函数，依赖 Microsoft 在线数据源 |
| `IMAGE` | Excel 365 单元格图片函数，WPS 支持不稳定 |
| `WEBSERVICE` | Web 请求类函数，跨环境差异大 |
| `FILTERXML` | XML 解析类函数，跨环境差异大 |
| `ENCODEURL` | URL 编码类函数，跨环境差异大 |

### 字节版文本函数

如 `LEFTB`、`RIGHTB`、`MIDB`、`LENB`、`FINDB`、`SEARCHB`、`TEXTJOINB`，和字符编码、双字节字符处理有关。能用普通文本函数时，优先用 `LEFT`、`RIGHT`、`MID`、`LEN`、`FIND`、`SEARCH`。

## 建议做区分的公式类型

| 类型 | 是否要区分 | 说明 |
|---|---|---|
| 正则提取/替换/判断 | 必须区分 | WPS `REGEXP`，Excel `REGEX...` |
| `MAP/LAMBDA` 中变量拼区域 | 建议区分 | WPS 写 `(x):A8` 更稳 |
| Web/股票/图片 | 必须区分 | Excel 服务依赖明显 |
| `GROUPBY/PIVOTBY` 复杂聚合 | 建议区分 | 函数默认可用，但嵌套数组、标题、汇总参数表现可能不同 |
| 合并单元格计算 | 建议单独测试 | WPS/Excel 对空值、溢出、引用的显示可能不同 |
| 动态数组溢出引用 `#` | 建议单独测试 | 旧版 Excel 或部分 WPS 环境可能不支持 |
| `@` 隐式交叉 | 建议避免 | 两边自动插入/显示行为可能不一致 |
| 基础求和、查找、文本截取 | 通常不必区分 | 如 `SUM`、`IF`、`XLOOKUP`、`LEFT`、`TEXTSPLIT` 等，仍需看版本 |

## 推荐跨环境写法

### 多条件筛选

```excel
=FILTER(A2:D100,(B2:B100="河北")*(C2:C100<>""))
```

### 合并多个表

```excel
=LET(a,VSTACK(表1:表3!A2:C100),FILTER(a,CHOOSECOLS(a,1)<>""))
```

### 向下填充空白

```excel
=SCAN(0,A2:A100,LAMBDA(x,y,IF(y<>"",y,x)))
```

### 分组汇总

```excel
=GROUPBY(B3:B100,C3:C100,HSTACK(ROWS,SUM),,0,,D3:D100<>"撤销")
```

### 文本拆分后计算

```excel
=TEXTSPLIT(A3,,",")*1
```

## 迁移检查清单

从 WPS 迁移到 Office Excel 前，检查：

```text
是否用了 REGEXP？
是否用了 WPS 特有函数或插件函数？
是否用了 GROUPBY 聚合返回数组？
是否用了合并单元格参与公式？
```

从 Office Excel 迁移到 WPS 前，检查：

```text
是否用了 REGEXEXTRACT / REGEXTEST / REGEXREPLACE？
是否用了 STOCKHISTORY / IMAGE？
是否用了 Web 类函数？
是否用了 x:A8 这种 LAMBDA 变量直接拼区域？
是否用了 # 溢出引用或 @ 隐式交叉？
是否用了 Excel 表格结构化引用？
```

## 动态数组环境测试建议

已勾选支持动态数组函数时，`GROUPBY/PIVOTBY` 可按默认可用处理。下面这些公式主要用于检查当前文件环境中的溢出、引用、聚合输出是否符合预期。

```excel
=FILTER(A2:A5,A2:A5<>"")
```

```excel
=MAP(A2:A5,LAMBDA(x,x))
```

```excel
=SCAN(0,A2:A5,LAMBDA(x,y,IF(y<>"",x+1,x)))
```

```excel
=GROUPBY(B2:B10,C2:C10,SUM,0,0)
```

```excel
=PIVOTBY(A2:A10,B2:B10,C2:C10,SUM,,0,,0)
```
