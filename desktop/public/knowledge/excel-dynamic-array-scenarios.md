# Excel/WPS 动态数组场景知识库

本文件只记录可直接检索复用的场景、例子、公式和踩坑点。优先给出 WPS 与 Office Excel 都可用的写法；涉及 WPS 独有函数时必须单独标注。

## 按类别合并明细文本

来源：
- https://mp.weixin.qq.com/s/01y_yKWX6CEP1clEDxmdAw

场景：A 列是姓名，B 列是部门，需要按部门汇总姓名，并用逗号合并。

示例数据：

| 姓名 | 部门 |
|---|---|
| 张三 | 财务 |
| 李四 | 销售 |
| 王五 | 财务 |
| 赵六 | 销售 |

目标结果：

| 部门 | 姓名汇总 |
|---|---|
| 财务 | 张三,王五 |
| 销售 | 李四,赵六 |

通用公式：

```excel
=LET(name,A2:A7,dept,B2:B7,u,UNIQUE(dept),HSTACK(u,MAP(u,LAMBDA(d,TEXTJOIN(",",,FILTER(name,dept=d))))))
```

注意：
- `UNIQUE(dept)` 先生成分组清单。
- `MAP(u,LAMBDA(d,...))` 按每个部门逐个筛选。
- `TEXTJOIN` 的第 2 参数留空或写 `TRUE`，用于忽略空值。
- 如果部门列含空白，先过滤：`FILTER(B2:B7,B2:B7<>"")`。

## 按数量把物品垂直展开

来源：
- https://mp.weixin.qq.com/s/6GP39CHCrvLN6RCvdWjUCQ

场景：A 列是物品，B 列是数量，需要按数量重复物品名称，生成一列明细。

示例数据：

| 物品 | 数量 |
|---|---:|
| A | 3 |
| B | 1 |
| C | 2 |

目标结果：

| 明细 |
|---|
| A |
| A |
| A |
| B |
| C |
| C |

通用公式，适合数量都大于 0：

```excel
=DROP(REDUCE(0,A1:A4,LAMBDA(x,y,VSTACK(x,IF(SEQUENCE(OFFSET(y,,1)),y)))),1)
```

更稳写法，可处理数量为 0 的行：

```excel
=LET(pos,SEQUENCE(SUM(B1:B4)),cum,SCAN(0,B1:B4,LAMBDA(a,b,a+b)),INDEX(A1:A4,XMATCH(pos,cum,1)))
```

注意：
- `REDUCE + VSTACK` 的思路是逐项追加数组。
- `OFFSET(y,,1)` 读取当前物品右侧的数量。
- 如果数量可能为 0，优先用 `SEQUENCE(SUM(...)) + SCAN + XMATCH` 版本。

## 按数量展开并生成组内序号

来源：
- https://mp.weixin.qq.com/s/7WopoauWEqp_s4QZ8k_Xug

场景：B 列是品名，C 列是重复次数，需要展开成明细，并给每个品名内部编号。

示例数据：

| 品名 | 次数 |
|---|---:|
| a | 3 |
| b | 2 |
| c | 1 |

目标结果：

| 组序 | 品名 |
|---:|---|
| 1 | a |
| 2 | a |
| 3 | a |
| 1 | b |
| 2 | b |
| 1 | c |

WPS 示例写法，使用 `REPTARRAY`，Office Excel 不通用：

```excel
=DROP(REDUCE(0,B3:B5,LAMBDA(x,y,VSTACK(x,HSTACK(SEQUENCE(OFFSET(y,,1)),REPTARRAY(y,OFFSET(y,,1)))))),1)
```

WPS / Office Excel 通用替代写法：

```excel
=LET(item,B3:B5,n,C3:C5,pos,SEQUENCE(SUM(n)),cum,SCAN(0,n,LAMBDA(a,b,a+b)),idx,XMATCH(pos,cum,1),prev,IF(idx=1,0,INDEX(cum,idx-1)),HSTACK(pos-prev,INDEX(item,idx)))
```

注意：
- `REPTARRAY` 是 WPS 独有函数，不能直接写入 Office Excel 通用公式。
- 通用替代公式用累计数量 `cum` 定位当前明细属于第几个品名。
- `pos-prev` 生成每个品名内部的 1、2、3 编号。

## 合并散落在区域中的小分表

来源：
- https://mp.weixin.qq.com/s/ZF930kQ3d618jMrhCST-jg

场景：多个 3 列小表散落在同一大区域中，每个小表都有相同表头，如“日期、品名、销量”，需要合并成一个连续总表。

示例数据形态：

| 日期 | 品名 | 销量 |  | 日期 | 品名 | 销量 |
|---|---|---:|---|---|---|---:|
| 3日 | a | 7 |  | 1日 | a | 10 |
| 3日 | b | 5 |  |  |  |  |
|  |  |  |  | 日期 | 品名 | 销量 |
|  |  |  |  | 2日 | a | 8 |

目标结果：

| 日期 | 品名 | 销量 |
|---|---|---:|
| 1日 | a | 10 |
| 2日 | a | 8 |
| 3日 | a | 7 |
| 3日 | b | 5 |

通用公式：

```excel
=LET(a,WRAPROWS(TOCOL(A2:G8,3),3),VSTACK(A2:C2,SORT(FILTER(a,TAKE(a,,1)<>"日期"),1,1)))
```

注意：
- `TOCOL(A2:G8,3)` 把大区域压成单列，并忽略空白和错误值。
- `WRAPROWS(...,3)` 每 3 个值还原成一行，因为小表固定为 3 列。
- `FILTER(a,TAKE(a,,1)<>"日期")` 删除重复表头。
- 该写法要求小表数据都严格是 3 列结构，且大区域内没有无关文本。

## 合并单元格分数向下补齐后分组求平均

来源：
- https://mp.weixin.qq.com/s/z6jPshdJhFjMn2kYit7RXw

场景：B 列分数存在合并单元格，合并区域下方实际为空，需要按员工统计平均分。

示例数据：

| 员工 | 分数 |
|---|---:|
| a | 85 |
| a |  |
| b | 80 |
| b | 70 |
| c | 75 |
| c | 60 |
| c |  |

目标结果：

| 员工 | 平均分 |
|---|---:|
| a | 85 |
| b | 75 |
| c | 65 |

通用公式：

```excel
=LET(name,A2:A8,score,SCAN(0,B2:B8,LAMBDA(x,y,IF(y="",x,y))),GROUPBY(name,score,LAMBDA(x,SUM(x)/ROWS(x)),,0))
```

注意：
- `SCAN` 用上一个非空分数补齐当前空白，相当于公式版“向下填充”。
- `GROUPBY` 的第 3 参数用 `LAMBDA` 自定义聚合逻辑。
- 如果姓名列有空白，先过滤姓名和分数数组，避免空白也成为一组。

## 生成两列清单的全部配对组合

来源：
- https://mp.weixin.qq.com/s/Joqw0eHu95_NarBv77otdw

场景：A 列是产品，D 列是部门，需要生成“每个产品 × 每个部门”的全部组合。

示例数据：

| 产品 |
|---|
| 产品A |
| 产品B |

| 部门 |
|---|
| 一部 |
| 二部 |

目标结果：

| 产品 | 部门 |
|---|---|
| 产品A | 一部 |
| 产品A | 二部 |
| 产品B | 一部 |
| 产品B | 二部 |

WPS 示例写法，使用 `REPTARRAY`，Office Excel 不通用：

```excel
=HSTACK(TOCOL(EXPAND(A2:A5,,ROWS(D2:D3),"")),REPTARRAY(D2:D3,ROWS(A2:A5)))
```

WPS / Office Excel 通用替代写法，每一行都填完整产品名：

```excel
=LET(p,A2:A5,d,D2:D3,pr,ROWS(p),dr,ROWS(d),seq,SEQUENCE(pr*dr),HSTACK(INDEX(p,ROUNDUP(seq/dr,0)),INDEX(d,MOD(seq-1,dr)+1)))
```

注意：
- `REPTARRAY` 是 WPS 独有函数，不能作为 Office Excel 通用公式。
- 通用替代公式用 `SEQUENCE(pr*dr)` 生成所有组合位置。
- `ROUNDUP(seq/dr,0)` 定位产品行，`MOD(seq-1,dr)+1` 定位部门行。

