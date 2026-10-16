# scores.json 数据约定与校验

## 最小结构

JSON顶层必须包含 `schema_version: "1.0"`、`batch` 和 `participants`。

`batch` 必填：

- `label`：本次批次名称，不得沿用旧批次；
- `year`：用户已确认届别，字符串；
- `norm`：当前报告常模原文，字符串；
- `score_type`：固定 `percentile`；
- `mean_includes_self`：本流程固定 `true`；
- `comparison_scope`：固定 `current_batch`；
- `expected_count`：用户指定人数，未指定可省略，由实际名单计算。

`participants` 是非空数组，每项必填：

- `id`：唯一内部识别码，不包含未获授权的身份信息；
- `alias`：用户认可的展示缩写；应唯一且只含字母、数字、下划线或连字符，以免文件覆盖。重名由用户指定区分名，不能猜姓名；
- `norm`：可省略，省略时使用batch.norm；存在则必须与batch一致；
- `scores`：HPI/HDS/MVPI三个字典，键为下列维度原名，值为0–100的整数；
- `sources`：三个量表的来源，各有 `file`（PDF路径/文件名字符串）、`overview_page`（1起算整数）；可附 `detail_pages`（维度→实际页码字典）；
- `crosscheck`：三个量表的核验记录，各有 `status` 和 `note`。只有确实进行了对照，才记 `matched`；只能检查一处来源时用 `single_source_reviewed` 并说明，脚本输出警告；`unresolved` 为阻断错误。

身份编号核验可在受限临时清单中完成，最终优先展示随机代号；缩写也可能识别人。不要将实际身份编号写进Skill包。

## 固定维度及顺序

HPI：调适、抱负、社交、人际敏感度、审慎、好奇、学习方式。

HDS：激动、多疑、谨慎、内敛、消极、自大、狡猾、戏剧化、幻想、苛求、恭顺。

MVPI：认可、权力、享乐、利他、归属、传统、保障、商业、美感、科学。

以上对应本次中文报告。新报告出现不同译名/维度名时，先依据其定义建立明确映射、保留原词；在确认相容前不能仅凭近义词自动替换。当前校验器会拒绝不同键名，这是为暴露差异而非说明新报告无效。

## 计算

同一维度均值 = 本批全部N个分值之和 / N，含本人。不丢弃缺分的人、不按维度偷偷改变分母；缺失必须解决或明确重新界定批次。

个人差值 = 个人分值 − 本批均值。最大差异取同量表内差值绝对值最大者；并列时报告所有并列项，PPT有限位置可按原维度顺序展示首项并在备注说明并列。

校验器用Fraction计算精确差值与并列，不用展示时四舍五入后的数值挑选最大差异。显示均值时十进制四舍五入至1位小数。

输出 `score_audit.json` 包含 `count`、`score_count`、`means`、`means_display`、逐人的差值与最大差异及警告。

## 执行

```bash
python scripts/validate_scores.py scores.json --out score_audit.json
```

成功退出0，结构/分值/核验阻断错误退出2。只依赖Python标准库，Python 3.10及以上。脚本不读取PDF，不能证明人工标为matched的记录真实；来源核验必须由AI实际阅读完成。

`tests/test_workflow.py`会在临时目录创建2个虚拟样本来测试。样本不对应真实受测者。


## v2补充：群体统计与输入边界

`group_stats.py`读取相同完整三表JSON。输出每维均值、中位值、最小/最大和有效N；中位值在偶数N时取中间两项平均，不提前取整。

默认不统计风险人数。`--display-rules`要求JSON含`confirmed: true`、`source`和三量表各自连续覆盖0–100且不重叠的`bands`；未经实际核对不要把示例中的false改成true。规则只用于当前经认可的展示，不宣称适用于所有版本。

角色仅在用户认可项目规则后，用`--include-project-roles --role-rule-source`明确启用；输出不带人员名单。角色占比人数可重叠，不能相加成人员总数。

`assets/input/scores.synthetic.json`是独立构造的虚拟样例，值并非任何真实人员；其来源名是虚构占位，不能把测试通过称为PDF核验通过。只有三表完整数据才能运行此校验器；只有HPI的任务按单表流程处理，不补造另外两表。

真实Excel/CSV完整分值表可先规范化为JSON，来源使用`kind: "table"`、真实来源文件名和`locator`（例如实际工作表/行号）；原报告语义缺失仍需披露，程序算术通过不构成完整心理解释。PDF来源省略kind时默认pdf，继续使用overview_page及detail_pages。

样例表来源：
```json
{"kind":"table","file":"authorized_scores.xlsx","locator":"Sheet1:实际行号与列名"}
```

此处的文件名和locator仅为格式示例，正式记录必须对应实际读取过的文件，不能原样复制。表格输入也须有已核对的常模、量表和真实交叉核验状态；只有一份表就标single_source_reviewed。
