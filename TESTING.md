# 测试与核验记录 · v2.1

验证日期：2026-10-15。

## 实际完成

- `python -m unittest discover -s tests -v`：**76项测试通过**。全部为合成测试，保留原41项，新增35项覆盖管理解读结构、群体/个人/对比的中位值选项、模板优先级、身份展示授权字段、标签映射、叠加区间、解读卡依据/场景/影响/观察缺失、假设与实际观察区分、专业词不一律禁用、12页模板几何与本地链接。
- `python scripts/release_check.py .`：**0项待处理**。该脚本仅做已定义的有限检查，不是隐私认证。
- 新增12页群体管理解读模板已由LibreOffice转为PDF，全部渲染成页图，在总览逐页检查；第4、8、11、12页另作放大检查。页面对象无越界或重复名称；未嵌入原报告媒体、工作簿、批注或外链。
- 新模板正文和备注为占位/使用说明，无真实人员、分数、名单、内部业务案例或公司标志。热力图六行和观察卡只是布局占位，不是实际人数。
- 原v2.0的个人1页模板、群体10页模板及两个预览与原ZIP逐字节一致。本次未声称重新渲染这两份未改动的模板。
- 新二进制模板和预览检查后更新发布哈希；哈希只说明版本一致。

## 测试范围与局限

`plan_report.py`只检查填写的配置约束，不能验证真实受众权限、用户身份、聊天确认或参考来源。

`narrative_check.py`只检查卡片结构和少量关键词。将任意文字填满字段也可能通过；它不能确认心理解释合理、业务影响成立、文字自然或支持建议有效。

本次没有重新计算真实个人报告，没有将分支原报告作为公共示例，没有检查远程仓库或Git历史，也没有进行任何GitHub上传、平台安装或Library覆盖。

未测试所有Hogan版本/PDF格式、OCR、所有Office客户端或组织访问控制。模板仍是空白结构，正式报告必须由执行者真实阅读资料、填充、复核、渲染和交付。

## 复现命令

```bash
python -m unittest discover -s tests -v
python scripts/release_check.py .
```

配置与文字结构的无真实数据演示：

```bash
python scripts/plan_report.py assets/input/batch-request.example.json --out "$RUN/synthetic_plan.json"
python scripts/narrative_check.py assets/input/analysis-card.example.json --out "$RUN/synthetic_narrative_audit.json"
```

`RUN`须由执行环境设置在仓库外；示例中的权限标志不构成真实授权，热力图区间仍未确认。统计/检查脚本仅用Python 3.10+标准库。重建空白PPT另需`python-pptx`。本次用当前环境LibreOffice及PyMuPDF渲染，不保证所有客户端逐像素一致。
