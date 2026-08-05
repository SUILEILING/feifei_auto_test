# NR/LTE 基站自动化测试框架 — 能力速览 & 一页 Demo

> 一页看懂：这套系统能干什么、怎么用、产出长什么样。
> 详细架构见 [架构总览.md](架构总览.md)。

---

## 30 秒电梯陈述

> **把 5G/LTE 基站射频产测从"人工点仪器、抄数、拼表格"变成一键无人值守。**
> 自动起停基站、扫频段/功率、测 BLER/发射功率，自动出带图表的 Excel 报告，全链路日志按用例归档可复盘。原本几小时的手动测试 → 无人值守自动完成。

---

## 核心卖点（一句话版）

| # | 卖点 | 为什么难/值钱 |
|---|---|---|
| 1 | **端到端一键化** | 覆盖"起基站→配置→测量→报告"最后一公里，不是半成品 |
| 2 | **完整日志不丢** | pty 伪终端抓全量基站日志，普通方式抓不全 |
| 3 | **智能重试 + 就绪判定** | AST 自动识别预期值，失败合并成一条记录 |
| 4 | **失败可视化** | 异常读数红色标出，不会被当成正常数据 |
| 5 | **全链路留痕** | SCPI/基站/信令/截图/核心网日志按 case 归档 |
| 6 | **专业报告** | 多脚本横向对比、最终数据汇总、图表自动生成 |

---

## 典型使用流程

```
1. 编辑 souren_config.py：
   PYTHON_SCRIPT_NAME = [
     {"script": "test_fixed_power", "nr_band": 77, "nr_bw": 100,
      "scs": 30, "range": "LOW", "lineLoss1": 25.0, "case_dir": "yc1100"},
     ... 想测多少配置就列多少 ...
   ]

2. Ubuntu 上确保 ubuntu_server.py 常驻（起基站/抓日志的服务端）

3. 客户端一条命令：
   python auto_test/main.py --run

4. 无人值守自动执行：
   起基站 → 配仪器/线损 → 扫频段×range×带宽 → 等 UE 连接 → 测 BLER/TXP → 停基站

5. 产出自动落到 log/execution_<时间戳>/：
   ✓ *_results_*.json     原始数据
   ✓ *_results_*.xlsx     执行汇总 + 详细记录 + 数据分析图表
   ✓ summary_charts_*.xlsx 多脚本横向对比 + 最终数据汇总
   ✓ 基站日志 / SCPI 日志 / 信令 & PVT/LTE 截图 / 核心网日志

6. 浏览器看历史：
   cd web_dashboard && python app.py  →  http://localhost:5000
```

---

## 测试脚本长什么样（极简、可读）

```python
def case_start():
    remote_gnb_start(mode="NSA")        # 一句话起基站
    remote_diag_start()                 # 起日志抓取
    ap.send("CONFigure:...线损/RF 配置...")

def case_body():
    for band in nr_band_list:
        for rng in ["Low", "Mid", "High"]:
            ap.send(f"...切频段 {band} / range {rng}...")
            ap.send("CONFigure:CELL1:NR:SIGN:CHANnel:SWITch")
            if not wait_for_ue_connected(ap):
                continue
            # 测量并自动提取数据点进图表
            ap.send("FETCh:NR:BLER:DL:RESult?", 7, True, "DL NR_BLER", x_label, ...)

def case_clear():
    remote_gnb_stop()                   # 一句话停基站
```

> 业务工程师只写"测什么"，"怎么跑/怎么抓日志/怎么出报告"框架全包。

---

## 产出示例（Excel 报告结构）

**单脚本报告 `*_results_*.xlsx`**
```
┌ 执行汇总 ─── 序号/执行时间/脚本/成功率/耗时/状态 ...
├ 详细执行记录 ─── 每一步命令/类型/状态/结果（失败标红）
└ 数据分析图表 ─── 每个指标一张柱状图，横坐标=频段/功率档
                    失败档位红色柱+红字，与真实值区分
```

**多脚本汇总 `summary_charts_*.xlsx`**
```
┌ 图表汇总 ─── 各配置各指标柱状图
├ 横向对比 ─── 多配置同指标并排对比
├ 汇总表格 ─── 数值表（每循环一列）
└ 最终数据汇总 ─── NR/LTE 功率 + BLER/TXP 关键指标一览
                    停止功率点状态标色（正常/回退后异常/失败）
```

---

## 适用场景

- 基站研发/产测：射频指标回归、频段/功率扫测
- 认证测试前的批量自检
- 版本迭代的自动化回归（改一版基站软件，一键重跑全套）

---

## 交付形态（可选）

| 形态 | 说明 |
|---|---|
| 源码 + 部署协助 | 适合硬件相近的团队直接用 |
| 配置化产品版 | 抽离硬编码后，换环境改配置即用（需产品化改造） |
| 定制适配 | 针对不同综测仪/基站厂商做硬件适配层 |

---

> 想深入了解实现细节、技术亮点与已知边界，见 [架构总览.md](架构总览.md)。
