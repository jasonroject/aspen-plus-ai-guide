# Aspen Plus AI 建模与调试指南

面向 AI 编程助手的 Aspen Plus 后台自动建模流程：通过 Windows Python 和 ActiveX/COM，从原生输入创建流程、调用求解器、检查日志与守恒，再保存并重开复算。

- [完整 AI 执行指南](GUIDE.zh-CN.md)
- [独立 Python 示例](smoke_background.py)
- [示例原生输入](examples/smoke.inp)
- [2026-10-16 本机验证报告](examples/verification_20261016.json)

## 快速开始

环境要求：Windows、已安装且许可证可用的 Aspen Plus，以及 Python。已验证版本为 Aspen Plus V15，COM ProgID 为 `Apwn.Document.41.0`；其他版本需要按指南重新探测接口。

```powershell
python -m pip install -r requirements.txt
python smoke_background.py --out ./runs/smoke_001
```

输出目录必须尚不存在。脚本将生成进料与单个 HEATER，后台求解，保存 `.bkp` 和 `.apwz`，再分别重开复算。退出码 0 且输出 `report.json` 的 `status=passed` 表示本示例通过。正式长任务应由外层监督器设置超时。

本次验证三次计算的出口均为 102 kg/hr，质量残差为零，未检出错误或警告。示例采用 IDEAL 物性，仅验证自动化调用链；实际工艺需要独立确定物性、反应和参数依据。

## 让其他 AI 使用

将 [指南](GUIDE.zh-CN.md) 提供给具备 Windows 命令与文件操作能力的 AI，并补充工艺目标、组分、流程、操作条件及数据来源。指南末尾提供可复制的启动提示词。

指南覆盖环境探测、原生文本与 COM 两条建模路线、变量树探索、单位检查、输入完整性、结果判定、自动调试、超时恢复和最终交付验收。

## 范围

针对 Aspen Plus。HYSYS、Aspen Custom Modeler 和 Aspen Dynamics 的接口不能直接套用。仓库不包含 Aspen 软件、许可证或研究项目数据。
