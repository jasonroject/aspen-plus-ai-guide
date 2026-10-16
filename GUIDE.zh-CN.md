# Aspen Plus 后台自动建模与调试：AI 执行指南

编写日期：2026-10-16。目标：让另一个具备 Windows 文件和命令执行能力的 AI，从工艺需求出发，创建可打开、可重新计算、可追溯的 Aspen Plus 模型。

本文针对 **Aspen Plus**。HYSYS、Aspen Custom Modeler、Aspen Dynamics 使用各自接口，不能直接套用本指南的 ProgID 和变量树。

## 1. 先建立正确的执行路线

主要路线：**Windows Python → Aspen Plus ActiveX/COM → 原生输入文件与变量树 → Aspen 求解器 → 原生日志与结果 → 保存并重开复算**。

后台运行指隐藏 Aspen 窗口、通过程序操作已安装的求解器。它仍依赖本机 Aspen 安装、许可证和可用的 Windows 会话。`SuppressDialogs=1` 不能保证屏蔽所有许可证或系统弹窗。

从零建模可以走两条路线，复杂模型通常组合使用：

| 路线 | 适用情况 | 执行动作 | 主要风险 |
|---|---|---|---|
| 原生文本建模 | 批量创建完整流程、反应网络、重复工艺结构 | 生成 `.inp`，用 `InitFromFile2` 导入 | 语法、数据库 ID、续行、导入后内容遗漏 |
| COM 变量树建模 | 探索对象结构、增加单元、参数修改、读取结果 | `InitNew2` 后添加块和物流，或打开副本后写树 | 节点尚未生成、表格行接口、单位、保存后身份丢失 |

优先把输入生成器作为唯一模型来源；需要 COM 后处理时，将修补步骤写成可重复函数。最终交付必须能由生成器和后处理脚本重建。

## 2. 本机已有证据与适用边界

以下是本地记录，不能推断为所有版本通用：

- 已安装目录：`C:\Program Files\AspenTech\Aspen Plus V15.0`；本地脚本使用 `Apwn.Document.41.0`。
- Python：`C:\Program Files\py\python.exe`；本次检查可导入 `pythoncom` 和 `win32com.client`。
- 历史验证覆盖空案例创建、原生输入生成、完整流程求解、输入完整性诊断和保存重开复算。
- 历史案例曾出现：保存文件重开后没有缓存结果，但重新计算正常。因此分别验收缓存持久化与可复算性。
- 本仓库发布独立最小示例及本次验证报告；历史研究项目文件不作为运行依赖。

本指南附带 `smoke_background.py`，是独立最小案例。2026-10-16 本次实测 `status=passed`，用时约 46 秒：初次计算、`.apwz` 重开复算和 `.bkp` 重开复算均得到 102 kg/hr，质量残差为 0，未检出错误或警告。证据见`examples/verification_20261016.json`。它采用 IDEAL 物性，只验证自动化，不作为实际物性选择示范。

## 3. 执行前：把任务变成模型契约

先生成 `model_spec.json`，至少写明：

1. 目标指标、稳态/动态边界、工艺范围和已知操作条件。
2. 组分清单：内部短名、数据库真实标识、化学身份；区分常规、非常规和假想组分。
3. 物性方法及选择依据；存在不同方法的工段时列明分配范围。
4. 模块类型、入口与出口物流、操作规定；每个单元的自由度由什么输入消除。
5. 所有数值的单位、质量/摩尔基准、干/湿基、绝压/表压。
6. 反应式、反应级数、浓度基准、相态、速率单位、A/Ea/Tref 定义及证据来源。
7. 未知量、采用的假设、参数有效范围、拟校准量。
8. 验收阈值：质量/元素/能量残差、关键结果容差、允许保留的警告。

必要数据缺失时，先完成环境检查和不依赖这些数据的最小结构；把依赖缺失数据的研究结论标记为待定。教学演示采用的假设值必须标注为演示。模型收敛不能替代实验或物性验证。

完成条件：每个实际写入模型的值都能对应到来源或明确假设，单位可换算，组分身份和反应基准无歧义。

## 4. 探测环境，不碰当前人工案例

在独立工作目录执行：

```powershell
python -c "import sys,pythoncom,win32com.client; print(sys.executable)"
Get-ChildItem 'Registry::HKEY_CLASSES_ROOT' |
    Where-Object PSChildName -Like 'Apwn.Document*' |
    Select-Object PSChildName
Get-Process AspenPlus,apmain -ErrorAction SilentlyContinue |
    Select-Object Id,ProcessName,StartTime
```

注册表和实际创建对象结果优先于版本号猜测；换机器重新探测。缺少 pywin32 时，安装到执行脚本的同一解释器，而不是另一个 Python。

使用 `DispatchEx` 创建独立自动化对象；记录自己的输入输出路径。即使使用 `DispatchEx`，仍须核对打开的文档归属。不要通过 `GetActiveObject` 或运行对象表随意接管人工案例。

推荐目录：

```text
AspenRuns/<case_id>/
  model_spec.json
  build.py
  source.inp
  environment.json
  attempts/001/...
  delivered/model.bkp
  delivered/model.apwz
  validation.json
  handoff.md
```

每次尝试使用新目录，保留输入哈希、运行时间、版本、日志、改动和结果。脚本只关闭自己创建的 COM 文档。

完成条件：解释器依赖可用、ProgID 可创建、许可证可初始化案例、用户原有案例未被改写。

## 5. 先跑通最小案例，再扩展工艺

随附脚本可直接执行，输出目录必须尚不存在：

```powershell
python smoke_background.py `
  --out F:/AspenRuns/smoke_001
```

该示例从文本生成三组分进料与一个 HEATER，随后导入、计算、检查质量守恒、保存 `.bkp/.apwz`，再分别打开两个文件重新计算并比较结果。通过标志是进程退出码 0 且 `report.json` 中 `status=passed`。同时阅读 `warning_lines`。

脚本的 `Run2(False)` 为同步调用。正式长任务由外层进程监督器设置墙钟超时；Python 循环不能打断阻塞中的 COM 调用。脚本本身不带强制杀进程逻辑。

完成条件：初次计算和两种保存文件的重开复算都通过，出口为 102 kg/hr，质量残差在设定容差内。

## 6. 从零创建原生结构

### 6.1 初始化与文件类型

```python
import pythoncom
import win32com.client

pythoncom.CoInitialize()
app = win32com.client.DispatchEx("Apwn.Document.41.0")
app.SuppressDialogs = 1
app.InitNew2()                    # 从空案例开始
app.Visible = False              # 初始化之后再设置
# app.InitFromFile2(inp_or_apwz)  # 原生输入文件或文档
# app.InitFromArchive2(bkp)       # 归档文件
```

每个使用 COM 的线程独立初始化/释放 COM。生产脚本在 `finally` 中执行对象清理和 `CoUninitialize()`，具体结构参见随附脚本。`Close()` 后服务可能已经断开，不要求再成功执行一次 `Quit()`。

### 6.2 原生文本路线

以随附脚本 `INPUT` 为最小语法样本，按以下顺序增量生成：

1. 单位、数据库、组分、物性方法。
2. FLOWSHEET 连接、进料、一个单元操作。
3. 逐段增加泵、加热、闪蒸、分离、反应。
4. 加入循环、设计规定、热集成。

每次导入立即比较实际组分、块名、物流名和连接与模型契约是否一致，然后导出 Aspen 解释后的 `.inp` 核对关键规定。**没有抛异常不等于导入完整。**

复杂语法优先查本机官方示例，或在独立的小案例中设置一个对象后导出原生文本，提取对应段落。官方示例目录：`C:\Program Files\AspenTech\Aspen Plus V15.0\GUI\Examples`，自动化示例为 `GUI\VBDotNetexample`；帮助主题为 *Using the Aspen Plus ActiveX Automation Server*。

历史上遇到过超过 78 列的反应段被读坏，导致流程表为空。把它作为本机排障线索：使用已验证的原生续行规则，避免超长行，并以导入后结构核对确认；不要在引号或标识符内部任意断行。

### 6.3 COM 变量树路线

以下添加块与端口形式出现在本机已有构建脚本中：

```python
tree = app.Tree
tree.FindNode(r"\Data\Blocks").Elements.Add("HX!HEATER")
tree.FindNode(r"\Data\Streams").Elements.Add("FEED")
tree.FindNode(r"\Data\Streams").Elements.Add("HOT")
tree.FindNode(r"\Data\Blocks\HX\Ports\F(IN)").Elements.Add("FEED")
# 出口端口先枚举 HX\Ports，确认实际名称，再添加 HOT。
```

端口名由模型类型决定。先枚举 `\Data\Blocks\<block>\Ports`，不能把 HEATER 的端口名复制到 FLASH、HEATX 或反应器上。

写任何参数前先做四件事：

1. `FindNode(path)` 确认存在；不存在时输出最近父节点的子节点。
2. 读取原值、`UnitString`，以及必要时的类型与枚举元数据。
3. 把目标数值显式换算到节点单位，再写 `Value`。
4. 读回新值与单位，和预期比较；有差异就停止该次构建。

对于不存在的组分流量行、分离规定行或反应表行，修改 `.Value` 不会自动创建全部结构。先完成组分和基准设置、使用该对象支持的行插入方式；仍无法可靠创建时，切换到已验证的原生文本生成路线。

完成条件：结构、连接、输入数值和单位与模型契约逐项一致；保存重开后仍一致。

## 7. 不知道接口或路径时怎么查

限定在独立案例中探索，采用“一个变量、一次变更、导出前后差异”的方法：

```python
parent = app.Tree.FindNode(r"\Data\Blocks\HX\Input")
for child in parent.Elements:
    try:
        print(child.Name, child.Value, child.UnitString)
    except Exception as exc:
        print(child.Name, type(exc).__name__)
```

遇到未知方法、参数或状态枚举，读取对象的 COM 类型信息：

```python
info = app._oleobj_.GetTypeInfo()
for i in range(info.GetTypeAttr().cFuncs):
    desc = info.GetFuncDesc(i)
    print(info.GetNames(desc.memid), info.GetDocumentation(desc.memid))
```

可对 `app.Engine`、树节点和集合对象重复此步骤。类型库还可枚举状态常量；可通过 `GetContainingTypeLib()` 枚举类型库中的状态常量。枚举值和方法签名以本机类型库为准，避免从其他版本照抄。

完成条件：把确认过的路径、单位、允许值及验证版本写入本项目 `node_map.json`；不能确认的项目保持未验证。

## 8. 运行前检查、运行与结果判定

### 8.1 运行前检查

本地 V15 代码使用以下检查；其他版本先核对类型库：

```python
data = app.Tree.FindNode(r"\Data")
flags = int(data.AttributeValue(12))  # 本地 HAP_COMPSTATUS
incomplete = bool(flags & 64)
next_item = data.NextIncomplete(0)
ready = bool(app.Engine.Ready)
```

`64` 在本地表示输入不完整，`128` 表示输入完整。它们是**位掩码**，不能用 `flags == 128` 判断。分别记录根节点和每个块的状态，定位缺项到具体块与输入表。

开算前要求输入完整且可运行。历史案例中，某些 SEP 分率缺项时引擎仍可能给出数值，但最终文档不满足输入完整性，所以不能把“算出了数”当作验收。

### 8.2 求解与日志

```python
app.Engine.Run2(False)  # 同步；返回 None 也可能是正常接口返回
app.Export(13, history_path)   # 本机 V15 历史日志
app.Export(4, solved_inp_path) # 本机 V15 原生输入
```

Export 数值常量需在换版本时重新核对。历史日志关键词扫描只是第一层筛查；最终同时检查原生运行状态、相关单元状态、结果和守恒。完整保留警告上下文。

异步路线只有在本机确认 `Run2` 参数与 `IsRunning` 行为后使用。轮询间隔建议 0.5–2 秒，记录心跳并设超时；`IsRunning=False` 只表示不在运行，不能证明成功。

`Engine.Ready` 是运行前检查的一部分。本地历史验证中，运行后出现 `Ready=False` 仍有正确结果，不应据此单独判为失败。

### 8.3 结果验收

至少保存以下证据：

- 实际输入完整性、原生运行状态和所有关键块状态，按本机状态枚举解释。
- 日志中的错误、严重错误、终止错误，以及每类警告的说明。
- 关键流股、组分流量、温压、相态、负荷；同时导出数值与单位。
- 全流程和单元质量衡算；含反应时增加元素衡算；有焓依据时增加能量衡算。
- 产率、纯度、转化率、相分率是否在物理范围内。

质量残差可采用：`abs(Σ入口−Σ出口) / max(Σ入口, m_floor)`。`m_floor` 是带流量单位的防零下限，由项目指定。能量衡算统一符号，计入热量、轴功和边界内所有流股；非常规或假想组分焓缺乏依据时，明确限制能耗结论。

常见结果候选路径：

```text
\Data\Streams\<S>\Output\MASSFLMX\MIXED
\Data\Streams\<S>\Output\MASSFLOW\MIXED\<CID>
\Data\Streams\<S>\Output\MOLEFLOW\MIXED\<CID>
\Data\Blocks\<B>\Output\QCALC
\Data\Results Summary\Run-Status\Output
```

这些是候选路径，必须先查存在性和单位。含非常规固体的流程需要覆盖实际子流股，不能只累计 MIXED。

## 9. 自动调试循环

固定循环：**保存失败现场 → 分类 → 提出可证伪原因 → 只改一个原因对应的设置 → 重跑 → 比较验收指标 → 接受或回滚**。

每次记录 `attempt.json`：原输入哈希、错误上下文、假设、修改、运行结果、守恒残差及是否接受。建议同一假设最多两次尝试、连续五次无改善则停止盲改并输出最小复现与明确缺口；次数可由项目调整。

| 症状 | 优先检查 | 下一动作 |
|---|---|---|
| ProgID 创建失败 | COM 注册、目标安装、解释器、具体 HRESULT | 重新探测注册项，保留异常；不改工艺模型 |
| 初始化卡住 | 许可证、首次启动或文件提示、后台引擎 | 外层超时取证，必要时显示自己创建的窗口查看提示 |
| 找不到节点 | 导入是否成功、父对象、组分/基准、表格行 | 枚举父节点，导出原生输入，创建缺少的结构 |
| 写入成功但数值不对 | 节点单位、质量/摩尔基准、温度标度 | 显式换算后写入并读回 |
| 输入不完整但能计算 | SEP 等规定表是否缺少新增组分行 | 用状态与 NextIncomplete 定位并按物理依据补全 |
| 导入后流程为空 | 文本语法、长行、数据库、原生诊断 | 缩成一个块的复现，逐段恢复 |
| 保存后组分错位 | CID、ANAME、数据库标识、用户组分 | 比较导出文本与重开组分身份；修复生成器 |
| 闪蒸/物性失败 | 组分身份、物性方法、温压有效范围、相数 | 单独做该混合物的物性与相平衡小案例 |
| 反应器积分失败 | 计量、速率维度、反应级数、零浓度附近行为 | 用单反应和独立动力学核算定位，再测试积分参数 |
| 循环不收敛 | 单程各单元是否可算、初值、tear stream | 先断开循环跑通，再逐步闭合、延续参数 |
| 精馏不收敛 | 进料相态、压力、规定自由度、初值 | 先降低问题复杂度并验证 VLE，再恢复完整规定 |
| 有结果但失衡 | 是否旧结果、错误块、遗漏子流股、反应配平 | 从新副本重新计算，逐块定位残差 |
| 警告减少但结果更差 | 是否提前终止或积分失败 | 比较完整状态和守恒，回滚该数值改动 |
| 保存后结果空白 | 是否保存缓存、是否能重新计算 | 重开再运行；分别报告缓存与可复算性 |

化学和数值修改分开记录。改变 Ea、反应式、组分身份、物性方法或收率属于模型变更，需要相应依据。只为消除警告而调整这些量，会改变要研究的问题。

本地实例：把 RPLUG 的非负积分开关打开后，某类警告消失，却导致积分失败。该尝试被回滚；因此不能用警告计数下降判断改善。

超时恢复时，仅操作经确认属于本次任务的文档/进程。不要按进程名批量结束所有 Aspen；也不要未经验证就结束与人工案例共享的引擎。若无法确认归属，保留现场并说明阻塞点。

## 10. 保存与交付验收

本机已有代码使用：

```python
app.WriteArchive2(archive_path, 1)
app.SaveAs2(document_path, True)
```

输出路径使用新文件名，随后确认文件存在且非空。关闭自己创建的文档，在新 COM 对象中重新打开**最终交付文件**，进行以下验收：

1. 运行前输入完整；结构、组分身份、关键参数与单位正确。
2. 重新计算，读取新日志，确认求解状态与关键块状态可接受。
3. 比较主要流量、组分量、温压、负荷和指标，满足预先定义的绝对/相对容差。
4. 重新进行守恒检查；所有保留警告都有影响说明。
5. 从构建脚本再次生成模型，确认修复已进入生成器，而非只存在于手工改过的文件。

如要求打开即显示结果，另行验收缓存持久化；不能从保存成功推断。若要求流程图可读，还要单独打开最终文档核对布局、连接显示和命名；数值通过不自动等于图面排版完成。

交付包含：模型契约、生成/运行脚本、原生输入、`.bkp`、原生文档、历史日志、带单位结果、验证报告、假设和限制、复现命令。

最终状态分开报告：`automation_pass`、`input_complete`、`solver_pass`、`balance_pass`、`reopen_pass`、`scientific_validation`。科学验证未完成时标记待验证，即使其他项已经通过。

## 11. 可直接交给其他 AI 的启动提示词

> 请按《Aspen后台自动建模与调试-AI执行指南.md》执行 Aspen Plus 自动建模。先读取项目约定和现有证据，探测本机 Python、COM 与版本，在独立目录创建自己的案例。先写明模型契约与未知量，运行最小后台测试；随后增量生成原生输入与流程结构，按实际节点单位写入并读回。每次运行检查输入完整性、原生状态、日志、关键结果和守恒。失败时保存现场，一次修改一个有依据的原因并记录对比。最终保存原生文件，关闭后在新实例打开最终交付文件重新计算并比较结果。把修复同步回生成器，交付脚本、模型、日志和验证报告。分别报告自动化、输入、求解、守恒、重开复算和科学验证状态；所有假设明确标注。工艺需求如下：〔填写目标、组分、流程、操作条件和数据来源〕。
