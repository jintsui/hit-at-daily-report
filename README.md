# HIT AT 效率诊断日报

自动生成港口水平运输（AT 无人集卡）效率诊断日报。读取车辆遥测、能耗、操作日志 CSV 与 AT 位移 XLSX，计算各项运营指标，产出一份**自包含的 HTML 报告**（数据与 ECharts 全部内联，双击即开，无需联网）。

## 报告内容

- 车队概览、箱量规模、时段分布
- 分时作业节奏（岸桥 BQ1–BQ4）、岸桥任务分配 · 每小时箱数、堆场任务分配 · 每小时箱数（堆场-贝位）
- 作业循环时长、各物理段速度与行驶里程
- 堆场贝位热力图（离散三段配色：0–1 绿 / 1–11 橙 / ≥11 红）
- 分时车队活跃度、能耗、故障码 × 车辆矩阵（含采样点数）
- AT 箱时 > 60 分钟 · 箱清单
- 全场轨迹与时间回放（Canvas 热力图，支持滚轮缩放、拖拽平移）
- 停车热点 · 按时段 × 区域（岸桥区 / 堆场区 / 其他）
- 中英文双语切换

## 目录结构

```
.
├── assets/echarts.min.js        # ECharts 本地内联资源
├── scripts/
│   ├── build_report.py          # 数据计算 + HTML 生成（核心）
│   ├── template.html            # 报告前端模板
│   └── server.py                # 本地服务 + 上传重算 API
├── 启动报告.bat                  # Windows 一键启动脚本
└── data/                        # 输入数据（ZIP / XLSX，默认已忽略）
```

## 环境要求

- Python 3.9+
- `pandas`、`numpy`

```bash
pip install pandas numpy
```

## 快速开始（推荐）

Windows 双击 `启动报告.bat`，脚本会自动检查依赖、拉起本地服务并打开浏览器：

```
http://127.0.0.1:8760/
```

在页面右上角「上传」拖入新的 ZIP 与 XLSX，即可自动重算并更新报告。

## 命令行用法

```bash
# 从 ZIP + XLSX 生成报告
python scripts/build_report.py --zip <数据ZIP> --xlsx <位移XLSX> --out <输出HTML>

# 从已解压的 CSV 目录 + XLSX 生成报告
python scripts/build_report.py --data-dir <CSV目录> --xlsx <位移XLSX> --out <输出HTML>

# 仅启动本地服务
python scripts/server.py --port 8760 --open
```

## 输入数据

| 文件 | 说明 |
| --- | --- |
| `HIT_AT*_VehicleData*.csv` | 车辆遥测：位置、车速、里程、SOC/SOH 等 |
| `*EQUIPMENT_ENERGY_USAGE*.csv` | 能耗与充放电统计 |
| `*user_operation_log*.csv` | 用户操作日志 |
| `(IT61D) Daily AT Movement*.xlsx` | 箱位移与作业锚点 |

## 说明

- 采样间隔按 `SAMPLE_SEC = 3.0` 秒近似，用于将采样点数换算为时长。
- 停车热点区域采用「作业锚点最近归类 + 远离锚点 > 20 记为其他」的近似口径，面板内已注明。
- 生成的报告 `*.html` 默认不纳入版本管理，请通过构建脚本产出。
