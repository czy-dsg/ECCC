# ECC 受压自感知曲线生成器

根据受压自感知的定性规律，生成一组同一加载过程对应的四联图：

1. 荷载–位移曲线
2. 电阻率–时间曲线
3. 相对电阻变化率（FCR）–时间曲线
4. 应力–时间曲线

并可导出对应的 Excel 数据表。

> **重要声明**
> 本项目的电阻响应是**示意模型**，依据自感知水泥基复合材料在单轴受压下的常见趋势构造：
> 压密阶段电阻率下降，接近峰值及峰后因损伤发展而回升。
> 它**不是实测数据**，不能作为试验结果、灵敏度指标或论文结论引用。
> 若用于论文，请替换为你自己的实测数据，或在图注中明确标注“示意图”。

## 环境要求

- Python 3.9+
- 依赖见 `requirements.txt`：

```bash
pip install -r requirements.txt
```

中文标签需要系统安装中文字体（Windows 自带微软雅黑/宋体即可）。其他系统会自动回退到可用字体，必要时可改为英文标签。

## 快速开始

使用内置示例曲线：

```bash
python src/ecc_curves.py --synthetic --outdir outputs
```

使用自己的荷载–位移数据（CSV，两列：位移 mm、荷载 kN，可带表头）：

```bash
python src/ecc_curves.py --input examples/sample_load_displacement.csv --outdir outputs
```

输出：

- `outputs/self_sensing_panels.png`：四联图，2400×1800，300 dpi
- `outputs/self_sensing_curves.xlsx`：数据表（time、displacement、load、stress、resistivity、FCR）

## 参数

| 参数 | 默认值 | 说明 |
|---|---|---|
| `--input` | 无 | 荷载–位移 CSV；省略则使用内置示例曲线 |
| `--outdir` | `outputs` | 输出目录 |
| `--rate` | `0.01` | 位移速率，mm/s；用于把位移换算为时间 |
| `--area` | `10000` | 截面面积，mm²（100×100 试件） |
| `--rho0` | `0.45` | 初始电阻率，Ω·m |
| `--noise` | `0.15,0.20,0.25` | 三段噪声标准差，单位为 ρ0 的百分比：压密段、峰前段、峰后段 |
| `--seed` | `42` | 随机种子，保证结果可复现 |
| `--peak-disp` | 自动 | 峰值位移，mm；默认取荷载最大处 |
| `--no-figure` | 关闭 | 只导出 Excel，不画图 |

## 计算关系

- 应力：σ(MPa) = P(kN) × 1000 / A(mm²)
- 电阻率：ρ = R·A/L（本仓库直接给出 ρ 的模型值）
- 相对电阻变化率：FCR = (ρ − ρ0) / ρ0 × 100%
- 时间：t = d / rate（仅当加载为等速位移时成立）

## 目录

```
src/ecc_curves.py                 生成器（模型、绘图、导出）
examples/sample_load_displacement.csv  示例荷载–位移数据
docs/preview.png                  生成效果预览
```

## 已知限制

- 电阻率锚点固定，未按材料配比、含水率、电极布置标定；
- 噪声为分段常数标准差，段边界可能出现小幅跳变；
- 未实现循环加载、卸载路径与温度补偿；
- 图示趋势来自文献定性规律，不代表某一具体配合比的定量结果。

## License

可自行选择。若公开分享，建议补充 MIT 或 CC BY 4.0 许可证。
