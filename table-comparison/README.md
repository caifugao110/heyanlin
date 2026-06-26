<p align="center">
  <h1 align="center">table-comparison</h1>
  <p align="center">
    <img src="https://img.shields.io/github/v/tag/caifugao110/heyanlin?filter=table-comparison%20V*&color=blue&label=version" alt="version">
    <img src="https://img.shields.io/badge/python-%3E%3D3.9-green" alt="python">
    <img src="https://img.shields.io/badge/license-MIT-yellow" alt="license">
    <img src="https://img.shields.io/badge/platform-Windows-lightgrey" alt="platform">
  </p>
  <p align="center">
    <i>A polished Python GUI tool for comparing Excel files and highlighting differences.</i>
  </p>
</p>

---

## 简介

**table-comparison** 是一款基于 Python 的桌面 GUI Excel 文件比较工具，用于比较两个 Excel 文件，自动标记差异内容，包括数值变化、新增行和删除行，帮助用户快速识别数据差异。

| 项目信息 |  |
| --- | --- |
| 作者 | **Tobin** |
| 项目地址 | [github.com/caifugao110/heyanlin/tree/master/table-comparison](https://github.com/caifugao110/heyanlin/tree/master/table-comparison) |
| 开源协议 | MIT |

---

## 功能特性

### 文件比较

- 支持选择 **基准文件**（改动前/原始文件）和 **比较文件**（改动后/新文件）
- 使用 `data_only=True` 模式只读取数据不加载公式
- 可视化预览前10行数据，点击选择表头所在行号（默认为第3行）
- 可视化勾选用于匹配行的关键字段列（1~6列），用于判断行的增删变化

### 颜色标记

- **数值变化**（黄色）
  - 颜色代码：`FFFF00`
  - 应用场景：当两个文件中匹配行的相同列值不同时
  - 标记方式：同时标记基准文件和比较文件中的差异单元格

- **删除行**（绿色）
  - 颜色代码：`00FF00`
  - 应用场景：基准文件中有但比较文件中没有的行（基准文件中的删除行）
  - 标记方式：整行标记为绿色

- **新增行**（红色）
  - 颜色代码：`FF0000`
  - 应用场景：比较文件中有但基准文件中没有的行（比较文件中的新增行）
  - 标记方式：整行标记为红色

### 结果输出

- 生成三个结果文件，保存到 `results` 文件夹：
  - 基准文件带标记：`原始文件名_my_比较结果_<时间戳>.xlsx`
  - 比较文件带标记：`原始文件名_from_比较结果_<时间戳>.xlsx`
  - 差异结果文件：`原始文件名_差异结果_<时间戳>.xlsx`
- 结果文件默认设置为只读属性，防止误修改
- 运行完成后自动打开生成的结果文件

### 主题支持

- **外观模式**
  - Light Mode：明亮主题
  - Dark Mode：暗黑主题
  - System Mode：跟随系统主题

- **颜色主题**
  - Blue：蓝色主题（默认）
  - Green：绿色主题
  - Dark Blue：深蓝主题

### 性能优化

- 预先将所有单元格值加载到内存中，提高后续访问速度
- GUI采用多线程设计，避免界面卡顿

---

## 工作流程

1. **文件加载**：选择基准文件（改动前/原始文件）和比较文件（改动后/新文件），使用 `data_only=True` 模式只读取数据不加载公式
2. **表头行号选择**：可视化预览前10行数据，点击选择表头所在行号（默认为第3行）
3. **特征列选择**：可视化勾选用于匹配行的关键字段列（1~6列），用于判断行的增删变化
4. **列名匹配**：按表头行读取列名，在基准文件和比较文件中建立列名到列号的映射
5. **行映射构建**：基于关键字段值组合进行行匹配；若关键字段不全则回退到全行内容匹配；若匹配行数不足一半则回退到按行号对齐
6. **单元格差异比较**：遍历匹配行对，比较非关键字段列的单元格值，不一致的标记为黄色，特征列内容的变化不视为数值变化
7. **新增/删除行标记**：基准文件有而比较文件无的行标记为绿色（删除行）；比较文件有而基准文件无的行标记为红色（新增行）
8. **差异结果文件生成**：以基准文件为模板，将比较文件中的新增行按位置插入，生成合并差异结果文件
9. **结果保存**：将三个结果文件保存到 `results` 文件夹，自动设置只读属性
10. **自动打开**：自动打开生成的三个结果文件

---

## 快速开始

### 环境要求

- Python >= 3.9
- Windows 操作系统

### 项目依赖

**外部库**

- **openpyxl**：用于Excel文件的读取、写入和样式设置
- **ttkbootstrap**：用于构建现代化的GUI界面（基于tkinter的主题框架）
- **pillow**：用于GUI界面的图标处理
- **customtkinter**：辅助UI组件支持

### 直接运行源码

```powershell
pip install -r requirements.txt
python .\app.py
```

---

## 构建

### 打包为单文件 exe

```powershell
.\scripts\build_exe.ps1
```

构建完成后保留产物：

```
dist\table-comparison.exe
```

> 构建脚本会自动创建临时虚拟环境、安装依赖、生成图标、调用 PyInstaller，并在结束后清理 `.venv`、`build`、spec 文件、缓存和临时报告等过程文件。

---

## 项目结构

```
table-comparison/
├── app.py                  # GUI 主程序，包含全部核心逻辑
├── assets/
│   └── app.ico             # 应用图标
├── scripts/
│   └── build_exe.ps1       # Windows 构建脚本
├── .gitignore
├── LICENSE
├── pyproject.toml
├── README.md
└── requirements.txt
```

| 条目 | 说明 |
| --- | --- |
| `app.py` | GUI 主程序，包含全部核心逻辑 |
| `assets/` | 图标资源 |
| `scripts/` | 构建脚本 |
| `pyproject.toml` | 项目元数据与依赖配置 |
| `requirements.txt` | pip 依赖清单 |

---

## 注意事项

1. **特征列**：特征列用于判断行的增删变化，特征列内容的变化不视为数值变化，最多支持6列
2. **文件格式**：仅支持 `.xlsx` 格式的Excel文件
3. **公式处理**：读取文件时仅加载数据，不加载公式
4. **结果文件**：自动生成带时间戳的结果文件，避免覆盖现有文件
5. **只读属性**：结果文件默认设置为只读，防止误修改

---

## License

MIT © Tobin