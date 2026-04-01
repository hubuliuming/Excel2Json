# Excel 导表工具（Python版）

---

# 📌 项目简介

本工具用于将 Excel 配置表自动导出为：

* JSON 数据文件
* C# 配置类文件

支持：

* 按 Excel 中每个 Sheet 单独导出
* 基于 `ClassName` 自动控制导出
* GUI + bat 双入口
* 中文路径支持（通过 Python 读取配置）
* 自动复制 JSON 到客户端目录

---

# 🧰 环境依赖（必须先安装）

## 1️⃣ Python

推荐版本：

```text
Python 3.8+
```

检查是否安装：

```bash
python --version
```

或：

```bash
py -3 --version
```

---

## 2️⃣ Python依赖库

安装：

```bash
pip install openpyxl
```

---

## 3️⃣ Windows 环境要求

* 支持 `.bat` 执行
* 支持 `.pyw` 运行（GUI）

如果 `.pyw` 双击打不开：

```bash
python excel_export_launcher.pyw
```

---

# 🧩 目录结构（推荐）

```text
motoconf/
├─ 游戏配置/                 # Excel 配置目录
│  ├─ xxx.xlsx
│  ├─ yyy.xlsx
│
├─ Tools/
│  ├─ excel_sheet_exporter.py
│  ├─ excel_export_launcher.pyw
│  ├─ excel_export_launcher_config.txt
│  ├─ run_export_from_config.bat
│  ├─ json/
│  └─ cs/
│
├─ client/
│  └─ JsonData/
```

---

# 📊 Excel 表规则（当前使用规则）

## ✅ 可导出 Sheet 条件

Sheet 第 1 行必须包含：

```text
ClassName
```

否则该 Sheet 会被忽略。

---

## 🧱 行结构

| 行号   | 内容                  |
| ---- | ------------------- |
| 第1行  | 字段名                 |
| 第2行  | 类型（ClassName列用于输出名） |
| 第3行+ | 数据                  |

---

## 🧾 示例

```text
id    name    desc    ClassName
int   string  string  LevelConfig
1     关卡1   新手关
2     关卡2   普通关
```

---

# 📦 输出规则

如果：

```text
ClassName列第2行 = LevelConfig
```

则生成：

```text
LevelConfig.json
LevelConfig.cs
```

---

# ⚠️ ClassName 列说明

* 用于控制是否导出
* 用于定义输出文件名
* 不参与字段生成

---

# 🔢 支持的数据类型

* int
* long
* float
* double
* bool
* string
* json
* 数组：`int[]` / `string[]`

---

# 🚀 使用方式

---

## 🖥️ 方法1：GUI（推荐配置）

### 启动

双击：

```text
excel_export_launcher.pyw
```

或：

```bash
python excel_export_launcher.pyw
```

---

### 操作流程

1. 配置路径
2. 点击 `Run Export`
3. 确认正常
4. 点击 `Save Config`

生成：

```text
excel_export_launcher_config.txt
```

---

## ⚡ 方法2：bat 一键执行（推荐日常使用）

双击：

```text
run_export_from_config.bat
```

---

## 🧠 原理

* bat 不解析配置（避免中文乱码）
* Python 读取 `excel_export_launcher_config.txt`
* 调用导表脚本执行

---

# ⚙️ 配置文件说明

文件：

```text
excel_export_launcher_config.txt
```

示例：

```text
excel_folder=C:/dev/U3D/motoconf/游戏配置
json_folder=C:/dev/U3D/motoconf/Tools/json
cs_folder=C:/dev/U3D/motoconf/Tools/cs
client_json_folder=C:/dev/U3D/client/JsonData
namespace=Game.Config
copy_to_client=1
```

---

# 📁 输出结果

## JSON

```text
Tools/json/*.json
```

## C#

```text
Tools/cs/*.cs
```

## Client 同步（如果存在）

```text
client/JsonData/
```

---

# ❗ 常见问题

---

## ❌ 导出0个Sheet

```text
[DONE] Exported 0 sheet(s)
```

原因：

* 没有 `ClassName`
* 拼写错误
* 没有填写输出名

---

## ❌ 类型错误

```text
unsupported type 'xxx'
```

原因：

* 类型写错
* 行顺序不符合规则

---

## ❌ 中文路径乱码

原因：

* bat 解析中文路径失败

解决：

✅ 使用配置文件 + Python读取（已解决）

---

## ❌ client目录不存在

不是错误：

* 会自动跳过复制
* 不影响导表

---

# 🔄 工作流

## 初次配置

1. 打开 GUI
2. 配路径
3. Run Export
4. Save Config

---

## 日常使用

```text
双击 run_export_from_config.bat
```

---

# 🧠 推荐实践

* GUI 负责配置
* bat 负责执行
* 配置统一由 txt 管理

---

# 🚀 后续可扩展

建议增加：

* 重名检测（防覆盖）
* 主键唯一性校验
* 外键校验
* Unity一键导表入口

---

# 🎯 总结

当前工具链已经具备：

* 自动导表
* 配置驱动
* GUI + bat 双入口
* 中文路径支持
* Unity接入基础能力

推荐使用方式：

👉 平时直接双击 bat 导表
👉 需要改路径时用 GUI 配置

---
