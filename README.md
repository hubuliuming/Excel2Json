# Excel 导表工具（Python版，含 JSON 上传）

## 项目简介

本工具用于将 Excel 配置表导出为：

- JSON 数据文件
- C# 配置类文件

并支持：

- 按每个 Excel 的每个 Sheet 单独导出
- 通过 `ClassName` 控制是否导出
- GUI 启动器执行导表
- bat 一键执行
- 导出后自动上传 `json_folder` 下所有 `.json` 到服务器
- client 输出目录不存在时自动跳过复制

---

## 环境依赖

### Python
推荐：

```text
Python 3.8+
```

检查：

```bash
python --version
```

或：

```bash
py -3 --version
```

### Python 依赖
安装：

```bash
pip install openpyxl
```

> 上传功能使用 Python 标准库 `urllib`，不需要额外安装 `requests`。

---

## 推荐目录结构

```text
motoconf/
├─ 游戏配置/
│  ├─ xxx.xlsx
│  └─ yyy.xlsx
├─ Tools/
│  ├─ excel_sheet_exporter.py
│  ├─ excel_export_launcher.pyw
│  ├─ excel_export_launcher_config.txt
│  ├─ run_export_from_config.bat
│  ├─ json/
│  └─ cs/
└─ client/
   └─ JsonData/
```

---

## Excel 规则

### 可导出 Sheet 条件
Sheet 第 1 行必须包含：

```text
ClassName
```

否则该 Sheet 会被忽略。

### 行结构
| 行号 | 内容 |
|---|---|
| 第1行 | 字段名 |
| 第2行 | 类型 |
| 第3行+ | 数据 |

### 示例

```text
id    name    desc    ClassName
int   string  string  LevelConfig
1     关卡1   新手关
2     关卡2   普通关
```

### 输出规则
如果 `ClassName` 列第 2 行值为：

```text
LevelConfig
```

则生成：

```text
LevelConfig.json
LevelConfig.cs
```

### ClassName 列说明
- 用于判定是否导出
- 用于决定输出文件名
- 不参与 JSON / C# 字段导出

---

## 支持的数据类型

- int
- long
- float
- double
- bool
- string
- json
- 数组类型：`int[]`、`string[]`

---

## 默认 C# 命名空间

默认命名空间为：

```text
GameConfig
```

---

## 上传服务器规则

当启用上传功能后，会遍历 `json_folder` 下所有 `.json` 文件，逐个上传到配置的服务器地址。

### 上传内容结构

每个 JSON 文件都会发送一个 HTTP POST 请求，请求体为 JSON：

```json
{
  "key": "当前json文件名",
  "desc": "测试",
  "context": "当前json文件内容"
}
```

### 字段说明

- `key`: 当前 `.json` 文件名，例如 `LevelConfig.json`
- `desc`: 固定值，默认 `测试`
- `context`: 当前 `.json` 文件完整文本内容

### 上传结果打印

每个文件都会依次打印：

- 成功：
```text
[UPLOAD OK] LevelConfig.json -> 200
```

- 失败：
```text
[UPLOAD FAIL] LevelConfig.json -> HTTP 500 | ...
```

最后汇总：

```text
[UPLOAD DONE] success=3, fail=1
```

---

## GUI 使用方式

### 启动

双击：

```text
excel_export_launcher.pyw
```

或：

```bash
python excel_export_launcher.pyw
```

### GUI 配置项

- Excel Folder
- JSON Folder
- C# Folder
- Client JSON Folder
- Namespace
- Upload all JSON files in JSON Folder after export
- Upload URL
- Upload Desc
- Upload Timeout

### 操作流程

1. 配置路径
2. 如需上传，勾选上传开关并填写 Upload URL
3. 点击 `Run Export`
4. 确认正常
5. 点击 `Save Config`

配置会保存到：

```text
excel_export_launcher_config.txt
```

---

## 配置文件说明

示例：

```text
excel_folder=C:/dev/U3D/motoconf/游戏配置
json_folder=C:/dev/U3D/motoconf/Tools/json
cs_folder=C:/dev/U3D/motoconf/Tools/cs
client_json_folder=C:/dev/U3D/client/JsonData
namespace=GameConfig
copy_to_client=1
enable_upload=1
upload_url=http://127.0.0.1:8080/upload
upload_desc=测试
upload_timeout=15
```

### 字段说明

- `excel_folder`: Excel 目录
- `json_folder`: JSON 输出目录
- `cs_folder`: C# 输出目录
- `client_json_folder`: 客户端 JSON 目录
- `namespace`: C# 命名空间
- `copy_to_client`: 是否复制 JSON 到客户端目录，`1` 开启
- `enable_upload`: 是否启用上传，`1` 开启
- `upload_url`: 上传接口地址
- `upload_desc`: 上传 payload 的 `desc`
- `upload_timeout`: 上传超时时间，单位秒

---

## bat 一键执行

双击：

```text
run_export_from_config.bat
```

bat 不直接解析中文配置，而是调用 Python 用 UTF-8 读取配置文件后执行导表和上传。

---

## 输出结果

### JSON 输出目录
```text
Tools/json/*.json
```

### C# 输出目录
```text
Tools/cs/*.cs
```

### Client 同步
如果 `client_json_folder` 存在，则自动复制 JSON。

---

## 常见问题

### 导出完成但没有文件
日志出现：

```text
[DONE] Exported 0 sheet(s).
```

说明没有符合规则的 Sheet。检查：

- 第 1 行是否存在 `ClassName`
- `ClassName` 拼写是否正确
- `ClassName` 列第 2 行是否写了输出名

### 上传没有执行
检查：

- GUI 里是否勾选了上传开关
- `upload_url` 是否填写
- bat 读取的配置文件里 `enable_upload=1`

### client 目录不存在
不是错误，会自动跳过复制。

### 中文路径乱码
当前方案已规避：配置文件由 Python 读取，不由 bat 直接解析。

---

## 推荐工作流

### 初次配置
1. 打开 GUI
2. 配置导表路径
3. 配置上传地址
4. Run Export
5. Save Config

### 日常使用
直接双击：

```text
run_export_from_config.bat
```

---

## 后续建议

下一步建议补充：

- 输出重名检测
- 主键重复校验
- 外键引用校验
- Unity 菜单一键导表
