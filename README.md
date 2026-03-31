# 修正说明

这版修正了旧表布局下 `ClassName` 列的识别问题。

之前脚本会去 `type_row` 找 `ClassName`，所以在旧布局里会找不到任何可导出 Sheet，最后显示：

`[DONE] Exported 0 sheet(s).`

现在逻辑改为：
- 在 `header_row` 查找 `ClassName` 列
- 用同列的 `type_row` 单元格值作为输出文件名

适用于当前规则：
- 第 1 行：字段名（包含 `ClassName`）
- 第 2 行：类型（`ClassName` 列下写输出名）
- 第 3 行开始：数据
