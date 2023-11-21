# 软件名称：DOC文档批量重命名

## 1. 简介

DOC文档批量重命名工具是一个用于批量处理多个 Word 文档文件的工具，它可以根据用户指定的表格、行和列来获取单元格内容，并将其作为新的文件名进行重命名。

## 2. 功能特点

- 支持选择多个 Word 文档文件进行批量处理。
- 可以指定要处理的表格、行和列。
- 可以使用正则表达式对单元格内容进行匹配和提取。
- 在处理完成后会弹出提示框进行通知。

## 3. 使用方法

### 3.1 选择文件

点击 "选择文件" 按钮，选择您要处理的 Word 文档文件。可以选择一个或多个文件。

### 3.2 设置表格和单元格

在 "指定表格" 输入框中输入要处理的表格索引（从1开始计数）。如果您想处理第一个表格，输入1；如果要处理第二个表格，输入2，依此类推。

在 "指定行" 和 "指定列" 输入框中分别输入要获取数据的单元格的行号和列号。行号和列号都是从1开始计数。

### 3.3 设置正则表达式（可选）

如果您想对单元格的内容进行匹配或提取，可以在 "正则表达式" 输入框中输入适当的正则表达式。程序将会使用该正则表达式对单元格内容进行匹配，并将第一个匹配项作为新的文件名。

### 3.4 执行批量重命名

点击 "批量重命名" 按钮开始执行批量处理操作。程序将依次打开选中的 Word 文档文件，获取指定单元格的内容，并根据该内容进行重命名操作。处理过程中，进度条将显示处理进度。

### 3.5 处理完成提示

当所有文件处理完成后，系统会弹出一个提示框，通知您任务完成。同时，程序将隐藏窗口并等待下一次操作。

## 4. 注意事项

- 请确保所选文件都是有效的 Word 文档文件（.docx 格式）。
- 在设置表格、行和列时，请确保输入的索引号正确。如果表格索引、行号或列号不存在，程序将忽略该项操作并继续执行下一个文件的处理。
- 使用正则表达式时，请确保您输入的正则表达式语法正确，以避免出现意外的匹配结果。
- 在使用本工具期间，请不要关闭主窗口，否则无法正常完成任务。
- 使用之前建议先单个测试效果是否符合预期，建议留存备份，确定无问题在删除备份，这是个好习惯。
