# PDF 导出字体

`NotoSansSC-Regular.ttf` 用于 PDF 导出中 Helvetica 无法编码的文字：中文等表单值的外观，以及扁平化附注。导出时只嵌入用到的字形。

- 来源：Google Fonts `ofl/notosanssc/NotoSansSC[wght].ttf`
  - SHA-256 `a3041811a78c361b1de50f953c805e0244951c21c5bd412f7232ef0d899af0da`
- 生成：使用 fontTools 4.60.1，命令：
  ```
  instantiateVariableFont(font, {"wght": 400})
  ```
  然后删除 `STAT`、`MVAR`、`HVAR`、`avar`、`gvar`、`fvar` 表。
  - 原因：可变字体的默认实例是 Thin（100），而 fontkit 只能子集化静态 TrueType 字形。
  - 产物 SHA-256 `06270f11c48d386531a73f78d08cb28c097b62ff4368b235a36bb1293dbf3f95`
- 许可：SIL Open Font License 1.1，见 `OFL.txt`。
  - 可以随软件分发，也可以嵌入导出的 PDF。
  - 修改后的版本不得使用保留字体名称（Reserved Font Name）。

不包含表情符号等字形。遇到这些字符时，导出会明确报错，而不是显示为空白。
