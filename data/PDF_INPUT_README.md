# PDF与分页文本输入

`data/`采用扁平命名，不按公司再建立子文件夹。

离线重放文件示例：

```text
data/001282_page_text.jsonl
data/301581_page_text.jsonl
data/920100_page_text.jsonl
```

文件名规则：

```text
<证券代码>_page_text.jsonl
```

原始招股说明书PDF未放入GitHub提交包。PDF文件名和SHA-256见
`data/pdf_manifest.csv`。需要从PDF重新解析时，可在仓库外准备PDF目录，
再通过对应阶段命令传入。
