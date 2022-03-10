    s
# doc2xlsx

> 下发通知文件为`word`格式，将回收大量含内嵌表格的`docx`文件。
> 
> 需汇总到`summmary.xlsx`并且根据`map.xlsx`映射添加其他字段内容。
> new: 根据类型输出分类汇集表格。
> new: 根据类型汇总附件到一起，包含所属单位等字段。
> new: 再根据固定类型二次分类然后二次汇总

```markdown
├── README.md
├── classify
│   └── Type
│       └── 分类型展品文件夹
│          └── 附件
│       └── type汇总表.docx
├── docx
│   └── GX-SB-XXXX-NumName
│       └── GX-SB-XXXX-NumName展品汇总表.docx
│       └── 展品文件夹含num
│          └── 附件
├── map
│   └── map.xlsx 
├── source
│   └── classification.py
│   └── classification_2.py
│   └── docx2xlsx_sum.py
│   └── merge_xlsx.py
│   └── path.json
│   └── summary_func.py
└── xlsx
    └── summary.xlsx
5 directories
```
