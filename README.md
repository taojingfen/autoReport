# autoReport

autoReport is a system to auto-generate PPT report.

## Design Concepts

autoReport iss designed to generate report from Excel file (.xlsx) with specific columns, it is construct by an *tkinter* module. It currently contains below steps:

For purification report:

| Module         | Function                 |
| -------------- | ------------------------ |
| coverPage.py   | 生成报告封面             |
| finalPage.py   | 生成报告每条蛋白最终信息 |
| processPage.py | 生成报告蛋白纯化的步骤   |
| stepPage.py    | 生成报告每步蛋白的信息   |
| sdsPage.py     | 生成报告SDS信息          |
| hplcPage.py    | 生成报告HPLC信息         |

