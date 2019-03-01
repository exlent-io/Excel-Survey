# Excel-Survey
Library which be able to manipulate xls, xlsx, ods


|  Project  | Language | Open Source | License | Platform | Feature |  |
|---|---|---|---|---|---|---|
| Apache POI | JAVA  | Apache License 2.0  |   |   |
| Microsoft/Open XML SDK  | C#  |   |   |   |
|   |   |   |   |   |
|   |   |   |   |   |
|   |   |   |   |   |
|   |   |   |   |   |
|   |   |   |   |   |
|   |   |   |   |   |
|   |   |   |   |   |

最近在摸office相關格式
目前研究下來，免費的
Java端Apache的Poi
C#端Windows的Open XML
是比較堪用的

Poi對於某些高階操作還是需要手動實現，像是跨Workbook複製Sheet，因為其實有些東西像是圖片和公式是Workbook層級的，就要一起複製到

Open XML


Aspose.Cells 



exceljs 讀寫檔pass through 會遺失跨儲存格的文字藝術師啥的樣式
xlsx-populate pass through沒事，問題是對於一些插入和複製之類的操作支援性超差
等等來試試看protobi/js-xlsx
再不行就用libreoffice SDK硬幹了



The Apache POI opensource project provides Java libraries for reading and writing Excel spreadsheet files. ExcelPackage is another open-source project that provides server-side generation of Microsoft Excel 2007 spreadsheets. PHPExcel is a PHP library that converts Excel5, Excel 2003, and Excel 2007 formats into objects for reading and writing within a web application. Excel Services is a current .NET developer tool that can enhance Excel's capabilities. Excel spreadsheets can be accessed from Python with xlrd and openpyxl. js-xlsx and js-xls can open Excel spreadsheets from JavaScript.

開源類別庫支援在 Microsoft Excel 應用程式以外的環境中開啟 Excel 試算表。 Excelize 是 Go 語言（Golang）編寫的一個用來操作 Office Excel 文件類別庫，可以使用它來讀取、寫入帶有複雜樣式的 XLSX 檔案。 Apache POI 開源專案提供用於讀取和寫入 Excel 試算表檔案的 Java 庫。 PHPExcel 是一個 PHP 語言的實現，可在 Web 應用中讀取 Excel5，Excel 2003 和 Excel 2007 格式的文件。 Excel Services 是使用 .NET 開發的工具。 使用 xlrd 和 openpyxl 可以使用 Python 存取 Excel試算表。 js-xlsx 和 js-xls 可以使用 JS 開啟 Excel 試算表。
