# CreateExcelWorkbook
Base class for creating Excel workbooks

# Class Diagram
:::mermaid
classDiagram
    class WorkbookBuilder
    WorkbookBuilder : +ClosedXML.Excel.XLWorkbook BookObject
    WorkbookBuilder : +string FilePath
    WorkbookBuilder : +WorksheetBuilder[] worksheets

    WorkbookBuilder : +AddWorksheetBuilder(WorksheetBuilder)
    WorkbookBuilder : +GetWorksheet(SheetIndex)
    WorkbookBuilder : +GetWorksheet(SheetName)
    WorkbookBuilder : +Build()

    WorkbookBuilder --* WorksheetBuilder

    <<interface>> WorksheetBuilder
    WorksheetBuilder : +Reflect()

    class ReportWorkbook
    ReportWorkbook -- WorkbookBuilder
    ReportWorkbook --* SummaryWorksheet
    ReportWorkbook --* DetailWorksheet

    class SummaryWorksheet
    SummaryWorksheet : +ClosedXML.Excel.XLWorksheet SheetObject
    SummaryWorksheet : +object SummaryData
    SummaryWorksheet : +Reflect()

    SummaryWorksheet ..|> WorksheetBuilder

    class DetailWorksheet
    DetailWorksheet : +ClosedXML.Excel.XLWorksheet SheetObject
    DetailWorksheet : +object DetailData
    DetailWorksheet : +Reflect()

    DetailWorksheet ..|> WorksheetBuilder
:::

