Attribute VB_Name = "Module1"
Option Explicit

'This is a PDF to Text conversion macro
Sub WordToPDF()

    'Define the variables
    Dim InputWordFileName As String
    Dim BaseFolderPath As String
    Dim WordExtension As String
    Dim OutputPDFFile As String
    
    Dim objWordApp As Word.Application
    Dim objMyWordFile As Word.document
    Set objWordApp = CreateObject("Word.Application")
    
    'Add Error Handler to catch error, if any
    On Error GoTo ErrHandler
    
    'Clear the cells where the error description and number will be updated, when encountered
    ThisWorkbook.Worksheets("Sheet1").Range("B3").Value = ""
    ThisWorkbook.Worksheets("Sheet1").Range("B4").Value = ""
    
    'Retrieve input values: Word filename and Basepath from cell B1 and B2
    InputWordFileName = ThisWorkbook.Worksheets("Sheet1").Range("B1").Value
    BaseFolderPath = ThisWorkbook.Worksheets("Sheet1").Range("B2").Value
    
    'Determine the file extension - .doc or .docx
    WordExtension = Right(InputWordFileName, Len(InputWordFileName) - InStrRev(InputWordFileName, "."))
    MsgBox ("Word Ext = " & WordExtension)
    
    'Open the word file
    Set objMyWordFile = objWordApp.documents.Open(BaseFolderPath & InputWordFileName)
    objWordApp.Visible = True
    
    'Create file name with .pdf extension
    OutputPDFFile = BaseFolderPath & Replace(objMyWordFile.Name, WordExtension, "pdf")
    
    'Convert word file to PDF and save it in the basepath
    objWordApp.activedocument.ExportAsFixedFormat OutputFileName:=OutputPDFFile, ExportFormat:=wdExportFormatPDF
    
    'Close the Word file and word application
    objMyWordFile.Close
    objWordApp.documents.Application.Quit
    
    Exit Sub
    
ErrHandler:
    'On error, save the error details
    ThisWorkbook.Worksheets("Sheet1").Range("B3").Value = Err.Description
    ThisWorkbook.Worksheets("Sheet1").Range("B4").Value = Err.Number
    
End Sub
