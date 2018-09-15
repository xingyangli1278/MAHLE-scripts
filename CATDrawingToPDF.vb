' This script converts 2D CATIA Drawings that the user selects into pdfs and saves them to a selected folder

Option Explicit

Sub CATMain()

    Dim fileSystem As Object
    Set fileSystem = CATIA.fileSystem

    ' Creates Excel Object to use the multi-select file dialog
    Dim excel As Object
    Set excel = CreateObject("Excel.Application")
    Dim files As Variant

    Dim outputFolder As String

    ' Prompts user to select the 2D CATIA Drawing files
    MsgBox ("Select the CATDrawings you want to convert.")
    files = excel.GetOpenFilename("CATDrawings (*.CATDrawing), *.CATDrawing", ,"Select the CATDrawings.", ,True)

    ' Prompts user to select where to save the converted pdf files
    MsgBox ("Select the output folder")
    outputFolder = GetFolderPath("Select the folder to which the PDFs will be saved (output)")


    Dim iFile As Object
    Dim iDoc, activeDoc As Document

    Dim  pdfName As String
    Dim catDrawingName As Integer
    Dim element As Variant

    ' Iterates through the list of selected files to apply conversion process
    For Each element In files

            Set iDoc = CATIA.Documents.Open(element)
            Set activeDoc = CATIA.ActiveDocument

            catDrawingName = Len(activeDoc.Name)
            pdfName = Left(activeDoc.Name, catDrawingName - 11)

            activeDoc.ExportData outputFolder & "\" & pdfName, "pdf"
            activeDoc.Close

    Next

End Sub

' EFFECTS: gets the name of the directory that the user selects in a file dialog
' RETURN TYPE: string
Public Function GetFolderPath(strTitle As String) As String

    Const WINDOW_HANDLE = 0
    Const NO_OPTIONS = &H1
    Dim objShellApp As Object
    Dim objFolder As Object
    Dim objFldrItem As Object

    Set objShellApp = CreateObject("Shell.Application")
    Set objFolder = objShellApp.BrowseForFolder(WINDOW_HANDLE, strTitle, NO_OPTIONS)
    Set objFldrItem = objFolder.Self

    GetFolderPath = objFldrItem.Path

End Function
