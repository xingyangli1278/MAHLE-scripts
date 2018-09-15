' This script allows the user to select multiple CATIA 3D models at once
' 	 to mass convert them into step file format

option Explicit

Sub CATMain()

	Dim iDoc,activedoc As Document
	Dim iFile As Object
	Dim fileSystem As Object
	Set fileSystem = CATIA.FileSystem
	Dim inputFolder, outputFolder As String

	Dim excel As Object
	' Creates an Excel object to use the multi-select file dialog
  Set excel = CreateObject("Excel.Application")
  Dim files as Variant

	' Prompts the user to select which files they want to convert
  MsgBox ("Select the CATIA Parts and Products you want to convert.")
 	files= excel.GetOpenFilename("CAT files (*.CATPart; *.CATProduct), *.CATPart;*.CATProduct" , ,"Select the CATIA Parts and Products you want to convert.", ,True)

  ' Prompts the user to select where to output the converted files
	MsgBox("Select the output folder")
	outputFolder=FindFolder("Select the folder to which the STP files wil be saved")
	Dim fileFolder As Object
	Dim stpName As String
	Dim nameLength As Integer
	Dim Element As Variant

	' Iterates through each file selected to apply conversion process
	For Each element In files
      Set iDoc=CATIA.Documents.Open(element)
      Set activedoc = CATIA.ActiveDocument
			nameLength=Len(activedoc.Name)
			If InStr(activedoc.Name,"CATPart") <>0 Then
				stpName=Left(activedoc.Name, nameLength-8)
				activedoc.ExportData outputFolder & "\" & stpName, "stp"
				activedoc.Close
			ElseIf InStr(activedoc.Name,"CATProduct") <>0 Then
				stpName=Left(activedoc.Name, nameLength-11)
				activedoc.ExportData outputFolder & "\" & stpName, "stp"
				activedoc.Close
			End If
	Next

End Sub


' EFFECTS: gets the name of a directory that the user chooses
' RETURN TYPE: string
Public Function FindFolder(prompt As String) As String
	Const WINDOW_HANDLE =0
	Const NO_OPTIONs = &H1
	Dim objShellApp As Object
	Dim objFolder As Object
	Dim objFldrItem As Object

	Set objShellApp = CreateObject("Shell.Application")
	Set objFolder = objShellApp.BrowseForFolder(WINDOW_HANDLE,prompt,NO_OPTIONS)
	Set objFldrItem =objFolder.Self
	FindFolder=objFldrItem.Path


End Function
