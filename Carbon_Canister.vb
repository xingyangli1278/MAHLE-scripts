
' This script takes in input from the user through an Excel design table and then
'   creates carbon canisters with the user specified dimensions and saves the models
'   to a specified directory.
Sub CATMAIN()

    Dim strTableFilePath As String
    Dim strPath As String
    Dim strRange As String

    Dim oDoc As Document
    Set oDoc = CATIA.Documents.Open("P:\xfer\Yang\Scripts\Carbon_Canister\BasePart.CATPart")

    MsgBox("Select the Excel sheet that will act as your design table.")
    strPath = CATIA.FileSelectionBox("Select design table:", "*xlsx", 0)
    If strPath = "" Then Exit Sub

   MsgBox("Select the folder where you will store the automated part files.")
    Dim outputPath As String
    outputPath= FindFolder("Select the save location:")
    Dim oActiveDoc As Document
    Set oActiveDoc = CATIA.ActiveDocument
    Dim oPart As Part
    Set oPart = oActiveDoc.Part

    Dim oRelations As Relations
    Set oRelations = oPart.Relations

    For i = 1 To oRelations.Count
        If oRelations.Item(i).Name = "DesignTable.1" Then
            oRelations.Remove oRelations.Item(i).Name
        End If
    Next


    ' Creates a relation between Excel sheet and design table in CATIA
    Dim oDesignTable As DesignTable
    Set oDesignTable = oRelations.CreateDesignTable("DesignTable.1", "", False, strPath)

    ' Links parameters in Excel sheet to model parameters in 3d model
    Dim oParameters As Parameters
    Set oParameters = oPart.Parameters
    Dim oTopLength As Length
    Set oTopLength = oParameters.Item("TopLength")
    oDesignTable.AddAssociation oTopLength, "TopLength"
    Dim oTopWidth As Length
    Set oTopWidth = oParameters.Item("TopWidth")
    oDesignTable.AddAssociation oTopWidth, "TopWidth"
    Dim oTRadiusTop As Length
    Set oTRadiusTop = oParameters.Item("TRadiusTop")
    oDesignTable.AddAssociation oTRadiusTop, "TRadiusTop"
	Dim oBRadiusTop As Length
    Set oBRadiusTop = oParameters.Item("BRadiusTop")
    oDesignTable.AddAssociation oBRadiusTop, "BRadiusTop"
    Dim oLRadiusTop As Length
    Set oLRadiusTop = oParameters.Item("LRadiusTop")
    oDesignTable.AddAssociation oLRadiusTop, "LRadiusTop"
	Dim oRRadiusTop As Length
    Set oRRadiusTop = oParameters.Item("RRadiusTop")
    oDesignTable.AddAssociation oRRadiusTop, "RRadiusTop"
    Dim oPadHeight As Length
    Set oPadHeight = oParameters.Item("PadHeight")
    oDesignTable.AddAssociation oPadHeight, "PadHeight"
    Dim oSpineHeight As Length
    Set oSpineHeight = oParameters.Item("SpineHeight")
    oDesignTable.AddAssociation oSpineHeight, "SpineHeight"

    Dim oBotLength As Length
    Set oBotLength = oParameters.Item("BotLength")
    oDesignTable.AddAssociation oBotLength, "BotLength"
    Dim oBotWidth As Length
    Set oBotWidth = oParameters.Item("BotWidth")
    oDesignTable.AddAssociation oBotWidth, "BotWidth"
    Dim oTRadiusBot As Length
    Set oTRadiusBot = oParameters.Item("TRadiusBot")
    oDesignTable.AddAssociation oTRadiusBot, "TRadiusBot"
	Dim oBRadiusBot As Length
    Set oBRadiusBot = oParameters.Item("BRadiusBot")
    oDesignTable.AddAssociation oBRadiusBot, "BRadiusBot"
    Dim oLRadiusBot As Length
    Set oLRadiusBot = oParameters.Item("LRadiusBot")
    oDesignTable.AddAssociation oLRadiusBot, "LRadiusBot"
	Dim oRRadiusBot As Length
    Set oRRadiusBot = oParameters.Item("RRadiusBot")
    oDesignTable.AddAssociation oRRadiusBot, "RRadiusBot"

	Dim R1 As Length
    Set R1 = oParameters.Item("R1")
    oDesignTable.AddAssociation R1, "R1"

	Dim R2 As Length
    Set R2 = oParameters.Item("R2")
    oDesignTable.AddAssociation R2, "R2"

	Dim R3 As Length
    Set R3 = oParameters.Item("R3")
    oDesignTable.AddAssociation R3, "R3"

	Dim R4 As Length
    Set R4 = oParameters.Item("R4")
    oDesignTable.AddAssociation R4, "R4"

	Dim FleeceOffset As Length
    Set FleeceOffset = oParameters.Item("FleeceOffset")
    oDesignTable.AddAssociation FleeceOffset, "FleeceOffset"

    ' opens Excel to output to the same design sheet
    Dim objExcel As Object
    Dim objWorkbook As Object
    Set objExcel = CreateObject("Excel.Application")
    Set objWorkbook = objExcel.Workbooks.Open(strPath)
    objExcel.Application.Visible = True

    Dim tempInteger As Integer
    tempInteger = 1 + oDesignTable.ConfigurationsNb
    strRange = "T2:" & "T" & tempInteger
    objExcel.Worksheets(1).Range(strRange).NumberFormat = "0.00"

    Set oPart = CATIA.ActiveDocument.Part
    Set oSPAWkb = CATIA.ActiveDocument.GetWorkbench("SPAWorkbench")
    Dim oBodies As bodies
    Set oBodies = oPart.bodies

    Dim partBody As body
    Dim prtBodyVolume As Double
    Dim x As Integer
    Dim strConfig As String
    Dim outputfile As String

    ' For each row in the Excel design sheet, use the parameters to change the 3D model
    ' and then output their volumes to the Excel design sheet
    For x = 1 To oDesignTable.ConfigurationsNb
    	    strConfig = CStr(x)
        oDesignTable.Configuration = x
        oPart.Update
        Set partBody = oBodies.Item("PartBody")
        Set objRef = oPart.CreateReferenceFromObject(partBody)
        Set objMeasurable = oSPAWkb.GetMeasurable(objRef)
        partBodyVolume = objMeasurable.Volume
        objExcel.Sheets("Sheet1").Cells(1 + x, 20).Value = partBodyVolume * 1000000000
	   outputfile=outputPath & "\" &"Config" & strConfig & ".CATPart"
        CATIA.ActiveDocument.SaveAs(outputfile)
	   Set oDoc=CATIA.Documents.Open( outputfile)
   	   Set oPart=CATIA.ActiveDocument.Part

    Next

End Sub

' EFFECTS: gets the name of the folder you want to find
' RETURN TYPE: string
Function FindFolder(ByRef prompt As String) As String
	Const WINDOW_HANDLE =0
	Const NO_OPTIONs = &H1
	Dim objShellApp As Object
	Dim objFolder As Object
	Dim objFldrItem As Object

  ' Opens the Windows Shell Application to use the folder dialog function
	Set objShellApp = CreateObject("Shell.Application")
	Set objFolder = objShellApp.BrowseForFolder(WINDOW_HANDLE,prompt,NO_OPTIONS)
	Set objFldrItem =objFolder.Self
	FindFolder=objFldrItem.Path


End Function
