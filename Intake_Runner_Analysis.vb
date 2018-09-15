' This script reads a intake runner model and calculates the cross-sectional area
'    of however many sections the user specifies and then exports the data to an Excel sheet.

Option Explicit

Public activeDoc As PartDocument
Public part1 As Part
Public partBody As Body
Public hybridBodies1 As HybridBodies
Public geoSet1, geoSet2 As HybridBody
Public hybridShapeFactory1 As HybridShapeFactory
Public planeIntersection As HybridShapeIntersection
Public intersectionFill As HybridShapeFill
Public relations As Relations

Public centerLine As Object 'Hybrid Shape Intersection
Public refPoint As HybridShapePointOnCurve

Public objExcel As Object
Public rowCounter As Integer

Sub CATMain()

    Dim path As String

    ' Prompts user to select an intake runner file and opens it
    MsgBox ("Select an intake runner part.")
    path=CATIA.FileSelectionBox("Select the runner part.","*.CATPart; *.CATProduct" , 0)
    CATIA.Documents.Open(path)
    Set activeDoc = CATIA.ActiveDocument

    Set part1 = activeDoc.Part
    Set hybridBodies1 = part1.HybridBodies
    Set hybridShapeFactory1 = part1.HybridShapeFactory

    ' Prompts the user to select the runner's centerline.
    MsgBox ("Select the runner's center line (intersect)")
    Set centerLine = SelectCenterLine("HybridShapeIntersection", "HybridShapeCurveExplicit","Select the center line (intersect)")

    ' Prompts user to select the runner body
    MsgBox ("Select the runner's PartBody")
    Set partBody = SelectBody("Body", "Select the PartBody")

    ' Creates the Geometrical Sets in which created geometry will be placed.
    Set geoSet1 = hybridBodies1.Add()
    geoSet1.Name = "Points and Planes"

    Set geoSet2 = hybridBodies1.Add()
    geoSet2.Name = "Cross Sectional Areas"


    ' Opens an Excel book to which data is exported.
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Workbooks.Add

	  objExcel.Sheets(1).Cells(1,1).Value="Section Number"
  	objExcel.Sheets(1).Cells(1,1).WrapText=True
    objExcel.Sheets(1).Cells(1, 2).Value = "DistFromRefPoint (mm)"
    objExcel.Sheets(1).Cells(1, 2).WrapText = True
    objExcel.Sheets(1).Cells(1, 3).Value = "Area (mm^2)"
    objExcel.Sheets(1).Cells(1, 3).WrapText = True

    Call AddPointOnCurve(centerLine)
    Call RepetitionPointsAndPlanes(centerLine,refPoint)

End Sub


' EFFECTS: Creates a reference point at one end of the runner's centerline
' RETURN TYPE: none
Public Sub AddPointOnCurve(ByRef centerLine As Object)

    Set refPoint = hybridShapeFactory1.AddNewPointOnCurveFromDistance(centerLine, 0, False)
    refPoint.DistanceType = 1

    geoSet1.AppendHybridShape refPoint
    part1.Update

End Sub


' EFFECTS: adds points onto the centerline determined by the user-inputted cuts to be made.
'             Then, creates a normal plane to the line at each point.
' RETURN TYPE: none
Public Sub RepetitionPointsAndPlanes(ByRef centerLine As Object, ByRef refPoint As Object)

    Dim refBody As Reference
    Set refBody = part1.CreateReferenceFromObject(partBody)

    Dim addPoint As hybridShapePointOnCurve
    Dim addPlane As HybridShapePlaneNormal
    Dim i, numCuts As Integer
    Dim division As Double

    numCuts = InputBox("Enter the number of cuts to be made: ")
    division = GetCurveLength(centerLine) / (numCuts + 1)

    '1st Plane on reference point
    Set addPlane = hybridShapeFactory1.AddNewPlaneNormal(centerLine, refPoint)
    geoSet1.AppendHybridShape addPlane
    part1.Update

    objExcel.Sheets(1).Cells(2, 1).Value = 1
    objExcel.Sheets(1).Cells(2, 2).Value = 0
    rowCounter = 2

    Call CreatePlaneBodyIntersection(addPlane, refBody,refPoint)

    For i = 1 To numCuts + 1

        rowCounter = rowCounter + 1

        Set addPoint = hybridShapeFactory1.AddNewPointOnCurveWithReferenceFromDistance(centerLine, refPoint, (i * division), False)
        addPoint.DistanceType = 1
        geoSet1.AppendHybridShape addPoint
        part1.Update

        Set addPlane = hybridShapeFactory1.AddNewPlaneNormal(centerLine, addPoint)
        geoSet1.AppendHybridShape addPlane
        part1.Update

        objExcel.Sheets(1).Cells(rowCounter, 1).Value = i+1
        objExcel.Sheets(1).Cells(rowCounter, 2).Value = i * division

        Call CreatePlaneBodyIntersection(addPlane, refBody, addPoint)

    Next

    objExcel.Application.Visible = True

End Sub


' EFFECTS: creates an intersection between each plane and the runner cross section at that point
' RETURN TYPE: none
Public Sub CreatePlaneBodyIntersection(ByRef refPlane As Object, ByRef refBody As Object, ByRef Point As Object)

    Set planeIntersection = hybridShapeFactory1.AddNewIntersection(refPlane, refBody)
    planeIntersection.PointType = 0
    'planeIntersection.SolidMode= 1
    'geoSet2.AppendHybridShape planeIntersection

    Dim ref1 As Reference
    Set ref1=part1.CreateReferenceFromObject(planeIntersection)
    Dim ref2 As Reference
    Set ref2= part1.CreateReferenceFromObject(Point)
    Dim Near As HybridBodyNear
    Set Near = hybridShapeFactory1.AddNewNear(ref1,ref2)
    geoSet2.AppendHybridShape Near
    part1.Update
    part1.InWorkObject=Near

    Call FillIntersection(Near)

End Sub

' EFFECTS: fills the intersection created to form a recognizable cross section
' RETURN TYPE: none
Public Sub FillIntersection(ByRef refNear As Object)

    Set intersectionFill = hybridShapeFactory1.AddNewFill()
    intersectionFill.AddBound refNear

    intersectionFill.Continuity = 1
    intersectionFill.Detection = 2
    intersectionFill.AdvancedTolerantMode = 2

    geoSet2.AppendHybridShape intersectionFill
    part1.Update

    Call ExportValueToExcel(intersectionFill)

End Sub

' EFFECTS: calculates the cross sectional areas and exports data to an excel sheet
' RETURN TYPE: none
Public Sub ExportValueToExcel(ByRef refFill As Object)

    Dim TheSPAWorkbench As Workbench
    Set TheSPAWorkbench = activeDoc.GetWorkbench("SPAWorkbench")

    Dim fillArea As Measurable
    Set fillArea = TheSPAWorkbench.GetMeasurable(refFill)

    objExcel.Sheets(1).Cells(rowCounter, 3).Value = fillArea.Area * 1000000

End Sub

' EFFECTS: allows the user to select the runner's centerline
' RETURN TYPE: object
Public Function SelectCenterLine(ByRef filterType1 As String, ByRef filterType2 As String, filterMessage As String) As Object

    Dim arrInputType(1) As Variant
    arrInputType(0) = filterType1
    arrInputType(1) = filterType2

    Dim selection1 As Object
    Set selection1 = activeDoc.Selection
    selection1.Clear

    Dim status As String
    status = selection1.SelectElement2(arrInputType, filterMessage, False)

    If status = "cancel" Then
        Exit Function
    Else
        Set SelectCenterLine = selection1.Item(1).Value
    End If

End Function


' EFFECTS: Allows the user to select the runner body
' RETURN TYPE: object
Public Function SelectBody(ByRef filterType As String, filterMessage As String) As Object

    Dim arrInputType(0) As Variant
    arrInputType(0) = filterType

    Dim selection1 As Object
    Set selection1 = activeDoc.Selection
    selection1.Clear

    Dim status As String
    status = selection1.SelectElement2(arrInputType, filterMessage, False)

    If status = "cancel" Then
        Exit Function
    Else
        Set SelectBody = selection1.Item(1).Value
    End If

End Function

' EFFECTS: gets the length of the centerline
' RETURN TYPE: double
Public Function GetCurveLength(ByRef centerLine As Object) As Double

    Dim TheSPAWorkbench As Workbench
    Set TheSPAWorkbench = activeDoc.GetWorkbench("SPAWorkbench")

    Dim curveLength As Measurable
    Set curveLength = TheSPAWorkbench.GetMeasurable(centerLine)

    GetCurveLength = curveLength.Length

End Function
