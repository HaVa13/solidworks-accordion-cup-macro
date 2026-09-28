Option Explicit

Const V_TARGET As Double = 250000
Const D_MAX As Double = 70
Const H_OPEN As Double = 80
Const N_PLEAT As Integer = 7
Const D_CREST As Double = 68
Const D_VALLEY As Double = 60
Const T_BODY As Double = 0.30
Const MM_TO_METERS As Double = 0.001
Const PI As Double = 3.14159265359

Sub Main()
    On Error GoTo ErrorHandler
    
    Dim swApp As Object
    Dim swPart As Object
    Dim swFeatMgr As Object
    Dim templatePath As String
    
    Set swApp = CreateObject("SldWorks.Application")
    swApp.Visible = True
    
    templatePath = "C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2025\templates\part.prtdot"
    
    Set swPart = swApp.NewDocument(templatePath, 0, 0, 0)
    
    If swPart Is Nothing Then
        MsgBox "Failed to create part.", vbCritical
        Exit Sub
    End If
    
    Set swFeatMgr = swPart.FeatureManager
    
    MsgBox "Creating Accordion Cup Model", vbInformation, "Start"
    
    Call CreateBaseSketch(swPart)
    Call CreateExtrusion(swPart, swFeatMgr)
    Call CreatePleatProfile(swPart, swFeatMgr)
    Call ApplyShellFeature(swPart, swFeatMgr)
    
    swPart.EditRebuild3
    
    Call VerifyVolume()
    
    MsgBox "Model Created!", vbInformation, "Done"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical
End Sub

Sub CreateBaseSketch(swPart As Object)
    Dim swSketch As Object
    Dim swSketchMgr As Object
    Dim dPt(2) As Double
    Dim dRadius As Double
    
    swPart.ClearSelection2 True
    swPart.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    
    Set swSketchMgr = swPart.SketchManager
    Set swSketch = swSketchMgr.CreateSketch(32)
    
    dPt(0) = 0
    dPt(1) = 0
    dPt(2) = 0
    dRadius = (D_MAX / 2) * MM_TO_METERS
    
    swSketch.AddCircle dPt(0), dPt(1), dPt(2), dRadius
    swPart.CloseSketch
    swPart.ClearSelection2 True
End Sub

Sub CreateExtrusion(swPart As Object, swFeatMgr As Object)
    Dim swExtrudeFeature As Object
    Dim dHeight As Double
    
    dHeight = H_OPEN * MM_TO_METERS
    
    swPart.SelectByID2 "Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
    
    Set swExtrudeFeature = swFeatMgr.FeatureExtrusion(False, dHeight, 0, 1, False, False, False, False, False, False, 0, 0, False, False, False, False, True, True, True, 0, 0, False)
    
    swPart.ClearSelection2 True
End Sub

Sub CreatePleatProfile(swPart As Object, swFeatMgr As Object)
    Dim swSketch As Object
    Dim swSketchMgr As Object
    Dim i As Integer
    Dim dAngle As Double
    Dim dAngleStep As Double
    Dim dRadiusCrest As Double
    Dim dRadiusValley As Double
    Dim dPt(2) As Double
    Dim dX1 As Double, dY1 As Double
    Dim dX2 As Double, dY2 As Double
    
    swPart.ClearSelection2 True
    
    dRadiusCrest = (D_CREST / 2) * MM_TO_METERS
    dRadiusValley = (D_VALLEY / 2) * MM_TO_METERS
    dAngleStep = 360 / N_PLEAT
    
    swPart.SelectByID2 "Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    
    Set swSketchMgr = swPart.SketchManager
    Set swSketch = swSketchMgr.CreateSketch(32)
    
    dPt(0) = 0
    dPt(1) = 0
    dPt(2) = 0
    
    swSketch.AddCircle dPt(0), dPt(1), dPt(2), dRadiusCrest
    swSketch.AddCircle dPt(0), dPt(1), dPt(2), dRadiusValley
    
    For i = 0 To N_PLEAT - 1
        dAngle = (i * dAngleStep) * PI / 180
        
        dX1 = dRadiusValley * Cos(dAngle)
        dY1 = dRadiusValley * Sin(dAngle)
        dX2 = dRadiusCrest * Cos(dAngle)
        dY2 = dRadiusCrest * Sin(dAngle)
        
        swSketch.AddCenterLine dX1, dY1, 0, dX2, dY2, 0
    Next i
    
    swPart.CloseSketch
    swPart.ClearSelection2 True
End Sub

Sub ApplyShellFeature(swPart As Object, swFeatMgr As Object)
    Dim swShellFeature As Object
    Dim dThickness As Double
    
    dThickness = T_BODY * MM_TO_METERS
    
    swPart.ClearSelection2 True
    swPart.SelectByID2 "", "FACE", 0, (H_OPEN * MM_TO_METERS), 0, True, 0, Nothing, 0
    
    Set swShellFeature = swFeatMgr.FeatureShell(dThickness, False, Nothing)
    
    swPart.ClearSelection2 True
End Sub

Sub VerifyVolume()
    Dim dVolumeCalc As Double
    Dim dRadiusAvg As Double
    Dim dVolumeDifference As Double
    Dim dPercentError As Double
    Dim dRequiredHeight As Double
    Dim msg As String
    
    dRadiusAvg = (D_CREST / 2 + D_VALLEY / 2) / 2
    dVolumeCalc = PI * (dRadiusAvg ^ 2) * H_OPEN
    
    dVolumeDifference = V_TARGET - dVolumeCalc
    dPercentError = (Abs(dVolumeDifference) / V_TARGET) * 100
    
    dRequiredHeight = CalculateHeightForVolume(V_TARGET)
    
    msg = "VOLUME ANALYSIS:" & vbCrLf & vbCrLf
    msg = msg & "Target: " & Format(V_TARGET, "0") & " mm^3" & vbCrLf
    msg = msg & "Calculated: " & Format(dVolumeCalc, "0.00") & " mm^3" & vbCrLf
    msg = msg & "Difference: " & Format(dVolumeDifference, "0.00") & " mm^3" & vbCrLf
    msg = msg & "Error: " & Format(dPercentError, "0.00") & "%" & vbCrLf & vbCrLf
    msg = msg & "For target volume, adjust H_OPEN to: " & Format(dRequiredHeight, "0.00") & " mm"
    
    MsgBox msg, vbInformation
End Sub

Function CalculateHeightForVolume(dTargetVolume As Double) As Double
    Dim dRadiusAvg As Double
    Dim dHeight As Double
    
    dRadiusAvg = (D_CREST / 2 + D_VALLEY / 2) / 2
    
    If dRadiusAvg > 0 Then
        dHeight = dTargetVolume / (PI * (dRadiusAvg ^ 2))
    Else
        dHeight = 0
    End If
    
    CalculateHeightForVolume = dHeight
End Function
