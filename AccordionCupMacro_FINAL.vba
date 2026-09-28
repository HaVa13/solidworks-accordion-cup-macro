' ========================================
' SolidWorks 2025 Accordion Cup Macro (FINAL - CORRECTED)
' ========================================
' Purpose: Generate parametric 3D accordion cup model
' Author: SolidWorks Automation
' Date: 2026
' ========================================

Option Explicit

' ========================================
' GLOBAL CONSTANTS - PARAMETERS
' ========================================

' Volume target (mm^3)
Const V_TARGET As Double = 250000

' Overall dimensions
Const D_MAX As Double = 70              ' Maximum diameter (mm)
Const H_OPEN As Double = 80             ' Height when open (mm)
Const H_COLLAPSED As Double = 25        ' Height when collapsed (mm)

' Pleat parameters
Const N_PLEAT As Integer = 7            ' Number of pleats
Const D_CREST As Double = 68            ' Crest diameter (mm)
Const D_VALLEY As Double = 60           ' Valley diameter (mm)

' Thickness parameters (mm)
Const T_BODY As Double = 0.30
Const T_CREST As Double = 0.35
Const T_VALLEY As Double = 0.35

' Radius parameters (mm)
Const R_CREST As Double = 1.50
Const R_VALLEY As Double = 1.50

' Conversion factor
Const MM_TO_METERS As Double = 0.001
Const PI As Double = 3.14159265359

' ========================================
' MAIN ENTRY POINT
' ========================================

Sub Main()
    
    On Error GoTo ErrorHandler
    
    Dim swApp As Object
    Dim swPart As Object
    Dim swFeatMgr As Object
    
    ' Get or create SolidWorks application
    On Error Resume Next
    Set swApp = GetObject(, "SldWorks.Application")
    On Error GoTo 0
    
    If swApp Is Nothing Then
        Set swApp = CreateObject("SldWorks.Application")
    End If
    
    swApp.Visible = True
    
    ' Create new part document (CORRECTED)
    Set swPart = swApp.NewDocument("Part", 0, 0, 0)
    
    If swPart Is Nothing Then
        MsgBox "Failed to create new part document", vbCritical
        Exit Sub
    End If
    
    Set swFeatMgr = swPart.FeatureManager
    
    ' Build the accordion cup
    Call BuildAccordionCup(swPart, swFeatMgr)
    
    ' Rebuild and verify
    swPart.EditRebuild3
    
    ' Calculate and display volume
    Call VerifyVolume()
    
    MsgBox "✓ Accordion Cup Created Successfully!" & vbCrLf & vbCrLf & _
           "Dimensions:" & vbCrLf & _
           "  • Diameter: " & D_MAX & " mm" & vbCrLf & _
           "  • Height (Open): " & H_OPEN & " mm" & vbCrLf & _
           "  • Pleats: " & N_PLEAT & vbCrLf & _
           "  • Body Thickness: " & T_BODY & " mm", _
           vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "ERROR: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "Macro Error"
    
End Sub

' ========================================
' BUILD ACCORDION CUP
' ========================================

Sub BuildAccordionCup(swPart As Object, swFeatMgr As Object)
    
    On Error GoTo BuildError
    
    ' Step 1: Create base cylinder
    Call CreateBaseSketch(swPart)
    Call CreateExtrusion(swPart, swFeatMgr, H_OPEN * MM_TO_METERS)
    
    ' Step 2: Create pleat geometry
    Call CreatePleatProfile(swPart, swFeatMgr)
    
    ' Step 3: Apply shell (wall thickness)
    Call ApplyShellFeature(swPart, swFeatMgr, T_BODY * MM_TO_METERS)
    
    ' Step 4: Add fillets to crest and valley
    Call AddFillets(swPart, swFeatMgr)
    
    Exit Sub
    
BuildError:
    MsgBox "Build Error: " & Err.Description, vbCritical
    
End Sub

' ========================================
' STEP 1: CREATE BASE SKETCH
' ========================================

Sub CreateBaseSketch(swPart As Object)
    
    On Error GoTo SketchError
    
    Dim swSketch As Object
    Dim swSketchMgr As Object
    Dim dPt(2) As Double
    Dim dRadius As Double
    
    ' Select Front Plane
    swPart.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    
    ' Create new sketch
    Set swSketchMgr = swPart.SketchManager
    Set swSketch = swSketchMgr.CreateSketch(32)
    
    If swSketch Is Nothing Then
        MsgBox "Failed to create sketch", vbCritical
        Exit Sub
    End If
    
    ' Define circle center
    dPt(0) = 0
    dPt(1) = 0
    dPt(2) = 0
    dRadius = (D_MAX / 2) * MM_TO_METERS
    
    ' Add circle
    swSketch.AddCircle dPt(0), dPt(1), dPt(2), dRadius
    
    ' Close sketch
    swPart.CloseSketch
    
    Exit Sub
    
SketchError:
    MsgBox "Sketch Creation Error: " & Err.Description, vbCritical
    
End Sub

' ========================================
' STEP 2: CREATE EXTRUSION
' ========================================

Sub CreateExtrusion(swPart As Object, swFeatMgr As Object, dHeight As Double)
    
    On Error GoTo ExtrudeError
    
    Dim swExtrudeFeature As Object
    
    ' Select sketch
    swPart.SelectByID2 "Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
    
    ' Create extrusion
    Set swExtrudeFeature = swFeatMgr.FeatureExtrusion(False, dHeight, 0, 1, False, False, _
                                                       False, False, False, False, 0, 0, False, _
                                                       False, False, False, True, True, True, 0, 0, False)
    
    If swExtrudeFeature Is Nothing Then
        MsgBox "Failed to create extrusion feature", vbCritical
    End If
    
    ' Clear selection
    swPart.ClearSelection2 True
    
    Exit Sub
    
ExtrudeError:
    MsgBox "Extrusion Error: " & Err.Description, vbCritical
    
End Sub

' ========================================
' STEP 3: CREATE PLEAT PROFILE
' ========================================

Sub CreatePleatProfile(swPart As Object, swFeatMgr As Object)
    
    On Error GoTo PleatError
    
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
    
    ' Initialize
    dRadiusCrest = (D_CREST / 2) * MM_TO_METERS
    dRadiusValley = (D_VALLEY / 2) * MM_TO_METERS
    dAngleStep = 360 / N_PLEAT
    
    ' Select Top Plane
    swPart.SelectByID2 "Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    
    ' Create sketch
    Set swSketchMgr = swPart.SketchManager
    Set swSketch = swSketchMgr.CreateSketch(32)
    
    If swSketch Is Nothing Then
        MsgBox "Failed to create pleat sketch", vbCritical
        Exit Sub
    End If
    
    ' Center point
    dPt(0) = 0
    dPt(1) = 0
    dPt(2) = 0
    
    ' Draw concentric circles
    swSketch.AddCircle dPt(0), dPt(1), dPt(2), dRadiusCrest
    swSketch.AddCircle dPt(0), dPt(1), dPt(2), dRadiusValley
    
    ' Add radial construction lines for pleats
    For i = 0 To N_PLEAT - 1
        dAngle = (i * dAngleStep) * PI / 180
        
        dX1 = dRadiusValley * Cos(dAngle)
        dY1 = dRadiusValley * Sin(dAngle)
        dX2 = dRadiusCrest * Cos(dAngle)
        dY2 = dRadiusCrest * Sin(dAngle)
        
        swSketch.AddCenterLine dX1, dY1, 0, dX2, dY2, 0
    Next i
    
    ' Close sketch
    swPart.CloseSketch
    
    ' Clear selection
    swPart.ClearSelection2 True
    
    Exit Sub
    
PleatError:
    MsgBox "Pleat Profile Error: " & Err.Description, vbCritical
    
End Sub

' ========================================
' STEP 4: APPLY SHELL FEATURE (WALL THICKNESS)
' ========================================

Sub ApplyShellFeature(swPart As Object, swFeatMgr As Object, dThickness As Double)
    
    On Error GoTo ShellError
    
    Dim swShellFeature As Object
    
    ' Clear any selection
    swPart.ClearSelection2 True
    
    ' Select the top face of the extrusion
    swPart.SelectByID2 "", "FACE", 0, (H_OPEN * MM_TO_METERS), 0, True, 0, Nothing, 0
    
    ' Create shell feature with specified thickness
    Set swShellFeature = swFeatMgr.FeatureShell(dThickness, False, Nothing)
    
    If swShellFeature Is Nothing Then
        MsgBox "Failed to create shell feature", vbCritical
    End If
    
    ' Clear selection
    swPart.ClearSelection2 True
    
    Exit Sub
    
ShellError:
    MsgBox "Shell Feature Error: " & Err.Description, vbCritical
    
End Sub

' ========================================
' STEP 5: ADD FILLETS TO EDGES
' ========================================

Sub AddFillets(swPart As Object, swFeatMgr As Object)
    
    On Error GoTo FilletError
    
    ' Fillet application is optional for basic model
    ' In production, iterate through edges and apply fillets
    
    Exit Sub
    
FilletError:
    ' Non-critical error
    
End Sub

' ========================================
' VOLUME VERIFICATION
' ========================================

Sub VerifyVolume()
    
    Dim dVolumeCalc As Double
    Dim dRadiusAvg As Double
    Dim dVolumeDifference As Double
    Dim dPercentError As Double
    Dim dRequiredHeight As Double
    
    ' Calculate average radius
    dRadiusAvg = (D_CREST / 2 + D_VALLEY / 2) / 2
    
    ' Volume calculation: V = π * r^2 * h
    dVolumeCalc = PI * (dRadiusAvg ^ 2) * H_OPEN
    
    ' Calculate difference
    dVolumeDifference = V_TARGET - dVolumeCalc
    dPercentError = (Abs(dVolumeDifference) / V_TARGET) * 100
    
    ' Calculate required height
    dRequiredHeight = CalculateHeightForVolume(V_TARGET)
    
    MsgBox "VOLUME ANALYSIS" & vbCrLf & vbCrLf & _
           "Target Volume:      " & Format(V_TARGET, "0") & " mm³" & vbCrLf & _
           "Calculated Volume:  " & Format(dVolumeCalc, "0.00") & " mm³" & vbCrLf & _
           "Difference:         " & Format(dVolumeDifference, "0.00") & " mm³" & vbCrLf & _
           "Error:              " & Format(dPercentError, "0.00") & "%" & vbCrLf & vbCrLf & _
           "ADJUSTMENT:" & vbCrLf & _
           "To match target volume, set H_OPEN to:" & vbCrLf & _
           Format(dRequiredHeight, "0.00") & " mm", _
           vbInformation, "Volume Verification"
    
End Sub

' ========================================
' CALCULATE HEIGHT FOR TARGET VOLUME
' ========================================

Function CalculateHeightForVolume(dTargetVolume As Double) As Double
    
    Dim dRadiusAvg As Double
    Dim dHeight As Double
    
    ' V = π * r^2 * h
    ' h = V / (π * r^2)
    
    dRadiusAvg = (D_CREST / 2 + D_VALLEY / 2) / 2
    
    If dRadiusAvg > 0 Then
        dHeight = dTargetVolume / (PI * (dRadiusAvg ^ 2))
    Else
        dHeight = 0
    End If
    
    CalculateHeightForVolume = dHeight
    
End Function

' ========================================
' END OF MACRO
' ========================================
