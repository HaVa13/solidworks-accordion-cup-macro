' ========================================
' SolidWorks 2025 Accordion Cup Macro
' ========================================
' Purpose: Generate parametric 3D accordion cup model
' Parameters: Target volume, dimensions, pleat count, thickness, radii
' ========================================

Option Explicit

' Global variables for SolidWorks application
Dim swApp As Object
Dim swModel As Object
Dim swPart As Object
Dim swFeatMgr As Object
Dim swSketchMgr As Object
Dim vSketch As Object
Dim swBody As Object

' ========================================
' PARAMETER DEFINITIONS
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

' ========================================
' MAIN SUBROUTINE
' ========================================

Sub CreateAccordionCup()
    
    ' Initialize SolidWorks
    Set swApp = CreateObject("SldWorks.Application")
    swApp.Visible = True
    
    ' Create new part document
    Dim swDoc As Object
    Set swDoc = swApp.NewDocument("Part", , 0, 0)
    Set swPart = swDoc
    Set swFeatMgr = swPart.FeatureManager
    
    ' Create base geometry
    Call CreateBaseProfile()
    Call CreatePleatGeometry()
    Call ApplyThickness()
    Call VerifyVolume()
    
    ' Rebuild and save
    swPart.EditRebuild3
    Call SaveDocument()
    
    MsgBox "Accordion cup model created successfully!", vbInformation, "SolidWorks Macro"
    
End Sub

' ========================================
' CREATE BASE PROFILE SKETCH
' ========================================

Sub CreateBaseProfile()
    
    Dim swSketch As Object
    Dim dPt(2) As Double
    Dim dCircRadius As Double
    
    ' Create sketch on XY plane
    swPart.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    Set swSketch = swFeatMgr.CreateSketch(32)
    
    ' Draw base circle (max diameter)
    dPt(0) = 0: dPt(1) = 0: dPt(2) = 0
    dCircRadius = D_MAX / 2
    
    swSketch.AddCircle dPt(0), dPt(1), dPt(2), dCircRadius
    
    ' Add construction centerlines
    dPt(0) = -D_MAX / 2: dPt(1) = 0: dPt(2) = 0
    Dim dPt2(2) As Double
    dPt2(0) = D_MAX / 2: dPt2(1) = 0: dPt2(2) = 0
    swSketch.AddCenterLine dPt(0), dPt(1), dPt(2), dPt2(0), dPt2(1), dPt2(2)
    
    ' Close sketch
    swPart.CloseSketch
    
    ' Create extrusion feature
    Call CreateExtrusion(H_OPEN)
    
End Sub

' ========================================
' CREATE EXTRUSION FEATURE
' ========================================

Sub CreateExtrusion(dHeight As Double)
    
    Dim swExtrudeFeature As Object
    Dim boolStatus As Boolean
    
    swPart.SelectByID2 "Sketch1", "SKETCH", 0, 0, 0, False, 0, Nothing, 0
    Set swExtrudeFeature = swFeatMgr.FeatureExtrusion(False, dHeight, 0, 1, False, _
                                                       False, False, False, False, False, _
                                                       0, 0, False, False, False, False, _
                                                       True, True, True, 0, 0, False)
    
    If swExtrudeFeature Is Nothing Then
        MsgBox "Error creating extrusion feature", vbCritical
    End If
    
End Sub

' ========================================
' CREATE PLEAT GEOMETRY
' ========================================

Sub CreatePleatGeometry()
    
    Dim i As Integer
    Dim dAngle As Double
    Dim dAngleStep As Double
    Dim dRadiusCrest As Double
    Dim dRadiusValley As Double
    
    dRadiusCrest = D_CREST / 2
    dRadiusValley = D_VALLEY / 2
    dAngleStep = 360 / N_PLEAT
    
    ' Create pleat profile sketch
    swPart.SelectByID2 "Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    Set vSketch = swFeatMgr.CreateSketch(32)
    
    ' Draw concentric circles for pleats
    Dim dPt(2) As Double
    dPt(0) = 0: dPt(1) = 0: dPt(2) = 0
    
    ' Crest circle (outer)
    vSketch.AddCircle dPt(0), dPt(1), dPt(2), dRadiusCrest
    
    ' Valley circle (inner)
    vSketch.AddCircle dPt(0), dPt(1), dPt(2), dRadiusValley
    
    ' Add radial construction lines for pleat distribution
    For i = 0 To N_PLEAT - 1
        dAngle = (i * dAngleStep) * 3.14159 / 180
        Dim dX1 As Double, dY1 As Double
        Dim dX2 As Double, dY2 As Double
        
        dX1 = dRadiusValley * Cos(dAngle)
        dY1 = dRadiusValley * Sin(dAngle)
        dX2 = dRadiusCrest * Cos(dAngle)
        dY2 = dRadiusCrest * Sin(dAngle)
        
        vSketch.AddCenterLine dX1, dY1, 0, dX2, dY2, 0
    Next i
    
    swPart.CloseSketch
    
End Sub

' ========================================
' APPLY WALL THICKNESS
' ========================================

Sub ApplyThickness()
    
    Dim swShellFeature As Object
    Dim vFace As Object
    Dim boolStatus As Boolean
    
    ' Select top face for shell operation
    swPart.SelectByID2 "", "FACE", 0, H_OPEN, 0, True, 0, Nothing, 0
    
    ' Create shell feature with variable thickness
    ' Body thickness
    Set swShellFeature = swFeatMgr.FeatureShell(T_BODY, False, Nothing)
    
    If swShellFeature Is Nothing Then
        MsgBox "Error creating shell feature", vbCritical
    End If
    
End Sub

' ========================================
' VERIFY VOLUME
' ========================================

Sub VerifyVolume()
    
    Dim swMass As Object
    Dim dVolume As Double
    Dim dVolumeCalc As Double
    
    ' Calculate theoretical volume
    ' Volume of cylinder: π * r^2 * h
    Dim dRadiusAvg As Double
    dRadiusAvg = (D_CREST / 2 + D_VALLEY / 2) / 2
    
    dVolumeCalc = 3.14159265 * (dRadiusAvg ^ 2) * H_OPEN
    
    MsgBox "Target Volume: " & V_TARGET & " mm³" & vbCrLf & _
           "Calculated Volume: " & Format(dVolumeCalc, "0.00") & " mm³" & vbCrLf & _
           "Difference: " & Format(Abs(dVolumeCalc - V_TARGET), "0.00") & " mm³", _
           vbInformation, "Volume Verification"
    
End Sub

' ========================================
' SAVE DOCUMENT
' ========================================

Sub SaveDocument()
    
    Dim sFilePath As String
    sFilePath = "C:\Users\Public\AccordionCup_2025.SLDPRT"
    
    swPart.SaveAs2 sFilePath, 0, True, False
    
    MsgBox "Document saved to: " & sFilePath, vbInformation, "Save Complete"
    
End Sub

' ========================================
' PARAMETER ADJUSTMENT FUNCTION
' ========================================

Function AdjustParameterByVolume(dTargetVol As Double) As Double
    
    ' Calculate required height adjustment for target volume
    ' V = π * r^2 * h
    ' h = V / (π * r^2)
    
    Dim dRadiusAvg As Double
    Dim dHeightRequired As Double
    
    dRadiusAvg = (D_CREST / 2 + D_VALLEY / 2) / 2
    dHeightRequired = dTargetVol / (3.14159265 * (dRadiusAvg ^ 2))
    
    AdjustParameterByVolume = dHeightRequired
    
End Function

' ========================================
' END OF MACRO
' ========================================
