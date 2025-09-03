' ===========================
' Flip-Stop Generator Macro
' SolidWorks VBA (SW2020+)
' ===========================
' Crée un assemblage simplifié : Base + Rail + Rack + Chariot + Pignon + Butée + 2 vérins
' D.O.F chariot : 1 (translation X). Ajoute ensuite manuellement le mate Rack&Pinion.
' Fichiers dans C:\FlipStopMacro\  (modifiable ci-dessous)

Option Explicit

' ---- CHEMIN DE SORTIE ----
Const OUT_DIR As String = "C:\FlipStopMacro\"

' ---- DIMENSIONS (mm) ----
' Base & rail
Const BASE_L As Double = 430#
Const BASE_W As Double = 60#
Const BASE_T As Double = 12#
Const RAIL_L As Double = 390#
Const RAIL_W As Double = 10#
Const RAIL_H As Double = 15#
' Rack (bloc droit qui sert d’arête pour le mate mécanique)
Const RACK_L As Double = 300#
Const RACK_W As Double = 10#
Const RACK_H As Double = 10#
' Chariot
Const CH_L As Double = 120#
Const CH_W As Double = 60#
Const CH_T As Double = 12#
' Pignon (représenté par un cylindre Ø20)
Const PINION_D As Double = 20#
Const PINION_T As Double = 15#
' Butée
Const BUT_L As Double = 40#
Const BUT_T As Double = 8#
Const BUT_H As Double = 40#
' Vérin vertical (DSNU simplifié)
Const V1_BODY_D As Double = 30#
Const V1_BODY_H As Double = 40#
Const V1_ROD_D  As Double = 12#
Const V1_ROD_EXT As Double = 50#
' Vérin frein (ADN simplifié)
Const V2_BLOC_L As Double = 40#
Const V2_BLOC_W As Double = 20#
Const V2_BLOC_H As Double = 25#
Const V2_ROD_D  As Double = 12#
Const V2_ROD_EXT As Double = 20#

' ---- API ----
Dim swApp As SldWorks.SldWorks

' ===================================================
' Helpers de création de pièce (bloc extrudé & cylindre)
' ===================================================
Function NewPart() As ModelDoc2
    Set NewPart = swApp.NewDocument(swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart), 0, 0, 0)
End Function

Sub SavePart(m As ModelDoc2, fullpath As String)
    Dim errs As Long, warns As Long
    m.SaveAs3 fullpath, 0, 0
End Sub

Function MakeBlock(ByVal L As Double, ByVal W As Double, ByVal T As Double, Optional ByVal MidPlane As Boolean = True) As ModelDoc2
    ' Crée un bloc centré au repère, extrudé en Z
    Dim p As ModelDoc2: Set p = NewPart()
    p.SketchManager.InsertSketch True
    p.SketchManager.CreateCenterRectangle 0#, 0#, 0#, L / 2000#, W / 2000#, 0#  ' en mètres
    p.FeatureManager.FeatureExtrusion2 True, False, IIf(MidPlane, 1, 0), 0, 0, T / 1000#, 0#, False, False, False, False, 0#, 0#, False, False, False, False, True, True, True, 0#, 0#, False
    Set MakeBlock = p
End Function

Function MakeCylinder(ByVal D As Double, ByVal H As Double) As ModelDoc2
    Dim p As ModelDoc2: Set p = NewPart()
    p.SketchManager.InsertSketch True
    p.SketchManager.CreateCircleByRadius 0#, 0#, 0#, (D / 2#) / 1000#
    p.FeatureManager.FeatureExtrusion2 True, False, 0, 0, 0, H / 1000#, 0#, False, False, False, False, 0#, 0#, False, False, False, False, True, True, True, 0#, 0#, False
    Set MakeCylinder = p
End Function

Sub AddThroughHole(p As ModelDoc2, ByVal x As Double, ByVal y As Double, ByVal dia As Double)
    ' Trou traversant depuis la face supérieure
    Dim sel As Boolean
    sel = p.Extension.SelectByID2("", "FACE", 0, 0, (CH_T / 1000#), False, 0, Nothing, 0)
    p.SketchManager.InsertSketch True
    p.SketchManager.CreateCircleByRadius x / 1000#, y / 1000#, 0#, (dia / 2#) / 1000#
    p.FeatureManager.FeatureCut3 True, False, False, 0, 0, CH_T / 1000# + 0.01, 0, False, False, False, False, 0#, 0#, False, False, False, False, False, True, True, True, True, False, 0#, 0#, False
End Sub

' ======================
' MAIN
' ======================
Sub main()
    Set swApp = Application.SldWorks
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(OUT_DIR) Then fso.CreateFolder OUT_DIR

    Dim m As ModelDoc2, fp As String

    ' ---- 1) BASE ----
    Set m = MakeBlock(BASE_L, BASE_W, BASE_T)
    m.Extension.SelectByID2("Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0): m.FeatureManager.InsertRefPlane 8, 0#, 0#, 0#, 0#, 0#
    fp = OUT_DIR & "01_Base.SLDPRT": Call SavePart(m, fp)

    ' ---- 2) RAIL (bloc) ----
    Set m = MakeBlock(RAIL_L, RAIL_W, RAIL_H)
    fp = OUT_DIR & "02_Rail.SLDPRT": Call SavePart(m, fp)

    ' ---- 3) RACK (bloc droit) ----
    Set m = MakeBlock(RACK_L, RACK_W, RACK_H)
    fp = OUT_DIR & "03_Rack.SLDPRT": Call SavePart(m, fp)

    ' ---- 4) CHARIOT (plaque + trou Ø10 pour pignon) ----
    Set m = MakeBlock(CH_L, CH_W, CH_T)
    Call AddThroughHole(m, -50#, -20#, 10#)
    fp = OUT_DIR & "04_Chariot.SLDPRT": Call SavePart(m, fp)

    ' ---- 5) PIGNON (cylindre Ø20×15) ----
    Set m = MakeCylinder(PINION_D, PINION_T)
    fp = OUT_DIR & "05_Pinion.SLDPRT": Call SavePart(m, fp)

    ' ---- 6) BUTEE (plaque verticale) ----
    Set m = MakeBlock(BUT_L, BUT_T, BUT_H)
    fp = OUT_DIR & "06_Butee.SLDPRT": Call SavePart(m, fp)

    ' ---- 7) VERIN VERTICAL (DSNU simplifié : corps + tige) ----
    ' Corps
    Set m = MakeCylinder(V1_BODY_D, V1_BODY_H)
    fp = OUT_DIR & "07a_Ver1_Body.SLDPRT": Call SavePart(m, fp)
    ' Tige
    Set m = MakeCylinder(V1_ROD_D, V1_ROD_EXT)
    fp = OUT_DIR & "07b_Ver1_Rod.SLDPRT": Call SavePart(m, fp)

    ' ---- 8) VERIN FREIN (bloc + tige) ----
    ' Bloc
    Set m = MakeBlock(V2_BLOC_L, V2_BLOC_W, V2_BLOC_H)
    fp = OUT_DIR & "08a_Ver2_Bloc.SLDPRT": Call SavePart(m, fp)
    ' Tige
    Set m = MakeCylinder(V2_ROD_D, V2_ROD_EXT)
    fp = OUT_DIR & "08b_Ver2_Rod.SLDPRT": Call SavePart(m, fp)

    ' ---- 9) ASSEMBLAGE ----
    Dim asmdoc As ModelDoc2
    Set asmdoc = swApp.NewDocument(swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateAssembly), 0, 0, 0)
    Dim swAssy As AssemblyDoc: Set swAssy = asmdoc

    Dim comp As Component2
    Dim path As String

    ' Insert Base (fixe à l’origine)
    path = OUT_DIR & "01_Base.SLDPRT"
    Set comp = swAssy.AddComponent5(path, 0, "", 0, 0, 0, 0) ' en mètres

    ' Insert Rail (posé sur la base)
    path = OUT_DIR & "02_Rail.SLDPRT"
    Set comp = swAssy.AddComponent5(path, 0, "", 0, 0, (BASE_T + RAIL_H / 2#) / 1000#, 0)

    ' Mate : rail posé sur base (Top faces coïncidentes)
    asmdoc.Extension.SelectByID2("Face<1>@02_Rail-1", "FACE", 0, 0, (BASE_T + RAIL_H) / 1000#, False, 1, Nothing, 0)
    asmdoc.Extension.SelectByID2("Face<2>@01_Base-1", "FACE", 0, 0, BASE_T / 2000#, True, 2, Nothing, 0)
    swAssy.AddMate3 swMateType_e.swMateCOINCIDENT, swMateAlign_e.swAlignCLOSEST, True, 0, 0, 0, 0, 0, 0, 0, 0, False, 0

    ' Insert Rack (à côté du rail)
    path = OUT_DIR & "03_Rack.SLDPRT"
    Set comp = swAssy.AddComponent5(path, 0, "", 0, (BASE_W / 2# - RACK_W / 2#) / 1000#, (BASE_T + RACK_H / 2#) / 1000#, 0)

    ' Insert Chariot (sur le rail, centré Y)
    path = OUT_DIR & "04_Chariot.SLDPRT"
    Set comp = swAssy.AddComponent5(path, 0, "", -0.06, 0, (BASE_T + RAIL_H + CH_T / 2#) / 1000#, 0)

    ' Mates chariot : posé sur rail + centré Y + orientation parallèle
    asmdoc.Extension.SelectByID2("Face<1>@04_Chariot-1", "FACE", 0, 0, (BASE_T + RAIL_H + CH_T) / 1000#, False, 1, Nothing, 0)
    asmdoc.Extension.SelectByID2("Face<1>@02_Rail-1", "FACE", 0, 0, (BASE_T + RAIL_H) / 1000#, True, 2, Nothing, 0)
    swAssy.AddMate3 swMateType_e.swMateCOINCIDENT, swMateAlign_e.swAlignCLOSEST, True, 0, 0, 0, 0, 0, 0, 0, 0, False, 0

    ' Centrer chariot en Y (plan Front de l’assy = plan Front du chariot)
    asmdoc.Extension.SelectByID2("Front Plane@04_Chariot-1", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
    asmdoc.Extension.SelectByID2("Front Plane@01_Base-1", "PLANE", 0, 0, 0, True, 2, Nothing, 0)
    swAssy.AddMate3 swMateType_e.swMateCOINCIDENT, swMateAlign_e.swAlignCLOSEST, True, 0, 0, 0, 0, 0, 0, 0, 0, False, 0
    ' Garder 1 DOF (translation X) : pas d’autre mate de position en X.

    ' Insert Pinion (au trou du chariot)
    path = OUT_DIR & "05_Pinion.SLDPRT"
    Set comp = swAssy.AddComponent5(path, 0, "", (-50#) / 1000#, (-20#) / 1000#, (BASE_T + RAIL_H + CH_T) / 1000#, 0)
    ' Concentric : axe cylindre pignon ↔ trou chariot
    asmdoc.Extension.SelectByID2("Face<1>@05_Pinion-1", "FACE", 0, 0, (BASE_T + RAIL_H + CH_T - PINION_T / 2#) / 1000#, False, 1, Nothing, 0)
    asmdoc.Extension.SelectByID2("Face<2>@04_Chariot-1", "FACE", (-50#) / 1000#, (-20#) / 1000#, (BASE_T + RAIL_H + CH_T / 2#) / 1000#, True, 2, Nothing, 0)
    swAssy.AddMate3 swMateType_e.swMateCONCENTRIC, swMateAlign_e.swAlignCLOSEST, True, 0, 0, 0, 0, 0, 0, 0, 0, False, 0
    ' Coïncider une face pour plaquer en Z
    asmdoc.Extension.SelectByID2("Face<2>@05_Pinion-1", "FACE", 0, 0, (BASE_T + RAIL_H + CH_T) / 1000#, False, 1, Nothing, 0)
    asmdoc.Extension.SelectByID2("Face<1>@04_Chariot-1", "FACE", 0, 0, (BASE_T + RAIL_H + CH_T) / 1000#, True, 2, Nothing, 0)
    swAssy.AddMate3 swMateType_e.swMateCOINCIDENT, swMateAlign_e.swAlignCLOSEST, True, 0, 0, 0, 0, 0, 0, 0, 0, False, 0

    ' Insert Butee (devant chariot)
    path = OUT_DIR & "06_Butee.SLDPRT"
    Set comp = swAssy.AddComponent5(path, 0, "", (-CH_L / 2#) / 1000#, 0#, (BASE_T + RAIL_H + CH_T + BUT_H / 2#) / 1000#, 0)
    ' Parallélisme butée/chariot
    asmdoc.Extension.SelectByID2("Right Plane@06_Butee-1", "PLANE", 0, 0, 0, False, 1, Nothing, 0)
    asmdoc.Extension.SelectByID2("Right Plane@04_Chariot-1", "PLANE", 0, 0, 0, True, 2, Nothing, 0)
    swAssy.AddMate3 swMateType_e.swMatePARALLEL, swMateAlign_e.swAlignCLOSEST, True, 0, 0, 0, 0, 0, 0, 0, 0, False, 0

    ' Vérin vertical (corps posé sur chariot, tige vers le haut) – mates simples
    path = OUT_DIR & "07a_Ver1_Body.SLDPRT"
    swAssy.AddComponent5 path, 0, "", (-20#) / 1000#, 0#, (BASE_T + RAIL_H + CH_T) / 1000#, 0
    path = OUT_DIR & "07b_Ver1_Rod.SLDPRT"
    swAssy.AddComponent5 path, 0, "", (-20#) / 1000#, 0#, (BASE_T + RAIL_H + CH_T + V1_BODY_H) / 1000#, 0

    ' Vérin frein (bloc + tige sur la base, à droite)
    path = OUT_DIR & "08a_Ver2_Bloc.SLDPRT"
    swAssy.AddComponent5 path, 0, "", 0.18, 0.02, (BASE_T + V2_BLOC_H / 2#) / 1000#, 0
    path = OUT_DIR & "08b_Ver2_Rod.SLDPRT"
    swAssy.AddComponent5 path, 0, "", 0.18, 0.02, (BASE_T + V2_BLOC_H) / 1000#, 0

    ' Enregistre l’assemblage
    Dim errs As Long, warns As Long
    asmdoc.SaveAs3 OUT_DIR & "FlipStop_Assembly.SLDASM", 0, 0

    MsgBox "Assemblage généré dans: " & OUT_DIR & vbCrLf & _
           "Ajoute maintenant le mate 'Rack and Pinion' entre '05_Pinion-1' (face cylindrique) et l'arête supérieure de '03_Rack-1'." & vbCrLf & _
           "Pitch = 62.832 mm/rev (pignon Ø20).", vbInformation, "Flip-Stop Generator"
End Sub
