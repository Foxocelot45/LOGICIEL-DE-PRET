VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReturnLoan 
   Caption         =   "Retour d'un prêt"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   OleObjectBlob   =   "ReturnLoan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReturnLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Données session
Private currentEmprunteur As String
Private currentTech As String
Private checkControls As New Collection

' Contrôles
Private tabControl As MSForms.MultiPage
Private btnQuit As MSForms.CommandButton

Public Sub InitializeWithData(emprunteur As String, tech As String)
    currentEmprunteur = emprunteur
    currentTech = tech
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Width = 760
        .Height = 630
        .BackColor = M_Core.COLOR_LIGHT
    End With
    
    CreateInterface
    LoadReturnData
End Sub

Private Sub CreateInterface()
    ' MultiPage avec 4 onglets
    Set tabControl = Me.Controls.Add("Forms.MultiPage.1", "tabControl")
    With tabControl
        .Left = 12
        .Top = 12
        .Width = 736
        .Height = 540
        .Style = fmTabStyleTabs
    End With
    
    ' Configuration onglets
    tabControl.Pages(0).Caption = "Retour Unitaire"
    tabControl.Pages.Add , "tab1"
    tabControl.Pages(1).Caption = "Tout Retourner"
    tabControl.Pages.Add , "tab2"
    tabControl.Pages(2).Caption = "Retour Cochage"
    tabControl.Pages.Add , "tab3"
    tabControl.Pages(3).Caption = "Scan Chaîne"
    
    ' Créer contenu de chaque onglet
    CreateTab0_Unitaire
    CreateTab1_ToutRetourner
    CreateTab2_Cochage
    CreateTab3_ScanChaine
    
    ' Bouton Quitter global
    Set btnQuit = Me.Controls.Add("Forms.CommandButton.1", "btnQuit")
    With btnQuit
        .Caption = "Quitter"
        .Left = 630
        .Top = 565
        .Width = 100
        .Height = 30
        .BackColor = RGB(135, 206, 250)
        .Font.Size = 10
    End With
End Sub

' =====================================================
' ONGLET 0: RETOUR UNITAIRE
' =====================================================

Private Sub CreateTab0_Unitaire()
    Dim page As MSForms.page
    Set page = tabControl.Pages(0)
    
    ' Liste prêts en cours
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = page.Controls.Add("Forms.ListBox.1", "lstPretsUnitaire")
    With lstPrets
        .Left = 12
        .Top = 24
        .Width = 700
        .Height = 350
        .Font.Size = 9
    End With
    
    ' Détails prêt sélectionné
    Dim frameDetails As MSForms.Frame
    Set frameDetails = page.Controls.Add("Forms.Frame.1", "frameDetails")
    With frameDetails
        .Caption = "Détails Prêt Sélectionné"
        .Left = 12
        .Top = 385
        .Width = 700
        .Height = 80
    End With
    
    ' Labels détails
    Dim lblDetail As MSForms.Label
    Set lblDetail = frameDetails.Controls.Add("Forms.Label.1", "lblDetail")
    With lblDetail
        .Left = 12
        .Top = 24
        .Width = 676
        .Height = 40
        .Font.Size = 9
        .Caption = "Sélectionnez un prêt ci-dessus"
    End With
    
    ' Bouton Valider retour
    Dim btnValider As MSForms.CommandButton
    Set btnValider = page.Controls.Add("Forms.CommandButton.1", "btnValiderUnitaire")
    With btnValider
        .Caption = "Validation modif"
        .Left = 12
        .Top = 475
        .Width = 340
        .Height = 40
        .BackColor = RGB(144, 238, 144)
        .Font.Size = 12
        .Font.Bold = True
    End With
End Sub

' =====================================================
' ONGLET 1: TOUT RETOURNER
' =====================================================

Private Sub CreateTab1_ToutRetourner()
    Dim page As MSForms.page
    Set page = tabControl.Pages(1)
    
    ' Titre
    Dim lblTitle As MSForms.Label
    Set lblTitle = page.Controls.Add("Forms.Label.1", "lblTitleTout")
    With lblTitle
        .Caption = "Retour en 1 CLIC - Tous les prêts de l'emprunteur"
        .Left = 12
        .Top = 24
        .Width = 700
        .Height = 30
        .Font.Size = 14
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = M_Core.COLOR_WARNING
    End With
    
    ' Info emprunteur
    Dim lblInfo As MSForms.Label
    Set lblInfo = page.Controls.Add("Forms.Label.1", "lblInfoTout")
    With lblInfo
        .Caption = "Emprunteur: " & currentEmprunteur
        .Left = 12
        .Top = 70
        .Width = 700
        .Height = 25
        .Font.Size = 12
    End With
    
    ' Liste prêts concernés
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = page.Controls.Add("Forms.ListBox.1", "lstPretsTout")
    With lstPrets
        .Left = 12
        .Top = 110
        .Width = 700
        .Height = 280
        .Font.Size = 9
    End With
    
    ' Stats
    Dim lblStats As MSForms.Label
    Set lblStats = page.Controls.Add("Forms.Label.1", "lblStatsTout")
    With lblStats
        .Left = 12
        .Top = 405
        .Width = 700
        .Height = 40
        .Font.Size = 11
        .Caption = "Prêts à retourner: 0"
    End With
    
    ' Bouton TOUT RETOURNER
    Dim btnTout As MSForms.CommandButton
    Set btnTout = page.Controls.Add("Forms.CommandButton.1", "btnToutRetourner")
    With btnTout
        .Caption = "🚀 TOUT RETOURNER"
        .Left = 150
        .Top = 455
        .Width = 400
        .Height = 60
        .BackColor = M_Core.COLOR_WARNING
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 16
        .Font.Bold = True
    End With
End Sub

' =====================================================
' ONGLET 2: RETOUR COCHAGE
' =====================================================

Private Sub CreateTab2_Cochage()
    Dim page As MSForms.page
    Set page = tabControl.Pages(2)
    
    ' Titre
    Dim lblTitle As MSForms.Label
    Set lblTitle = page.Controls.Add("Forms.Label.1", "lblTitleCochage")
    With lblTitle
        .Caption = "Retour par COCHAGE - Sélection multiple"
        .Left = 12
        .Top = 12
        .Width = 700
        .Height = 25
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    ' Boutons sélection
    Dim btnToutCocher As MSForms.CommandButton
    Set btnToutCocher = page.Controls.Add("Forms.CommandButton.1", "btnToutCocher")
    With btnToutCocher
        .Caption = "✓ Tout cocher"
        .Left = 12
        .Top = 45
        .Width = 120
        .Height = 30
        .Font.Size = 9
    End With
    
    Dim btnToutDecocher As MSForms.CommandButton
    Set btnToutDecocher = page.Controls.Add("Forms.CommandButton.1", "btnToutDecocher")
    With btnToutDecocher
        .Caption = "✗ Tout décocher"
        .Left = 140
        .Top = 45
        .Width = 120
        .Height = 30
        .Font.Size = 9
    End With
    
    ' Frame scrollable pour checkboxes
    Dim frameScroll As MSForms.Frame
    Set frameScroll = page.Controls.Add("Forms.Frame.1", "frameCheckboxes")
    With frameScroll
        .Left = 12
        .Top = 85
        .Width = 700
        .Height = 350
        .ScrollBars = fmScrollBarsVertical
        .BorderStyle = fmBorderStyleSingle
    End With
    
    ' Timer/Stats
    Dim lblTimer As MSForms.Label
    Set lblTimer = page.Controls.Add("Forms.Label.1", "lblTimerCochage")
    With lblTimer
        .Left = 12
        .Top = 445
        .Width = 700
        .Height = 25
        .Font.Size = 11
        .Caption = "Sélectionnés: 0 | Gain temps: 0%"
    End With
    
    ' Bouton Valider
    Dim btnValider As MSForms.CommandButton
    Set btnValider = page.Controls.Add("Forms.CommandButton.1", "btnValiderCochage")
    With btnValider
        .Caption = "✓ VALIDER RETOURS COCHÉS"
        .Left = 200
        .Top = 480
        .Width = 320
        .Height = 40
        .BackColor = M_Core.COLOR_SUCCESS
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 12
        .Font.Bold = True
    End With
End Sub

' =====================================================
' ONGLET 3: SCAN CHAÎNE
' =====================================================

Private Sub CreateTab3_ScanChaine()
    Dim page As MSForms.page
    Set page = tabControl.Pages(3)
    
    ' Titre
    Dim lblTitle As MSForms.Label
    Set lblTitle = page.Controls.Add("Forms.Label.1", "lblTitleScan")
    With lblTitle
        .Caption = "Mode SCAN À LA CHAÎNE"
        .Left = 12
        .Top = 12
        .Width = 700
        .Height = 30
        .Font.Size = 14
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = RGB(173, 216, 230)
    End With
    
    ' Instructions
    Dim lblInstructions As MSForms.Label
    Set lblInstructions = page.Controls.Add("Forms.Label.1", "lblInstructionsScan")
    With lblInstructions
        .Caption = "1. Cliquez sur DÉMARRER" & vbCrLf & _
                   "2. Scannez les QR Codes à la chaîne" & vbCrLf & _
                   "3. Cliquez STOP quand terminé"
        .Left = 12
        .Top = 55
        .Width = 700
        .Height = 60
        .Font.Size = 11
    End With
    
    ' Champ scan
    Dim txtScan As MSForms.TextBox
    Set txtScan = page.Controls.Add("Forms.TextBox.1", "txtScanQR")
    With txtScan
        .Left = 12
        .Top = 130
        .Width = 700
        .Height = 30
        .Font.Size = 14
        .BackColor = RGB(255, 255, 200)
    End With
    
    ' Stats temps réel
    Dim lblStatsScan As MSForms.Label
    Set lblStatsScan = page.Controls.Add("Forms.Label.1", "lblStatsScan")
    With lblStatsScan
        .Left = 12
        .Top = 175
        .Width = 700
        .Height = 60
        .Font.Size = 12
        .Font.Bold = True
        .Caption = "Articles retournés: 0" & vbCrLf & _
                    "Temps moyen/article: -- sec"
    End With
    
    ' Liste articles scannés
    Dim lstScanned As MSForms.ListBox
    Set lstScanned = page.Controls.Add("Forms.ListBox.1", "lstScanned")
    With lstScanned
        .Left = 12
        .Top = 250
        .Width = 700
        .Height = 200
        .Font.Size = 9
    End With
    
    ' Boutons contrôle
    Dim btnStart As MSForms.CommandButton
    Set btnStart = page.Controls.Add("Forms.CommandButton.1", "btnStartScan")
    With btnStart
        .Caption = "▶ DÉMARRER"
        .Left = 150
        .Top = 465
        .Width = 180
        .Height = 50
        .BackColor = M_Core.COLOR_SUCCESS
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    Dim btnStop As MSForms.CommandButton
    Set btnStop = page.Controls.Add("Forms.CommandButton.1", "btnStopScan")
    With btnStop
        .Caption = "■ STOP"
        .Left = 370
        .Top = 465
        .Width = 180
        .Height = 50
        .BackColor = M_Core.COLOR_DANGER
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 14
        .Font.Bold = True
        .Enabled = False
    End With
End Sub

' =====================================================
' CHARGEMENT DONNÉES
' =====================================================

Private Sub LoadReturnData()
    ' Charger onglet Unitaire
    LoadTab0Data
    
    ' Charger onglet Tout Retourner
    LoadTab1Data
    
    ' Charger onglet Cochage avec checkboxes
    LoadTab2Data
End Sub

Private Sub LoadTab0Data()
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = tabControl.Pages(0).Controls("lstPretsUnitaire")
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    lstPrets.Clear
    Dim i As Long
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 3).Value = currentEmprunteur And wsPrets.Cells(i, 15).Value = "" Then
            lstPrets.AddItem
            lstPrets.List(lstPrets.ListCount - 1, 0) = wsPrets.Cells(i, 4).Value ' Date
            lstPrets.List(lstPrets.ListCount - 1, 1) = wsPrets.Cells(i, 6).Value ' Article
            lstPrets.List(lstPrets.ListCount - 1, 2) = wsPrets.Cells(i, 7).Value ' Qté
        End If
    Next i
    lstPrets.ColumnCount = 3
    lstPrets.ColumnWidths = "100;450;50"
End Sub

Private Sub LoadTab1Data()
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = tabControl.Pages(1).Controls("lstPretsTout")
    
    Dim lblStats As MSForms.Label
    Set lblStats = tabControl.Pages(1).Controls("lblStatsTout")
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    lstPrets.Clear
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 3).Value = currentEmprunteur And wsPrets.Cells(i, 15).Value = "" Then
            lstPrets.AddItem wsPrets.Cells(i, 4).Value & " - " & wsPrets.Cells(i, 6).Value & " (x" & wsPrets.Cells(i, 7).Value & ")"
            count = count + 1
        End If
    Next i
    
    lblStats.Caption = "Prêts à retourner: " & count & " | Gain temps: ~95% (vs " & Format(count * 90, "0") & " sec traditionnels)"
End Sub

Private Sub LoadTab2Data()
    Dim frameCheckboxes As MSForms.Frame
    Set frameCheckboxes = tabControl.Pages(2).Controls("frameCheckboxes")
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    ' Nettoyer checkboxes existantes
    Set checkControls = New Collection
    
    Dim yPos As Long
    yPos = 10
    Dim index As Long
    index = 0
    
    Dim i As Long
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 3).Value = currentEmprunteur And wsPrets.Cells(i, 15).Value = "" Then
            ' Créer checkbox
            Dim chk As MSForms.CheckBox
            Set chk = frameCheckboxes.Controls.Add("Forms.CheckBox.1", "chk" & index)
            With chk
                .Left = 10
                .Top = yPos
                .Width = 650
                .Height = 20
                .Caption = wsPrets.Cells(i, 4).Value & " - " & wsPrets.Cells(i, 6).Value & " (x" & wsPrets.Cells(i, 7).Value & ")"
                .Font.Size = 9
                .Tag = i - 1 ' Ligne Excel (0-based pour compatibilité)
            End With
            
            checkControls.Add chk
            yPos = yPos + 25
            index = index + 1
        End If
    Next i
End Sub

' =====================================================
' ÉVÉNEMENTS
' =====================================================

Private Sub btnValiderUnitaire_Click()
    MsgBox "Retour unitaire en cours", vbInformation
End Sub

Private Sub btnToutRetourner_Click()
    Call M_Business.RetournerTousPretsEmprunteur(currentEmprunteur, currentTech)
    Me.Hide
    LoanHub.Show
End Sub

Private Sub btnToutCocher_Click()
    Dim i As Long
    For i = 1 To checkControls.count
        checkControls(i).Value = True
    Next i
    UpdateCochageStats
End Sub

Private Sub btnToutDecocher_Click()
    Dim i As Long
    For i = 1 To checkControls.count
        checkControls(i).Value = False
    Next i
    UpdateCochageStats
End Sub

Private Sub btnValiderCochage_Click()
    Dim lstDummy As MSForms.ListBox
    Set lstDummy = tabControl.Pages(0).Controls("lstPretsUnitaire") ' Utiliser ListBox avec données
    
    Call M_Business.ValiderRetoursCoches(lstDummy, checkControls, currentTech)
    Me.Hide
    LoanHub.Show
End Sub

Private Sub UpdateCochageStats()
    Dim count As Long
    Dim i As Long
    For i = 1 To checkControls.count
        If checkControls(i).Value = True Then count = count + 1
    Next i
    
    Dim lblTimer As MSForms.Label
    Set lblTimer = tabControl.Pages(2).Controls("lblTimerCochage")
    lblTimer.Caption = "Sélectionnés: " & count & " | Gain temps estimé: " & Format((1 - (count * 5 / (count * 90))) * 100, "0") & "%"
End Sub

Private Sub btnStartScan_Click()
    M_Business.RetourCount = 0
    M_Business.RetourStartTime = Timer
    
    Dim btnStart As MSForms.CommandButton
    Dim btnStop As MSForms.CommandButton
    Set btnStart = tabControl.Pages(3).Controls("btnStartScan")
    Set btnStop = tabControl.Pages(3).Controls("btnStopScan")
    
    btnStart.Enabled = False
    btnStop.Enabled = True
    
    Dim txtScan As MSForms.TextBox
    Set txtScan = tabControl.Pages(3).Controls("txtScanQR")
    txtScan.SetFocus
End Sub

Private Sub btnStopScan_Click()
    Dim btnStart As MSForms.CommandButton
    Dim btnStop As MSForms.CommandButton
    Set btnStart = tabControl.Pages(3).Controls("btnStartScan")
    Set btnStop = tabControl.Pages(3).Controls("btnStopScan")
    
    btnStart.Enabled = True
    btnStop.Enabled = False
    
    MsgBox "Scan terminé!" & vbCrLf & "Articles retournés: " & M_Business.RetourCount, vbInformation
    
    If M_Business.RetourCount > 0 Then
        Call M_Business.SendBatchReturnEmail(currentEmprunteur, M_Business.RetourCount, currentTech)
    End If
End Sub

Private Sub btnQuit_Click()
    Me.Hide
    LoanHub.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnQuit_Click
    End If
End Sub
