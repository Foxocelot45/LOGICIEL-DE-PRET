VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Returnloan 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Returnloan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Returnloan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Donnees session
Private currentEmprunteur As String
Private currentTech As String
Private checkControls As New Collection
Private arrPretsData As Variant

' Controles
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
        .Caption = "Retour d'un pret"
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
    tabControl.Pages(3).Caption = "Scan Chaine"
    
    ' Creer contenu de chaque onglet
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

Private Sub CreateTab0_Unitaire()
    Dim page As MSForms.page
    Set page = tabControl.Pages(0)
    
    ' Liste prets en cours
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = page.Controls.Add("Forms.ListBox.1", "lstPretsUnitaire")
    With lstPrets
        .Left = 12
        .Top = 24
        .Width = 700
        .Height = 300
        .Font.Size = 9
    End With
    
    ' Frame details pret selectionne
    Dim frameDetails As MSForms.Frame
    Set frameDetails = page.Controls.Add("Forms.Frame.1", "frameDetails")
    With frameDetails
        .Caption = "Details Pret Selectionne"
        .Left = 12
        .Top = 335
        .Width = 700
        .Height = 100
    End With
    
    ' TextBox date retour
    Dim lblDateRetour As MSForms.Label
    Set lblDateRetour = frameDetails.Controls.Add("Forms.Label.1", "lblDateRetour")
    With lblDateRetour
        .Caption = "Date retour:"
        .Left = 12
        .Top = 25
        .Width = 80
        .Height = 20
    End With
    
    Dim txtDateRetour As MSForms.TextBox
    Set txtDateRetour = frameDetails.Controls.Add("Forms.TextBox.1", "txtDateRetour")
    With txtDateRetour
        .Left = 100
        .Top = 24
        .Width = 120
        .Height = 22
        .BackColor = RGB(255, 255, 200)
    End With
    
    ' Bouton Date&heure
    Dim btnDateHeure As MSForms.CommandButton
    Set btnDateHeure = frameDetails.Controls.Add("Forms.CommandButton.1", "btnDateHeure")
    With btnDateHeure
        .Caption = "Date&heure"
        .Left = 230
        .Top = 22
        .Width = 80
        .Height = 26
    End With
    
    ' TextBox Commentaires
    Dim lblCommentaires As MSForms.Label
    Set lblCommentaires = frameDetails.Controls.Add("Forms.Label.1", "lblCommentaires")
    With lblCommentaires
        .Caption = "Commentaires:"
        .Left = 330
        .Top = 25
        .Width = 90
        .Height = 20
    End With
    
    Dim txtCommentaires As MSForms.TextBox
    Set txtCommentaires = frameDetails.Controls.Add("Forms.TextBox.1", "txtCommentaires")
    With txtCommentaires
        .Left = 430
        .Top = 24
        .Width = 258
        .Height = 22
    End With
    
    ' Label Technicien retour
    Dim lblTechRetour As MSForms.Label
    Set lblTechRetour = frameDetails.Controls.Add("Forms.Label.1", "lblTechRetour")
    With lblTechRetour
        .Caption = "Technicien: " & currentTech
        .Left = 12
        .Top = 60
        .Width = 300
        .Height = 20
        .Font.Bold = True
    End With
    
    ' Bouton Valider retour
    Dim btnValider As MSForms.CommandButton
    Set btnValider = page.Controls.Add("Forms.CommandButton.1", "btnValiderUnitaire")
    With btnValider
        .Caption = "Validation modif"
        .Left = 250
        .Top = 450
        .Width = 220
        .Height = 50
        .BackColor = RGB(144, 238, 144)
        .Font.Size = 12
        .Font.Bold = True
    End With
End Sub

Private Sub CreateTab1_ToutRetourner()
    Dim page As MSForms.page
    Set page = tabControl.Pages(1)
    
    ' Titre
    Dim lblTitle As MSForms.Label
    Set lblTitle = page.Controls.Add("Forms.Label.1", "lblTitleTout")
    With lblTitle
        .Caption = "Retour en 1 CLIC - Tous les prets de l'emprunteur"
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
    
    ' Liste prets concernes
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
        .Caption = "Prets a retourner: 0"
    End With
    
    ' Bouton TOUT RETOURNER
    Dim btnTout As MSForms.CommandButton
    Set btnTout = page.Controls.Add("Forms.CommandButton.1", "btnToutRetourner")
    With btnTout
        .Caption = "TOUT RETOURNER"
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

Private Sub CreateTab2_Cochage()
    Dim page As MSForms.page
    Set page = tabControl.Pages(2)
    
    ' Titre
    Dim lblTitle As MSForms.Label
    Set lblTitle = page.Controls.Add("Forms.Label.1", "lblTitleCochage")
    With lblTitle
        .Caption = "Retour par COCHAGE - Selection multiple"
        .Left = 12
        .Top = 12
        .Width = 700
        .Height = 25
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    ' Boutons selection
    Dim btnToutCocher As MSForms.CommandButton
    Set btnToutCocher = page.Controls.Add("Forms.CommandButton.1", "btnToutCocher")
    With btnToutCocher
        .Caption = "Tout cocher"
        .Left = 12
        .Top = 45
        .Width = 120
        .Height = 30
        .Font.Size = 9
    End With
    
    Dim btnToutDecocher As MSForms.CommandButton
    Set btnToutDecocher = page.Controls.Add("Forms.CommandButton.1", "btnToutDecocher")
    With btnToutDecocher
        .Caption = "Tout decocher"
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
        .Caption = "Selectionnes: 0 | Gain temps: 0%"
    End With
    
    ' Bouton Valider
    Dim btnValider As MSForms.CommandButton
    Set btnValider = page.Controls.Add("Forms.CommandButton.1", "btnValiderCochage")
    With btnValider
        .Caption = "VALIDER RETOURS COCHES"
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

Private Sub CreateTab3_ScanChaine()
    Dim page As MSForms.page
    Set page = tabControl.Pages(3)
    
    ' Titre
    Dim lblTitle As MSForms.Label
    Set lblTitle = page.Controls.Add("Forms.Label.1", "lblTitleScan")
    With lblTitle
        .Caption = "Mode SCAN A LA CHAINE"
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
        .Caption = "1. Cliquez sur DEMARRER" & vbCrLf & _
                   "2. Scannez les QR Codes a la chaine" & vbCrLf & _
                   "3. Cliquez STOP quand termine"
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
    
    ' Stats temps reel
    Dim lblStatsScan As MSForms.Label
    Set lblStatsScan = page.Controls.Add("Forms.Label.1", "lblStatsScan")
    With lblStatsScan
        .Left = 12
        .Top = 175
        .Width = 700
        .Height = 60
        .Font.Size = 12
        .Font.Bold = True
        .Caption = "Articles retournes: 0" & vbCrLf & _
                    "Temps moyen/article: -- sec"
    End With
    
    ' Liste articles scannes
    Dim lstScanned As MSForms.ListBox
    Set lstScanned = page.Controls.Add("Forms.ListBox.1", "lstScanned")
    With lstScanned
        .Left = 12
        .Top = 250
        .Width = 700
        .Height = 200
        .Font.Size = 9
    End With
    
    ' Boutons controle
    Dim btnStart As MSForms.CommandButton
    Set btnStart = page.Controls.Add("Forms.CommandButton.1", "btnStartScan")
    With btnStart
        .Caption = "DEMARRER"
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
        .Caption = "STOP"
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

Private Sub LoadReturnData()
    LoadTab0Data
    LoadTab1Data
    LoadTab2Data
End Sub

Private Sub LoadTab0Data()
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = tabControl.Pages(0).Controls("lstPretsUnitaire")
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    lstPrets.Clear
    
    ReDim tempArr(1 To 100, 1 To 4) As Variant
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = 2 To wsPrets.Cells(wsPrets.Rows.count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 3).value = currentEmprunteur And wsPrets.Cells(i, 15).value = "" Then
            count = count + 1
            tempArr(count, 1) = wsPrets.Cells(i, 4).value
            tempArr(count, 2) = wsPrets.Cells(i, 6).value
            tempArr(count, 3) = wsPrets.Cells(i, 7).value
            tempArr(count, 4) = i
        End If
    Next i
    
    If count > 0 Then
        ReDim Preserve tempArr(1 To count, 1 To 4)
        arrPretsData = tempArr
        
        lstPrets.List = tempArr
        lstPrets.ColumnCount = 4
        lstPrets.ColumnWidths = "100;450;50;0"
    End If
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
    For i = 2 To wsPrets.Cells(wsPrets.Rows.count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 3).value = currentEmprunteur And wsPrets.Cells(i, 15).value = "" Then
            lstPrets.AddItem wsPrets.Cells(i, 4).value & " - " & wsPrets.Cells(i, 6).value & " (x" & wsPrets.Cells(i, 7).value & ")"
            count = count + 1
        End If
    Next i
    
    lblStats.Caption = "Prets a retourner: " & count & " | Gain temps: ~95% (vs " & Format(count * 90, "0") & " sec traditionnels)"
End Sub

Private Sub LoadTab2Data()
    Dim frameCheckboxes As MSForms.Frame
    Set frameCheckboxes = tabControl.Pages(2).Controls("frameCheckboxes")
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    Set checkControls = New Collection
    
    Dim yPos As Long
    yPos = 10
    Dim index As Long
    index = 0
    
    Dim i As Long
    For i = 2 To wsPrets.Cells(wsPrets.Rows.count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 3).value = currentEmprunteur And wsPrets.Cells(i, 15).value = "" Then
            Dim chk As MSForms.CheckBox
            Set chk = frameCheckboxes.Controls.Add("Forms.CheckBox.1", "chk" & index)
            With chk
                .Left = 10
                .Top = yPos
                .Width = 650
                .Height = 20
                .Caption = wsPrets.Cells(i, 4).value & " - " & wsPrets.Cells(i, 6).value & " (x" & wsPrets.Cells(i, 7).value & ")"
                .Font.Size = 9
                .Tag = i
            End With
            
            checkControls.Add chk
            yPos = yPos + 25
            index = index + 1
        End If
    Next i
End Sub

Private Sub btnValiderUnitaire_Click()
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = tabControl.Pages(0).Controls("lstPretsUnitaire")
    
    If lstPrets.ListIndex < 0 Then
        MsgBox "Selectionnez un pret a retourner", vbExclamation
        Exit Sub
    End If
    
    Dim txtDateRetour As MSForms.TextBox
    Set txtDateRetour = tabControl.Pages(0).Controls("frameDetails").Controls("txtDateRetour")
    
    Dim txtCommentaires As MSForms.TextBox
    Set txtCommentaires = tabControl.Pages(0).Controls("frameDetails").Controls("txtCommentaires")
    
    If Trim(txtDateRetour.value) = "" Then
        MsgBox "Veuillez renseigner la date de retour", vbExclamation
        Exit Sub
    End If
    
    Dim rowNum As Long
    rowNum = arrPretsData(lstPrets.ListIndex + 1, 4)
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    With wsPrets
        .Cells(rowNum, 15).value = txtDateRetour.value
        .Cells(rowNum, 13).value = currentTech
        .Cells(rowNum, 14).value = "Termine"
        .Cells(rowNum, 12).value = txtCommentaires.value
    End With
    
    MsgBox "Pret retourne avec succes!", vbInformation
    
    LoadTab0Data
End Sub

Private Sub btnDateHeure_Click()
    Dim txtDateRetour As MSForms.TextBox
    Set txtDateRetour = tabControl.Pages(0).Controls("frameDetails").Controls("txtDateRetour")
    txtDateRetour.value = Format(Now, "DD/MM/YYYY HH:MM:SS")
End Sub

Private Sub lstPretsUnitaire_Click()
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = tabControl.Pages(0).Controls("lstPretsUnitaire")
    
    If lstPrets.ListIndex >= 0 Then
        Dim txtDateRetour As MSForms.TextBox
        Set txtDateRetour = tabControl.Pages(0).Controls("frameDetails").Controls("txtDateRetour")
        txtDateRetour.value = Format(Now, "DD/MM/YYYY HH:MM:SS")
    End If
End Sub

Private Sub btnToutRetourner_Click()
    Call M_Business.RetournerTousPretsEmprunteur(currentEmprunteur, currentTech)
    Me.Hide
    Loanhub.Show
End Sub

Private Sub btnToutCocher_Click()
    Dim i As Long
    For i = 1 To checkControls.count
        checkControls(i).value = True
    Next i
    UpdateCochageStats
End Sub

Private Sub btnToutDecocher_Click()
    Dim i As Long
    For i = 1 To checkControls.count
        checkControls(i).value = False
    Next i
    UpdateCochageStats
End Sub

Private Sub btnValiderCochage_Click()
    M_Business.RetourStartTime = Timer
    M_Business.RetourCount = 0
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    Dim dateRetour As Date
    dateRetour = Now
    
    Application.ScreenUpdating = False
    
    Dim i As Long
    Dim chk As MSForms.CheckBox
    Dim rowNum As Long
    
    For i = 1 To checkControls.count
        Set chk = checkControls(i)
        If chk.value = True Then
            rowNum = CLng(chk.Tag)
            
            wsPrets.Cells(rowNum, 15).value = dateRetour
            wsPrets.Cells(rowNum, 13).value = currentTech
            wsPrets.Cells(rowNum, 14).value = "Termine"
            M_Business.RetourCount = M_Business.RetourCount + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    If M_Business.RetourCount > 0 Then
        Dim elapsedTime As Double
        Dim traditionalTime As Double
        elapsedTime = Timer - M_Business.RetourStartTime
        traditionalTime = M_Business.RetourCount * 90
        
        Dim gain As Double
        gain = (1 - (elapsedTime / traditionalTime)) * 100
        
        MsgBox M_Business.RetourCount & " articles retournes!" & vbCrLf & vbCrLf & _
               "Temps reel: " & Format(elapsedTime, "0") & " sec" & vbCrLf & _
               "Temps traditionnel: " & Format(traditionalTime, "0") & " sec" & vbCrLf & _
               "GAIN: -" & Format(gain, "0") & "%", vbInformation
        
        Call M_Business.SendBatchReturnEmail(currentEmprunteur, M_Business.RetourCount, currentTech)
        
        Me.Hide
        Loanhub.Show
    End If
End Sub

Private Sub UpdateCochageStats()
    Dim count As Long
    Dim i As Long
    For i = 1 To checkControls.count
        If checkControls(i).value = True Then count = count + 1
    Next i
    
    Dim lblTimer As MSForms.Label
    Set lblTimer = tabControl.Pages(2).Controls("lblTimerCochage")
    lblTimer.Caption = "Selectionnes: " & count & " | Gain temps estime: " & Format((1 - (count * 5 / (count * 90))) * 100, "0") & "%"
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
    
    Dim lstScanned As MSForms.ListBox
    Set lstScanned = tabControl.Pages(3).Controls("lstScanned")
    lstScanned.Clear
    
    Dim txtScan As MSForms.TextBox
    Set txtScan = tabControl.Pages(3).Controls("txtScanQR")
    txtScan.value = ""
    txtScan.SetFocus
End Sub

Private Sub txtScanQR_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        Dim txtScan As MSForms.TextBox
        Set txtScan = tabControl.Pages(3).Controls("txtScanQR")
        
        Dim qrCode As String
        qrCode = Trim(txtScan.value)
        
        If qrCode <> "" Then
            Call M_Business.TraiterScanQR(qrCode, currentTech)
            
            Dim lstScanned As MSForms.ListBox
            Set lstScanned = tabControl.Pages(3).Controls("lstScanned")
            lstScanned.AddItem qrCode & " - Retourne"
            
            Dim lblStats As MSForms.Label
            Set lblStats = tabControl.Pages(3).Controls("lblStatsScan")
            
            Dim avgTime As Double
            If M_Business.RetourCount > 0 Then
                avgTime = (Timer - M_Business.RetourStartTime) / M_Business.RetourCount
            End If
            
            lblStats.Caption = "Articles retournes: " & M_Business.RetourCount & vbCrLf & _
                               "Temps moyen/article: " & Format(avgTime, "0.0") & " sec"
        End If
        
        txtScan.value = ""
        KeyAscii = 0
    End If
End Sub

Private Sub btnStopScan_Click()
    Dim btnStart As MSForms.CommandButton
    Dim btnStop As MSForms.CommandButton
    Set btnStart = tabControl.Pages(3).Controls("btnStartScan")
    Set btnStop = tabControl.Pages(3).Controls("btnStopScan")
    
    btnStart.Enabled = True
    btnStop.Enabled = False
    
    MsgBox "Scan termine!" & vbCrLf & "Articles retournes: " & M_Business.RetourCount, vbInformation
    
    If M_Business.RetourCount > 0 Then
        Call M_Business.SendBatchReturnEmail(currentEmprunteur, M_Business.RetourCount, currentTech)
    End If
End Sub

Private Sub btnQuit_Click()
    Me.Hide
    Loanhub.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnQuit_Click
    End If
End Sub

