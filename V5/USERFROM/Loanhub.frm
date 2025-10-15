VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Loanhub 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Loanhub.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Loanhub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private arrEmprunteurs As Variant
Private arrPrets As Variant
Private selectedEmprunteur As String
Private selectedEmail As String
Private selectedTech As String

Private lblTitle As MSForms.Label
Private txtSearch As MSForms.TextBox
Private lstEmprunteurs As MSForms.ListBox
Private lblEmail As MSForms.Label
Private txtEmail As MSForms.TextBox
Private btnVersEmprunteur As MSForms.CommandButton
Private frameTech As MSForms.Frame
Private btnTech1 As MSForms.CommandButton
Private btnTech2 As MSForms.CommandButton
Private btnTech3 As MSForms.CommandButton
Private btnTech4 As MSForms.CommandButton
Private btnTech5 As MSForms.CommandButton
Private framePretsEnCours As MSForms.Frame
Private lstPretsEnCours As MSForms.ListBox
Private btnCreer As MSForms.CommandButton
Private btnRetourner As MSForms.CommandButton
Private btnModifier As MSForms.CommandButton
Private btnQuit As MSForms.CommandButton

Private Sub UserForm_Initialize()
    With Me
        .Width = 840
        .Height = 630
        .BackColor = M_Core.COLOR_LIGHT
    End With
    
    CreateInterface
    LoadData
End Sub

Private Sub CreateInterface()
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lblTitle
        .Caption = "Recherche par: Emprunteur (NOM_PRENOM)"
        .Left = 12
        .Top = 12
        .Width = 300
        .Height = 20
        .Font.Size = 10
    End With
    
    Set txtSearch = Me.Controls.Add("Forms.TextBox.1", "txtSearch")
    With txtSearch
        .Left = 12
        .Top = 36
        .Width = 300
        .Height = 24
        .BackColor = RGB(255, 255, 200)
        .Font.Size = 10
    End With
    
    Set lstEmprunteurs = Me.Controls.Add("Forms.ListBox.1", "lstEmprunteurs")
    With lstEmprunteurs
        .Left = 12
        .Top = 72
        .Width = 300
        .Height = 300
        .Font.Size = 9
    End With
    
    Set lblEmail = Me.Controls.Add("Forms.Label.1", "lblEmail")
    With lblEmail
        .Caption = "Email"
        .Left = 12
        .Top = 384
        .Width = 80
        .Height = 20
    End With
    
    Set txtEmail = Me.Controls.Add("Forms.TextBox.1", "txtEmail")
    With txtEmail
        .Left = 100
        .Top = 384
        .Width = 212
        .Height = 24
        .Font.Size = 9
    End With
    
    Set btnVersEmprunteur = Me.Controls.Add("Forms.CommandButton.1", "btnVersEmprunteur")
    With btnVersEmprunteur
        .Caption = "Vers emprunteur"
        .Left = 12
        .Top = 420
        .Width = 140
        .Height = 30
        .Font.Size = 9
    End With
    
    Set frameTech = Me.Controls.Add("Forms.Frame.1", "frameTech")
    With frameTech
        .Caption = "technicien"
        .Left = 330
        .Top = 12
        .Width = 480
        .Height = 120
    End With
    
    Dim techs As Variant
    techs = M_Core.GetTechniciens()
    Dim techLabels As Variant
    techLabels = Array("FL", "TP", "ND", "SP", "DJ")
    
    Set btnTech1 = frameTech.Controls.Add("Forms.CommandButton.1", "btnTech1")
    With btnTech1
        .Caption = techLabels(0)
        .Left = 12
        .Top = 30
        .Width = 80
        .Height = 60
        .Tag = techs(0)
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    Set btnTech2 = frameTech.Controls.Add("Forms.CommandButton.1", "btnTech2")
    With btnTech2
        .Caption = techLabels(1)
        .Left = 108
        .Top = 30
        .Width = 80
        .Height = 60
        .Tag = techs(1)
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    Set btnTech3 = frameTech.Controls.Add("Forms.CommandButton.1", "btnTech3")
    With btnTech3
        .Caption = techLabels(2)
        .Left = 204
        .Top = 30
        .Width = 80
        .Height = 60
        .Tag = techs(2)
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    Set btnTech4 = frameTech.Controls.Add("Forms.CommandButton.1", "btnTech4")
    With btnTech4
        .Caption = techLabels(3)
        .Left = 300
        .Top = 30
        .Width = 80
        .Height = 60
        .Tag = techs(3)
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    Set btnTech5 = frameTech.Controls.Add("Forms.CommandButton.1", "btnTech5")
    With btnTech5
        .Caption = techLabels(4)
        .Left = 396
        .Top = 30
        .Width = 80
        .Height = 60
        .Tag = techs(4)
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    Set framePretsEnCours = Me.Controls.Add("Forms.Frame.1", "framePretsEnCours")
    With framePretsEnCours
        .Caption = "emprunt en cours"
        .Left = 330
        .Top = 150
        .Width = 480
        .Height = 300
    End With
    
    Set lstPretsEnCours = framePretsEnCours.Controls.Add("Forms.ListBox.1", "lstPretsEnCours")
    With lstPretsEnCours
        .Left = 12
        .Top = 24
        .Width = 456
        .Height = 260
        .Font.Size = 9
    End With
    
    Set btnCreer = Me.Controls.Add("Forms.CommandButton.1", "btnCreer")
    With btnCreer
        .Caption = "Créer" & vbCrLf & "un prêt"
        .Left = 330
        .Top = 480
        .Width = 140
        .Height = 80
        .BackColor = RGB(255, 228, 196)
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    Set btnRetourner = Me.Controls.Add("Forms.CommandButton.1", "btnRetourner")
    With btnRetourner
        .Caption = "Retourner" & vbCrLf & "un prêt"
        .Left = 490
        .Top = 480
        .Width = 140
        .Height = 80
        .BackColor = RGB(144, 238, 144)
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    Set btnModifier = Me.Controls.Add("Forms.CommandButton.1", "btnModifier")
    With btnModifier
        .Caption = "Modifier" & vbCrLf & "un prêt"
        .Left = 650
        .Top = 480
        .Width = 140
        .Height = 80
        .BackColor = RGB(255, 192, 203)
        .Font.Size = 12
        .Font.Bold = True
    End With
    
    Set btnQuit = Me.Controls.Add("Forms.CommandButton.1", "btnQuit")
    With btnQuit
        .Caption = "Quitter"
        .Left = 690
        .Top = 570
        .Width = 100
        .Height = 30
        .BackColor = RGB(135, 206, 250)
        .Font.Size = 10
    End With
End Sub

Private Sub LoadData()
    arrEmprunteurs = M_Core.LoadDataToArray("Tableau1")
    If IsArray(arrEmprunteurs) Then
        M_Core.PopulateListBox lstEmprunteurs, arrEmprunteurs
        lstEmprunteurs.ColumnCount = UBound(arrEmprunteurs, 2)
        lstEmprunteurs.ColumnWidths = "0;150;0;0;0;0;0;0;0;0"
    End If
End Sub

Private Sub txtSearch_Change()
    If Trim(txtSearch.value) = "" Then
        M_Core.PopulateListBox lstEmprunteurs, arrEmprunteurs
    Else
        Dim filtered As Variant
        filtered = M_Core.SearchArrayWildcard(arrEmprunteurs, 2, "*" & txtSearch.value & "*")
        If IsArray(filtered) Then
            M_Core.PopulateListBox lstEmprunteurs, filtered
        Else
            lstEmprunteurs.Clear
        End If
    End If
End Sub

Private Sub lstEmprunteurs_Click()
    If lstEmprunteurs.ListIndex >= 0 Then
        selectedEmprunteur = lstEmprunteurs.List(lstEmprunteurs.ListIndex, 1)
        selectedEmail = lstEmprunteurs.List(lstEmprunteurs.ListIndex, 5)
        txtEmail.value = selectedEmail
        LoadPretsEnCours
    End If
End Sub

Private Sub LoadPretsEnCours()
    If selectedEmprunteur = "" Then Exit Sub
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    Dim tempArr() As Variant
    ReDim tempArr(1 To 100, 1 To 3)
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = 2 To wsPrets.Cells(wsPrets.Rows.count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 3).value = selectedEmprunteur And wsPrets.Cells(i, 15).value = "" Then
            count = count + 1
            tempArr(count, 1) = wsPrets.Cells(i, 4).value
            tempArr(count, 2) = wsPrets.Cells(i, 6).value
            tempArr(count, 3) = wsPrets.Cells(i, 7).value
        End If
    Next i
    
    If count > 0 Then
        ReDim Preserve tempArr(1 To count, 1 To 3)
        lstPretsEnCours.Clear
        lstPretsEnCours.List = tempArr
        lstPretsEnCours.ColumnCount = 3
        lstPretsEnCours.ColumnWidths = "100;250;50"
    Else
        lstPretsEnCours.Clear
    End If
End Sub

Private Sub btnTech1_Click()
    selectedTech = btnTech1.Tag
    HighlightTechButton btnTech1
End Sub

Private Sub btnTech2_Click()
    selectedTech = btnTech2.Tag
    HighlightTechButton btnTech2
End Sub

Private Sub btnTech3_Click()
    selectedTech = btnTech3.Tag
    HighlightTechButton btnTech3
End Sub

Private Sub btnTech4_Click()
    selectedTech = btnTech4.Tag
    HighlightTechButton btnTech4
End Sub

Private Sub btnTech5_Click()
    selectedTech = btnTech5.Tag
    HighlightTechButton btnTech5
End Sub

Private Sub HighlightTechButton(btn As MSForms.CommandButton)
    btnTech1.BackColor = &H8000000F
    btnTech2.BackColor = &H8000000F
    btnTech3.BackColor = &H8000000F
    btnTech4.BackColor = &H8000000F
    btnTech5.BackColor = &H8000000F
    btn.BackColor = M_Core.COLOR_SUCCESS
End Sub

Private Sub btnCreer_Click()
    If selectedEmprunteur = "" Then
        MsgBox "Sélectionnez un emprunteur", vbExclamation
        Exit Sub
    End If
    If selectedTech = "" Then
        MsgBox "Sélectionnez un technicien", vbExclamation
        Exit Sub
    End If
    
    Me.Hide
    Createloan.InitializeWithData selectedEmprunteur, selectedEmail, selectedTech
    Createloan.Show
End Sub

Private Sub btnRetourner_Click()
    If selectedEmprunteur = "" Then
        MsgBox "Sélectionnez un emprunteur", vbExclamation
        Exit Sub
    End If
    If selectedTech = "" Then
        MsgBox "Sélectionnez un technicien", vbExclamation
        Exit Sub
    End If
    
    Me.Hide
    Returnloan.InitializeWithData selectedEmprunteur, selectedTech
    Returnloan.Show
End Sub

Private Sub btnModifier_Click()
    MsgBox "Modification en cours de développement", vbInformation
End Sub

Private Sub btnVersEmprunteur_Click()
    Managedata.Show
    Managedata.SwitchToTab 1
End Sub

Private Sub btnQuit_Click()
    Me.Hide
    Mainmenu.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnQuit_Click
    End If
End Sub
