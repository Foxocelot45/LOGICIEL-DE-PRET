VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateLoan 
   Caption         =   "sortis d'objet"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "CreateLoan.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Donnees session
Private currentEmprunteur As String
Private currentEmail As String
Private currentTech As String
Private currentRaison As String
Private currentArticleID As String

' Controles
Private lblEmprunteur As MSForms.Label
Private txtEmprunteur As MSForms.TextBox
Private lblNumPret As MSForms.Label
Private txtNumPret As MSForms.TextBox
Private frameEnregistrement As MSForms.Frame
Private lblDateHeure As MSForms.Label
Private txtDateHeure As MSForms.TextBox
Private btnDate1 As MSForms.CommandButton
Private lblDateRetour As MSForms.Label
Private txtDateRetour As MSForms.TextBox
Private btnDate2 As MSForms.CommandButton
Private frameRaisons As MSForms.Frame
Private btnRaison1 As MSForms.CommandButton
Private btnRaison2 As MSForms.CommandButton
Private btnRaison3 As MSForms.CommandButton
Private btnRaison4 As MSForms.CommandButton
Private btnRaison5 As MSForms.CommandButton
Private btnRaison6 As MSForms.CommandButton
Private lblRaisonPret As MSForms.Label
Private txtRaisonPret As MSForms.TextBox
Private lblQRCode As MSForms.Label
Private btnQRCode As MSForms.CommandButton
Private lblObjet As MSForms.Label
Private txtObjet As MSForms.TextBox
Private lblQtePris As MSForms.Label
Private txtQtePris As MSForms.TextBox
Private btnQte1 As MSForms.CommandButton
Private btnQte2 As MSForms.CommandButton
Private btnQte3 As MSForms.CommandButton
Private btnQte4 As MSForms.CommandButton
Private btnQte5 As MSForms.CommandButton
Private lblQRCodeValue As MSForms.Label
Private txtQRCodeValue As MSForms.TextBox
Private lblTechnicien As MSForms.Label
Private txtTechnicien As MSForms.TextBox
Private btnValidation As MSForms.CommandButton
Private btnQuitter As MSForms.CommandButton

Public Sub InitializeWithData(emprunteur As String, email As String, tech As String)
    currentEmprunteur = emprunteur
    currentEmail = email
    currentTech = tech
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Width = 720
        .Height = 675
        .BackColor = M_Core.COLOR_LIGHT
    End With
    
    CreateInterface
    PopulateDefaults
End Sub

Private Sub CreateInterface()
    ' Emprunteur
    Set lblEmprunteur = Me.Controls.Add("Forms.Label.1", "lblEmprunteur")
    With lblEmprunteur
        .Caption = "Emprunteur" & vbCrLf & "(NOM_PRENOM)"
        .Left = 24
        .Top = 24
        .Width = 80
        .Height = 30
        .Font.Size = 9
    End With
    
    Set txtEmprunteur = Me.Controls.Add("Forms.TextBox.1", "txtEmprunteur")
    With txtEmprunteur
        .Left = 120
        .Top = 24
        .Width = 360
        .Height = 24
        .Font.Size = 10
        .Locked = True
        .BackColor = RGB(240, 240, 240)
    End With
    
    ' N pret
    Set lblNumPret = Me.Controls.Add("Forms.Label.1", "lblNumPret")
    With lblNumPret
        .Caption = "N du pret"
        .Left = 500
        .Top = 24
        .Width = 60
        .Height = 24
    End With
    
    Set txtNumPret = Me.Controls.Add("Forms.TextBox.1", "txtNumPret")
    With txtNumPret
        .Left = 570
        .Top = 24
        .Width = 100
        .Height = 24
        .Font.Size = 11
        .Font.Bold = True
        .Locked = True
        .BackColor = RGB(255, 255, 255)
    End With
    
    ' Frame Enregistrement
    Set frameEnregistrement = Me.Controls.Add("Forms.Frame.1", "frameEnregistrement")
    With frameEnregistrement
        .Caption = "Enregistrement"
        .Left = 24
        .Top = 60
        .Width = 646
        .Height = 90
    End With
    
    Set lblDateHeure = frameEnregistrement.Controls.Add("Forms.Label.1", "lblDateHeure")
    With lblDateHeure
        .Caption = "date et heure du" & vbCrLf & "pret"
        .Left = 12
        .Top = 24
        .Width = 90
        .Height = 30
        .Font.Size = 9
    End With
    
    Set txtDateHeure = frameEnregistrement.Controls.Add("Forms.TextBox.1", "txtDateHeure")
    With txtDateHeure
        .Left = 110
        .Top = 30
        .Width = 160
        .Height = 24
        .BackColor = RGB(255, 255, 200)
        .Font.Size = 10
    End With
    
    Set btnDate1 = frameEnregistrement.Controls.Add("Forms.CommandButton.1", "btnDate1")
    With btnDate1
        .Caption = "Date&heure"
        .Left = 280
        .Top = 28
        .Width = 80
        .Height = 28
        .Font.Size = 9
    End With
    
    Set lblDateRetour = frameEnregistrement.Controls.Add("Forms.Label.1", "lblDateRetour")
    With lblDateRetour
        .Caption = "date du retour" & vbCrLf & "prevu"
        .Left = 380
        .Top = 24
        .Width = 80
        .Height = 30
        .Font.Size = 9
    End With
    
    Set txtDateRetour = frameEnregistrement.Controls.Add("Forms.TextBox.1", "txtDateRetour")
    With txtDateRetour
        .Left = 470
        .Top = 30
        .Width = 100
        .Height = 24
        .BackColor = RGB(255, 255, 255)
        .Font.Size = 10
    End With
    
    Set btnDate2 = frameEnregistrement.Controls.Add("Forms.CommandButton.1", "btnDate2")
    With btnDate2
        .Caption = "Date"
        .Left = 575
        .Top = 28
        .Width = 55
        .Height = 28
        .Font.Size = 9
    End With
    
    ' Raisons 6 boutons
    Dim raisons As Variant
    raisons = M_Core.GetRaisonsPret()
    
    Set btnRaison1 = Me.Controls.Add("Forms.CommandButton.1", "btnRaison1")
    With btnRaison1
        .Caption = raisons(0)
        .Left = 85
        .Top = 170
        .Width = 75
        .Height = 45
        .BackColor = RGB(144, 238, 144)
        .Font.Size = 10
        .Font.Bold = True
        .Tag = raisons(0)
    End With
    
    Set btnRaison2 = Me.Controls.Add("Forms.CommandButton.1", "btnRaison2")
    With btnRaison2
        .Caption = raisons(1)
        .Left = 165
        .Top = 170
        .Width = 75
        .Height = 45
        .BackColor = RGB(144, 238, 144)
        .Font.Size = 10
        .Font.Bold = True
        .Tag = raisons(1)
    End With
    
    Set btnRaison3 = Me.Controls.Add("Forms.CommandButton.1", "btnRaison3")
    With btnRaison3
        .Caption = raisons(2)
        .Left = 245
        .Top = 170
        .Width = 75
        .Height = 45
        .BackColor = RGB(144, 238, 144)
        .Font.Size = 10
        .Font.Bold = True
        .Tag = raisons(2)
    End With
    
    Set btnRaison4 = Me.Controls.Add("Forms.CommandButton.1", "btnRaison4")
    With btnRaison4
        .Caption = raisons(3)
        .Left = 325
        .Top = 170
        .Width = 75
        .Height = 45
        .BackColor = RGB(144, 238, 144)
        .Font.Size = 10
        .Font.Bold = True
        .Tag = raisons(3)
    End With
    
    Set btnRaison5 = Me.Controls.Add("Forms.CommandButton.1", "btnRaison5")
    With btnRaison5
        .Caption = raisons(4)
        .Left = 405
        .Top = 170
        .Width = 75
        .Height = 45
        .BackColor = RGB(144, 238, 144)
        .Font.Size = 10
        .Font.Bold = True
        .Tag = raisons(4)
    End With
    
    Set btnRaison6 = Me.Controls.Add("Forms.CommandButton.1", "btnRaison6")
    With btnRaison6
        .Caption = raisons(5)
        .Left = 485
        .Top = 170
        .Width = 75
        .Height = 45
        .BackColor = RGB(144, 238, 144)
        .Font.Size = 10
        .Font.Bold = True
        .Tag = raisons(5)
    End With
    
    Set lblRaisonPret = Me.Controls.Add("Forms.Label.1", "lblRaisonPret")
    With lblRaisonPret
        .Caption = "raison du pret"
        .Left = 24
        .Top = 230
        .Width = 90
        .Height = 20
        .Font.Size = 9
    End With
    
    Set txtRaisonPret = Me.Controls.Add("Forms.TextBox.1", "txtRaisonPret")
    With txtRaisonPret
        .Left = 85
        .Top = 225
        .Width = 390
        .Height = 24
        .BackColor = RGB(255, 255, 200)
        .Font.Size = 10
    End With
    
    ' QR Code
    Set btnQRCode = Me.Controls.Add("Forms.CommandButton.1", "btnQRCode")
    With btnQRCode
        .Caption = "Choix par QRCode"
        .Left = 133
        .Top = 280
        .Width = 250
        .Height = 60
        .BackColor = RGB(173, 216, 230)
        .Font.Size = 16
        .Font.Bold = True
    End With
    
    Set lblObjet = Me.Controls.Add("Forms.Label.1", "lblObjet")
    With lblObjet
        .Caption = "Objet"
        .Left = 24
        .Top = 360
        .Width = 50
        .Height = 20
        .Font.Size = 9
    End With
    
    Set txtObjet = Me.Controls.Add("Forms.TextBox.1", "txtObjet")
    With txtObjet
        .Left = 85
        .Top = 355
        .Width = 390
        .Height = 24
        .BackColor = RGB(255, 255, 200)
        .Font.Size = 10
    End With
    
    ' Quantite
    Set lblQtePris = Me.Controls.Add("Forms.Label.1", "lblQtePris")
    With lblQtePris
        .Caption = "Qte pris"
        .Left = 24
        .Top = 400
        .Width = 50
        .Height = 20
        .Font.Size = 9
    End With
    
    Set txtQtePris = Me.Controls.Add("Forms.TextBox.1", "txtQtePris")
    With txtQtePris
        .Left = 85
        .Top = 395
        .Width = 60
        .Height = 24
        .BackColor = RGB(255, 255, 255)
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    ' Boutons quantites rapides
    Dim qtes As Variant
    qtes = M_Core.GetQuantitesRapides()
    
    Set btnQte1 = Me.Controls.Add("Forms.CommandButton.1", "btnQte1")
    With btnQte1
        .Caption = qtes(0)
        .Left = 155
        .Top = 393
        .Width = 40
        .Height = 28
        .Font.Size = 10
        .Tag = qtes(0)
    End With
    
    Set btnQte2 = Me.Controls.Add("Forms.CommandButton.1", "btnQte2")
    With btnQte2
        .Caption = qtes(1)
        .Left = 200
        .Top = 393
        .Width = 40
        .Height = 28
        .Font.Size = 10
        .Tag = qtes(1)
    End With
    
    Set btnQte3 = Me.Controls.Add("Forms.CommandButton.1", "btnQte3")
    With btnQte3
        .Caption = qtes(2)
        .Left = 245
        .Top = 393
        .Width = 40
        .Height = 28
        .Font.Size = 10
        .Tag = qtes(2)
    End With
    
    Set btnQte4 = Me.Controls.Add("Forms.CommandButton.1", "btnQte4")
    With btnQte4
        .Caption = qtes(3)
        .Left = 290
        .Top = 393
        .Width = 40
        .Height = 28
        .Font.Size = 10
        .Tag = qtes(3)
    End With
    
    Set btnQte5 = Me.Controls.Add("Forms.CommandButton.1", "btnQte5")
    With btnQte5
        .Caption = qtes(4)
        .Left = 335
        .Top = 393
        .Width = 40
        .Height = 28
        .Font.Size = 10
        .Tag = qtes(4)
    End With
    
    ' QRCode value
    Set lblQRCodeValue = Me.Controls.Add("Forms.Label.1", "lblQRCodeValue")
    With lblQRCodeValue
        .Caption = "QRCode"
        .Left = 24
        .Top = 440
        .Width = 50
        .Height = 20
        .Font.Size = 9
    End With
    
    Set txtQRCodeValue = Me.Controls.Add("Forms.TextBox.1", "txtQRCodeValue")
    With txtQRCodeValue
        .Left = 85
        .Top = 435
        .Width = 390
        .Height = 24
        .Font.Size = 10
    End With
    
    ' Technicien
    Set lblTechnicien = Me.Controls.Add("Forms.Label.1", "lblTechnicien")
    With lblTechnicien
        .Caption = "Technicien depart"
        .Left = 24
        .Top = 480
        .Width = 100
        .Height = 20
        .Font.Size = 9
    End With
    
    Set txtTechnicien = Me.Controls.Add("Forms.TextBox.1", "txtTechnicien")
    With txtTechnicien
        .Left = 130
        .Top = 475
        .Width = 345
        .Height = 24
        .Font.Size = 10
        .Locked = True
        .BackColor = RGB(240, 240, 240)
    End With
    
    ' Boutons action
    Set btnValidation = Me.Controls.Add("Forms.CommandButton.1", "btnValidation")
    With btnValidation
        .Caption = "Validation"
        .Left = 80
        .Top = 540
        .Width = 180
        .Height = 60
        .BackColor = RGB(255, 255, 200)
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    Set btnQuitter = Me.Controls.Add("Forms.CommandButton.1", "btnQuitter")
    With btnQuitter
        .Caption = "Quitter"
        .Left = 360
        .Top = 540
        .Width = 180
        .Height = 60
        .BackColor = RGB(135, 206, 250)
        .Font.Size = 14
        .Font.Bold = True
    End With
End Sub

Private Sub PopulateDefaults()
    txtEmprunteur.Value = currentEmprunteur
    txtTechnicien.Value = currentTech
    txtDateHeure.Value = Format(Now, "DD/MM/YYYY HH:MM:SS")
    
    ' Generer N pret
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    Dim nextID As Long
    nextID = wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
    txtNumPret.Value = nextID
End Sub

Private Sub btnDate1_Click()
    txtDateHeure.Value = Format(Now, "DD/MM/YYYY HH:MM:SS")
End Sub

Private Sub btnDate2_Click()
    txtDateRetour.Value = Format(Date + 7, "DD/MM/YYYY")
End Sub

Private Sub btnRaison1_Click()
    txtRaisonPret.Value = btnRaison1.Tag
    currentRaison = btnRaison1.Tag
End Sub

Private Sub btnRaison2_Click()
    txtRaisonPret.Value = btnRaison2.Tag
    currentRaison = btnRaison2.Tag
End Sub

Private Sub btnRaison3_Click()
    txtRaisonPret.Value = btnRaison3.Tag
    currentRaison = btnRaison3.Tag
End Sub

Private Sub btnRaison4_Click()
    txtRaisonPret.Value = btnRaison4.Tag
    currentRaison = btnRaison4.Tag
End Sub

Private Sub btnRaison5_Click()
    txtRaisonPret.Value = btnRaison5.Tag
    currentRaison = btnRaison5.Tag
End Sub

Private Sub btnRaison6_Click()
    txtRaisonPret.Value = btnRaison6.Tag
    currentRaison = btnRaison6.Tag
End Sub

Private Sub btnQte1_Click()
    txtQtePris.Value = btnQte1.Tag
End Sub

Private Sub btnQte2_Click()
    txtQtePris.Value = btnQte2.Tag
End Sub

Private Sub btnQte3_Click()
    txtQtePris.Value = btnQte3.Tag
End Sub

Private Sub btnQte4_Click()
    txtQtePris.Value = btnQte4.Tag
End Sub

Private Sub btnQte5_Click()
    txtQtePris.Value = btnQte5.Tag
End Sub

Private Sub btnQRCode_Click()
    MsgBox "Scannez le QR Code et validez", vbInformation
    txtQRCodeValue.SetFocus
End Sub

Private Sub txtQRCodeValue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        SearchArticleByQR txtQRCodeValue.Value
        KeyAscii = 0
    End If
End Sub

Private Sub SearchArticleByQR(qrCode As String)
    Dim wsArticles As Worksheet
    Set wsArticles = ThisWorkbook.Worksheets("articles")
    
    Dim i As Long
    For i = 2 To wsArticles.Cells(wsArticles.Rows.Count, 1).End(xlUp).Row
        If wsArticles.Cells(i, 3).Value = qrCode Then
            txtObjet.Value = wsArticles.Cells(i, 2).Value
            currentArticleID = wsArticles.Cells(i, 1).Value
            Exit Sub
        End If
    Next i
    
    MsgBox "QR Code non trouve", vbExclamation
End Sub

Private Sub btnValidation_Click()
    ' Validation
    If Not M_Core.ValidateRequiredFields( _
        txtEmprunteur, "Emprunteur", _
        txtRaisonPret, "Raison", _
        txtObjet, "Objet", _
        txtQtePris, "Quantite", _
        txtTechnicien, "Technicien" _
    ) Then Exit Sub
    
    ' Verifier que l'article a ete scanne ID existant
    If currentArticleID = "" Then
        MsgBox "Veuillez scanner un article avec le QR Code", vbExclamation
        Exit Sub
    End If
    
    ' Creer pret
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    Dim newRow As Long
    newRow = wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row + 1
    
    With wsPrets
        .Cells(newRow, 1).Value = txtNumPret.Value
        .Cells(newRow, 2).Value = currentTech
        .Cells(newRow, 3).Value = currentEmprunteur
        .Cells(newRow, 4).Value = Now
        .Cells(newRow, 5).Value = currentArticleID
        .Cells(newRow, 6).Value = txtObjet.Value
        .Cells(newRow, 7).Value = txtQtePris.Value
        .Cells(newRow, 8).Value = txtDateRetour.Value
        .Cells(newRow, 9).Value = currentRaison
    End With
    
    ' Email
    Dim loanData As Object
    Set loanData = CreateObject("Scripting.Dictionary")
    loanData("ID") = txtNumPret.Value
    loanData("Email") = currentEmail
    loanData("Emprunteur") = currentEmprunteur
    loanData("Technicien") = currentTech
    loanData("Date") = txtDateHeure.Value
    loanData("Raison") = currentRaison
    loanData("Article") = txtObjet.Value
    loanData("Quantite") = txtQtePris.Value
    loanData("RetourPrevu") = txtDateRetour.Value
    
    Call M_Business.SendLoanEmail(loanData)
    
    MsgBox "Pret cree avec succes!", vbInformation
    
    Me.Hide
    LoanHub.Show
End Sub

Private Sub btnQuitter_Click()
    Me.Hide
    LoanHub.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnQuitter_Click
    End If
End Sub
