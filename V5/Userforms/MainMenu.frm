VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainMenu 
   Caption         =   "ACCUEIL - Gestion Prêts ESAD"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10800
   OleObjectBlob   =   "MainMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Références contrôles créés dynamiquement
Private lblTitle As MSForms.Label
Private lblVersion As MSForms.Label
Private frameMiniDashboard As MSForms.Frame
Private lblPretsEnCours As MSForms.Label
Private lblAlertes As MSForms.Label
Private btnLoanHub As MSForms.CommandButton
Private btnManageData As MSForms.CommandButton
Private btnDashboard As MSForms.CommandButton
Private btnExport As MSForms.CommandButton
Private btnQuit As MSForms.CommandButton

Private Sub UserForm_Initialize()
    ' Configuration fenêtre
    With Me
        .Width = 720
        .Height = 540
        .BackColor = M_Core.COLOR_LIGHT
        .Caption = "ACCUEIL - Gestion Prêts ESAD"
    End With
    
    ' Créer interface
    CreateTitle
    CreateMiniDashboard
    CreateMainButtons
    CreateFooter
End Sub

Private Sub CreateTitle()
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lblTitle
        .Caption = "GESTION DES PRÊTS"
        .Left = 50
        .Top = 30
        .Width = 620
        .Height = 50
        .Font.Name = "Segoe UI"
        .Font.Size = 24
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .ForeColor = M_Core.COLOR_DARK
        .BackColor = RGB(255, 255, 200)
        .BorderStyle = fmBorderStyleSingle
    End With
End Sub

Private Sub CreateMiniDashboard()
    ' Frame conteneur
    Set frameMiniDashboard = Me.Controls.Add("Forms.Frame.1", "frameMiniDashboard")
    With frameMiniDashboard
        .Caption = "Aperçu Rapide"
        .Left = 50
        .Top = 100
        .Width = 620
        .Height = 80
        .BackColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    ' Stats prêts en cours
    Set lblPretsEnCours = frameMiniDashboard.Controls.Add("Forms.Label.1", "lblPretsEnCours")
    With lblPretsEnCours
        .Caption = "Prêts en cours : " & GetActiveLoansCount()
        .Left = 15
        .Top = 25
        .Width = 280
        .Height = 25
        .Font.Size = 11
        .Font.Bold = False
    End With
    
    ' Stats alertes
    Set lblAlertes = frameMiniDashboard.Controls.Add("Forms.Label.1", "lblAlertes")
    With lblAlertes
        Dim alertCount As Long
        alertCount = GetAlertCount()
        .Caption = "Alertes : " & alertCount
        .Left = 320
        .Top = 25
        .Width = 280
        .Height = 25
        .Font.Size = 11
        .Font.Bold = False
        If alertCount > 0 Then
            .ForeColor = M_Core.COLOR_DANGER
            .Font.Bold = True
        End If
    End With
End Sub

Private Sub CreateMainButtons()
    ' Bouton Gestion Prêts
    Set btnLoanHub = Me.Controls.Add("Forms.CommandButton.1", "btnLoanHub")
    With btnLoanHub
        .Caption = "🎯 GESTION PRÊTS"
        .Left = 50
        .Top = 210
        .Width = 300
        .Height = 100
        .BackColor = M_Core.COLOR_PRIMARY
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    ' Bouton Gestion Données
    Set btnManageData = Me.Controls.Add("Forms.CommandButton.1", "btnManageData")
    With btnManageData
        .Caption = "📋 ARTICLES & EMPRUNTEURS"
        .Left = 370
        .Top = 210
        .Width = 300
        .Height = 100
        .BackColor = M_Core.COLOR_SUCCESS
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    ' Bouton Dashboard
    Set btnDashboard = Me.Controls.Add("Forms.CommandButton.1", "btnDashboard")
    With btnDashboard
        .Caption = "📊 STATISTIQUES"
        .Left = 50
        .Top = 330
        .Width = 300
        .Height = 100
        .BackColor = RGB(230, 126, 34)
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    ' Bouton Export
    Set btnExport = Me.Controls.Add("Forms.CommandButton.1", "btnExport")
    With btnExport
        .Caption = "📤 EXPORT INVENTAIRE"
        .Left = 370
        .Top = 330
        .Width = 300
        .Height = 100
        .BackColor = RGB(142, 68, 173)
        .ForeColor = RGB(255, 255, 255)
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
    End With
    
    ' Bouton Quitter
    Set btnQuit = Me.Controls.Add("Forms.CommandButton.1", "btnQuit")
    With btnQuit
        .Caption = "Sortie"
        .Left = 570
        .Top = 460
        .Width = 100
        .Height = 35
        .BackColor = RGB(189, 195, 199)
        .Font.Name = "Segoe UI"
        .Font.Size = 10
    End With
End Sub

Private Sub CreateFooter()
    Set lblVersion = Me.Controls.Add("Forms.Label.1", "lblVersion")
    With lblVersion
        .Caption = "Prêt matériel_V5 ESAD"
        .Left = 50
        .Top = 470
        .Width = 300
        .Height = 20
        .Font.Size = 9
        .ForeColor = RGB(127, 140, 141)
    End With
End Sub

' =====================================================
' ÉVÉNEMENTS BOUTONS
' =====================================================

Private Sub btnLoanHub_Click()
    Me.Hide
    LoanHub.Show
End Sub

Private Sub btnManageData_Click()
    Me.Hide
    ManageData.Show
End Sub

Private Sub btnDashboard_Click()
    Me.Hide
    Dashboard.Show
End Sub

Private Sub btnExport_Click()
    Call M_Business.ExportInventaireComplet
End Sub

Private Sub btnQuit_Click()
    Unload Me
End Sub

' =====================================================
' FONCTIONS UTILITAIRES
' =====================================================

Private Function GetActiveLoansCount() As Long
    On Error Resume Next
    Dim ws As Worksheet
    Dim i As Long, count As Long
    Set ws = ThisWorkbook.Worksheets("prets")
    count = 0
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 15).Value = "" Then count = count + 1
    Next i
    GetActiveLoansCount = count
End Function

Private Function GetAlertCount() As Long
    On Error Resume Next
    Dim ws As Worksheet
    Dim i As Long, count As Long
    Dim daysElapsed As Long
    Set ws = ThisWorkbook.Worksheets("prets")
    count = 0
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(i, 15).Value = "" Then
            daysElapsed = DateDiff("d", ws.Cells(i, 4).Value, Now)
            If daysElapsed >= 15 Then count = count + 1
        End If
    Next i
    GetAlertCount = count
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnQuit_Click
    End If
End Sub
