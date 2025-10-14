VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Dashboard 
   Caption         =   "Tableau de Bord - Statistiques"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   OleObjectBlob   =   "Dashboard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Dashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Contrôles
Private lblTitle As MSForms.Label
Private frameStatsGlobales As MSForms.Frame
Private lblTotalArticles As MSForms.Label
Private lblPretsEnCours As MSForms.Label
Private lblTauxUtilisation As MSForms.Label
Private lblTotalPrets As MSForms.Label
Private frameAlertes As MSForms.Frame
Private lblAlertesCritiques As MSForms.Label
Private lblAlertesAvertissement As MSForms.Label
Private lstAlertesDetail As MSForms.ListBox
Private frameTopArticles As MSForms.Frame
Private lstTopArticles As MSForms.ListBox
Private btnRefresh As MSForms.CommandButton
Private btnExportStats As MSForms.CommandButton
Private btnQuit As MSForms.CommandButton

Private Sub UserForm_Initialize()
    With Me
        .Width = 760
        .Height = 630
        .BackColor = M_Core.COLOR_LIGHT
    End With
    
    CreateInterface
    LoadStats
End Sub

Private Sub CreateInterface()
    ' Titre
    Set lblTitle = Me.Controls.Add("Forms.Label.1", "lblTitle")
    With lblTitle
        .Caption = "📊 TABLEAU DE BORD - STATISTIQUES"
        .Left = 12
        .Top = 12
        .Width = 736
        .Height = 40
        .Font.Name = "Segoe UI"
        .Font.Size = 18
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
        .BackColor = M_Core.COLOR_PRIMARY
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Frame Stats Globales
    Set frameStatsGlobales = Me.Controls.Add("Forms.Frame.1", "frameStatsGlobales")
    With frameStatsGlobales
        .Caption = "Statistiques Globales"
        .Left = 12
        .Top = 65
        .Width = 360
        .Height = 160
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = RGB(255, 255, 255)
    End With
    
    Set lblTotalArticles = frameStatsGlobales.Controls.Add("Forms.Label.1", "lblTotalArticles")
    With lblTotalArticles
        .Caption = "Total Articles : --"
        .Left = 15
        .Top = 30
        .Width = 330
        .Height = 25
        .Font.Size = 11
        .Font.Bold = False
    End With
    
    Set lblPretsEnCours = frameStatsGlobales.Controls.Add("Forms.Label.1", "lblPretsEnCours")
    With lblPretsEnCours
        .Caption = "Prêts en cours : --"
        .Left = 15
        .Top = 60
        .Width = 330
        .Height = 25
        .Font.Size = 11
        .Font.Bold = False
    End With
    
    Set lblTauxUtilisation = frameStatsGlobales.Controls.Add("Forms.Label.1", "lblTauxUtilisation")
    With lblTauxUtilisation
        .Caption = "Taux d'utilisation : --%"
        .Left = 15
        .Top = 90
        .Width = 330
        .Height = 25
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = M_Core.COLOR_PRIMARY
    End With
    
    Set lblTotalPrets = frameStatsGlobales.Controls.Add("Forms.Label.1", "lblTotalPrets")
    With lblTotalPrets
        .Caption = "Total Prêts historiques : --"
        .Left = 15
        .Top = 120
        .Width = 330
        .Height = 25
        .Font.Size = 11
        .Font.Bold = False
    End With
    
    ' Frame Alertes
    Set frameAlertes = Me.Controls.Add("Forms.Frame.1", "frameAlertes")
    With frameAlertes
        .Caption = "🚨 Alertes"
        .Left = 388
        .Top = 65
        .Width = 360
        .Height = 160
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = RGB(255, 240, 240)
    End With
    
    Set lblAlertesCritiques = frameAlertes.Controls.Add("Forms.Label.1", "lblAlertesCritiques")
    With lblAlertesCritiques
        .Caption = "🔴 Prêts > 30 jours : --"
        .Left = 15
        .Top = 35
        .Width = 330
        .Height = 25
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = M_Core.COLOR_DANGER
    End With
    
    Set lblAlertesAvertissement = frameAlertes.Controls.Add("Forms.Label.1", "lblAlertesAvertissement")
    With lblAlertesAvertissement
        .Caption = "🟠 Prêts > 15 jours : --"
        .Left = 15
        .Top = 70
        .Width = 330
        .Height = 25
        .Font.Size = 11
        .Font.Bold = False
        .ForeColor = M_Core.COLOR_WARNING
    End With
    
    Dim lblNote As MSForms.Label
    Set lblNote = frameAlertes.Controls.Add("Forms.Label.1", "lblNote")
    With lblNote
        .Caption = "Voir détail ci-dessous"
        .Left = 15
        .Top = 110
        .Width = 330
        .Height = 20
        .Font.Size = 9
        .ForeColor = RGB(127, 140, 141)
    End With
    
    ' Frame Top Articles
    Set frameTopArticles = Me.Controls.Add("Forms.Frame.1", "frameTopArticles")
    With frameTopArticles
        .Caption = "📈 Articles les plus prêtés (TOP 10)"
        .Left = 12
        .Top = 240
        .Width = 360
        .Height = 260
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    Set lstTopArticles = frameTopArticles.Controls.Add("Forms.ListBox.1", "lstTopArticles")
    With lstTopArticles
        .Left = 12
        .Top = 30
        .Width = 336
        .Height = 215
        .Font.Size = 9
    End With
    
    ' Liste Alertes détaillées
    Dim frameAlertesDetail As MSForms.Frame
    Set frameAlertesDetail = Me.Controls.Add("Forms.Frame.1", "frameAlertesDetail")
    With frameAlertesDetail
        .Caption = "Détail des Alertes"
        .Left = 388
        .Top = 240
        .Width = 360
        .Height = 260
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    Set lstAlertesDetail = frameAlertesDetail.Controls.Add("Forms.ListBox.1", "lstAlertesDetail")
    With lstAlertesDetail
        .Left = 12
        .Top = 30
        .Width = 336
        .Height = 215
        .Font.Size = 8
    End With
    
    ' Boutons
    Set btnRefresh = Me.Controls.Add("Forms.CommandButton.1", "btnRefresh")
    With btnRefresh
        .Caption = "🔄 Actualiser"
        .Left = 12
        .Top = 515
        .Width = 140
        .Height = 40
        .BackColor = M_Core.COLOR_PRIMARY
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    Set btnExportStats = Me.Controls.Add("Forms.CommandButton.1", "btnExportStats")
    With btnExportStats
        .Caption = "📤 Exporter Rapport"
        .Left = 170
        .Top = 515
        .Width = 180
        .Height = 40
        .BackColor = M_Core.COLOR_SUCCESS
        .ForeColor = RGB(255, 255, 255)
        .Font.Size = 11
        .Font.Bold = True
    End With
    
    Set btnQuit = Me.Controls.Add("Forms.CommandButton.1", "btnQuit")
    With btnQuit
        .Caption = "Retour Menu"
        .Left = 620
        .Top = 515
        .Width = 128
        .Height = 40
        .BackColor = RGB(189, 195, 199)
        .Font.Size = 10
    End With
End Sub

Private Sub LoadStats()
    Dim stats As Object
    Set stats = M_Business.GetDashboardStats()
    
    If Not stats Is Nothing Then
        lblTotalArticles.Caption = "Total Articles : " & stats("TotalArticles")
        lblPretsEnCours.Caption = "Prêts en cours : " & stats("PretsEnCours")
        lblTauxUtilisation.Caption = "Taux d'utilisation : " & stats("TauxUtilisation")
        lblTotalPrets.Caption = "Total Prêts historiques : " & stats("TotalPrets")
        
        lblAlertesCritiques.Caption = "🔴 Prêts > 30 jours : " & stats("PretsDepasses")
        lblAlertesAvertissement.Caption = "🟠 Prêts > 15 jours : " & stats("PretsAvertissement")
    End If
    
    LoadTopArticles
    LoadAlertesDetail
End Sub

Private Sub LoadTopArticles()
    ' Top 10 articles
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    Dim dictArticles As Object
    Set dictArticles = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim articleNom As String
    
    ' Compter prêts par article
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        articleNom = wsPrets.Cells(i, 6).Value
        If articleNom <> "" Then
            If dictArticles.Exists(articleNom) Then
                dictArticles(articleNom) = dictArticles(articleNom) + 1
            Else
                dictArticles.Add articleNom, 1
            End If
        End If
    Next i
    
    ' Trier et afficher top 10
    lstTopArticles.Clear
    
    Dim maxCount As Long
    Dim maxArticle As String
    Dim rank As Long
    rank = 1
    
    Do While dictArticles.count > 0 And rank <= 10
        maxCount = 0
        maxArticle = ""
        
        ' Trouver max
        Dim key As Variant
        For Each key In dictArticles.Keys
            If dictArticles(key) > maxCount Then
                maxCount = dictArticles(key)
                maxArticle = key
            End If
        Next key
        
        If maxArticle <> "" Then
            lstTopArticles.AddItem rank & ". " & maxArticle & " (" & maxCount & " prêts)"
            dictArticles.Remove maxArticle
            rank = rank + 1
        End If
    Loop
End Sub

Private Sub LoadAlertesDetail()
    lstAlertesDetail.Clear
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    Dim i As Long
    Dim daysElapsed As Long
    Dim emprunteur As String
    Dim article As String
    Dim datePret As Date
    
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 15).Value = "" Then ' En cours
            datePret = wsPrets.Cells(i, 4).Value
            daysElapsed = DateDiff("d", datePret, Now)
            
            If daysElapsed >= 15 Then
                emprunteur = wsPrets.Cells(i, 3).Value
                article = wsPrets.Cells(i, 6).Value
                
                Dim icon As String
                If daysElapsed >= 30 Then
                    icon = "🔴"
                Else
                    icon = "🟠"
                End If
                
                lstAlertesDetail.AddItem icon & " " & daysElapsed & "j - " & emprunteur & " - " & article
            End If
        End If
    Next i
    
    If lstAlertesDetail.ListCount = 0 Then
        lstAlertesDetail.AddItem "✅ Aucune alerte"
    End If
End Sub

' =====================================================
' ÉVÉNEMENTS
' =====================================================

Private Sub btnRefresh_Click()
    LoadStats
    MsgBox "Statistiques actualisées!", vbInformation
End Sub

Private Sub btnExportStats_Click()
    MsgBox "Export rapport statistiques en développement", vbInformation
    ' À implémenter: Export vers Excel ou PDF
End Sub

Private Sub btnQuit_Click()
    Me.Hide
    MainMenu.Show
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        btnQuit_Click
    End If
End Sub
