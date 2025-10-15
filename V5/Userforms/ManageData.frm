VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ManageData 
   Caption         =   "Gestion Donnees"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11400
   OleObjectBlob   =   "ManageData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ManageData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' Controles
Private tabControl As MSForms.MultiPage
Private btnQuit As MSForms.CommandButton

Private Sub UserForm_Initialize()
    With Me
        .Width = 760
        .Height = 630
        .BackColor = M_Core.COLOR_LIGHT
    End With
    
    CreateInterface
    LoadData
End Sub

Private Sub CreateInterface()
    ' MultiPage 3 onglets
    Set tabControl = Me.Controls.Add("Forms.MultiPage.1", "tabControl")
    With tabControl
        .Left = 12
        .Top = 12
        .Width = 736
        .Height = 540
    End With
    
    tabControl.Pages(0).Caption = "Articles"
    tabControl.Pages.Add , "tab1"
    tabControl.Pages(1).Caption = "Emprunteurs"
    tabControl.Pages.Add , "tab2"
    tabControl.Pages(2).Caption = "Articles en Pret"
    
    CreateTab0_Articles
    CreateTab1_Emprunteurs
    CreateTab2_ArticlesEnPret
    
    ' Bouton Quitter
    Set btnQuit = Me.Controls.Add("Forms.CommandButton.1", "btnQuit")
    With btnQuit
        .Caption = "Retour Menu"
        .Left = 630
        .Top = 565
        .Width = 100
        .Height = 30
        .Font.Size = 10
    End With
End Sub

Private Sub CreateTab0_Articles()
    Dim page As MSForms.page
    Set page = tabControl.Pages(0)
    
    ' Recherche
    Dim lblSearch As MSForms.Label
    Set lblSearch = page.Controls.Add("Forms.Label.1", "lblSearchArticle")
    With lblSearch
        .Caption = "Rechercher:"
        .Left = 12
        .Top = 12
        .Width = 80
        .Height = 20
    End With
    
    Dim txtSearch As MSForms.TextBox
    Set txtSearch = page.Controls.Add("Forms.TextBox.1", "txtSearchArticle")
    With txtSearch
        .Left = 100
        .Top = 12
        .Width = 300
        .Height = 24
        .BackColor = RGB(255, 255, 200)
    End With
    
    ' Liste articles
    Dim lstArticles As MSForms.ListBox
    Set lstArticles = page.Controls.Add("Forms.ListBox.1", "lstArticles")
    With lstArticles
        .Left = 12
        .Top = 48
        .Width = 700
        .Height = 300
        .Font.Size = 9
    End With
    
    ' Boutons CRUD
    Dim btnNew As MSForms.CommandButton
    Set btnNew = page.Controls.Add("Forms.CommandButton.1", "btnNewArticle")
    With btnNew
        .Caption = "Nouveau"
        .Left = 12
        .Top = 360
        .Width = 100
        .Height = 30
        .BackColor = M_Core.COLOR_SUCCESS
        .ForeColor = RGB(255, 255, 255)
    End With
    
    Dim btnEdit As MSForms.CommandButton
    Set btnEdit = page.Controls.Add("Forms.CommandButton.1", "btnEditArticle")
    With btnEdit
        .Caption = "Modifier"
        .Left = 120
        .Top = 360
        .Width = 100
        .Height = 30
        .BackColor = M_Core.COLOR_WARNING
        .ForeColor = RGB(255, 255, 255)
    End With
    
    Dim btnDelete As MSForms.CommandButton
    Set btnDelete = page.Controls.Add("Forms.CommandButton.1", "btnDeleteArticle")
    With btnDelete
        .Caption = "Supprimer"
        .Left = 228
        .Top = 360
        .Width = 100
        .Height = 30
        .BackColor = M_Core.COLOR_DANGER
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Frame details
    Dim frameDetails As MSForms.Frame
    Set frameDetails = page.Controls.Add("Forms.Frame.1", "frameDetailsArticle")
    With frameDetails
        .Caption = "Details Article"
        .Left = 12
        .Top = 400
        .Width = 700
        .Height = 120
    End With
    
    Dim lblDetails As MSForms.Label
    Set lblDetails = frameDetails.Controls.Add("Forms.Label.1", "lblDetailsArticle")
    With lblDetails
        .Left = 12
        .Top = 24
        .Width = 676
        .Height = 80
        .Caption = "Selectionnez un article"
        .Font.Size = 9
    End With
End Sub

Private Sub CreateTab1_Emprunteurs()
    Dim page As MSForms.page
    Set page = tabControl.Pages(1)
    
    ' Recherche
    Dim lblSearch As MSForms.Label
    Set lblSearch = page.Controls.Add("Forms.Label.1", "lblSearchEmp")
    With lblSearch
        .Caption = "Rechercher:"
        .Left = 12
        .Top = 12
        .Width = 80
        .Height = 20
    End With
    
    Dim txtSearch As MSForms.TextBox
    Set txtSearch = page.Controls.Add("Forms.TextBox.1", "txtSearchEmp")
    With txtSearch
        .Left = 100
        .Top = 12
        .Width = 300
        .Height = 24
        .BackColor = RGB(255, 255, 200)
    End With
    
    ' Liste emprunteurs
    Dim lstEmp As MSForms.ListBox
    Set lstEmp = page.Controls.Add("Forms.ListBox.1", "lstEmprunteurs")
    With lstEmp
        .Left = 12
        .Top = 48
        .Width = 700
        .Height = 300
        .Font.Size = 9
    End With
    
    ' Boutons CRUD
    Dim btnNew As MSForms.CommandButton
    Set btnNew = page.Controls.Add("Forms.CommandButton.1", "btnNewEmp")
    With btnNew
        .Caption = "Nouveau"
        .Left = 12
        .Top = 360
        .Width = 100
        .Height = 30
        .BackColor = M_Core.COLOR_SUCCESS
        .ForeColor = RGB(255, 255, 255)
    End With
    
    Dim btnEdit As MSForms.CommandButton
    Set btnEdit = page.Controls.Add("Forms.CommandButton.1", "btnEditEmp")
    With btnEdit
        .Caption = "Modifier"
        .Left = 120
        .Top = 360
        .Width = 100
        .Height = 30
        .BackColor = M_Core.COLOR_WARNING
        .ForeColor = RGB(255, 255, 255)
    End With
    
    Dim btnDelete As MSForms.CommandButton
    Set btnDelete = page.Controls.Add("Forms.CommandButton.1", "btnDeleteEmp")
    With btnDelete
        .Caption = "Supprimer"
        .Left = 228
        .Top = 360
        .Width = 100
        .Height = 30
        .BackColor = M_Core.COLOR_DANGER
        .ForeColor = RGB(255, 255, 255)
    End With
    
    ' Frame details
    Dim frameDetails As MSForms.Frame
    Set frameDetails = page.Controls.Add("Forms.Frame.1", "frameDetailsEmp")
    With frameDetails
        .Caption = "Details Emprunteur"
        .Left = 12
        .Top = 400
        .Width = 700
        .Height = 120
    End With
    
    Dim lblDetails As MSForms.Label
    Set lblDetails = frameDetails.Controls.Add("Forms.Label.1", "lblDetailsEmp")
    With lblDetails
        .Left = 12
        .Top = 24
        .Width = 676
        .Height = 80
        .Caption = "Selectionnez un emprunteur"
        .Font.Size = 9
    End With
End Sub

Private Sub CreateTab2_ArticlesEnPret()
    Dim page As MSForms.page
    Set page = tabControl.Pages(2)
    
    ' Titre
    Dim lblTitle As MSForms.Label
    Set lblTitle = page.Controls.Add("Forms.Label.1", "lblTitleEnPret")
    With lblTitle
        .Caption = "Articles Actuellement en Pret"
        .Left = 12
        .Top = 12
        .Width = 700
        .Height = 25
        .Font.Size = 12
        .Font.Bold = True
        .TextAlign = fmTextAlignCenter
    End With
    
    ' Filtres
    Dim lblFiltre As MSForms.Label
    Set lblFiltre = page.Controls.Add("Forms.Label.1", "lblFiltre")
    With lblFiltre
        .Caption = "Filtrer par emprunteur:"
        .Left = 12
        .Top = 50
        .Width = 130
        .Height = 20
    End With
    
    Dim txtFiltre As MSForms.TextBox
    Set txtFiltre = page.Controls.Add("Forms.TextBox.1", "txtFiltreEnPret")
    With txtFiltre
        .Left = 150
        .Top = 48
        .Width = 250
        .Height = 24
        .BackColor = RGB(255, 255, 200)
    End With
    
    Dim btnRefresh As MSForms.CommandButton
    Set btnRefresh = page.Controls.Add("Forms.CommandButton.1", "btnRefreshEnPret")
    With btnRefresh
        .Caption = "Actualiser"
        .Left = 410
        .Top = 46
        .Width = 80
        .Height = 28
    End With
    
    ' Liste prets
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = page.Controls.Add("Forms.ListBox.1", "lstArticlesEnPret")
    With lstPrets
        .Left = 12
        .Top = 85
        .Width = 700
        .Height = 420
        .Font.Size = 9
    End With
    
    ' Stats
    Dim lblStats As MSForms.Label
    Set lblStats = page.Controls.Add("Forms.Label.1", "lblStatsEnPret")
    With lblStats
        .Left = 500
        .Top = 48
        .Width = 212
        .Height = 24
        .Font.Size = 9
        .Caption = "Total prets en cours: 0"
    End With
End Sub

Private Sub LoadData()
    LoadArticlesData
    LoadEmpData
    LoadArticlesEnPretData
End Sub

Private Sub LoadArticlesData()
    Dim lstArticles As MSForms.ListBox
    Set lstArticles = tabControl.Pages(0).Controls("lstArticles")
    
    Dim arr As Variant
    arr = M_Core.LoadDataToArray("Tableau4")
    
    If IsArray(arr) Then
        M_Core.PopulateListBox lstArticles, arr
        lstArticles.ColumnCount = UBound(arr, 2)
        lstArticles.ColumnWidths = "0;200;100;150;150;80;0;0;0;0"
    End If
End Sub

Private Sub LoadEmpData()
    Dim lstEmp As MSForms.ListBox
    Set lstEmp = tabControl.Pages(1).Controls("lstEmprunteurs")
    
    Dim arr As Variant
    arr = M_Core.LoadDataToArray("Tableau1")
    
    If IsArray(arr) Then
        M_Core.PopulateListBox lstEmp, arr
        lstEmp.ColumnCount = UBound(arr, 2)
        lstEmp.ColumnWidths = "0;200;150;150;0;200;0;0;0"
    End If
End Sub

Private Sub LoadArticlesEnPretData()
    Dim lstPrets As MSForms.ListBox
    Set lstPrets = tabControl.Pages(2).Controls("lstArticlesEnPret")
    
    Dim lblStats As MSForms.Label
    Set lblStats = tabControl.Pages(2).Controls("lblStatsEnPret")
    
    Dim wsPrets As Worksheet
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    
    lstPrets.Clear
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 15).Value = "" Then
            lstPrets.AddItem
            lstPrets.List(lstPrets.ListCount - 1, 0) = wsPrets.Cells(i, 3).Value ' Emprunteur
            lstPrets.List(lstPrets.ListCount - 1, 1) = wsPrets.Cells(i, 6).Value ' Article
            lstPrets.List(lstPrets.ListCount - 1, 2) = wsPrets.Cells(i, 7).Value ' Qte
            lstPrets.List(lstPrets.ListCount - 1, 3) = wsPrets.Cells(i, 4).Value ' Date
            count = count + 1
        End If
    Next i
    
    lstPrets.ColumnCount = 4
    lstPrets.ColumnWidths = "200;300;50;100"
    
    lblStats.Caption = "Total prets en cours: " & count
End Sub

Public Sub SwitchToTab(tabIndex As Long)
    tabControl.Value = tabIndex
End Sub

' =====================================================
' EVENEMENTS
' =====================================================

Private Sub btnNewArticle_Click()
    MsgBox "Formulaire creation article en developpement", vbInformation
End Sub

Private Sub btnEditArticle_Click()
    MsgBox "Formulaire modification article en developpement", vbInformation
End Sub

Private Sub btnDeleteArticle_Click()
    MsgBox "Suppression article en developpement", vbInformation
End Sub

Private Sub btnNewEmp_Click()
    MsgBox "Formulaire creation emprunteur en developpement", vbInformation
End Sub

Private Sub btnEditEmp_Click()
    MsgBox "Formulaire modification emprunteur en developpement", vbInformation
End Sub

Private Sub btnDeleteEmp_Click()
    MsgBox "Suppression emprunteur en developpement", vbInformation
End Sub

Private Sub btnRefreshEnPret_Click()
    LoadArticlesEnPretData
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
