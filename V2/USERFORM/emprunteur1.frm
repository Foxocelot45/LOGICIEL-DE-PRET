VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} emprunteur1 
   Caption         =   "Liste des emprunteurs"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   21855
   OleObjectBlob   =   "emprunteur1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "emprunteur1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Text
Dim f, Rng, TblBD(), NbCol, NomTableau, TblBD2(), NbCol2, NomTableau2

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
  Unload emprunteur
'MENU_GENERAL.Show


End Sub

Private Sub CommandButton2_Click()
Fonction.Show
End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub Frame1_Click()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Initialize()
 On Error Resume Next
   NomTableau = "Tableau1"                                   ' à Adapter
   NbCol = Range(NomTableau).Columns.Count
   TblBD = Range(NomTableau).Resize(, NbCol + 1).Value       ' Array: + rapide
   For i = 1 To UBound(TblBD): TblBD(i, NbCol + 1) = i: Next i  ' No enregistrement
   Me.ListBox1.List = TblBD
   Me.ListBox1.ColumnCount = NbCol + 1
   Me.ListBox1.ColumnWidths = "0;150;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0"
   '--- ComboBox choix colonne filtre
   Me.ComboChoixColFiltre.List = Application.Transpose(Range(NomTableau).Offset(-1).Resize(1))
   Me.ComboTri.List = Application.Transpose(Range(NomTableau).Offset(-1).Resize(1))
   Me.ComboChoixColFiltre.ListIndex = 0
   Me.LabelColFiltre.Caption = "Filtre:" & Me.ComboChoixColFiltre
   '--- Combobox recherche
   Set d = CreateObject("scripting.dictionary")
   For i = 1 To UBound(TblBD)
     d(TblBD(i, 1)) = ""
   Next i
   temp = d.keys
   Tri temp, LBound(temp), UBound(temp)
   Me.ComboBoxRech.List = temp
   '--- Labels
   TblTitre = Application.Transpose(Range(NomTableau).Offset(-1).Resize(1))
   For i = 1 To NbCol
     Me("label" & i) = TblTitre(i, 1)
   Next i
   For i = NbCol + 1 To 18
      Me("label" & i).Visible = False: Me("TextBox" & i).Visible = False
   Next i
   '--- non standard  pour alimenter les comboboxs
   'Me.TextBox11.List = Array("Etablissement1", "Etablissement2", "Etablissement3", "Etablissement4")
   'Me.TextBox12.List = Array("Prestataire1", "Prestataire2", "Prestataire3", "Prestataire4")
   
Dim tablo2 As Variant
With Feuil3 'type service
derligne = .Range("a65536").End(xlUp).Row
tablo2 = .Range("a2:a" & derligne)
textbox3.List = tablo2
End With

Dim tablo3 As Variant 'Fonction
With Feuil4 'fonction
derligne = .Range("a65536").End(xlUp).Row
tablo3 = .Range("a2:a" & derligne)
textbox4.List = tablo3
End With


'      NomTableau2 = "Tableau4"                                   ' à Adapter
'   NbCol2 = Range(NomTableau2).Columns.Count
'   TblBD2 = Range(NomTableau2).Resize(, NbCol2 + 1).Value       ' Array: + rapide
'   For j = 1 To UBound(TblBD): TblBD(j, NbCol2 + 1) = j: Next j  ' No enregistrement
'   Me.ListBox2.List = TblBD2
'   Me.ListBox2.ColumnCount = NbCol2 + 1
'   Me.ListBox2.ColumnWidths = "0;150;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0"
'   '--- ComboBox choix colonne filtre
'   Me.ComboChoixColFiltre2.List = Application.Transpose(Range(NomTableau2).Offset(-1).Resize(1))
'   Me.ComboTri2.List = Application.Transpose(Range(NomTableau2).Offset(-1).Resize(1))
'   Me.ComboChoixColFiltre2.ListIndex = 0
'   Me.LabelColFiltre2.Caption = "Filtre:" & Me.ComboChoixColFiltre2
'   '--- Combobox recherche
'   Set e = CreateObject("scripting.dictionary")
'   For j = 1 To UBound(TblBD2)
'     e(TblBD2(j, 1)) = ""
'   Next j
'   temp2 = e.keys
'   Tri temp2, LBound(temp2), UBound(temp2)
'   Me.ComboBoxRech2.List = temp2
''   --- Labels
'   TblTitre2 = Application.Transpose(Range(NomTableau2).Offset(-1).Resize(1))
'   For j = 1 To NbCol2
'     Me("label" & j) = TblTitre(j, 1)
'   Next j
'   For j = NbCol + 1 To 18
'      Me("label" & j).Visible = False: Me("TextBox" & j).Visible = False
'   Next j
''
   
   
End Sub
Private Sub ComboChoixColFiltre_click()
  Me.LabelColFiltre.Caption = "Filtre:" & Me.ComboChoixColFiltre
  Me.LabelRech.Caption = "Recherche par:" & Me.ComboChoixColFiltre
  Set Titre = Range(NomTableau).Offset(-1).Resize(1) 'Rng.Offset(-1).Resize(1)
  colFiltre = Application.Match(Me.ComboChoixColFiltre, Titre, 0)
  Set d = CreateObject("scripting.dictionary")
  For i = 1 To UBound(TblBD)
    d(TblBD(i, colFiltre)) = ""
  Next i
  temp = d.keys
  Tri temp, LBound(temp), UBound(temp)
  Me.ComboBoxRech.List = temp
End Sub
Private Sub ComboTri_click()
  Dim Tbl()
  colTri = Me.ComboTri.ListIndex
  Tbl = Me.ListBox1.List
  TriMultiCol Tbl, LBound(Tbl), UBound(Tbl), colTri
  Me.ListBox1.List = Tbl
End Sub
Private Sub TextBoxRech_Change()
  colRecherche = Me.ComboChoixColFiltre.ListIndex + 1
  clé = "*" & Me.TextBoxRech & "*": n = 0
  Dim Tbl()
  For i = 1 To UBound(TblBD)
    If TblBD(i, colRecherche) Like clé Then
        n = n + 1: ReDim Preserve Tbl(1 To UBound(TblBD, 2), 1 To n)
        For k = 1 To UBound(TblBD, 2): Tbl(k, n) = TblBD(i, k): Next k
     End If
  Next i
  If n > 0 Then Me.ListBox1.Column = Tbl Else Me.ListBox1.Clear
End Sub
Private Sub ComboBoxRech_Change()
  colRecherche = Me.ComboChoixColFiltre.ListIndex + 1
  clé = Me.ComboBoxRech: n = 0
  Dim Tbl()
  For i = 1 To UBound(TblBD)
    If TblBD(i, colRecherche) Like clé Then
        n = n + 1: ReDim Preserve Tbl(1 To UBound(TblBD, 2), 1 To n)
        For k = 1 To UBound(TblBD, 2): Tbl(k, n) = TblBD(i, k): Next k
     End If
  Next i
  If n > 0 Then Me.ListBox1.Column = Tbl Else Me.ListBox1.Clear
End Sub
Private Sub B_SupFilte_Click()
   Me.ListBox1.List = TblBD
End Sub

Private Sub ListBox1_Click()
  For i = 1 To NbCol
    Me("textbox" & i) = Me.ListBox1.Column(i - 1)
  Next i
  Me.Enreg = Me.ListBox1.Column(i - 1)
End Sub
Private Sub B_valid_Click()
  Enreg = Me.Enreg
'TextBox11.Value = TextBox4.Value & " " & TextBox5.Value 'nom & prénom dans textbox 11
  For c = 1 To NbCol
   If Not Range(NomTableau).Item(Enreg, c).HasFormula Then
     tmp = Me("textbox" & c)
     If IsNumeric(Replace(tmp, ".", ",")) And InStr(tmp, " ") = 0 Then
        tmp = Replace(tmp, ".", ",")
        Range(NomTableau).Item(Enreg, c) = CDbl(tmp)
     Else
         If IsDate(tmp) Then
           Range(NomTableau).Item(Enreg, c) = CDate(tmp)
         Else
           Range(NomTableau).Item(Enreg, c) = tmp
         End If
     End If
    Else
     Range(NomTableau).Item(Enreg - 1, c).Copy
     Range(NomTableau).Item(Enreg, c).PasteSpecial Paste:=xlPasteFormats
    End If
  Next c
  UserForm_Initialize
  'raz
End Sub
Sub raz()
    For k = 1 To NbCol
      Me("textBox" & k) = ""
    Next k
    Me.TextBox1.SetFocus
End Sub
Private Sub B_ajout_Click()
 raz
 Me.Enreg = Range(NomTableau).Rows.Count + 1
 Me.TextBox1.SetFocus
 TextBox1 = Enreg.Value
End Sub
Private Sub B_sup_Click()
If Me.Enreg <> "" Then
  If MsgBox("Etes vous sûr de supprimer " & Me.TextBox1 & "?", vbYesNo) = vbYes Then
    Range(NomTableau).Rows(Me.Enreg).Delete
    Me.Enreg = ""
    UserForm_Initialize
    raz
    Me.Enreg = Range(NomTableau).Rows.Count + 1
  End If
 End If
End Sub
Sub Tri(a, gauc, droi) ' Quick sort
  ref = a((gauc + droi) \ 2)
  g = gauc: d = droi
  Do
    Do While a(g) < ref: g = g + 1: Loop
    Do While ref < a(d): d = d - 1: Loop
    If g <= d Then
      temp = a(g): a(g) = a(d): a(d) = temp
      g = g + 1: d = d - 1
    End If
  Loop While g <= d
  If g < droi Then Tri a, g, droi
  If gauc < d Then Tri a, gauc, d
End Sub
Sub TriMultiCol(a, gauc, droi, colTri) ' Quick sort
  ref = a((gauc + droi) \ 2, colTri)
  g = gauc: d = droi
  Do
    Do While a(g, colTri) < ref: g = g + 1: Loop
    Do While ref < a(d, colTri): d = d - 1: Loop
    If g <= d Then
       For c = LBound(a, 2) To UBound(a, 2)
          temp = a(g, c): a(g, c) = a(d, c): a(d, c) = temp
       Next
       g = g + 1: d = d - 1
     End If
   Loop While g <= d
   If g < droi Then TriMultiCol a, g, droi, colTri
   If gauc < d Then TriMultiCol a, gauc, d, colTri
End Sub



