VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} qrcode 
   Caption         =   "UserForm1"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8385
   OleObjectBlob   =   "qrcode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "qrcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Text
Dim f, Rng, TblBD(), NbCol, NomTableau

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()
  Unload qrcode
'MENU_GENERAL.Show


End Sub

Private Sub CommandButton2_Click()
textbox4.Value = Now
End Sub

Private Sub CommandButton3_Click()

End Sub

Private Sub CommandButton4_Click()
If textbox2 = "" Then MsgBox "Nom d'objet non renseigné": textbox2.SetFocus: Exit Sub
pretobjet.TextBox5 = textbox2
pretobjet.TextBox7 = TextBox7

  Unload qrcode
End Sub

Private Sub CommandButton5_Click()
articles.Show
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub UserForm_Initialize()
 On Error Resume Next
   NomTableau = "Tableau4"                                   ' à Adapter
   NbCol = Range(NomTableau).Columns.Count
   TblBD = Range(NomTableau).Resize(, NbCol + 1).Value       ' Array: + rapide
   For i = 1 To UBound(TblBD): TblBD(i, NbCol + 1) = i: Next i  ' No enregistrement
   Me.ListBox1.List = TblBD
   Me.ListBox1.ColumnCount = NbCol + 1
   Me.ListBox1.ColumnWidths = "0;100;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0"
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
   
'Dim tablo2 As Variant
'With Feuil3 'type service
'derligne = .Range("a65536").End(xlUp).Row
'tablo2 = .Range("a2:a" & derligne)
'TextBox3.List = tablo2
'End With
'
'Dim tablo3 As Variant 'Fonction
'With Feuil4 'fonction
'derligne = .Range("a65536").End(xlUp).Row
'tablo3 = .Range("a2:a" & derligne)
'TextBox4.List = tablo3
'End With
Frame1.Visible = False
CommandButton4.Visible = False
CommandButton5.Visible = False
    Me.TextBoxRech.SetFocus
   
   
   
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
  
     If Me.ListBox1.ListIndex < Me.ListBox1.ListCount - 1 Then
   Me.ListBox1.ListIndex = Me.ListBox1.ListIndex + 1
   
  End If
  If TextBoxRech.Value = TextBox7.Value Then
  Frame1.Visible = True
  CommandButton4.Visible = True
  End If
    If TextBoxRech.Value <> TextBox7.Value Then
CommandButton5.Visible = True
  End If
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
    Me.TextBoxRech.SetFocus
End Sub
Private Sub B_ajout_Click()
 raz
 Me.Enreg = Range(NomTableau).Rows.Count + 1
 Me.TextBoxRech.SetFocus
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





