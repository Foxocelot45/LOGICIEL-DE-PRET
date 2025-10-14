VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} articles 
   Caption         =   "UserForm1"
   ClientHeight    =   11715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16230
   OleObjectBlob   =   "articles.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "articles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Text
Dim g, NomTableau4, TabBD(), ColCombo4(), colVisu4(), colInterro(), NcolVisu4, NbCol4, NcolInt4, Choix()



Private Sub b_recup_Click()

End Sub

Private Sub UserForm_Initialize()
Set g = Sheets("articles")
Set Rng = g.Range("A2:j" & g.[A65000].End(xlUp).Row)     ' à adapter
 NomTableau4 = "Tableau4"
' ActiveWorkbook.Names.Add Name:=NomTableau4                                    ' A adapter
'ActiveWorkbook.Names.Add Name:=NomTableau4, RefersTo:=Rng                                     ' A adapter
 NbCol4 = Range(NomTableau4).Columns.Count
 
 '---- A adapter
 TabBD = Range(NomTableau4).Resize(, NbCol4 + 1).Value              ' Array: + rapide
 For ii = 1 To UBound(TabBD): TabBD(ii, NbCol4 + 1) = ii: Next ii      ' No enregistrement
 ColCombo4 = Array(2)                                 ' A adapter (1 à 6 colonnes maxi)
 colVisu4 = Array(2, 4, 7)      ' Colonnes ListBox (à adapter)
 colInterro = Array(2, 4, 7)   ' colonnes à interroger (adapter)
 '----
 
 NcolInt4 = UBound(colInterro) + 1
 Me.ListBox10.List = TabBD
 For ii = UBound(ColCombo4) + 1 To 5
   Me("combobox" & ii + 1).Visible = False: Me("labelCbx" & ii + 1).Visible = False
 Next ii
 For c = 1 To UBound(ColCombo4) + 1: Me("combobox" & c) = "*": Next c
 For c = 1 To UBound(ColCombo4) + 1: ListeCol c: Next c
 For ii = 1 To UBound(ColCombo4) + 1:  Me("labelCbx" & ii) = Range(NomTableau4).Offset(-1).Item(1, ColCombo4(ii - 1)): Next ii
 Me.ListBox10.ColumnCount = NbCol4 + 1
 Me.ListBox10.ColumnWidths = "0;100;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0"
 '-- en têtes de colonnes ListBox
 EnteteListBox           ' Supprimer sur Excel 2013
 '-- labels textbox
 LabelsTextBox
 For ii = NbCol4 + 1 To 40: Me("textbox" & ii).Visible = False: Next ii
 For ii = NbCol4 + 1 To 40: Me("label" & ii).Visible = False: Next ii
 '-- colTri
 Me.ComboTri10.List = Application.Transpose(Range(NomTableau4).Offset(-1).Resize(1))  ' Ordre tri
 Affiche

'B_ajout_Click
End Sub
 Sub EnteteListBox()
   X = Me.ListBox10.Left + 8
   y = Me.ListBox10.Top - 20
   For c = 1 To NbCol4
     pos = Application.Match(c, colVisu4, 0)
     If Not IsError(pos) Then
       k = c
       Set Lab = Me.Controls.Add("Forms.Label.1")
       Lab.Caption = Range(NomTableau4).Offset(-1).Item(1, c)
       Lab.Top = y
       Lab.Left = X
       Lab.Height = 24
       Lab.Width = Range(NomTableau4).Columns(c).Width * 1#
       X = X + Range(NomTableau4).Columns(c).Width * 1
       tempcol = tempcol & Range(NomTableau4).Columns(c).Width * 1# & ";"
     Else
       X = X + 0
       tempcol = tempcol & 0 & ";"
     End If
   Next c
   tempcol = tempcol & "20"
   On Error Resume Next
   Me.ListBox10.ColumnWidths = tempcol
   On Error GoTo 0
 End Sub
Sub LabelsTextBox()
   For c = 1 To NbCol4
'      Me("textbox" & c).Width = Range(NomTableau4).Columns(c).Width * 1.3
      tmp = Range(NomTableau4).Offset(-1).Item(1, c)
      Me("label" & c).Caption = tmp
      lg = Len(tmp): If Len(tmp) > 11 Then lg = 11
      Me("label" & c).Width = lg * 6
   Next
End Sub
Sub ListeCol(noCol)
  Set d = CreateObject("Scripting.Dictionary")
  d.CompareMode = vbTextCompare
  For i = 1 To UBound(TabBD)
     ok = True
     For Cb = 0 To UBound(ColCombo4)
       colBD = ColCombo4(Cb)
       If Cb + 1 <> noCol Then
         If Not TabBD(i, colBD) Like Me("comboBox" & Cb + 1) Then ok = False
       End If
     Next Cb
     If ok Then
       tmp = TabBD(i, ColCombo4(noCol - 1))
       d(tmp) = ""
     End If
   Next i
   d("*") = ""
   temp = d.keys
   Tri temp, LBound(temp), UBound(temp)
   Me("ComboBox" & noCol).List = temp
End Sub
Private Sub B_tout_Click()
  For i = 1 To 6: Me("combobox" & i) = "*": Next i
End Sub
Private Sub ListBox10_Click()
  For i = 1 To NbCol4
    tmp = Me.ListBox10.Column(i - 1)
    If Not IsError(tmp) Then Me("textbox" & i) = tmp
  Next i
  Me.Enreg10 = Me.ListBox10.Column(NbCol4)
End Sub
Sub Affiche()
  Dim Tbl()
  cbx1 = Me.ComboBox1: cbx2 = Me.ComboBox2:  cbx3 = Me.ComboBox3:  cbx4 = Me.ComboBox4: cbx5 = Me.ComboBox5: cbx6 = Me.ComboBox6
  n = 0
  Cb = Array(1, 1, 1, 1, 1, 1)
  For i = 0 To UBound(ColCombo4): Cb(i) = ColCombo4(i): Next i
  For i = 1 To UBound(TabBD)
    If TabBD(i, Cb(0)) Like cbx1 And TabBD(i, Cb(1)) Like cbx2 _
       And TabBD(i, Cb(2)) Like cbx3 And TabBD(i, Cb(3)) Like cbx4 And TabBD(i, Cb(4)) Like cbx5 And TabBD(i, Cb(5)) Like cbx6 Then
        n = n + 1: ReDim Preserve Tbl(1 To NbCol4 + 1, 1 To n)
        c = 0
        For c = 1 To NbCol4: Tbl(c, n) = TabBD(i, c): Next c
        'Tbl(6, n) = Format(TabBD(i, 6), "hh:mm")
        Tbl(c, n) = TabBD(i, NbCol4 + 1)
    End If
  Next i
  If n > 0 Then
     Me.ListBox10.Column = Tbl
  Else
     Me.ListBox10.Clear
  End If
  Gchoix
End Sub
Private Sub ComboBox1_DropButtonClick()
   ListeCol 1
End Sub
Private Sub ComboBox2_DropButtonClick()
   ListeCol 2
End Sub
Private Sub ComboBox3_DropButtonClick()
  ListeCol 3
End Sub
Private Sub ComboBox4_DropButtonClick()
  ListeCol 4
End Sub
Private Sub ComboBox5_DropButtonClick()
  ListeCol 5
End Sub
Private Sub ComboBox6_DropButtonClick()
  ListeCol 6
End Sub
Private Sub ComboBox1_Change()
  Affiche
End Sub
Private Sub ComboBox2_Change()
 Affiche
End Sub
Private Sub ComboBox3_Change()
  Affiche
End Sub
Private Sub ComboBox4_Change()
  Affiche
End Sub
Private Sub ComboBox5_Change()
  Affiche
End Sub
Private Sub ComboBox6_Change()
  Affiche
End Sub

Private Sub B_valid_Click()
  Enreg = Me.Enreg10
  For c = 1 To NbCol4
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   
   If Not Range(NomTableau4).Item(Enreg, c).HasFormula Then
     tmp = Me("textbox" & c)
     If IsNumeric(Replace(tmp, ".", ",")) And InStr(tmp, " ") = 0 Then
        tmp = Replace(tmp, ".", ",")
        Range(NomTableau4).Item(Enreg, c) = CDbl(tmp)
     Else
         If IsDate(tmp) Then
           Range(NomTableau4).Item(Enreg, c) = CDate(tmp)
         Else
           Range(NomTableau4).Item(Enreg, c) = tmp
         End If
     End If
    Else
     Range(NomTableau4).Item(Enreg - 1, c).Copy
     Range(NomTableau4).Item(Enreg, c).PasteSpecial Paste:=xlPasteFormats
    End If
  Next c
  UserForm_Initialize
  raz
End Sub
Private Sub B_ajout_Click()
 raz
 Me.Enreg10 = Range(NomTableau4).Rows.Count + 1
 TextBox1.Value = Me.Enreg10
 Me.textbox2.SetFocus
End Sub
Private Sub B_sup_Click()
If Me.Enreg <> "" Then
  If MsgBox("Etes vous sûr de suppimer " & Me.TextBox1 & "?", vbYesNo) = vbYes Then
    [Tableau1].Rows(Me.Enreg).Delete
    Me.Enreg = ""
    UserForm_Initialize
    raz
    Me.Enreg = Range(NomTableau4).Rows.Count + 1
  End If
 End If
End Sub
Sub raz()
    For k = 1 To NbCol4
      Me("textBox" & k) = ""
    Next k
    Me.TextBox1.SetFocus
End Sub
'Private Sub B_duplique_Click()
'  Me.Enreg = Range(NomTableau4).Rows.Count + 1
'  B_valid_Click
'End Sub
Sub Tri(a, gauc, droi) ' Quick sort
 ref = CStr(a((gauc + droi) \ 2))
 g = gauc: d = droi
 Do
  Do While CStr(a(g)) < ref: g = g + 1: Loop
  Do While ref < CStr(a(d)): d = d - 1: Loop
  If g <= d Then
    temp = a(g): a(g) = a(d): a(d) = temp
    g = g + 1: d = d - 1
  End If
 Loop While g <= d
 If g < droi Then Call Tri(a, g, droi)
 If gauc < d Then Call Tri(a, gauc, d)
End Sub
Private Sub ComboMenu_click()
  nomcontrole = Me.TextBoxActive
  Me(nomcontrole) = Me.ComboMenu.Value
  Me.ComboMenu.Visible = False
End Sub
Private Sub ComboTri10_click()
  Dim Tbl()
  colTri = Me.ComboTri10.ListIndex
  Tbl = Me.ListBox10.List
  TriMultiCol Tbl, LBound(Tbl), UBound(Tbl), colTri
  Me.ListBox10.List = Tbl
End Sub
Sub TriMultiCol(a(), gauc, droi, colTri) ' Quick sort
  Dim colD, colF, ref, g, d, c, temp
  colD = LBound(a, 2): colF = UBound(a, 2)
  ref = a((gauc + droi) \ 2, colTri)
  g = gauc: d = droi
  Do
    Do While a(g, colTri) < ref: g = g + 1: Loop
    Do While ref < a(d, colTri): d = d - 1: Loop
    If g <= d Then
      For c = colD To colF
        temp = a(g, c): a(g, c) = a(d, c): a(d, c) = temp
      Next
      g = g + 1: d = d - 1
    End If
  Loop While g <= d
  If g < droi Then TriMultiCol a, g, droi, colTri
  If gauc < d Then TriMultiCol a, gauc, d, colTri
End Sub
Sub Gchoix()
  '-- génération de choix()
  BDListBox = Me.ListBox10.List

ReDim Choix(1 To UBound(BDListBox) + 1)
  col = UBound(BDListBox, 2)
  For i = LBound(BDListBox) To UBound(BDListBox)
     For Each k In colInterro
       Choix(i + 1) = Choix(i + 1) & BDListBox(i, k - 1) & "|"
     Next k
     Choix(i + 1) = Choix(i + 1) & BDListBox(i, col) & "|" ' no enreg
  Next i
  Me.TextBoxRech10 = ""
End Sub
Private Sub TextBoxRech10_Change()
  If Me.TextBoxRech10 <> "" Then
     mots = Split(Trim(Me.TextBoxRech10), " ")
     Tbl = Choix
     For i = LBound(mots) To UBound(mots)
        Tbl = Filter(Tbl, mots(i), True, vbTextCompare)
     Next i
     If UBound(Tbl) > -1 Then
        Dim b(): ReDim b(1 To UBound(Tbl) + 1, 1 To NbCol4 + 1)
        For i = LBound(Tbl) To UBound(Tbl)
          a = Split(Tbl(i), "|")
          j = a(NcolInt4)
          For c = 1 To NbCol4: b(i + 1, c) = TabBD(j, c): Next c
          b(i + 1, c) = j
        Next i
        Me.ListBox10.List = b
     Else
       Me.ListBox10.Clear
     End If
  Else
     Affiche
     'UserForm_Initialize
  End If
End Sub
Private Sub B_prédent_Click()
 If Me.ListBox10.ListIndex > 0 Then
    Me.ListBox10.ListIndex = Me.ListBox10.ListIndex - 1
 End If
End Sub

Private Sub B_suivant_Click()
 If Me.ListBox10.ListIndex < Me.ListBox10.ListCount - 1 Then
    Me.ListBox10.ListIndex = Me.ListBox10.ListIndex + 1
 End If
End Sub


