VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pret 
   Caption         =   "Gestion des prêts"
   ClientHeight    =   10890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15615
   OleObjectBlob   =   "pret.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "pret"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Text
Dim f, Rng, TblBD(), NbCol, NomTableau, TblBD2(), NbCol2, NomTableau2, g, NomTableau4, TabBD(), ColCombo4(), colVisu4(), colInterro(), NcolVisu4, NbCol4, NcolInt4, Choix()



Private Sub CommandButton1_Click()
  Unload pret
'MENU_GENERAL.Show


End Sub

Private Sub CommandButton10_Click()
If ComboBox7 = "" Then MsgBox "le nom du technicien n'est pas renseigné": ComboBox7.SetFocus: Exit Sub
retourpretobjet.ComboBox1 = textbox2
retourpretobjet.ComboBox2 = "En cours"
'retourpretobjet.Textbox13 = ComboBox7
retourpretobjet.Show
End Sub

Private Sub CommandButton11_Click()
raz
UserForm_Initialize

'  If n > 0 Then Me.ListBox1.Column = Tbl Else Me.ListBox1.Clear
'     If Me.ListBox1.ListIndex < Me.ListBox1.ListCount - 1 Then
'   Me.ListBox1.ListIndex = Me.ListBox1.ListIndex + 1
'  End If
  
End Sub

Private Sub CommandButton12_Click()
retourpretobjet.ComboBox1 = textbox2
retourpretobjet.ComboBox2 = "En cours"
retourpretobjet.TextBox13 = ComboBox7
End Sub

Private Sub CommandButton13_Click()
If ComboBox7 = "" Then MsgBox "le nom du technicien n'est pas renseigné": ComboBox7.SetFocus: Exit Sub
modifpret.ComboBox1 = textbox2
modifpret.ComboBox2 = "En cours"
modifpret.TextBox13 = ComboBox7
modifpret.Show
End Sub

Private Sub CommandButton14_Click()
modifpret.ComboBox1 = textbox2
modifpret.ComboBox2 = "En cours"
modifpret.TextBox13 = ComboBox7
End Sub

Private Sub CommandButton15_Click()

  Set f2 = Sheets("résultat")
  f2.Cells.ClearContents
  a = Me.ListBox10.List
  f2.[a2].Resize(UBound(a) + 1, UBound(a, 2) + 1) = a
  c = 0
  For c = 1 To NbCol4
     f2.Cells(1, c) = Range(NomTableau4).Offset(-1).Item(1, c)
  Next
  f2.Cells.EntireColumn.AutoFit

End Sub

Private Sub CommandButton16_Click()

Dim ret1 As Integer
Dim URLto
Set g = Sheets("résultat")
URLto = "mailto:" & "mdelooze@esad-orleans.fr" _
& ";" _
& "mdelooze@esad-orleans.fr" _
& " ?Subject=" & "pret n°" & Enreg & " _ " & textbox3 & " _ " & textbox2 _
& "&Body=" _
& "Créé par " _
& TextBox8 _
& Chr(13) & Chr(10) _
& "_ le " & textbox4 _
& " _ " & TextBox9 & " / " & TextBox10 & " / " & TextBox11 & " / " _
& Chr(13) & Chr(10) _
& ". Commentaire = " _
& g.Range("A2:M" & g.[A65000].End(xlUp).Row)

ActiveWorkbook.FollowHyperlink Address:=URLto
ret1 = MsgBox("Avez-vous validé l'envoi du BT par mail ?", vbYesNo)
If ret1 = vbNo Then
    Exit Sub
End If

'Sheets("Feuil1").Range(Cells(1, 1), Cells(ListBox1.ListCount, 1)) = ListBox1.List


'
'Dim objOutlook As Outlook.Application
'Dim objOutlookMsg As Outlook.MailItem
'
'Set objOutlookMsg = objOutlook.createitem(olMailItem)
'
'With objOutlookMsg
'       .to = "mdelooze@esad-orleans.fr"
'       .Subject = TextBox2.Value
'       .CC = TextBox4.Value
'       .Attachments.Add (fic)
'       .body = NbCol4 = Range(NomTableau4).Columns.Count
'       .Send
'
'End With
'Set objOutlookMsg = Nothing
'Set g = Sheets("résultat")
'
'
'Dim MonOutlook As Object
'  Dim MonMessage As Object
'  Dim body As String
'  Set g = Sheets("résultat")
' g.Range("A2:M" & g.[A65000].End(xlUp).Row).Copy
'
'  Set MonOutlook = CreateObject("Outlook.Application")
'  Set MonMessage = MonOutlook.createitem(0)
'  MonMessage.to = "mdelooze@esad-orleans.fr"
'  MonMessage.Subject = "Code magasin"
'    body = "Bonjour,"
'    body = body & Chr(13) & Chr(10)
'    body = "Veuillez trouver ci-dessous les codes magasins que vous recherchez" & Chr(13) & Chr(10)
'    body = Selection.Paste
'    body = body & Chr(13) & Chr(10)
'
'  MonMessage.body = body
'  MonMessage.Display
'  Set MonOutlook = Nothing

End Sub

Private Sub CommandButton17_Click()
emprunteur.Show
End Sub

Private Sub CommandButton3_Click()
If ComboBox7 = "" Then MsgBox "le nom du technicien n'est pas renseigné": ComboBox7.SetFocus: Exit Sub
If TextBox6 = "" Then MsgBox "l'adresse mail n'est pas renseigné, merci de la renseigner dans emprunteur ": TextBox6.SetFocus: Exit Sub

pretobjet.B_ajout = True


pretobjet.textbox3 = textbox2.Value
pretobjet.TextBox9 = ComboBox7.Value
pretobjet.TextBox15 = TextBox6.Value
pretobjet.Show
End Sub

'Private Sub CommandButton4_Click()
'ComboBox7.Value = "DELOOZE_MICHAEL"
'End Sub

Private Sub CommandButton4_Click()
ComboBox7.Value = "LIMMELETTE_FLORIAN"
End Sub

Private Sub CommandButton5_Click()
ComboBox7.Value = "POLVECHE_THEO"
End Sub

Private Sub CommandButton6_Click()
ComboBox7.Value = "DURIEUX_NOE"
End Sub

Private Sub CommandButton7_Click()
ComboBox7.Value = "PARROD_STEPHANE"
End Sub

Private Sub CommandButton8_Click()
ComboBox7.Value = "JUGI_DAVID"
End Sub

Private Sub CommandButton9_Click()
pretobjet.B_ajout = True

pretobjet.textbox3 = textbox2.Value
pretobjet.TextBox9 = ComboBox7.Value
End Sub

Private Sub TextBox40_Change()

End Sub

Private Sub Frame1_Click()

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

'Dim tablo3 As Variant 'tech
'With Feuil8 '
'derligne = .Range("a65536").End(xlUp).Row
'tablo3 = .Range("a2:a" & derligne)
'ComboBox7.List = tablo3
'End With
'Me.ListBox10.Visible = False
'Me.CommandButton10.Visible = False
   Set g = Sheets("prets")
Set Rng = g.Range("A2:M" & g.[A65000].End(xlUp).Row)     ' à adapter
 NomTableau4 = "Tableau10"
' ActiveWorkbook.Names.Add Name:=NomTableau4                                    ' A adapter
'ActiveWorkbook.Names.Add Name:=NomTableau4, RefersTo:=Rng                                     ' A adapter
 NbCol4 = Range(NomTableau4).Columns.Count
 
 '---- A adapter
 TabBD = Range(NomTableau4).Resize(, NbCol4 + 1).Value              ' Array: + rapide
 For ii = 1 To UBound(TabBD): TabBD(ii, NbCol4 + 1) = ii: Next ii      ' No enregistrement
 ColCombo4 = Array(3, 14)                                   ' A adapter (1 à 6 colonnes maxi)
 colVisu4 = Array(4, 5, 14)      ' Colonnes ListBox (à adapter)
 colInterro = Array(4, 5, 14)   ' colonnes à interroger (adapter)
 '----
 NcolInt4 = UBound(colInterro) + 1
 Me.ListBox10.List = TabBD
 
 For ii = UBound(ColCombo4) + 1 To 5
   Me("combobox" & ii + 1).Visible = False: Me("labelCbx" & ii + 1).Visible = False
 Next ii
 For c = 1 To UBound(ColCombo4) + 1: Me("combobox" & c) = "*": Next c
 For c = 1 To UBound(ColCombo4) + 1: ListeCol c: Next c
' For ii = 1 To UBound(ColCombo4) + 1:  Me("labelCbx" & ii) = Range(NomTableau4).Offset(-1).Item(1, ColCombo4(ii - 1)): Next ii
' Me.ListBox10.ColumnCount = NbCol4 + 1
  Me.ListBox10.ColumnCount = NbCol4
' Me.ListBox10.ColumnWidths = "0;150;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0"
 '-- en têtes de colonnes ListBox
 EnteteListBox           ' Supprimer sur Excel 2013
 '-- labels textbox
' LabelsTextBox
' For ii = NbCol4 + 40 To 50: Me("textbox" & ii).Visible = False: Next ii
' For ii = NbCol4 + 40 To 50: Me("label" & ii).Visible = False: Next ii
 '-- colTri
 Me.ComboTri10.List = Application.Transpose(Range(NomTableau4).Offset(-1).Resize(1))  ' Ordre tri
 Affiche
'   Me.ComboBox2 = "en cours"
Me.ListBox10.Clear

Me.ComboChoixColFiltre = "Emprunteur (NOM_PRENOM)"
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


 If Me.ListBox1.ListIndex < Me.ListBox1.ListCount - 1 Then
   Me.ListBox1.ListIndex = Me.ListBox1.ListIndex + 1
  End If

End Sub
Private Sub B_SupFilte_Click()
   Me.ListBox1.List = TblBD
End Sub

Private Sub ListBox1_Click()
On Error Resume Next
  For i = 1 To NbCol
    Me("textbox" & i) = Me.ListBox1.Column(i - 1)
  Next i
  Me.Enreg = Me.ListBox1.Column(i - 1)
  Me.ComboBox1 = Me.textbox2

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
     Me.ComboBox2 = "En cours"
'     Me.CommandButton10.Visible = True
'     Me.ListBox10.Visible = True

  Else
     Me.ListBox10.Clear
'     Me.CommandButton10.Visible = False
'     Me.ListBox10.Visible = False
  End If

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
'Private Sub ComboBox1_Change()
'  Affiche
'End Sub
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
   tempcol = tempcol & "10"
   On Error Resume Next
   Me.ListBox10.ColumnWidths = tempcol
   On Error GoTo 0
 End Sub
Sub LabelsTextBox()
   For c = 1 To NbCol4
      Me("textbox" & c).Width = Range(NomTableau4).Columns(c).Width * 1.3
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
Private Sub ListBox10_Click()
  For i = 1 To NbCol4
    tmp = Me.ListBox10.Column(i - 1)
'    If Not IsError(tmp) Then Me("textbox" & i) = tmp
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

Sub Gchoix()
 On Error Resume Next
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
'  Me.TextBoxRech10 = ""
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


