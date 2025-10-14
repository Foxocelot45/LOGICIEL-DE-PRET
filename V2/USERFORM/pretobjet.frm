VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} pretobjet 
   Caption         =   "sorti d'objet"
   ClientHeight    =   11265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11610
   OleObjectBlob   =   "pretobjet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "pretobjet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Text
Dim f, Rng, TblBD(), NbCol, NomTableau

Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()

'MENU_GENERAL.Show
pret.ListBox10.Clear
'pret.raz
pret.CommandButton11 = True
'pret.TextBoxRech = TextBox3
'  If n > 0 Then pret.ListBox1.Column = Tbl Else pret.ListBox1.Clear
'     If pret.ListBox1.ListIndex < pret.ListBox1.ListCount - 1 Then
'   pret.ListBox1.ListIndex = pret.ListBox1.ListIndex + 1
'  End If
Unload pretobjet
End Sub

Private Sub CommandButton10_Click()
textbox2.Value = "DIPLOME"

End Sub

Private Sub CommandButton11_Click()
textbox2.Value = "EXPO"
End Sub

Private Sub CommandButton12_Click()
textbox2.Value = "PONCTUEL"
End Sub

Private Sub CommandButton13_Click()
textbox2.Value = "WORKSHOP"
End Sub

Private Sub CommandButton14_Click()
textbox2.Value = "PERMANENT"
End Sub

Private Sub CommandButton15_Click()

'Dim ret1 As Integer
'Dim URLto
'URLto = "mailto:" & "mdelooze@esad-orleans.fr" _
'& ";" _
'& "mdelooze@esad-orleans.fr" _
'& " ?Subject=" & "pret n°" & Enreg & " _ " & TextBox3 & " _ " & TextBox2 _
'& "&Body=" _
'& "Créé par " _
'& TextBox9 _
'& vbCrLf _
'& "_ le " & TextBox4 _
'& Chr(13) & Chr(10) _
'& ". Objet = " & TextBox5 _
'& ". Quantité = " & TextBox6 _
'& ". retour prévu le : = " & TextBox8
'
''& ListBox1.List

Set olApp = CreateObject("Outlook.application")
        Set Mail = olApp.CreateItem(olMailItem)
        With Mail
            .display
            .To = TextBox15
            .Subject = "PRET_n°" & Enreg & " _ " & textbox3
            .HTMLBody = "Un prêt a été crée par " & TextBox9 & " le " & textbox4 & "<br>" & _
                        "<br>" & _
                        "La raison du prêt est : " & textbox2 & "<br>" & _
                        "L'objet emprunté est: " & TextBox5 & "<br>" & _
                        "La quantité empruntée est de : " & TextBox6 & "<br>" & _
                        "La date de retour prévue est: " & TextBox8 & "<br>" & _
                        "<br>" & _
                        "Pour toute demande, veuillez contacter la regie à l'adresse suivante  : " & "gestionstockregie@esad-orleans.fr" & "<br>" & _
                        .HTMLBody
            .display
            .send

        End With

'ret1 = MsgBox("Avez-vous validé l'envoi du BT par mail ?", vbYesNo)
'
' If ret1 = vbYes Then

'                        Label9 & " : " & TextBox9 & "<br>" & _
'                        Label10 & " : " & TextBox10 & "<br>" & _
'                        Label11 & " : " & TextBox11 & "<br>" & _
'                        Label19 & " : " & ComboBox1 & "<br>" & _


'ActiveWorkbook.FollowHyperlink Address:=URLto


ret1 = MsgBox("Avez-vous validé l'envoi du BT par mail ?", vbYesNo)
If ret1 = vbNo Then
    Exit Sub
End If
'End If
End Sub

Private Sub CommandButton2_Click()
textbox4.Value = Now
End Sub

Private Sub CommandButton3_Click()
Dim Pose As String
Pose = Me.Top + TextBox8.Top + TextBox8.Height + 26
Pose = Pose & ";" & Me.Left + TextBox7.Left + 6
Pose = Me.Top + Me.TextBox8.Top + (Me.TextBox8.Height * 2)
Pose = Pose & ";" & Me.Left + Me.TextBox8.Left
TextBox8.Value = Calendrier.Chargement(TextBox8.Value, Pose)
End Sub

Private Sub CommandButton4_Click()

qrcode.ComboChoixColFiltre = "QRCode"
qrcode.Show

End Sub

Private Sub CommandButton5_Click()
TextBox6.Value = "5"

End Sub

Private Sub CommandButton6_Click()
TextBox6.Value = "10"

End Sub

Private Sub CommandButton7_Click()
TextBox6.Value = "15"

End Sub

Private Sub CommandButton8_Click()
TextBox6.Value = "20"


End Sub

Private Sub CommandButton9_Click()
textbox2.Value = "BILAN"
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub TextBox3_Change()

End Sub

Private Sub TextBox5_Change()

End Sub

Private Sub UserForm_Initialize()
 On Error Resume Next
   NomTableau = "Tableau10"                                   ' à Adapter
   NbCol = Range(NomTableau).Columns.Count
   TblBD = Range(NomTableau).Resize(, NbCol + 1).Value       ' Array: + rapide
   For i = 1 To UBound(TblBD): TblBD(i, NbCol + 1) = i: Next i  ' No enregistrement
   Me.ListBox1.List = TblBD
   Me.ListBox1.ColumnCount = NbCol + 1
   Me.ListBox1.ColumnWidths = "0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0;0"
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
   
   TextBox6.Value = "1"
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
If TextBox15 = "" Then MsgBox "Email non renseigné": TextBox15.SetFocus: Exit Sub
If textbox3 = "" Then MsgBox "Emprunteur non renseigné": textbox3.SetFocus: Exit Sub
If textbox4 = "" Then MsgBox "Date&heure non renseigné": textbox4.SetFocus: Exit Sub
If TextBox8 = "" Then MsgBox "Date de retour prévu non renseignée": TextBox8.SetFocus: Exit Sub
If textbox2 = "" Then MsgBox "Raison du prêt non renseignée": textbox2.SetFocus: Exit Sub
If TextBox5 = "" Then MsgBox "Objet non renseigné": TextBox5.SetFocus: Exit Sub
If TextBox6 = "" Then MsgBox "Quantité non renseignée": TextBox6.SetFocus: Exit Sub
  TextBox14.Value = "En cours"
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
  
Set olApp = CreateObject("Outlook.application")
        Set Mail = olApp.CreateItem(olMailItem)
        With Mail
            .display
            .To = TextBox15
            .Subject = "PRET_n°" & Enreg & " _ " & textbox3
            .HTMLBody = "Un prêt a été créé par " & TextBox9 & " le " & textbox4 & "<br>" & _
                        "<br>" & _
                        "La raison du prêt est : " & textbox2 & "<br>" & _
                        "L'objet emprunté est : " & TextBox5 & "<br>" & _
                        "La quantité empruntée est de : " & TextBox6 & "<br>" & _
                        "La date de retour prévue est : " & TextBox8 & "<br>" & _
                        "<br>" & _
                        "Pour toute demande, veuillez contacter la regie à l'adresse suivante  : " & "gestionstockregie@esad-orleans.fr" & "<br>" & _
                        .HTMLBody
            .display
            .send

        End With
ret1 = MsgBox("Avez-vous validé l'envoi du BT par mail ?", vbYesNo)
If ret1 = vbNo Then
    Exit Sub
End If


'  UserForm_Initialize
  'raz

pret.CommandButton9 = True

TextBox5.Value = ""
TextBox6.Value = "1"
TextBox7.Value = ""
End Sub
Sub raz()
    For k = 1 To NbCol
      Me("textBox" & k) = ""
    Next k
    Me.TextBox1.SetFocus
End Sub
Private Sub B_ajout_Click()
' raz
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




