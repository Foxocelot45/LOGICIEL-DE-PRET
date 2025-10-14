VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendrier 
   Caption         =   "Calendrier"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2775
   OleObjectBlob   =   "Calendrier.frx":0000
End
Attribute VB_Name = "Calendrier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
Option Base 1
#If VBA7 Then
    Dim hWnd As LongPtr, Style As LongPtr
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
 #Else
    Dim hWnd As Long, Style As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If


Dim Charge As Boolean
Dim OldAn As Integer, OldMois As Integer, Decaler As Integer
Dim OldDate As String
Dim EvitSub As Boolean
Public Function Chargement(Optional Mydate As String = "", Optional Pose As String = "0;0") As String

Dim t

OldDate = Mydate
t = Split(Pose, ";")
Me.Top = t(0): Me.Left = t(1)
hWnd = FindWindow(vbNullString, Me.Caption)
 Style = GetWindowLong(hWnd, -16) And Not &HC00000
 SetWindowLong hWnd, -16, Style
 DrawMenuBar hWnd
 Me.Height = Me.Height - 17
If Mydate <> "" And Mydate <> "?" Then Me.Tag = Mydate Else Me.Tag = Date

EvitSub = True
    CBox_Mois.ListIndex = Mid$(Me.Tag, 4, 2) - 1: OldMois = CBox_Mois.ListIndex
    CBox_An.ListIndex = Right$(Me.Tag, 4) - 1950: OldAn = CBox_An.ListIndex
EvitSub = False

MajControle
Me.Show vbModal
On Error Resume Next
Chargement = Me.Tag
Unload Me

End Function
Sub MajControle()

Dim laDate As Date
Dim j As Integer
Dim m As Integer
Dim trouve As Boolean
Dim i  As Integer

Charge = False
laDate = CDate("01/" & Format(Me.Tag, "mm/yyyy"))
j = Weekday(laDate)

For i = 1 To 42
    m = i Mod 7
    Me.Controls("D" & i).Caption = ""
    Me.Controls("D" & i).Tag = ""
    Me.Controls("D" & i).SpecialEffect = fmSpecialEffectFlat
    Me.Controls("D" & i).ForeColor = &H800000
    Me.Controls("D" & i).BorderStyle = fmBorderStyleNone
    Me.Controls("D" & i).BackColor = &HFFFFFF
    
    If j = m + 1 And Not trouve Then
        trouve = True
        Me.Controls("D" & i).Enabled = True
        Me.Controls("D" & i).Caption = Format(laDate, "dd")
        Me.Controls("D" & i).Tag = laDate
    Else
        If i > 1 Then
            If Me.Controls("D" & i - 1).Tag = "" Then
                Me.Controls("D" & i).Enabled = False
            Else
                Me.Controls("D" & i).Caption = Format(CDate(Me.Controls("D" & i - 1).Tag) + 1, "dd")
                Me.Controls("D" & i).Tag = CDate(Me.Controls("D" & i - 1).Tag) + 1
                Me.Controls("D" & i).Enabled = True
            End If
        Else
            Me.Controls("D" & i).Enabled = False
        End If
    End If

    If Me.Controls("D" & i).Tag <> "" Then
        If Month(CDate(Me.Controls("D" & i).Tag)) <> Month(Me.Tag) Then
            Me.Controls("D" & i).Caption = ""
            Me.Controls("D" & i).Tag = ""
            Me.Controls("D" & i).Enabled = False
        End If
    End If


    If Me.Controls("D" & i).Tag <> "" Then
        If OldDate <> "" Then
            If CDate(Me.Controls("D" & i).Tag) = CDate(OldDate) Then
                'Me.Controls("D" & I).SpecialEffect = fmSpecialEffectSunken
                'Me.Controls("D" & I).BackColor = &H8000000F
                Me.Controls("D" & i).BackColor = &H8000000A
           Else
                Me.Controls("D" & i).SpecialEffect = fmSpecialEffectFlat
           End If
        End If
    End If
    
    If Me.Controls("D" & i).Tag <> "" Then
        If CDate(Me.Controls("D" & i).Tag) = Date Then
            Me.Controls("D" & i).BorderStyle = fmBorderStyleSingle
            Me.Controls("D" & i).BorderColor = &HFF&
        End If
    End If
    
Next
Charge = True

End Sub
Private Sub Cmd_CeJour_Click()

Me.Tag = Date: Me.Hide

End Sub
Private Sub Cmd_Echap_Click()

Me.Tag = OldDate: Me.Hide

End Sub
Private Sub Cmd_NonDate_Click()

Me.Tag = "?": Me.Hide

End Sub
Private Sub Cmd_Suppr_Click()

Me.Tag = "": Me.Hide

End Sub

Private Sub Label1_Click()
Me.Tag = Date: Me.Hide
End Sub

Private Sub UserForm_Initialize()

Dim i

CBox_Mois.AddItem "Janvier"
CBox_Mois.AddItem "Février"
CBox_Mois.AddItem "Mars"
CBox_Mois.AddItem "Avril"
CBox_Mois.AddItem "Mai"
CBox_Mois.AddItem "Juin"
CBox_Mois.AddItem "Juillet"
CBox_Mois.AddItem "Août"
CBox_Mois.AddItem "Septembre"
CBox_Mois.AddItem "Octobre"
CBox_Mois.AddItem "Novembre"
CBox_Mois.AddItem "Décembre"
For i = 1950 To 2050: CBox_An.AddItem i: Next
Me.Cmd_CeJour.Caption = "Today: " & Date
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

If CloseMode = vbFormControlMenu Then
    MsgBox "Vous ne pouvez pas utiliser ce bouton de fermeture." & Chr(13) & Chr(13) & "Veuillez clicker sur la commande ou sur la touche Echap"
    Cancel = True
End If

End Sub
Private Sub CBox_An_Change()

If EvitSub Then Exit Sub

Decaler = CBox_An.ListIndex - OldAn
OldAn = CBox_An.ListIndex
ModifierDate Decaler * 12

End Sub
Private Sub CBox_Mois_Change()

If EvitSub Then Exit Sub

Decaler = CBox_Mois.ListIndex - OldMois
OldMois = CBox_Mois.ListIndex
ModifierDate Decaler

End Sub
Sub ModifierDate(i As Integer)

Dim j As Byte: Dim m As Byte: Dim y As Integer

 j = Day(Me.Tag): m = Month(Me.Tag): y = Year(Me.Tag)
If i > 11 Or i < -11 Then y = y + i / 12 Else m = m + i
If Charge Then Me.Tag = j & "/" & m & "/" & y
Do Until IsDate(Me.Tag)
    j = j - 1
    If Charge Then Me.Tag = j & "/" & m & "/" & y
Loop
MajControle

End Sub
Private Sub D1_Click()

If Charge Then Me.Tag = D1.Tag: Me.Hide

End Sub
Private Sub D2_Click()

If Charge Then Me.Tag = D2.Tag: Me.Hide

End Sub
Private Sub D3_Click()

If Charge Then Me.Tag = D3.Tag: Me.Hide

End Sub
Private Sub D4_Click()

If Charge Then Me.Tag = D4.Tag: Me.Hide

End Sub
Private Sub D5_Click()

If Charge Then Me.Tag = D5.Tag: Me.Hide

End Sub
Private Sub d6_Click()

If Charge Then Me.Tag = D6.Tag: Me.Hide

End Sub
Private Sub D7_Click()

If Charge Then Me.Tag = D7.Tag: Me.Hide

End Sub
Private Sub D8_Click()

If Charge Then Me.Tag = D8.Tag: Me.Hide

End Sub
Private Sub D9_Click()

If Charge Then Me.Tag = D9.Tag: Me.Hide

End Sub
Private Sub D10_Click()

If Charge Then Me.Tag = D10.Tag: Me.Hide

End Sub
Private Sub D11_Click()

If Charge Then Me.Tag = D11.Tag: Me.Hide

End Sub
Private Sub D12_Click()

If Charge Then Me.Tag = D12.Tag: Me.Hide

End Sub
Private Sub D13_Click()

If Charge Then Me.Tag = D13.Tag: Me.Hide

End Sub
Private Sub D14_Click()

If Charge Then Me.Tag = D14.Tag: Me.Hide

End Sub
Private Sub D15_Click()

If Charge Then Me.Tag = D15.Tag: Me.Hide

End Sub
Private Sub D16_Click()

If Charge Then Me.Tag = D16.Tag: Me.Hide

End Sub
Private Sub D17_Click()

If Charge Then Me.Tag = D17.Tag: Me.Hide

End Sub
Private Sub D18_Click()

If Charge Then Me.Tag = D18.Tag: Me.Hide

End Sub
Private Sub D19_Click()

If Charge Then Me.Tag = D19.Tag: Me.Hide

End Sub
Private Sub D20_Click()

If Charge Then Me.Tag = D20.Tag: Me.Hide

End Sub
Private Sub D21_Click()

If Charge Then Me.Tag = D21.Tag: Me.Hide

End Sub
Private Sub D22_Click()

If Charge Then Me.Tag = D22.Tag: Me.Hide

End Sub
Private Sub D23_Click()

If Charge Then Me.Tag = D23.Tag: Me.Hide

End Sub
Private Sub D24_Click()

If Charge Then Me.Tag = D24.Tag: Me.Hide

End Sub
Private Sub D25_Click()

If Charge Then Me.Tag = D25.Tag: Me.Hide

End Sub
Private Sub D26_Click()

If Charge Then Me.Tag = D26.Tag: Me.Hide

End Sub
Private Sub D27_Click()

If Charge Then Me.Tag = D27.Tag: Me.Hide

End Sub
Private Sub D28_Click()

If Charge Then Me.Tag = D28.Tag: Me.Hide

End Sub
Private Sub D29_Click()

If Charge Then Me.Tag = D29.Tag: Me.Hide

End Sub
Private Sub D30_Click()

If Charge Then Me.Tag = D30.Tag: Me.Hide

End Sub
Private Sub D31_Click()

If Charge Then Me.Tag = D31.Tag: Me.Hide

End Sub
Private Sub D32_Click()

If Charge Then Me.Tag = D32.Tag: Me.Hide

End Sub
Private Sub D33_Click()

If Charge Then Me.Tag = D33.Tag: Me.Hide

End Sub
Private Sub D34_Click()

If Charge Then Me.Tag = D34.Tag: Me.Hide

End Sub
Private Sub D35_Click()

If Charge Then Me.Tag = D35.Tag: Me.Hide

End Sub
Private Sub D36_Click()

If Charge Then Me.Tag = D36.Tag: Me.Hide

End Sub
Private Sub D37_Click()

If Charge Then Me.Tag = D37.Tag: Me.Hide

End Sub
Private Sub D38_Click()

If Charge Then Me.Tag = D38.Tag: Me.Hide

End Sub
Private Sub D39_Click()

If Charge Then Me.Tag = D39.Tag: Me.Hide

End Sub
Private Sub D40_Click()

If Charge Then Me.Tag = D40.Tag: Me.Hide

End Sub
Private Sub D41_Click()

If Charge Then Me.Tag = D41.Tag: Me.Hide

End Sub
Private Sub D42_Click()

If Charge Then Me.Tag = D42.Tag: Me.Hide

End Sub

