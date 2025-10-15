Attribute VB_Name = "Module1"
Option Explicit

' =====================================================
' MODULE PRINCIPAL - LANCEMENT APPLICATION V5
' =====================================================

' Variables globales session
Public g_AppVersion As String
Public g_CurrentUser As String
Public g_CurrentTech As String
Public g_SessionStart As Date

Sub Auto_Open()
    On Error GoTo ErrorHandler
    
    ' Initialisation
    g_AppVersion = "5.0"
    g_SessionStart = Now
    
    ' Vérifier environnement
    If Not VerifierEnvironnement() Then
        MsgBox "Erreur: Environnement non valide. L'application ne peut pas démarrer.", vbCritical, "Erreur Critique"
        Exit Sub
    End If
    
    ' Lancer menu principal
    MainMenu.Show
    
    Exit Sub

ErrorHandler:
    M_Core.LogError "Auto_Open", Err.Description
    MsgBox "Erreur au lancement: " & Err.Description & vbCrLf & vbCrLf & _
           "Contactez le régisseur général.", vbCritical, "Erreur Critique"
End Sub

Function VerifierEnvironnement() As Boolean
    Dim requiredSheets As Variant
    Dim requiredTables As Variant
    Dim ws As Worksheet
    Dim missing As String
    Dim i As Long
    
    VerifierEnvironnement = False
    
    ' Feuilles obligatoires
    requiredSheets = Array("accueil", "emprunteurs", "prets", "articles", "service", "fonction", "tech", "résultat")
    
    For i = 0 To UBound(requiredSheets)
        If Not WorksheetExists(CStr(requiredSheets(i))) Then
            missing = missing & vbCrLf & "- " & requiredSheets(i)
        End If
    Next i
    
    If missing <> "" Then
        MsgBox "Feuilles Excel manquantes:" & missing, vbCritical
        Exit Function
    End If
    
    ' Tables nommées obligatoires
    requiredTables = Array("Tableau1", "Tableau10", "Tableau4")
    missing = ""
    
    For i = 0 To UBound(requiredTables)
        If Not M_Core.TableExists(CStr(requiredTables(i))) Then
            missing = missing & vbCrLf & "- " & requiredTables(i)
        End If
    Next i
    
    If missing <> "" Then
        MsgBox "Tables nommées manquantes:" & missing, vbCritical
        Exit Function
    End If
    
    ' Tout OK
    VerifierEnvironnement = True
End Function

Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = (Not ws Is Nothing)
    On Error GoTo 0
End Function
