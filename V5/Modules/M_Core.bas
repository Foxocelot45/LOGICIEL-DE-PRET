Attribute VB_Name = "M_Core"
Option Explicit
Option Compare Text

' =====================================================
' MODULE CORE - FONCTIONS COMMUNES V5
' =====================================================

' Constantes application
Public Const APP_NAME As String = "Gestion Prêts ESAD v5"
Public Const EMAIL_REGIE As String = "gestionstockregie@esad-orleans.fr"

' Palette couleurs moderne
Public Const COLOR_PRIMARY As Long = 2854940    ' RGB(41, 128, 185)  - Bleu
Public Const COLOR_SUCCESS As Long = 5287936    ' RGB(39, 174, 96)   - Vert
Public Const COLOR_WARNING As Long = 42495      ' RGB(243, 156, 18)  - Orange
Public Const COLOR_DANGER As Long = 3684408     ' RGB(192, 57, 43)   - Rouge
Public Const COLOR_DARK As Long = 5855577       ' RGB(52, 73, 94)    - Gris foncé
Public Const COLOR_LIGHT As Long = 15658734     ' RGB(236, 240, 241) - Gris clair

' =====================================================
' CHARGEMENT DONNÉES
' =====================================================

Public Function LoadDataToArray(tableName As String) As Variant
    ' Charge une table Excel dans un array VBA
    ' Ajoute colonne index ligne pour retrouver position Excel
    
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Dim arr As Variant
    Dim nbRows As Long, nbCols As Long
    Dim i As Long
    
    Set rng = Range(tableName)
    nbRows = rng.Rows.Count
    nbCols = rng.Columns.Count
    
    ' Charger dans array avec colonne index supplémentaire
    ReDim arr(1 To nbRows, 1 To nbCols + 1)
    
    Dim srcArr As Variant
    srcArr = rng.Value
    
    For i = 1 To nbRows
        Dim j As Long
        For j = 1 To nbCols
            arr(i, j) = srcArr(i, j)
        Next j
        arr(i, nbCols + 1) = i ' Index ligne Excel
    Next i
    
    LoadDataToArray = arr
    Exit Function
    
ErrorHandler:
    LogError "LoadDataToArray(" & tableName & ")", Err.Description
    LoadDataToArray = Array()
End Function

Public Sub PopulateListBox(lst As MSForms.ListBox, data As Variant, Optional visibleCols As Variant)
    ' Remplit une ListBox avec un array
    ' visibleCols: array des indices colonnes à afficher (1-based)
    
    On Error GoTo ErrorHandler
    
    If Not IsArray(data) Then Exit Sub
    If UBound(data, 1) < 1 Then Exit Sub
    
    lst.Clear
    lst.List = data
    lst.ColumnCount = UBound(data, 2)
    
    ' Masquer colonnes si spécifié
    If Not IsMissing(visibleCols) Then
        Dim widths As String
        Dim i As Long
        For i = 1 To UBound(data, 2)
            If IsInArray(i, visibleCols) Then
                widths = widths & "150;" ' Largeur visible
            Else
                widths = widths & "0;" ' Colonne cachée
            End If
        Next i
        lst.ColumnWidths = Left(widths, Len(widths) - 1)
    End If
    
    Exit Sub
    
ErrorHandler:
    LogError "PopulateListBox", Err.Description
End Sub

Public Function SearchArrayWildcard(arr As Variant, colIndex As Long, searchTerm As String) As Variant
    ' Recherche dans array avec wildcard (*)
    ' Retourne sous-array filtré
    
    On Error GoTo ErrorHandler
    
    If Not IsArray(arr) Then
        SearchArrayWildcard = Array()
        Exit Function
    End If
    
    Dim results() As Variant
    Dim resultCount As Long
    resultCount = 0
    
    ReDim results(1 To UBound(arr, 1), 1 To UBound(arr, 2))
    
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        If arr(i, colIndex) Like searchTerm Then
            resultCount = resultCount + 1
            Dim j As Long
            For j = 1 To UBound(arr, 2)
                results(resultCount, j) = arr(i, j)
            Next j
        End If
    Next i
    
    If resultCount = 0 Then
        SearchArrayWildcard = Array()
        Exit Function
    End If
    
    ' Redimensionner au nombre réel de résultats
    ReDim finalResults(1 To resultCount, 1 To UBound(arr, 2))
    For i = 1 To resultCount
        For j = 1 To UBound(arr, 2)
            finalResults(i, j) = results(i, j)
        Next j
    Next i
    
    SearchArrayWildcard = finalResults
    Exit Function
    
ErrorHandler:
    LogError "SearchArrayWildcard", Err.Description
    SearchArrayWildcard = Array()
End Function

' =====================================================
' VALIDATION
' =====================================================

Public Function ValidateRequiredFields(ParamArray fields() As Variant) As Boolean
    ' Valide que des champs ne sont pas vides
    ' Usage: ValidateRequiredFields(ctrl1, "Nom", ctrl2, "Email", ...)
    
    ValidateRequiredFields = True
    Dim missing As String
    Dim i As Long
    
    For i = LBound(fields) To UBound(fields) Step 2
        Dim ctrl As Object
        Set ctrl = fields(i)
        Dim fieldName As String
        fieldName = CStr(fields(i + 1))
        
        If Trim(ctrl.Value & "") = "" Then
            If missing <> "" Then missing = missing & ", "
            missing = missing & fieldName
            ValidateRequiredFields = False
        End If
    Next i
    
    If Not ValidateRequiredFields Then
        MsgBox "Champs obligatoires manquants:" & vbCrLf & missing, vbExclamation, APP_NAME
    End If
End Function

' =====================================================
' UTILITAIRES
' =====================================================

Public Function TableExists(tableName As String) As Boolean
    On Error Resume Next
    Dim testRange As Range
    Set testRange = Range(tableName)
    TableExists = (Not testRange Is Nothing)
    On Error GoTo 0
End Function

Public Sub LogError(source As String, description As String)
    ' Log erreurs dans feuille cachée ou Debug
    Debug.Print Now & " | ERROR | " & source & " | " & description
    
    ' Optionnel: écrire dans feuille logs
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("logs")
    If Not ws Is Nothing Then
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(lastRow, 1).Value = Now
        ws.Cells(lastRow, 2).Value = source
        ws.Cells(lastRow, 3).Value = description
    End If
    On Error GoTo 0
End Sub

Public Function IsInArray(value As Variant, arr As Variant) As Boolean
    ' Vérifie si valeur existe dans array
    IsInArray = False
    If Not IsArray(arr) Then Exit Function
    
    Dim elem As Variant
    For Each elem In arr
        If elem = value Then
            IsInArray = True
            Exit Function
        End If
    Next elem
End Function

Public Function GetTechniciens() As Variant
    ' Retourne array des techniciens
    GetTechniciens = Array("LIMMELETTE_FLORIAN", "POLVECHE_THEO", "DURIEUX_NOE", "PARROD_STEPHANE", "JUGI_DAVID")
End Function

Public Function GetRaisonsPret() As Variant
    ' Retourne array des raisons de prêt
    GetRaisonsPret = Array("DIPLOME", "BILAN", "EXPO", "PONCTUEL", "WORKSHOP", "PERMANENT")
End Function

Public Function GetQuantitesRapides() As Variant
    ' Retourne array des quantités rapides
    GetQuantitesRapides = Array(1, 5, 10, 15, 20)
End Function

' =====================================================
' THÈME VISUEL
' =====================================================

Public Sub ApplyModernTheme(frm As Object)
    ' Applique le thème moderne à un UserForm
    frm.BackColor = COLOR_LIGHT
End Sub
