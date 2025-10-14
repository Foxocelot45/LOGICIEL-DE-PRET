Attribute VB_Name = "M_Business"
Option Explicit
Option Compare Text

' =====================================================
' MODULE BUSINESS - LOGIQUE MÉTIER V5
' =====================================================

' Variables module pour retours groupés
Public RetourCount As Long
Public RetourStartTime As Double

' =====================================================
' EXPORT INVENTAIRE
' =====================================================

Public Sub ExportInventaireComplet()
    On Error GoTo ErrorHandler
    
    Dim wsExport As Worksheet
    Dim wsArticles As Worksheet
    Dim wsPrets As Worksheet
    Dim lastRow As Long, i As Long, exportRow As Long
    Dim articleID As String
    Dim statut As String
    Dim emprunteur As String
    Dim datePret As String
    Dim pretDict As Object
    
    Application.ScreenUpdating = False
    
    ' Créer dictionnaire des prêts actifs
    Set pretDict = CreateObject("Scripting.Dictionary")
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    Set wsArticles = ThisWorkbook.Worksheets("articles")
    
    ' Scanner prêts en cours (colonne 15 vide)
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 15).Value = "" Then
            articleID = CStr(wsPrets.Cells(i, 5).Value)
            If Not pretDict.Exists(articleID) Then
                pretDict(articleID) = Array( _
                    wsPrets.Cells(i, 3).Value, _
                    wsPrets.Cells(i, 4).Value _
                )
            End If
        End If
    Next i
    
    ' Créer/vider feuille export
    On Error Resume Next
    Set wsExport = ThisWorkbook.Worksheets("Inventaire_" & Format(Now, "YYYYMMDD_HHMMSS"))
    If Not wsExport Is Nothing Then wsExport.Delete
    On Error GoTo ErrorHandler
    
    Set wsExport = ThisWorkbook.Worksheets.Add
    wsExport.Name = "Inventaire_" & Format(Now, "YYYYMMDD_HHMMSS")
    
    ' En-têtes
    With wsExport
        .Cells(1, 1).Value = "ID"
        .Cells(1, 2).Value = "Article"
        .Cells(1, 3).Value = "QR Code"
        .Cells(1, 4).Value = "Famille"
        .Cells(1, 5).Value = "Emplacement"
        .Cells(1, 6).Value = "État"
        .Cells(1, 7).Value = "Statut Prêt"
        .Cells(1, 8).Value = "Emprunteur"
        .Cells(1, 9).Value = "Date Prêt"
        
        .Range("A1:I1").Font.Bold = True
        .Range("A1:I1").Interior.Color = M_Core.COLOR_DARK
        .Range("A1:I1").Font.Color = RGB(255, 255, 255)
    End With
    
    ' Remplir données
    exportRow = 2
    For i = 2 To wsArticles.Cells(wsArticles.Rows.Count, 1).End(xlUp).Row
        articleID = CStr(wsArticles.Cells(i, 1).Value)
        
        With wsExport
            .Cells(exportRow, 1).Value = wsArticles.Cells(i, 1).Value ' ID
            .Cells(exportRow, 2).Value = wsArticles.Cells(i, 2).Value ' Article
            .Cells(exportRow, 3).Value = wsArticles.Cells(i, 3).Value ' QR
            .Cells(exportRow, 4).Value = wsArticles.Cells(i, 4).Value ' Famille
            .Cells(exportRow, 5).Value = wsArticles.Cells(i, 5).Value ' Emplacement
            .Cells(exportRow, 6).Value = wsArticles.Cells(i, 6).Value ' État
            
            ' Statut prêt
            If pretDict.Exists(articleID) Then
                .Cells(exportRow, 7).Value = "EN PRÊT"
                .Cells(exportRow, 8).Value = pretDict(articleID)(0) ' Emprunteur
                .Cells(exportRow, 9).Value = pretDict(articleID)(1) ' Date
                ' Coloration orange
                .Range(.Cells(exportRow, 1), .Cells(exportRow, 9)).Interior.Color = M_Core.COLOR_WARNING
            Else
                .Cells(exportRow, 7).Value = "DISPONIBLE"
                ' Coloration verte
                .Range(.Cells(exportRow, 1), .Cells(exportRow, 9)).Interior.Color = M_Core.COLOR_SUCCESS
            End If
        End With
        
        exportRow = exportRow + 1
    Next i
    
    ' Finalisation
    wsExport.Columns.AutoFit
    wsExport.Range("A1:I1").AutoFilter
    wsExport.Activate
    
    Application.ScreenUpdating = True
    
    MsgBox "Export terminé!" & vbCrLf & vbCrLf & _
           "Feuille: " & wsExport.Name & vbCrLf & _
           "Articles: " & (exportRow - 2) & vbCrLf & _
           "En prêt: " & pretDict.Count, vbInformation, M_Core.APP_NAME
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    M_Core.LogError "ExportInventaireComplet", Err.Description
    MsgBox "Erreur lors de l'export: " & Err.Description, vbCritical
End Sub

' =====================================================
' RETOURS GROUPÉS
' =====================================================

Public Sub RetournerTousPretsEmprunteur(emprunteur As String, technicien As String)
    ' Méthode 1: Tout retourner (1 clic)
    
    On Error GoTo ErrorHandler
    
    If MsgBox("Retourner TOUS les prêts de " & emprunteur & " ?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    RetourStartTime = Timer
    RetourCount = 0
    
    Dim wsPrets As Worksheet
    Dim i As Long
    Dim dateRetour As Date
    
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    dateRetour = Now
    
    Application.ScreenUpdating = False
    
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 3).Value = emprunteur And wsPrets.Cells(i, 15).Value = "" Then
            wsPrets.Cells(i, 15).Value = dateRetour
            wsPrets.Cells(i, 13).Value = technicien
            wsPrets.Cells(i, 14).Value = "Terminé"
            RetourCount = RetourCount + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    If RetourCount > 0 Then
        Dim elapsedTime As Double
        elapsedTime = Timer - RetourStartTime
        
        ' Email
        SendBatchReturnEmail emprunteur, RetourCount, technicien
        
        MsgBox RetourCount & " prêts retournés en " & Format(elapsedTime, "0.0") & " sec!" & vbCrLf & _
               "Gain temps: ~95% vs méthode traditionnelle", vbInformation
    Else
        MsgBox "Aucun prêt en cours pour cet emprunteur.", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    M_Core.LogError "RetournerTousPretsEmprunteur", Err.Description
    MsgBox "Erreur: " & Err.Description, vbCritical
End Sub

Public Sub ValiderRetoursCoches(lstReturns As MSForms.ListBox, checkControls As Collection, technicien As String)
    ' Méthode 2: Retour par cochage
    
    On Error GoTo ErrorHandler
    
    RetourStartTime = Timer
    RetourCount = 0
    
    Dim wsPrets As Worksheet
    Dim dateRetour As Date
    Dim i As Long, rowNum As Long
    Dim chk As MSForms.CheckBox
    
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    dateRetour = Now
    
    Application.ScreenUpdating = False
    
    ' Parcourir checkboxes cochées
    For i = 0 To lstReturns.ListCount - 1
        Set chk = checkControls(i + 1)
        If chk.Value = True Then
            ' Récupérer ligne Excel depuis dernière colonne de ListBox
            rowNum = lstReturns.List(i, lstReturns.ColumnCount - 1)
            
            wsPrets.Cells(rowNum + 1, 15).Value = dateRetour
            wsPrets.Cells(rowNum + 1, 13).Value = technicien
            wsPrets.Cells(rowNum + 1, 14).Value = "Terminé"
            RetourCount = RetourCount + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    
    If RetourCount > 0 Then
        Dim elapsedTime As Double
        Dim traditionalTime As Double
        elapsedTime = Timer - RetourStartTime
        traditionalTime = RetourCount * 90 ' 90 sec par retour méthode traditionnelle
        
        Dim gain As Double
        gain = (1 - (elapsedTime / traditionalTime)) * 100
        
        MsgBox RetourCount & " articles retournés!" & vbCrLf & vbCrLf & _
               "Temps réel: " & Format(elapsedTime, "0") & " sec" & vbCrLf & _
               "Temps traditionnel: " & Format(traditionalTime, "0") & " sec" & vbCrLf & _
               "GAIN: -" & Format(gain, "0") & "%", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    M_Core.LogError "ValiderRetoursCoches", Err.Description
    MsgBox "Erreur: " & Err.Description, vbCritical
End Sub

Public Sub TraiterScanQR(qrCode As String, technicien As String)
    ' Méthode 3: Scan QR à la chaîne
    
    On Error GoTo ErrorHandler
    
    Dim wsPrets As Worksheet
    Dim wsArticles As Worksheet
    Dim i As Long
    Dim articleID As String
    Dim found As Boolean
    
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    Set wsArticles = ThisWorkbook.Worksheets("articles")
    
    ' Trouver article par QR
    found = False
    For i = 2 To wsArticles.Cells(wsArticles.Rows.Count, 1).End(xlUp).Row
        If wsArticles.Cells(i, 3).Value = qrCode Then
            articleID = wsArticles.Cells(i, 1).Value
            found = True
            Exit For
        End If
    Next i
    
    If Not found Then
        Beep
        Exit Sub
    End If
    
    ' Retourner prêt en cours de cet article
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 5).Value = articleID And wsPrets.Cells(i, 15).Value = "" Then
            wsPrets.Cells(i, 15).Value = Now
            wsPrets.Cells(i, 13).Value = technicien
            wsPrets.Cells(i, 14).Value = "Terminé"
            RetourCount = RetourCount + 1
            
            ' Double beep succès
            Beep
            Application.Wait Now + TimeValue("00:00:00.2")
            Beep
            Exit For
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    M_Core.LogError "TraiterScanQR", Err.Description
End Sub

' =====================================================
' EMAILS OUTLOOK
' =====================================================

Public Sub SendLoanEmail(loanData As Variant)
    ' Email création prêt
    
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim Mail As Object
    
    Set olApp = CreateObject("Outlook.Application")
    Set Mail = olApp.CreateItem(0)
    
    With Mail
        .Display
        .To = loanData("Email")
        .Subject = "PRÊT n°" & loanData("ID") & " - " & loanData("Emprunteur")
        .HTMLBody = "<html><body>" & _
                    "<h2>Nouveau Prêt ESAD</h2>" & _
                    "<p>Un prêt a été créé par <b>" & loanData("Technicien") & "</b> le " & loanData("Date") & "</p>" & _
                    "<table border='1' cellpadding='5'>" & _
                    "<tr><td><b>Raison</b></td><td>" & loanData("Raison") & "</td></tr>" & _
                    "<tr><td><b>Article</b></td><td>" & loanData("Article") & "</td></tr>" & _
                    "<tr><td><b>Quantité</b></td><td>" & loanData("Quantite") & "</td></tr>" & _
                    "<tr><td><b>Retour prévu</b></td><td>" & loanData("RetourPrevu") & "</td></tr>" & _
                    "</table>" & _
                    "<p>Contact régie: <a href='mailto:" & M_Core.EMAIL_REGIE & "'>" & M_Core.EMAIL_REGIE & "</a></p>" & _
                    "</body></html>"
        .Display
        '.Send
    End With
    
    Exit Sub
    
ErrorHandler:
    M_Core.LogError "SendLoanEmail", Err.Description
    MsgBox "Impossible d'envoyer l'email. Vérifiez qu'Outlook est ouvert.", vbExclamation
End Sub

Public Sub SendBatchReturnEmail(emprunteur As String, count As Long, technicien As String)
    ' Email retours groupés
    
    On Error GoTo ErrorHandler
    
    Dim olApp As Object
    Dim Mail As Object
    Dim emailAddr As String
    
    ' Récupérer email emprunteur
    emailAddr = GetEmailEmprunteur(emprunteur)
    If emailAddr = "" Then Exit Sub
    
    Set olApp = CreateObject("Outlook.Application")
    Set Mail = olApp.CreateItem(0)
    
    With Mail
        .Display
        .To = emailAddr
        .Subject = "RETOURS GROUPÉS - " & emprunteur
        .HTMLBody = "<html><body>" & _
                    "<h2>Retours Groupés ESAD</h2>" & _
                    "<p><b>" & count & " articles</b> ont été retournés le " & Format(Now, "DD/MM/YYYY à HH:MM") & "</p>" & _
                    "<p>Technicien: <b>" & technicien & "</b></p>" & _
                    "<p>Merci pour votre collaboration.</p>" & _
                    "<p>Contact régie: <a href='mailto:" & M_Core.EMAIL_REGIE & "'>" & M_Core.EMAIL_REGIE & "</a></p>" & _
                    "</body></html>"
        .Display
        '.Send
    End With
    
    Exit Sub
    
ErrorHandler:
    M_Core.LogError "SendBatchReturnEmail", Err.Description
End Sub

Private Function GetEmailEmprunteur(nom As String) As String
    ' Récupère email depuis table emprunteurs
    Dim wsEmp As Worksheet
    Dim i As Long
    
    Set wsEmp = ThisWorkbook.Worksheets("emprunteurs")
    For i = 2 To wsEmp.Cells(wsEmp.Rows.Count, 1).End(xlUp).Row
        If wsEmp.Cells(i, 2).Value = nom Then
            GetEmailEmprunteur = wsEmp.Cells(i, 6).Value
            Exit Function
        End If
    Next i
End Function

' =====================================================
' STATISTIQUES DASHBOARD
' =====================================================

Public Function GetDashboardStats() As Object
    ' Retourne dictionnaire avec statistiques
    
    On Error GoTo ErrorHandler
    
    Dim stats As Object
    Set stats = CreateObject("Scripting.Dictionary")
    
    Dim wsPrets As Worksheet
    Dim wsArticles As Worksheet
    Dim i As Long
    
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    Set wsArticles = ThisWorkbook.Worksheets("articles")
    
    ' Compter prêts en cours
    Dim preEnCours As Long
    preEnCours = 0
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 15).Value = "" Then
            preEnCours = preEnCours + 1
        End If
    Next i
    
    ' Stats globales
    stats("TotalArticles") = wsArticles.Cells(wsArticles.Rows.Count, 1).End(xlUp).Row - 1
    stats("PretsEnCours") = preEnCours
    stats("TauxUtilisation") = Format((preEnCours / stats("TotalArticles")) * 100, "0.0") & "%"
    stats("TotalPrets") = wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row - 1
    
    ' Alertes
    stats("PretsDepasses") = CountOverdueLoans(30)
    stats("PretsAvertissement") = CountOverdueLoans(15)
    
    Set GetDashboardStats = stats
    Exit Function
    
ErrorHandler:
    M_Core.LogError "GetDashboardStats", Err.Description
    Set GetDashboardStats = Nothing
End Function

Private Function CountOverdueLoans(daysThreshold As Long) As Long
    Dim wsPrets As Worksheet
    Dim i As Long, count As Long
    Dim datePret As Date
    Dim daysElapsed As Long
    
    Set wsPrets = ThisWorkbook.Worksheets("prets")
    count = 0
    
    For i = 2 To wsPrets.Cells(wsPrets.Rows.Count, 1).End(xlUp).Row
        If wsPrets.Cells(i, 15).Value = "" Then ' En cours
            datePret = wsPrets.Cells(i, 4).Value
            daysElapsed = DateDiff("d", datePret, Now)
            If daysElapsed >= daysThreshold Then count = count + 1
        End If
    Next i
    
    CountOverdueLoans = count
End Function
