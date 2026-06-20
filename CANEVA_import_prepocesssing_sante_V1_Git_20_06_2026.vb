'========================================================
' PROCEDURE UNIQUE
'========================================================
Sub procedure_complete()

    Dim chemin As String
    Dim fichier As String

    Dim wbSource As Workbook
    Dim wsSource As Worksheet

    Dim wbData As Workbook
    Dim wsData As Worksheet

    Dim derniereLigne As Long
    Dim derniereLigneData As Long

    Dim plage As Range
    Dim cellule As Range

    Dim colProduit As Long
    Dim colDateNaissance As Long
    Dim colTelephone As Long
    Dim colMethode As Long
    Dim colMode As Long
    Dim colCalcul As Long

    Dim i As Long
    Dim produit As String

    '--------------------------------------------
    ' DOSSIER
    '--------------------------------------------
    chemin = "C:\Users\harouna.ouili\Desktop\Imports\"
    If Right(chemin, 1) <> "\" Then chemin = chemin & "\"

    '--------------------------------------------
    ' FICHIER DESTINATION
    '--------------------------------------------
    Set wbData = Workbooks.Add
    Set wsData = wbData.Sheets(1)
    wsData.Name = "Data"

    '--------------------------------------------
    ' IMPORT DES FICHIERS
    '--------------------------------------------
    fichier = Dir(chemin & "*.xls*")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Do While fichier <> ""

        Set wbSource = Workbooks.Open(chemin & fichier)
        Set wsSource = wbSource.Sheets(1)

        derniereLigne = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

        'Trim global
        Set plage = wsSource.Range("A1:Z" & derniereLigne)

        For Each cellule In plage
            If Not IsEmpty(cellule.Value) Then
                cellule.Value = Trim(cellule.Value)
            End If
        Next cellule

        derniereLigneData = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
        If derniereLigneData = 1 And wsData.Range("A1").Value = "" Then derniereLigneData = 0

        If derniereLigneData = 0 Then
            wsSource.Range("A1:Z" & derniereLigne).Copy
            wsData.Range("A1").PasteSpecial xlPasteValues
        Else
            wsSource.Range("A2:Z" & derniereLigne).Copy
            wsData.Range("A" & derniereLigneData + 1).PasteSpecial xlPasteValues
        End If

        wbSource.Close False
        fichier = Dir()

    Loop

    Application.CutCopyMode = False

    '--------------------------------------------
    ' HEADERS
    '--------------------------------------------
    wsData.Rows(1).Delete

    wsData.Range("A1:Z1").Value = Array( _
        "ProductName", "PolicyId", "AgentId", "FirstName", "LastName", _
        "BirthDate", "PhoneNumber", "NationalId", "Adress", "City", _
        "EmailAddress", "PaymentMethod", "Frequence", "InceptionDate", _
        "ExpiryDate", "PremiumCalculationMode", "SubscriberFirstName", _
        "SubscriberLastName", "SubscriberPhoneNumber", "SubscriberDateOfBirth", _
        "SubscriberNationalIdNumber", "SubscriberEmail", "SubscriberCity", _
        "SubscriberPostalAddress", "SubscriberPostalBox", "PromotionCode")

    '--------------------------------------------
    ' COLONNES
    '--------------------------------------------
    derniereLigne = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row

    colProduit = trouverColonne(wsData, "ProductName")
    colDateNaissance = trouverColonne(wsData, "BirthDate")
    colTelephone = trouverColonne(wsData, "PhoneNumber")
    colMethode = trouverColonne(wsData, "PaymentMethod")
    colMode = trouverColonne(wsData, "Frequence")
    colCalcul = trouverColonne(wsData, "PremiumCalculationMode")

    '================================================
    ' TRAITEMENT
    '================================================
    For i = 2 To derniereLigne

        '----------------------------
        ' PRODUIT
        '----------------------------
        If colProduit > 0 Then
            produit = UCase(Replace(wsData.Cells(i, colProduit).Value, " ", ""))
            wsData.Cells(i, colProduit).Value = produit
        End If

        '----------------------------
        ' DATE NAISSANCE (CORRIGÉ)
        '----------------------------
        If colDateNaissance > 0 Then
            wsData.Cells(i, colDateNaissance).Value = _
                ConvertirDateDMY(wsData.Cells(i, colDateNaissance).Value)

            wsData.Cells(i, colDateNaissance).NumberFormat = "dd/mm/yyyy"
        End If

        '----------------------------
        ' TELEPHONE
        '----------------------------
        If colTelephone > 0 Then
            If wsData.Cells(i, colTelephone).Value <> "" Then
                If Left(wsData.Cells(i, colTelephone).Value, 3) <> "226" Then
                    wsData.Cells(i, colTelephone).Value = "226" & wsData.Cells(i, colTelephone).Value
                End If
            End If
        End If

        '----------------------------
        ' FIXES METIER
        '----------------------------
        If colMethode > 0 Then wsData.Cells(i, colMethode).Value = "MoneyTransfer"
        If colMode > 0 Then wsData.Cells(i, colMode).Value = "YEARLY"
        If colCalcul > 0 Then wsData.Cells(i, colCalcul).Value = "DEFAULT_MODE"

    Next i

    wsData.Columns.AutoFit

    wbData.SaveAs chemin & "Data.xlsx"

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Traitement terminé avec succès"

End Sub

'========================================================
' CONVERSION DATE SÉCURISÉE (dd/mm/yyyy)
'========================================================
Function ConvertirDateDMY(valeur As Variant) As Variant

    Dim parts() As String
    Dim d As Integer, m As Integer, y As Integer

    If IsDate(valeur) Then
        'force format français
        ConvertirDateDMY = CDate(valeur)
        Exit Function
    End If

    If InStr(valeur, "/") > 0 Then
        parts = Split(valeur, "/")

        If UBound(parts) = 2 Then
            d = Val(parts(0))
            m = Val(parts(1))
            y = Val(parts(2))

            ConvertirDateDMY = DateSerial(y, m, d)
            Exit Function
        End If
    End If

    ConvertirDateDMY = valeur

End Function

'========================================================
' RECHERCHE COLONNE
'========================================================
Function trouverColonne(ws As Worksheet, nomColonne As String) As Long

    Dim cellule As Range

    For Each cellule In ws.Rows(1).Cells
        If Trim(UCase(cellule.Value)) = Trim(UCase(nomColonne)) Then
            trouverColonne = cellule.Column
            Exit Function
        End If
    Next cellule

    trouverColonne = 0

End Function

