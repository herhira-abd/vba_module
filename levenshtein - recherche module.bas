Attribute VB_Name = "Module1"
Function Levenshtein(s1 As String, s2 As String) As Long
    Dim i As Long, j As Long
    Dim matrix() As Long
    Dim cost As Long
    
    ' Ajuster pour les majuscules/minuscules et supprimer les espaces en trop
    s1 = Trim(UCase(s1))
    s2 = Trim(UCase(s2))

    ' Initialiser la matrice
    ReDim matrix(0 To Len(s1), 0 To Len(s2))
    
    For i = 0 To Len(s1)
        matrix(i, 0) = i
    Next i
    For j = 0 To Len(s2)
        matrix(0, j) = j
    Next j
    
    ' Calculer les distances
    For i = 1 To Len(s1)
        For j = 1 To Len(s2)
            If Mid(s1, i, 1) = Mid(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            matrix(i, j) = Application.Min(matrix(i - 1, j) + 1, _
                                           matrix(i, j - 1) + 1, _
                                           matrix(i - 1, j - 1) + cost)
        Next j
    Next i

    Levenshtein = matrix(Len(s1), Len(s2))
End Function

Sub CompareLists()
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    
    Set ws = ThisWorkbook.Sheets("recherche")
        idx_page = ws.Cells(1, 3).Value 'cellule c1
    Set ws2 = ThisWorkbook.Sheets(idx_page) ' Modifier selon la feuille de travail
    
    Dim tx1  As String, tx2 As String
    Dim i As Long, j As Long
    Dim lastRowA As Long, lastRowB As Long, firstrow As Integer, rs_col As String
    Dim threshold As Long ' Tolérance pour les erreurs de distance


    lastRowA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row ' Dernière ligne de la colonne A
    lastRowB = ws2.Cells(ws2.Rows.Count, 1).End(xlUp).Row ' Dernière ligne de la colonne B
    threshold = ws.Cells(3, 3).Value ' Tolérance de distance d'édition
    
    firstrow = ws.Cells(5, 3).Value 'cellule c5 valeur
    rs_col = ws.Cells(7, 3).Value    'cellule c7 valeur
    col_rs = Range(rs_col & "1").Column
    
    
    For i = firstrow To lastRowA
     If ws.Cells(i, col_rs).Value = "" Then
        tx1 = ws.Cells(i, 1).Value
        For j = 1 To lastRowB
            tx2 = ws2.Cells(j, 1).Value
               
            If Levenshtein(tx1, tx2) <= threshold Then
                ws.Cells(i, col_rs).Value = ws2.Cells(j, 1).Value ' Écrire la correspondance
                Exit For
            End If
        Next j
      End If
    Next i
End Sub


