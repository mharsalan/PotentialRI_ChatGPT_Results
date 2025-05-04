Sub StratifiedRandomSample()

    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim dataArray As Variant
    Dim panelCounts As Object, panelIndices As Object
    Dim sampleSizes As Object
    Dim totalSamples As Long
    Dim selectedIndices As Object, allIndices As Object
    Dim randIndex As Long, temp As Long

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" if your sheet has a different name

    ' Find the last row of data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Read the data into an array (assuming data starts from row 2 with headers in row 1)
    dataArray = ws.Range("A2:B" & lastRow).Value

    ' Initialize dictionaries/objects to store panel counts and indices
    Set panelCounts = CreateObject("Scripting.Dictionary")
    Set panelIndices = CreateObject("Scripting.Dictionary")
    Set sampleSizes = CreateObject("Scripting.Dictionary")
    Set selectedIndices = CreateObject("Scripting.Dictionary")
    Set allIndices = CreateObject("Scripting.Dictionary")

    ' Initialize panel counts and store row indices for each panel
    panelCounts("A") = 0: panelIndices("A") = "": allIndices("A") = ""
    panelCounts("B") = 0: panelIndices("B") = "": allIndices("B") = ""
    panelCounts("C") = 0: panelIndices("C") = "": allIndices("C") = ""
    panelCounts("D") = 0: panelIndices("D") = "": allIndices("D") = ""

    ' Count panels and store all row indices
    For i = 1 To UBound(dataArray, 1)
        If Not IsError(dataArray(i, 2)) And Not IsNull(dataArray(i, 2)) Then ' Check for errors or nulls
            panel = UCase(Trim(dataArray(i, 2)))
            If panel = "A" Or panel = "B" Or panel = "C" Or panel = "D" Then
                panelCounts(panel) = panelCounts(panel) + 1
                allIndices(panel) = allIndices(panel) & i & ","
            End If
        End If
    Next i

    totalSamples = 100

    ' Calculate sample sizes (proportional allocation)
    For Each key In panelCounts.Keys()
        sampleSizes(key) = Round((panelCounts(key) / (lastRow - 1)) * totalSamples) 'lastRow -1 since dataArray starts from row 2
    Next key

    ' Adjust sample sizes if needed to reach exactly 100 (prioritize panels A and D)
    Dim total As Long
    total = 0
    For Each key In sampleSizes.Keys()
        total = total + sampleSizes(key)
    Next key

    Dim diff As Long
    diff = totalSamples - total

    If diff <> 0 Then
        If sampleSizes.Exists("A") Then sampleSizes("A") = sampleSizes("A") + 1: diff = diff - 1
        If diff <> 0 And sampleSizes.Exists("D") Then sampleSizes("D") = sampleSizes("D") + 1: diff = diff - 1
        If diff <> 0 And sampleSizes.Exists("C") Then sampleSizes("C") = sampleSizes("C") + 1: diff = diff - 1
        If diff <> 0 And sampleSizes.Exists("B") Then sampleSizes("B") = sampleSizes("B") + 1: diff = diff - 1
    End If


    ' Select random indices for each panel
    For Each key In panelIndices.Keys()
        Dim indices() As String
        indices = Split(Left(allIndices(key), Len(allIndices(key)) - 1), ",") ' Remove trailing comma
        Dim n As Long: n = sampleSizes(key)
        Dim m As Long: m = UBound(indices) + 1

        If m > 0 Then ' Ensure there are indices to select from
            For i = 1 To n
                Randomize
                randIndex = Int((m - i + 1) * Rnd + 1)
                selectedIndices(key) = selectedIndices(key) & indices(randIndex - 1) & ","
                ' Swap the selected index with the last unselected index
                If randIndex < m - i + 1 Then
                    temp = indices(randIndex - 1)
                    indices(randIndex - 1) = indices(m - i)
                    indices(m - i) = temp
                End If
            Next i
        End If
    Next key

    ' Highlight selected rows and unselected rows
    Dim isSelected As Boolean
    Dim rowIndex As Long
    Dim selectedIndicesArray() As String
    Dim selectedIndex As Variant

    For i = 2 To lastRow
        isSelected = False
        For Each key In selectedIndices.Keys()
            If selectedIndices(key) <> "" Then
                selectedIndicesArray = Split(Left(selectedIndices(key), Len(selectedIndices(key)) - 1), ",")
                For Each selectedIndex In selectedIndicesArray
                    If i = CInt(selectedIndex) + 1 Then ' +1 because dataArray starts from row 2
                        isSelected = True
                        Exit For
                    End If
                Next selectedIndex
            End If
            If isSelected Then Exit For
        Next key

        If isSelected Then
            ws.Rows(i).Interior.Color = RGB(255, 204, 204) ' Light red for selected
        Else
            ws.Rows(i).Interior.Color = RGB(220, 220, 220) ' Light gray for unselected
        End If
    Next i

    MsgBox "Stratified random sampling complete. " & total & " rows colored.", vbInformation

End Sub