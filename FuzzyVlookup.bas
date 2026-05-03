REM  *****  BASIC  *****

Option VBASupport 1
Option Explicit

Type RankInfo
    Offset          As Long
    Percentage      As Single
End Type


'*************************************
'** Return a % match on two strings **
'*************************************
Function FuzzyPercent(ByVal String1 As String, _
                     ByVal String2 As String, _
                     Optional Algorithm As Variant, _
                     Optional Normalised As Variant) As Single
    Dim intLen1 As Integer, intLen2 As Integer
    Dim intScore As Integer
    Dim intTotScore As Integer
    Dim intScoreAlg1 As Integer
    Dim intTotScoreAlg1 As Integer
    Dim intScoreAlg2 As Integer
    Dim intTotScoreAlg2 As Integer
    Dim sngPercentAlg1 As Single
    Dim sngPercentAlg2 As Single
    Dim sngPercentEdit As Single

    Dim intAlgorithm As Integer
    Dim blnNormalised As Boolean

    ' Optional args from Calc formulas are Variants; default explicitly.
    If IsMissing(Algorithm) Or IsEmpty(Algorithm) Then
        intAlgorithm = 3
    Else
        intAlgorithm = CInt(Algorithm)
    End If
    If intAlgorithm < 1 Or intAlgorithm > 3 Then intAlgorithm = 3

    If IsMissing(Normalised) Or IsEmpty(Normalised) Then
        blnNormalised = FALSE
    Else
        blnNormalised = CBool(Normalised)
    End If

    '-------------------------------------------------------
    '-- If strings haven't been normalised, normalise them --
    '-------------------------------------------------------
    If blnNormalised = FALSE Then
        String1 = LCase$(Trim(String1))
        String2 = LCase$(Trim(String2))
    End If

    ' Error handling for empty strings
    If Len(String1) = 0 Or Len(String2) = 0 Then
        FuzzyPercent = 0
        Exit Function
    End If

    '----------------------------------------------
    '-- Give 100% match if strings exactly equal --
    '----------------------------------------------
    If String1 = String2 Then
        FuzzyPercent = 1
        Exit Function
    End If

    intLen1 = Len(String1)
    intLen2 = Len(String2)

    '----------------------------------------
    '-- Give 0% match if string length < 2 --
    '----------------------------------------
    If intLen1 < 2 Or intLen2 < 2 Then
        FuzzyPercent = 0
        Exit Function
    End If

    intScore = 0
    intTotScore = 0
    intScoreAlg1 = 0
    intTotScoreAlg1 = 0
    intScoreAlg2 = 0
    intTotScoreAlg2 = 0
    sngPercentAlg1 = 0
    sngPercentAlg2 = 0
    sngPercentEdit = 0

    '--------------------------------------------------------
    '-- If Algorithm = 1 or 3, Search for single characters --
    '--------------------------------------------------------
    If (intAlgorithm And 1) <> 0 Then
        FuzzyAlg1 String1, String2, intScoreAlg1, intTotScoreAlg1
        If intLen1 < intLen2 Then FuzzyAlg1 String2, String1, intScoreAlg1, intTotScoreAlg1
        If intTotScoreAlg1 > 0 Then sngPercentAlg1 = intScoreAlg1 / intTotScoreAlg1
    End If

    '-----------------------------------------------------------
    '-- If Algorithm = 2 or 3, Search for pairs, triplets etc. --
    '-----------------------------------------------------------
    If (intAlgorithm And 2)<> 0 Then
    	FuzzyAlg2 String1, String2, intScoreAlg2, intTotScoreAlg2
        If intLen1 < intLen2 Then FuzzyAlg2 String2, String1, intScoreAlg2, intTotScoreAlg2
        If intTotScoreAlg2 > 0 Then sngPercentAlg2 = intScoreAlg2 / intTotScoreAlg2
    End If

    If intAlgorithm = 1 Then
        FuzzyPercent = sngPercentAlg1
    ElseIf intAlgorithm = 2 Then
        FuzzyPercent = sngPercentAlg2
    Else
        ' Edit-distance similarity is strong for small typos (insert/delete/replace).
        sngPercentEdit = NormalizedEditSimilarity(String1, String2)

        ' Combined mode should not penalize close typo matches when one algorithm
        ' already indicates a strong match.
        If sngPercentAlg1 > sngPercentAlg2 Then
            FuzzyPercent = sngPercentAlg1
        Else
            FuzzyPercent = sngPercentAlg2
        End If
        If sngPercentEdit > FuzzyPercent Then FuzzyPercent = sngPercentEdit
    End If

End Function


Private Function NormalizedEditSimilarity(ByVal String1 As String, ByVal String2 As String) As Single
    Dim len1 As Integer, len2 As Integer
    Dim maxLen As Integer
    Dim i As Integer, j As Integer
    Dim cost As Integer
    Dim deletion As Integer, insertion As Integer, substitution As Integer
    Dim dist() As Integer
    Dim editDistance As Integer

    len1 = Len(String1)
    len2 = Len(String2)

    If len1 = 0 And len2 = 0 Then
        NormalizedEditSimilarity = 1
        Exit Function
    End If

    ReDim dist(0 To len1, 0 To len2)

    For i = 0 To len1
        dist(i, 0) = i
    Next i
    For j = 0 To len2
        dist(0, j) = j
    Next j

    For i = 1 To len1
        For j = 1 To len2
            If Mid$(String1, i, 1) = Mid$(String2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If

            deletion = dist(i - 1, j) + 1
            insertion = dist(i, j - 1) + 1
            substitution = dist(i - 1, j - 1) + cost

            dist(i, j) = deletion
            If insertion < dist(i, j) Then dist(i, j) = insertion
            If substitution < dist(i, j) Then dist(i, j) = substitution
        Next j
    Next i

    editDistance = dist(len1, len2)
    maxLen = len1
    If len2 > maxLen Then maxLen = len2

    If maxLen = 0 Then
        NormalizedEditSimilarity = 1
    Else
        NormalizedEditSimilarity = 1 - (editDistance / maxLen)
    End If
End Function


Private Sub FuzzyAlg1(ByVal String1 As String, _
        ByVal String2 As String, _
        ByRef Score As Integer, _
        ByRef TotScore As Integer)
    Dim intLen1     As Integer, intPos As Integer, intPtr As Integer, intStartPos As Integer
    Dim foundChars() As Boolean ' Keep track of found characters
    
    intLen1 = Len(String1)
    TotScore = TotScore + intLen1        'update total possible score
    ReDim foundChars(1 To Len(String2)) ' Initialize foundChars array
    
    intPos = 0
    For intPtr = 1 To intLen1
        intStartPos = intPos + 1
        intPos = InStr(intStartPos, String2, Mid$(String1, intPtr, 1))
        If intPos > 0 Then
            If intPos > intStartPos + 3 Then
                intPos = intStartPos
            Else
                If Not foundChars(intPos) Then ' Check if character was already found
                    Score = Score + 1
                    foundChars(intPos) = True ' Mark character as found
                End If
            End If
        Else
            intPos = intStartPos
        End If
    Next intPtr
End Sub


Private Sub FuzzyAlg2(ByVal String1 As String, _
                      ByVal String2 As String, _
                      ByRef Score As Integer, _
                      ByRef TotScore As Integer)
    Dim intCurLen As Integer, intLen1 As Integer, intTo As Integer, intPtr As Integer, intPos As Integer, i as Integer
    Dim strWork As String
    Dim corruptPositions() As Boolean ' Array to track corrupted positions

    intLen1 = Len(String1)
    strWork = String2 ' Create the copy once

    For intCurLen = 2 To intLen1
        ' Score contribution for this pass: Int(intLen1 / intCurLen) represents the
        ' maximum number of non-overlapping substrings of length intCurLen that can
        ' fit in String1.  Array is sized Len(strWork)+intLen1 to prevent out-of-bounds
        ' writes when a match starts near the end of strWork and spans intCurLen-1
        ' positions past the last character.
        ReDim corruptPositions(1 To Len(strWork) + intLen1)

        intTo = intLen1 - intCurLen + 1
        TotScore = TotScore + Int(intLen1 / intCurLen)

        For intPtr = 1 To intTo Step intCurLen
            intPos = InStr(strWork, Mid$(String1, intPtr, intCurLen))
            If intPos > 0 Then
                If Not corruptPositions(intPos) Then
                    For i = 0 to intCurLen -1
                        corruptPositions(intPos + i) = True
                    Next i
                    Score = Score + 1
                End If
            End If
        Next intPtr
    Next intCurLen
End Sub


Function FuzzyVLookup(ByVal LookupValue As String, _
                     ByVal TableArray As Variant, _
                     ByVal IndexNum As Integer, _
                     Optional NFPercent As Variant, _
                     Optional Rank As Variant, _
                     Optional Algorithm As Variant) As Variant
    On Error GoTo ErrorHandler

    Dim lRow As Long
    Dim sngMinPercent As Single
    Dim sngCurPercent As Single
    Dim intBestMatchPtr As Long
    Dim sortedRanks() As RankInfo
    Dim strListString as String
    Dim vCurValue As Variant
    Dim nCols As Long
    Dim intRank As Integer
    Dim intAlgorithm As Integer
    Dim arrData As Variant
    Dim rowLB As Long, rowUB As Long
    Dim colLB As Long, colUB As Long
    Dim retCol As Long
    Dim relRow As Long
    Dim blnHaveData As Boolean

    LookupValue = LCase$(Trim(LookupValue))

    ' Normalize TableArray into a 2D array.
    ' With Option VBASupport 1, Calc passes a range reference as a VBA-style
    ' Range object (TypeName = "Range").  Its .Value property gives a 2D array.
    ' A plain array can be used directly.
    blnHaveData = False
    If IsArray(TableArray) Then
        arrData = TableArray
        blnHaveData = True
    ElseIf IsObject(TableArray) Then
        If TableArray Is Nothing Then
            FuzzyVLookup = "*** TableArray is invalid ***"
            Exit Function
        End If
        If TypeName(TableArray) = "Range" Then
            ' VBA-compatible Range object — .Value returns a 2D Variant array
            arrData = TableArray.Value
            blnHaveData = True
        Else
            ' UNO SheetCellRange — use getDataArray()
            On Error GoTo RangeArrayError
            arrData = TableArray.getDataArray()
            blnHaveData = True
            On Error GoTo ErrorHandler
        End If
    End If

    If blnHaveData = False Then
        FuzzyVLookup = "*** TableArray must be a CellRange or range array ***"
        Exit Function
    End If

    rowLB = LBound(arrData, 1)
    rowUB = UBound(arrData, 1)
    colLB = LBound(arrData, 2)
    colUB = UBound(arrData, 2)
    nCols = colUB - colLB + 1

    ' Assign defaults for optional parameters (conversion-safe)
    intRank = 1
    If Not (IsMissing(Rank) Or IsEmpty(Rank)) Then
        If IsNumeric(Rank) Then intRank = CInt(Rank)
    End If

    intAlgorithm = 3
    If Not (IsMissing(Algorithm) Or IsEmpty(Algorithm)) Then
        If IsNumeric(Algorithm) Then intAlgorithm = CInt(Algorithm)
    End If
    If intAlgorithm < 1 Or intAlgorithm > 3 Then intAlgorithm = 3

    If IndexNum < 0 Then
        FuzzyVLookup = "*** IndexNum must be greater than or equal to 0 ***"
        Exit Function
    End If

    If intRank < 1 Then
        FuzzyVLookup = "*** 'Rank' must be an integer > 0 ***"
        Exit Function
    End If

    sngMinPercent = 0.05
    If Not (IsMissing(NFPercent) Or IsEmpty(NFPercent)) Then
        If Not IsNumeric(NFPercent) Then
            FuzzyVLookup = "*** 'NFPercent' must be numeric and > 0 and <= 1 ***"
            Exit Function
        End If
        sngMinPercent = CSng(NFPercent)
        If (sngMinPercent <= 0) Or (sngMinPercent > 1) Then
            FuzzyVLookup = "*** 'NFPercent' must be a percentage > 0 and <= 1 ***"            
            Exit Function
        End If
    End If

    If IndexNum > nCols And IndexNum > 0 Then
        FuzzyVLookup = "*** IndexNum out of bounds ***"
        Exit Function
    End If

    ReDim sortedRanks(1 To intRank)

    lRow = rowLB
    Do While lRow <= rowUB
        vCurValue = arrData(lRow, colLB)

        ' Skip blank cells rather than halting — the table may have gaps
        If Trim(CStr(vCurValue)) <> "" Then
            strListString = LCase$(Trim(vCurValue))

            sngCurPercent = FuzzyPercent(String1:=LookupValue, _
                                          String2:=strListString, _
                                          Algorithm:=intAlgorithm, _
                                          Normalised:=True)

            If sngCurPercent >= sngMinPercent Then
                ' Insert into sortedRanks using binary search
                InsertSortedRank sortedRanks, intRank, lRow, sngCurPercent
            End If
        End If

        lRow = lRow + 1
    Loop

    If sortedRanks(intRank).Percentage < sngMinPercent Then
        FuzzyVLookup = CVErr(2042)
    Else
        intBestMatchPtr = sortedRanks(intRank).Offset
        If IndexNum > 0 Then
            retCol = colLB + IndexNum - 1
            FuzzyVLookup = arrData(intBestMatchPtr, retCol)
        Else
            relRow = intBestMatchPtr - rowLB + 1
            FuzzyVLookup = relRow
        End If
    End If

    Exit Function
    RangeArrayError:
    FuzzyVLookup = "*** TableArray must be a CellRange or range array ***"
    Exit Function
    ErrorHandler:
    ' Return a #VALUE! error rather than showing a MsgBox, which is disruptive
    ' in cell-formula context and broken in headless environments.
    FuzzyVLookup = CVErr(2036)
End Function


Private Sub InsertSortedRank(ByRef ranks() As RankInfo, ByVal rankSize As Long, ByVal row As Long, ByVal percentage As Single)
    Dim i As Long, j As Long
    For i = 1 To rankSize
        If percentage > ranks(i).Percentage Then
            For j = rankSize To i + 1 Step -1
                ranks(j) = ranks(j - 1)
            Next j
            ranks(i).Offset = row
            ranks(i).Percentage = percentage
            Exit Sub
        End If
    Next i
End Sub


Sub TestFuzzyVLookup
    Dim oDoc As Object
    Dim oSheets As Object
    Dim oSheet As Object
    Dim oCell As Object
    Dim vResult As Variant
    Dim LookupValue As String
    Dim TableArray As CellRange
    Dim IndexNum As Integer
    Dim NFPercent As Single
    Dim Rank As Integer
    Dim Algorithm As Integer
    Dim msg As String
    Const TEST_SHEET_NAME As String = "FuzzyVLookupTest"

    ' Use a dedicated test sheet — never touch the user's active sheet
    oDoc = ThisComponent
    oSheets = oDoc.Sheets
    If oSheets.hasByName(TEST_SHEET_NAME) Then
        oSheet = oSheets.getByName(TEST_SHEET_NAME)
    Else
        oSheets.insertNewByName(TEST_SHEET_NAME, oSheets.Count)
        oSheet = oSheets.getByName(TEST_SHEET_NAME)
    End If

    ' Clear text and numeric values in the test sheet
    oSheet.clearContents(7)

    ' Define a small dataset in the sheet
    oSheet.getCellByPosition(0, 0).String = "Name"
    oSheet.getCellByPosition(1, 0).String = "Age"
    oSheet.getCellByPosition(2, 0).String = "City"

    oSheet.getCellByPosition(0, 1).String = "William"
    oSheet.getCellByPosition(1, 1).String = "25"
    oSheet.getCellByPosition(2, 1).String = "New York"

    oSheet.getCellByPosition(0, 2).String = "John"
    oSheet.getCellByPosition(1, 2).String = "30"
    oSheet.getCellByPosition(2, 2).String = "Los Angeles"

    oSheet.getCellByPosition(0, 3).String = "Anna"
    oSheet.getCellByPosition(1, 3).String = "28"
    oSheet.getCellByPosition(2, 3).String = "Chicago"

    oSheet.getCellByPosition(0, 4).String = "Michael"
    oSheet.getCellByPosition(1, 4).String = "35"
    oSheet.getCellByPosition(2, 4).String = "Houston"

    ' Define the lookup parameters
    LookupValue = "Willam" ' Intentionally misspelled to test fuzzy matching
    'TableArray = oSheet.getCellByPosition(0, 0) ' Top-left cell of the table
    TableArray = oSheet.getCellRangeByName("A2:C5") 
    IndexNum = 2 ' Return the "Age" column
    NFPercent = 0.5 ' Minimum match percentage (50%)
    Rank = 1 ' Return the best match
    Algorithm = 3 ' Use both algorithms
    ' Call the FuzzyVLookup function
    vResult = FuzzyVLookup(LookupValue, TableArray, IndexNum, NFPercent, Rank, Algorithm)

    ' Display the result
    If IsError(vResult) Then
        msg = "No match found for '" & LookupValue & "'."
    Else
        msg = "Match found for '" & LookupValue & "': " & vResult
    End If

    MsgBox msg, 0, "FuzzyVLookup Test Result"
End Sub

