REM  *****  BASIC  *****

Option VBASupport 1
Option Explicit

' Error constants (LibreOffice Calc error codes)
Const ERR_NA As Long = 2042                ' #N/A error (no match found)
Const ERR_VALUE As Long = 2036             ' #VALUE! error (invalid input)
' Default parameter values
Const DEFAULT_MIN_PERCENT As Single = 0.05
Const DEFAULT_ALGORITHM As Integer = 1
Const DEFAULT_RANK As Integer = 1
' Jaro-Winkler tuning constants
Const JARO_WINKLER_PREFIX_SCALE As Single = 0.1 ' Standard Winkler prefix weight
Const JARO_WINKLER_MAX_PREFIX_LEN As Integer = 4 ' Max prefix length for Winkler adjustment

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
    Dim intAlgorithm As Integer
    Dim blnNormalised As Boolean

    ' Optional args from Calc formulas are Variants; default explicitly.
    If IsMissing(Algorithm) Or IsEmpty(Algorithm) Then
        intAlgorithm = DEFAULT_ALGORITHM
    Else
        intAlgorithm = CInt(Algorithm)
    End If
    If intAlgorithm < 1 Or intAlgorithm > 2 Then intAlgorithm = DEFAULT_ALGORITHM

    If IsMissing(Normalised) Or IsEmpty(Normalised) Then
        blnNormalised = False
    Else
        blnNormalised = CBool(Normalised)
    End If

    '-------------------------------------------------------
    '-- If strings haven't been normalised, normalise them --
    '-------------------------------------------------------
    If blnNormalised = False Then
        String1 = LCase$(Trim(String1))
        String2 = LCase$(Trim(String2))
        ' Tokenize and sort to handle reversed names (e.g. "John Smith" vs "Smith, John")
        String1 = TokenizeAndSort(String1)
        String2 = TokenizeAndSort(String2)
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

    '----------------------------------------
    '-- Give 0% match if string length < 2 --
    '----------------------------------------
    If Len(String1) < 2 Or Len(String2) < 2 Then
        FuzzyPercent = 0
        Exit Function
    End If

    '--------------------------------------------------------
    '-- Algorithm 1 (default): Jaro-Winkler similarity      --
    '-- Algorithm 2: Normalized Levenshtein distance       --
    '--------------------------------------------------------
    If intAlgorithm = 1 Then
        FuzzyPercent = JaroWinklerSimilarity(String1, String2)
    Else
        FuzzyPercent = LevenshteinSimilarity(String1, String2)
    End If

End Function


Private Function TokenizeAndSort(ByVal str As String) As String
    Dim tokens() As String
    ' Split on spaces and commas, replace commas with spaces first
    tokens = Split(Replace(str, ",", " "), " ")
    
    ' Filter out empty tokens
    Dim cleanTokens() As String
    ReDim cleanTokens(0 To UBound(tokens))
    Dim tokenCount As Integer
    tokenCount = 0
    Dim i As Integer
    For i = 0 To UBound(tokens)
        If Trim(tokens(i)) <> "" Then
            cleanTokens(tokenCount) = tokens(i)
            tokenCount = tokenCount + 1
        End If
    Next i
    
    If tokenCount = 0 Then
        TokenizeAndSort = ""
        Exit Function
    End If
    
    ' Trim to actual token count
    ReDim Preserve cleanTokens(0 To tokenCount - 1)
    
    ' Bubble sort tokens alphabetically (small arrays, simple sort is fine)
    Dim j As Integer, temp As String
    For i = 0 To tokenCount - 1
        For j = i + 1 To tokenCount - 1
            If cleanTokens(i) > cleanTokens(j) Then
                temp = cleanTokens(i)
                cleanTokens(i) = cleanTokens(j)
                cleanTokens(j) = temp
            End If
        Next j
    Next i
    
    TokenizeAndSort = Join(cleanTokens, " ")
End Function


Private Function JaroWinklerSimilarity(ByVal s1 As String, ByVal s2 As String) As Single
    Dim len1 As Integer, len2 As Integer
    len1 = Len(s1)
    len2 = Len(s2)
    
    If len1 = 0 And len2 = 0 Then
        JaroWinklerSimilarity = 1
        Exit Function
    End If
    If len1 = 0 Or len2 = 0 Then
        JaroWinklerSimilarity = 0
        Exit Function
    End If
    
    ' Max distance for matching characters: floor(max(len1, len2) / 2) - 1
    Dim maxDist As Integer
    maxDist = len1
    If len2 > len1 Then maxDist = len2
    maxDist = Int(maxDist / 2) - 1
    If maxDist < 0 Then maxDist = 0
    
    ' Track matching characters in each string
    Dim match1() As Boolean, match2() As Boolean
    ReDim match1(1 To len1) As Boolean
    ReDim match2(1 To len2) As Boolean
    
    ' Count matching characters
    Dim m As Integer
    m = 0
    Dim i As Integer, j As Integer
    For i = 1 To len1
        Dim startPos As Integer, endPos As Integer
        startPos = 1
        If i - maxDist > 1 Then startPos = i - maxDist
        endPos = len2
        If i + maxDist < len2 Then endPos = i + maxDist
        
        For j = startPos To endPos
            If Not match2(j) Then
                If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                    match1(i) = True
                    match2(j) = True
                    m = m + 1
                    Exit For
                End If
            End If
        Next j
    Next i
    
    If m = 0 Then
        JaroWinklerSimilarity = 0
        Exit Function
    End If
    
    ' Count transpositions (mismatched order matches)
    Dim t As Integer
    t = 0
    Dim k As Integer
    k = 1
    For i = 1 To len1
        If match1(i) Then
            Do While k <= len2 And Not match2(k)
                k = k + 1
            Loop
            If k <= len2 Then
                If Mid$(s1, i, 1) <> Mid$(s2, k, 1) Then
                    t = t + 1
                End If
                k = k + 1
            End If
        End If
    Next i
    t = t \ 2  ' Integer division (transpositions are counted twice)
    
    ' Calculate Jaro distance
    Dim jaro As Single
    jaro = (m / len1 + m / len2 + (m - t) / m) / 3
    
    ' Winkler prefix adjustment
    Dim l As Integer
    l = 0
    For i = 1 To JARO_WINKLER_MAX_PREFIX_LEN
        If i <= len1 And i <= len2 Then
            If Mid$(s1, i, 1) = Mid$(s2, i, 1) Then
                l = l + 1
            Else
                Exit For
            End If
        Else
            Exit For
        End If
    Next i
    
    ' Final Jaro-Winkler similarity
    JaroWinklerSimilarity = jaro + l * JARO_WINKLER_PREFIX_SCALE * (1 - jaro)
End Function


Private Function LevenshteinSimilarity(ByVal s1 As String, ByVal s2 As String) As Single
    Dim len1 As Integer, len2 As Integer
    Dim maxLen As Integer
    Dim i As Integer, j As Integer
    Dim cost As Integer
    Dim deletion As Integer, insertion As Integer, substitution As Integer
    Dim dist() As Integer
    Dim editDistance As Integer

    len1 = Len(s1)
    len2 = Len(s2)

    If len1 = 0 And len2 = 0 Then
        LevenshteinSimilarity = 1
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
            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
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
        LevenshteinSimilarity = 1
    Else
        LevenshteinSimilarity = 1 - (editDistance / maxLen)
    End If
End Function


Function FuzzyVLookup(ByVal LookupValue As String, _
                     ByVal TableArray As Variant, _
                     ByVal IndexNum As Integer, _
                     Optional NFPercent As Variant, _
                     Optional Rank As Variant, _
                     Optional Algorithm As Variant) As Variant
    On Error GoTo ErrorHandler
    
    Dim lRow As Long
    Dim minPercent As Single
    Dim curPercent As Single
    Dim bestMatchPtr As Long
    Dim sortedRanks() As RankInfo
    Dim listString As String
    Dim curValue As Variant
    Dim nCols As Long
    Dim intRank As Integer
    Dim intAlgorithm As Integer
    Dim arrData As Variant
    Dim rowLB As Long, rowUB As Long
    Dim colLB As Long, colUB As Long
    Dim retCol As Long
    Dim relRow As Long
    Dim haveData As Boolean
    
    LookupValue = LCase$(Trim(LookupValue))
    
    ' Normalize TableArray into a 2D array.
    ' With Option VBASupport 1, Calc passes a range reference as a VBA-style
    ' Range object (TypeName = "Range").  Its .Value property gives a 2D array.
    ' A plain array can be used directly.
    haveData = False
    If IsArray(TableArray) Then
        arrData = TableArray
        haveData = True
    ElseIf IsObject(TableArray) Then
        If TableArray Is Nothing Then
            FuzzyVLookup = "*** TableArray is invalid ***"
            Exit Function
        End If
        If TypeName(TableArray) = "Range" Then
            ' VBA-compatible Range object — .Value returns a 2D Variant array
            arrData = TableArray.Value
            haveData = True
        Else
            ' UNO SheetCellRange — use getDataArray()
            On Error GoTo RangeArrayError
            arrData = TableArray.getDataArray()
            haveData = True
            On Error GoTo ErrorHandler
        End If
    End If
    
    If haveData = False Then
        FuzzyVLookup = "*** TableArray must be a CellRange or range array ***"
        Exit Function
    End If
    
    rowLB = LBound(arrData, 1)
    rowUB = UBound(arrData, 1)
    colLB = LBound(arrData, 2)
    colUB = UBound(arrData, 2)
    nCols = colUB - colLB + 1
    
    ' Assign defaults for optional parameters (conversion-safe)
    intRank = DEFAULT_RANK
    If Not (IsMissing(Rank) Or IsEmpty(Rank)) Then
        If IsNumeric(Rank) Then intRank = CInt(Rank)
    End If
    
    intAlgorithm = DEFAULT_ALGORITHM
    If Not (IsMissing(Algorithm) Or IsEmpty(Algorithm)) Then
        If IsNumeric(Algorithm) Then intAlgorithm = CInt(Algorithm)
    End If
    If intAlgorithm < 1 Or intAlgorithm > 2 Then intAlgorithm = DEFAULT_ALGORITHM
    
    If IndexNum < 0 Then
        FuzzyVLookup = "*** IndexNum must be greater than or equal to 0 ***"
        Exit Function
    End If
    
    If intRank < 1 Then
        FuzzyVLookup = "*** 'Rank' must be an integer > 0 ***"
        Exit Function
    End If
    
    minPercent = DEFAULT_MIN_PERCENT
    If Not (IsMissing(NFPercent) Or IsEmpty(NFPercent)) Then
        If Not IsNumeric(NFPercent) Then
            FuzzyVLookup = "*** 'NFPercent' must be numeric and > 0 and <= 1 ***"
            Exit Function
        End If
        minPercent = CSng(NFPercent)
        If (minPercent <= 0) Or (minPercent > 1) Then
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
        curValue = arrData(lRow, colLB)
        
        ' Skip blank cells rather than halting — the table may have gaps
        If Trim(CStr(curValue)) <> "" Then
            listString = LCase$(Trim(curValue))
            
            curPercent = FuzzyPercent(String1:=LookupValue, _
                                          String2:=listString, _
                                          Algorithm:=intAlgorithm, _
                                          Normalised:=True)
            
            If curPercent >= minPercent Then
                ' Insert into sortedRanks using binary search
                InsertSortedRank sortedRanks, intRank, lRow, curPercent
            End If
        End If
        
        lRow = lRow + 1
    Loop
    
    If sortedRanks(intRank).Percentage < minPercent Then
        FuzzyVLookup = CVErr(ERR_NA)
    Else
        bestMatchPtr = sortedRanks(intRank).Offset
        If IndexNum > 0 Then
            retCol = colLB + IndexNum - 1
            FuzzyVLookup = arrData(bestMatchPtr, retCol)
        Else
            relRow = bestMatchPtr - rowLB + 1
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
    FuzzyVLookup = CVErr(ERR_VALUE)
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
    Dim oProbeCell As Object
    Dim vResult As Variant
    Dim sFormula As String
    Dim sArgSep As String
    Dim nErr As Long
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
    
    ' --- Test by writing the formula to a cell and reading the result ---
    ' This simulates a real user call, which is necessary to get the
    ' special VBA-style Range object that FuzzyVLookup expects.
    oProbeCell = oSheet.getCellByPosition(6, 0) ' G1 (test output)
    oCell = oSheet.getCellByPosition(5, 0) ' F1 (formula cell)
    
    ' Detect Calc argument separator for locale-agnostic formula construction.
    sArgSep = ";"
    oProbeCell.setFormula("=SUM(1;2)")
    nErr = oProbeCell.getError()
    If nErr <> 0 Then sArgSep = ","
    
    sFormula = "=FUZZYVLOOKUP(""Willam""" & sArgSep & " A2:C5" & sArgSep & " 2" & sArgSep & " 1/2" & sArgSep & " 1" & sArgSep & " 1)"
    oCell.setFormula(sFormula)
    
    ' Read the result from the cell
    vResult = oCell.getString()
    If vResult = "" Then vResult = oCell.getValue()
    
    ' Display result in test sheet cell instead of MsgBox (headless-friendly)
    If oCell.getError() <> 0 Then
        msg = "Test failed. Formula returned error code: " & oCell.getError()
        oProbeCell.String = msg
    Else
        msg = "Match found for 'Willam': " & vResult
        oProbeCell.String = msg
    End If
End Sub