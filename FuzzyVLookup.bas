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
Const JARO_WINKLER_PREFIX_SCALE As Single = 0.1
Const JARO_WINKLER_MAX_PREFIX_LEN As Integer = 4
' Cache size limits (for 32GB RAM system)
Const MAX_NORM_CACHE As Long = 500000      ' 500k normalized strings ~50 MB
Const MAX_SCORE_CACHE As Long = 10000000   ' 10M score entries ~2.5 GB
' Blocking: number of chars for block key (prefix matching)
Const BLOCK_KEY_LENGTH As Integer = 3

' Module-level caches
Dim normCache As Collection        ' Normalized strings: key=original, value=normalized
Dim scoreCache As Collection       ' Fuzzy scores: key=norm1|norm2|algo, value=score
Dim scoreCacheCount As Long       ' Track score cache size for eviction
Dim cachesInitialized As Boolean  ' Track if caches have been initialized


'*************************************
'** Initialize caches if needed      **
'*************************************
Sub InitCaches()
    If Not cachesInitialized Then
        Set normCache = New Collection
        Set scoreCache = New Collection
        scoreCacheCount = 0
        cachesInitialized = True
    End If
End Sub


'*************************************
'** Get normalized string from cache **
'** Returns cached value or computes  **
'** and caches the result            **
'*************************************
Function GetNormalized(ByVal s As String) As String
    Dim normalized As Variant
    Dim cacheKey As String
    Dim found As Boolean
    
    InitCaches
    
    ' Use original string as cache key
    cacheKey = s
    
    ' Try to get from cache
    found = TryGetFromNormCache(cacheKey, normalized)
    
    If found Then
        GetNormalized = CStr(normalized)
        Exit Function
    End If
    
    ' Not in cache, compute normalization
    GetNormalized = LCase$(Trim(s))
    GetNormalized = TokenizeAndSort(GetNormalized)
    
    ' Add to cache if under limit
    If normCache.Count < MAX_NORM_CACHE Then
        TryAddToNormCache cacheKey, GetNormalized
    End If
End Function


' Helper: isolated error handling for collection lookup
Private Function TryGetFromNormCache(ByVal cacheKey As String, ByRef value As Variant) As Boolean
    On Error GoTo NotFound
    value = normCache.Item(cacheKey)
    TryGetFromNormCache = True
    On Error GoTo 0
    Exit Function
NotFound:
    TryGetFromNormCache = False
    Resume DoneNotFound
DoneNotFound:
    On Error GoTo 0
End Function


' Helper: isolated error handling for collection add
Private Sub TryAddToNormCache(ByVal cacheKey As String, ByVal value As String)
    On Error GoTo AddFailed
    normCache.Add value, cacheKey
    On Error GoTo 0
    Exit Sub
AddFailed:
    ' Key already exists or other error - silently ignore
    Resume DoneAddFailed
DoneAddFailed:
    On Error GoTo 0
End Sub


'*************************************
'** Get cached fuzzy score if exists **
'** Returns True and score if cached **
'*************************************
Function GetCachedScore(ByVal key As String, ByRef score As Single) As Boolean
    Dim cachedVal As Variant
    
    InitCaches
    
    If TryGetFromScoreCache(key, cachedVal) Then
        score = CSng(cachedVal)
        GetCachedScore = True
    Else
        GetCachedScore = False
    End If
End Function


' Helper: isolated error handling for collection lookup
Private Function TryGetFromScoreCache(ByVal key As String, ByRef value As Variant) As Boolean
    On Error GoTo NotFound
    value = scoreCache.Item(key)
    TryGetFromScoreCache = True
    On Error GoTo 0
    Exit Function
NotFound:
    TryGetFromScoreCache = False
    Resume DoneNotFound
DoneNotFound:
    On Error GoTo 0
End Function


'*************************************
'** Add fuzzy score to cache         **
'** Evicts all if over limit         **
'*************************************
Sub AddCachedScore(ByVal key As String, ByVal score As Single)
    InitCaches
    
    ' Evict all if over limit (simple strategy)
    If scoreCacheCount >= MAX_SCORE_CACHE Then
        Set scoreCache = New Collection
        scoreCacheCount = 0
    End If
    
    If TryAddToScoreCache(key, score) Then
        scoreCacheCount = scoreCacheCount + 1
    End If
End Sub


' Helper: isolated error handling for collection add
Private Function TryAddToScoreCache(ByVal key As String, ByVal score As Single) As Boolean
    On Error GoTo AddFailed
    scoreCache.Add score, key
    TryAddToScoreCache = True
    On Error GoTo 0
    Exit Function
AddFailed:
    TryAddToScoreCache = False
    Resume DoneAddFailed
DoneAddFailed:
    On Error GoTo 0
End Function


'*************************************
'** Build cache key for score cache **
'*************************************
Function BuildScoreKey(ByVal norm1 As String, ByVal norm2 As String, ByVal algo As Integer) As String
    BuildScoreKey = norm1 & "|" & norm2 & "|" & algo
End Function


'*************************************
'** Compute block key from string    **
'** Uses first N chars of normalized**
'** string for blocking optimization**
'*************************************
Function GetBlockKey(ByVal normalized As String) As String
    If Len(normalized) >= BLOCK_KEY_LENGTH Then
        GetBlockKey = Left$(normalized, BLOCK_KEY_LENGTH)
    Else
        GetBlockKey = normalized
    End If
End Function


Type RankInfo
    Offset          As Long
    Percentage      As Single
End Type


'*************************************
'** Return a % match on two strings **
'*************************************
Function FuzzyPercent(ByVal string1 As String, _
                     ByVal string2 As String, _
                     Optional algorithm As Variant, _
                     Optional normalised As Variant) As Single
    Dim algo As Integer
    Dim isNormalised As Boolean
    Dim norm1 As String, norm2 As String
    Dim scoreKey As String
    Dim cachedScore As Single

    ' Optional args from Calc formulas are Variants; default explicitly.
    If IsMissing(algorithm) Or IsEmpty(algorithm) Then
        algo = DEFAULT_ALGORITHM
    Else
        algo = CInt(algorithm)
    End If
    If algo < 1 Or algo > 2 Then algo = DEFAULT_ALGORITHM

    If IsMissing(normalised) Or IsEmpty(normalised) Then
        isNormalised = False
    Else
        isNormalised = CBool(normalised)
    End If

    '-------------------------------------------------------
    '-- If strings haven't been normalised, normalise them --
    '-- Uses cache to avoid re-processing                 --
    '-------------------------------------------------------
    If isNormalised = False Then
        norm1 = GetNormalized(string1)
        norm2 = GetNormalized(string2)
    Else
        norm1 = string1
        norm2 = string2
    End If

    ' Error handling for empty strings
    If Len(norm1) = 0 Or Len(norm2) = 0 Then
        FuzzyPercent = 0
        Exit Function
    End If

    '----------------------------------------------
    '-- Give 100% match if strings exactly equal --
    '----------------------------------------------
    If norm1 = norm2 Then
        FuzzyPercent = 1
        Exit Function
    End If

    '----------------------------------------
    '-- Give 0% match if string length < 2 --
    '----------------------------------------
    If Len(norm1) < 2 Or Len(norm2) < 2 Then
        FuzzyPercent = 0
        Exit Function
    End If

    '--------------------------------------------------------
    '-- Check score cache before computing                  --
    '--------------------------------------------------------
    scoreKey = BuildScoreKey(norm1, norm2, algo)
    If GetCachedScore(scoreKey, cachedScore) Then
        FuzzyPercent = cachedScore
        Exit Function
    End If

    '--------------------------------------------------------
    '-- Algorithm 1 (default): Jaro-Winkler similarity      --
    '-- Algorithm 2: Normalized Levenshtein distance       --
    '--------------------------------------------------------
    If algo = 1 Then
        FuzzyPercent = JaroWinklerSimilarity(norm1, norm2)
    Else
        FuzzyPercent = LevenshteinSimilarity(norm1, norm2)
    End If

    ' Cache the computed score
    AddCachedScore scoreKey, FuzzyPercent
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


Function FuzzyVLookup(ByVal lookupValue As String, _
                     ByVal tableArray As Variant, _
                     ByVal indexNum As Integer, _
                     Optional nfPercent As Variant, _
                     Optional rank As Variant, _
                     Optional algorithm As Variant) As Variant
    On Error GoTo ErrorHandler

    Dim row As Long
    Dim minPercent As Single
    Dim curPercent As Single
    Dim bestMatchPtr As Long
    Dim sortedRanks() As RankInfo
    Dim listString As String
    Dim curValue As Variant
    Dim nCols As Long
    Dim rankNum As Integer
    Dim algo As Integer
    Dim arrData As Variant
    Dim rowLB As Long, rowUB As Long
    Dim colLB As Long, colUB As Long
    Dim retCol As Long
    Dim relRow As Long
    Dim haveData As Boolean
    
    ' Normalized lookup value and its block key for blocking optimization
    Dim normLookup As String
    Dim lookupBlockKey As String

    ' Initialize caches
    InitCaches

    lookupValue = LCase$(Trim(lookupValue))
    normLookup = GetNormalized(lookupValue)
    lookupBlockKey = GetBlockKey(normLookup)

    ' Normalize TableArray into a 2D array.
    ' With Option VBASupport 1, Calc passes a range reference as a VBA-style
    ' Range object (TypeName = "Range").  Its .Value property gives a 2D array.
    ' A plain array can be used directly.
    haveData = False
    If IsArray(tableArray) Then
        arrData = tableArray
        haveData = True
    ElseIf IsObject(tableArray) Then
        If tableArray Is Nothing Then
            FuzzyVLookup = "*** TableArray is invalid ***"
            Exit Function
        End If
        If TypeName(tableArray) = "Range" Then
            ' VBA-compatible Range object — .Value returns a 2D Variant array
            arrData = tableArray.Value
            haveData = True
        Else
            ' UNO SheetCellRange — use getDataArray()
            On Error GoTo RangeArrayError
            arrData = tableArray.getDataArray()
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
    rankNum = DEFAULT_RANK
    If Not (IsMissing(rank) Or IsEmpty(rank)) Then
        If IsNumeric(rank) Then rankNum = CInt(rank)
    End If

    algo = DEFAULT_ALGORITHM
    If Not (IsMissing(algorithm) Or IsEmpty(algorithm)) Then
        If IsNumeric(algorithm) Then algo = CInt(algorithm)
    End If
    If algo < 1 Or algo > 2 Then algo = DEFAULT_ALGORITHM

    If indexNum < 0 Then
        FuzzyVLookup = "*** IndexNum must be greater than or equal to 0 ***"
        Exit Function
    End If

    If rankNum < 1 Then
        FuzzyVLookup = "*** 'Rank' must be an integer > 0 ***"
        Exit Function
    End If

    minPercent = DEFAULT_MIN_PERCENT
    If Not (IsMissing(nfPercent) Or IsEmpty(nfPercent)) Then
        If Not IsNumeric(nfPercent) Then
            FuzzyVLookup = "*** 'NFPercent' must be numeric and > 0 and <= 1 ***"
            Exit Function
        End If
        minPercent = CSng(nfPercent)
        If (minPercent <= 0) Or (minPercent > 1) Then
            FuzzyVLookup = "*** 'NFPercent' must be a percentage > 0 and <= 1 ***"
            Exit Function
        End If
    End If

    If indexNum > nCols And indexNum > 0 Then
        FuzzyVLookup = "*** IndexNum out of bounds ***"
        Exit Function
    End If

    ReDim sortedRanks(1 To rankNum)

    row = rowLB
    Do While row <= rowUB
        curValue = arrData(row, colLB)

        ' Skip blank cells rather than halting — the table may have gaps
        If Trim(CStr(curValue)) <> "" Then
            listString = LCase$(Trim(curValue))
            listString = GetNormalized(listString)
            
            ' Blocking optimization: only compare if block keys match
            ' This dramatically reduces comparisons for large datasets
            Dim tableBlockKey As String
            tableBlockKey = GetBlockKey(listString)
            
            If tableBlockKey = lookupBlockKey Then
                curPercent = FuzzyPercent(String1:=normLookup, _
                                              String2:=listString, _
                                              Algorithm:=algo, _
                                              Normalised:=True)

                If curPercent >= minPercent Then
                    ' Insert into sortedRanks
                    InsertSortedRank sortedRanks, rankNum, row, curPercent
                End If
            End If
        End If

        row = row + 1
    Loop

    If sortedRanks(rankNum).Percentage < minPercent Then
        FuzzyVLookup = CVErr(ERR_NA)
    Else
        bestMatchPtr = sortedRanks(rankNum).Offset
        If indexNum > 0 Then
            retCol = colLB + indexNum - 1
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
    Dim doc As Object
    Dim sheets As Object
    Dim sheet As Object
    Dim cell As Object
    Dim probeCell As Object
    Dim result As Variant
    Dim formula As String
    Dim argSep As String
    Dim errCode As Long
    Dim msg As String
    Const TEST_SHEET_NAME As String = "FuzzyVLookupTest"

    ' Use a dedicated test sheet — never touch the user's active sheet
    doc = ThisComponent
    sheets = doc.Sheets
    If sheets.hasByName(TEST_SHEET_NAME) Then
        sheet = sheets.getByName(TEST_SHEET_NAME)
    Else
        sheets.insertNewByName(TEST_SHEET_NAME, sheets.Count)
        sheet = sheets.getByName(TEST_SHEET_NAME)
    End If

    ' Clear text and numeric values in the test sheet
    sheet.clearContents(7)

    ' Reset caches so each test run starts fresh
    Set normCache = New Collection
    Set scoreCache = New Collection
    scoreCacheCount = 0
    cachesInitialized = True

    ' Define a small dataset in the sheet
    sheet.getCellByPosition(0, 0).String = "Name"
    sheet.getCellByPosition(1, 0).String = "Age"
    sheet.getCellByPosition(2, 0).String = "City"

    sheet.getCellByPosition(0, 1).String = "William"
    sheet.getCellByPosition(1, 1).String = "25"
    sheet.getCellByPosition(2, 1).String = "New York"

    sheet.getCellByPosition(0, 2).String = "John"
    sheet.getCellByPosition(1, 2).String = "30"
    sheet.getCellByPosition(2, 2).String = "Los Angeles"

    sheet.getCellByPosition(0, 3).String = "Anna"
    sheet.getCellByPosition(1, 3).String = "28"
    sheet.getCellByPosition(2, 3).String = "Chicago"

    sheet.getCellByPosition(0, 4).String = "Michael"
    sheet.getCellByPosition(1, 4).String = "35"
    sheet.getCellByPosition(2, 4).String = "Houston"

    ' --- Test by writing the formula to a cell and reading the result ---
    ' This simulates a real user call, which is necessary to get the
    ' special VBA-style Range object that FuzzyVLookup expects.
    probeCell = sheet.getCellByPosition(6, 0) ' G1 (test output)
    cell = sheet.getCellByPosition(5, 0) ' F1 (formula cell)

    ' Detect Calc argument separator for locale-agnostic formula construction.
    argSep = ";"
    probeCell.setFormula("=SUM(1;2)")
    errCode = probeCell.getError()
    If errCode <> 0 Then argSep = ","

    formula = "=FUZZYVLOOKUP(""Willam""" & argSep & " A2:C5" & argSep & " 2" & argSep & " 0.5" & argSep & " 1" & argSep & " 1)"
    cell.setFormula(formula)

    ' Read the result from the cell
    result = cell.getString()
    If result = "" Then result = cell.getValue()

    ' Display result in test sheet cell instead of MsgBox (headless-friendly)
    If cell.getError() <> 0 Then
        msg = "Test failed. Formula returned error code: " & cell.getError()
        probeCell.String = msg
    Else
        msg = "Match found for 'Willam': " & result
        probeCell.String = msg
    End If
End Sub