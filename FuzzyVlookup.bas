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
                     Optional Algorithm As Integer, _
                     Optional Normalised As Boolean) As Single
    Dim intLen1 As Integer, intLen2 As Integer
    Dim intCurLen As Integer
    Dim intTo As Integer
    Dim intPos As Integer
    Dim intPtr As Integer
    Dim intScore As Integer
    Dim intTotScore As Integer
    Dim intStartPos As Integer
    Dim strWork As String

    ' Assign default values manually
    If IsMissing(Algorithm) Then Algorithm = 3
    If IsMissing(Normalised) Then Normalised = FALSE

    '-------------------------------------------------------
    '-- If strings haven't been normalised, normalise them --
    '-------------------------------------------------------
    If Normalised = FALSE Then
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
    If intLen1 < 2 Then
        FuzzyPercent = 0
        Exit Function
    End If

    intTotScore = 0        'initialise total possible score
    intScore = 0           'initialise current score

    '--------------------------------------------------------
    '-- If Algorithm = 1 or 3, Search for single characters --
    '--------------------------------------------------------
    If (Algorithm And 1) <> 0 Then
        FuzzyAlg1 String1, String2, intScore, intTotScore
        If intLen1 < intLen2 Then FuzzyAlg1 String2, String1, intScore, intTotScore
    End If

    '-----------------------------------------------------------
    '-- If Algorithm = 2 or 3, Search for pairs, triplets etc. --
    '-----------------------------------------------------------
    If (Algorithm And 2)<> 0 Then
    	FuzzyAlg2 String1, String2, intScore, intTotScore
        If intLen1 < intLen2 Then FuzzyAlg2 String2, String1, intScore, intTotScore
    End If

    If intTotScore > 0 Then
        FuzzyPercent = intScore / intTotScore
    Else
        FuzzyPercent = 0
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
        ReDim corruptPositions(1 To Len(strWork)) ' Reset corrupted positions

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
                     ByVal TableArray As CellRange, _
                     ByVal IndexNum As Integer, _
                     Optional NFPercent As Single, _
                     Optional Rank As Integer, _
                     Optional Algorithm As Integer, _
                     Optional AdditionalCols As Integer) As Variant
    On Error GoTo ErrorHandler

    Dim oSheet As Object
    Dim oCell As Object
    Dim lEndRow As Long
    Dim lRow As Long
    Dim lCol as Long
    Dim sngMinPercent As Single
    Dim sngCurPercent As Single
    Dim intBestMatchPtr As Long
    Dim sortedRanks() As RankInfo
    Dim strListString as String
    Dim vCurValue As Variant
    Dim lastCol As Long

    LookupValue = LCase$(Trim(LookupValue))
    oSheet = ThisComponent.CurrentController.ActiveSheet

    ' Parameter validation
    If TableArray Is Nothing Then
        FuzzyVLookup = "*** TableArray is invalid ***"
        Exit Function
    End If
	 
    If TypeName(TableArray) <> "SheetCellRange" Then
        MsgBox TypeName(TableArray)
        FuzzyVLookup = "*** TableArray must be a CellRange ***"
        Exit Function
    End If

    If IndexNum < 0 Then
        FuzzyVLookup = "*** IndexNum must be greater than or equal to 0 ***"
        Exit Function
    End If

    If Rank < 1 Then
        FuzzyVLookup = "*** 'Rank' must be an integer > 0 ***"
        Exit Function
    End If

    If IsMissing(NFPercent) Then
        sngMinPercent = 0.05
    Else
        If (NFPercent <= 0) Or (NFPercent > 1) Then
            FuzzyVLookup = "*** 'NFPercent' must be a percentage > 0 and <= 1 ***"            
            Exit Function
        End If
        sngMinPercent = NFPercent
    End If

    'Find the last column of the table
    ' Set TableArray = oSheet.getCellRangeByName(TableArray)
    lastCol = TableArray.RangeAddress.EndColumn

    If IndexNum > (lastCol - TableArray.RangeAddress.StartColumn + 1) And IndexNum > 0 Then
        FuzzyVLookup = "*** IndexNum out of bounds ***"
        Exit Function
    End If
    'End validation.

    ReDim sortedRanks(1 To Rank)

    lEndRow = TableArray.RangeAddress.EndRow
    lRow = TableArray.RangeAddress.StartRow
    lCol = TableArray.RangeAddress.StartColumn

    Do While lRow <= lEndRow
        oCell = oSheet.getCellByPosition(lCol, lRow)
        vCurValue = oCell.String
        If vCurValue = "" Then Exit Do

        strListString = LCase$(Trim(vCurValue))

        sngCurPercent = FuzzyPercent(String1:=LookupValue, _
                                      String2:=strListString, _
                                      Algorithm:=Algorithm, _
                                      Normalised:=True)

        If sngCurPercent >= sngMinPercent Then
            ' Insert into sortedRanks using binary search
            InsertSortedRank sortedRanks, Rank, lRow, sngCurPercent
        End If

        lRow = lRow + 1
    Loop

    If sortedRanks(Rank).Percentage < sngMinPercent Then
        FuzzyVLookup = CVErr(2042)
    Else
        intBestMatchPtr = sortedRanks(Rank).Offset
        If IndexNum > 0 Then
            If lCol + IndexNum - 1 <= oSheet.Columns.Count Then
                FuzzyVLookup = oSheet.getCellByPosition(lCol + IndexNum - 1, intBestMatchPtr).String
            Else
                FuzzyVLookup = "*** IndexNum out of bounds ***"
            End If
        Else
            FuzzyVLookup = intBestMatchPtr - TableArray.RangeAddress.StartRow + 1
        End If
    End If

    Exit Function
    ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation, "FuzzyVLookup Error"
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
    Dim oSheet As Object
    Dim oCell As Object
    Dim vResult As Variant
    Dim LookupValue As String
    Dim TableArray As CellRange
    Dim IndexNum As Integer
    Dim NFPercent As Single
    Dim Rank As Integer
    Dim Algorithm As Integer
    Dim AdditionalCols As Integer
    Dim msg As String

    ' Get the active sheet
    oSheet = ThisComponent.CurrentController.ActiveSheet

    ' Clear the sheet (optional)
    oSheet.clearContents(0)

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
    AdditionalCols = 0 ' No additional columns

    ' Call the FuzzyVLookup function
    vResult = FuzzyVLookup(LookupValue, TableArray, IndexNum, NFPercent, Rank, Algorithm, AdditionalCols)

    ' Display the result
    If IsError(vResult) Then
        msg = "No match found for '" & LookupValue & "'."
    Else
        msg = "Match found for '" & LookupValue & "': " & vResult
    End If

    MsgBox msg, 0, "FuzzyVLookup Test Result"
End Sub

