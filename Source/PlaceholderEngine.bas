Attribute VB_Name = "PlaceholderEngine"
Option Explicit

Private Const PLACEHOLDER_PREFIX As String = "<<"
Private Const PLACEHOLDER_SUFFIX As String = ">>"

' Replaces one placeholder throughout the Word document body.
Public Function ReplacePlaceholder(ByVal Doc As Object, ByVal Placeholder As String, ByVal Value As String) As Boolean
    Dim searchText As String
    Dim replacementCount As Long

    On Error GoTo ErrorHandler

    If Doc Is Nothing Then
        WriteLog "Placeholder replacement skipped because document is not available."
        ReplacePlaceholder = False
        Exit Function
    End If

    searchText = NormalizePlaceholder(Placeholder)

    If Len(searchText) = 0 Then
        WriteLog "Placeholder replacement skipped because placeholder is blank."
        ReplacePlaceholder = False
        Exit Function
    End If

    replacementCount = ReplaceTextInBody(Doc, searchText, Value)
    WriteLog "Placeholder replaced: " & searchText & " -> " & CStr(replacementCount) & " occurrence(s)."

    ReplacePlaceholder = True
    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.ReplacePlaceholder", Err.Number, Err.Description
    ReplacePlaceholder = False
End Function

' Replaces placeholders using Excel headers from row 1 and values from one data row.
Public Function ReplaceAllPlaceholders(ByVal Doc As Object, ByVal Headers As Variant, ByVal RowValues As Variant) As Boolean
    Dim itemCount As Long
    Dim itemIndex As Long
    Dim headerText As String
    Dim valueText As String

    On Error GoTo ErrorHandler

    If Doc Is Nothing Then
        WriteLog "Bulk placeholder replacement skipped because document is not available."
        ReplaceAllPlaceholders = False
        Exit Function
    End If

    itemCount = GetPairCount(Headers, RowValues)

    If itemCount = 0 Then
        WriteLog "Bulk placeholder replacement skipped because headers or row values are empty."
        ReplaceAllPlaceholders = False
        Exit Function
    End If

    For itemIndex = 1 To itemCount
        headerText = ValueToText(GetIndexedValue(Headers, itemIndex))
        valueText = ValueToText(GetIndexedValue(RowValues, itemIndex))

        If Len(headerText) > 0 Then
            If Not ReplacePlaceholder(Doc, headerText, valueText) Then
                ReplaceAllPlaceholders = False
                Exit Function
            End If
        End If
    Next itemIndex

    WriteLog "Bulk placeholder replacement completed for " & CStr(itemCount) & " field(s)."
    ReplaceAllPlaceholders = True
    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.ReplaceAllPlaceholders", Err.Number, Err.Description
    ReplaceAllPlaceholders = False
End Function

' Compatibility helper for existing dictionary-based workflow code.
Public Function ReplacePlaceholders(ByVal Doc As Object, ByVal Data As Object) As Boolean
    Dim key As Variant

    On Error GoTo ErrorHandler

    If Doc Is Nothing Or Data Is Nothing Then
        WriteLog "Dictionary placeholder replacement skipped because document or data is not available."
        ReplacePlaceholders = False
        Exit Function
    End If

    For Each key In Data.Keys
        If Not ReplacePlaceholder(Doc, CStr(key), ValueToText(Data(key))) Then
            ReplacePlaceholders = False
            Exit Function
        End If
    Next key

    WriteLog "Dictionary placeholder replacement completed."
    ReplacePlaceholders = True
    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.ReplacePlaceholders", Err.Number, Err.Description
    ReplacePlaceholders = False
End Function

' Returns True when a placeholder exists in the Word document body.
Public Function PlaceholderExists(ByVal Doc As Object, ByVal Placeholder As String) As Boolean
    Dim searchText As String

    On Error GoTo ErrorHandler

    If Doc Is Nothing Then
        PlaceholderExists = False
        Exit Function
    End If

    searchText = NormalizePlaceholder(Placeholder)

    If Len(searchText) = 0 Then
        PlaceholderExists = False
        Exit Function
    End If

    PlaceholderExists = TextExistsInBody(Doc, searchText)
    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.PlaceholderExists", Err.Number, Err.Description
    PlaceholderExists = False
End Function

' Counts supported placeholders in the Word document body.
Public Function GetPlaceholderCount(ByVal Doc As Object) As Long
    Dim bodyText As String
    Dim searchStart As Long
    Dim openPosition As Long
    Dim closePosition As Long
    Dim placeholderCount As Long

    On Error GoTo ErrorHandler

    If Doc Is Nothing Then
        GetPlaceholderCount = 0
        Exit Function
    End If

    bodyText = CStr(Doc.Content.Text)
    searchStart = 1

    Do
        openPosition = InStr(searchStart, bodyText, PLACEHOLDER_PREFIX, vbTextCompare)

        If openPosition = 0 Then
            Exit Do
        End If

        closePosition = InStr(openPosition + Len(PLACEHOLDER_PREFIX), bodyText, PLACEHOLDER_SUFFIX, vbTextCompare)

        If closePosition = 0 Then
            Exit Do
        End If

        placeholderCount = placeholderCount + 1
        searchStart = closePosition + Len(PLACEHOLDER_SUFFIX)
    Loop

    GetPlaceholderCount = placeholderCount
    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.GetPlaceholderCount", Err.Number, Err.Description
    GetPlaceholderCount = 0
End Function

' Validates that the supplied Word document body contains at least one supported placeholder.
Public Function ValidateTemplate(ByVal Doc As Object) As Boolean
    On Error GoTo ErrorHandler

    ValidateTemplate = (GetPlaceholderCount(Doc) > 0)
    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.ValidateTemplate", Err.Number, Err.Description
    ValidateTemplate = False
End Function

Private Function ReplaceTextInBody(ByVal Doc As Object, ByVal SearchText As String, ByVal ReplacementText As String) As Long
    Dim targetRange As Object
    Dim replacementCount As Long

    On Error GoTo ErrorHandler

    Set targetRange = Doc.Content.Duplicate

    With targetRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = SearchText
        .Replacement.Text = ReplacementText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute(Replace:=wdReplaceOne)
            replacementCount = replacementCount + 1
            targetRange.Collapse wdCollapseEnd
            targetRange.End = Doc.Content.End
        Loop
    End With

    ReplaceTextInBody = replacementCount
    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.ReplaceTextInBody", Err.Number, Err.Description
    ReplaceTextInBody = 0
End Function

Private Function TextExistsInBody(ByVal Doc As Object, ByVal SearchText As String) As Boolean
    Dim targetRange As Object

    On Error GoTo ErrorHandler

    Set targetRange = Doc.Content.Duplicate

    With targetRange.Find
        .ClearFormatting
        .Text = SearchText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        TextExistsInBody = .Execute
    End With

    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.TextExistsInBody", Err.Number, Err.Description
    TextExistsInBody = False
End Function

Private Function NormalizePlaceholder(ByVal Placeholder As String) As String
    On Error GoTo ErrorHandler

    Placeholder = Trim$(Placeholder)

    If Len(Placeholder) = 0 Then
        NormalizePlaceholder = vbNullString
    ElseIf Left$(Placeholder, Len(PLACEHOLDER_PREFIX)) = PLACEHOLDER_PREFIX _
            And Right$(Placeholder, Len(PLACEHOLDER_SUFFIX)) = PLACEHOLDER_SUFFIX Then
        NormalizePlaceholder = Placeholder
    Else
        NormalizePlaceholder = PLACEHOLDER_PREFIX & Placeholder & PLACEHOLDER_SUFFIX
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.NormalizePlaceholder", Err.Number, Err.Description
    NormalizePlaceholder = vbNullString
End Function

Private Function GetPairCount(ByVal Headers As Variant, ByVal RowValues As Variant) As Long
    Dim headerCount As Long
    Dim valueCount As Long

    On Error GoTo ErrorHandler

    headerCount = GetValueCount(Headers)
    valueCount = GetValueCount(RowValues)

    If headerCount < valueCount Then
        GetPairCount = headerCount
    Else
        GetPairCount = valueCount
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.GetPairCount", Err.Number, Err.Description
    GetPairCount = 0
End Function

Private Function GetValueCount(ByVal Values As Variant) As Long
    On Error GoTo TryOneDimensionalArray

    If IsArray(Values) Then
        GetValueCount = UBound(Values, 2) - LBound(Values, 2) + 1
        Exit Function
    End If

TryOneDimensionalArray:
    On Error GoTo TryCollection

    If IsArray(Values) Then
        GetValueCount = UBound(Values) - LBound(Values) + 1
        Exit Function
    End If

TryCollection:
    On Error GoTo ErrorHandler
    GetValueCount = Values.Count
    Exit Function

ErrorHandler:
    GetValueCount = 0
End Function

Private Function GetIndexedValue(ByVal Values As Variant, ByVal ItemIndex As Long) As Variant
    On Error GoTo TryOneDimensionalArray

    If IsArray(Values) Then
        GetIndexedValue = Values(LBound(Values, 1), LBound(Values, 2) + ItemIndex - 1)
        Exit Function
    End If

TryOneDimensionalArray:
    On Error GoTo TryCollection

    If IsArray(Values) Then
        GetIndexedValue = Values(LBound(Values) + ItemIndex - 1)
        Exit Function
    End If

TryCollection:
    On Error GoTo ErrorHandler
    GetIndexedValue = Values(ItemIndex)
    Exit Function

ErrorHandler:
    GetIndexedValue = Empty
End Function

Private Function ValueToText(ByVal Value As Variant) As String
    On Error GoTo ErrorHandler

    If IsError(Value) Or IsNull(Value) Or IsEmpty(Value) Then
        ValueToText = vbNullString
    Else
        ValueToText = CStr(Value)
    End If

    Exit Function

ErrorHandler:
    WriteErrorLog "PlaceholderEngine.ValueToText", Err.Number, Err.Description
    ValueToText = vbNullString
End Function
