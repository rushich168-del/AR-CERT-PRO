Attribute VB_Name = "PlaceholderEngine"
Option Explicit

Private Const PLACEHOLDER_PREFIX As String = "<<"
Private Const PLACEHOLDER_SUFFIX As String = ">>"
Private Const PLACEHOLDER_WILDCARD_PATTERN As String = "\<\<[!\<\>]@\>\>"
Private Const SHAPE_TYPE_GROUP As Long = 6

' Replaces all placeholders in the supplied Word document using values from a Dictionary.
' Each dictionary key maps to a placeholder in the format <<Key>>.
Public Function ReplacePlaceholders(ByVal document As Document, ByVal data As Object) As Boolean
    Dim key As Variant

    On Error GoTo ErrorHandler

    If document Is Nothing Or data Is Nothing Then
        ReplacePlaceholders = False
        Exit Function
    End If

    For Each key In data.Keys
        ReplacePlaceholder document, CStr(key), ValueToText(data(key))
    Next key

    ReplacePlaceholders = True
    Exit Function

ErrorHandler:
    ReplacePlaceholders = False
End Function

' Replaces a single placeholder throughout the supplied Word document.
' The placeholder argument can be provided as Name or <<Name>>.
Public Sub ReplacePlaceholder(ByVal document As Document, ByVal placeholder As String, ByVal replacement As String)
    Dim searchText As String

    On Error GoTo ErrorHandler

    If document Is Nothing Then
        Exit Sub
    End If

    searchText = NormalizePlaceholder(placeholder)

    If Len(searchText) = 0 Then
        Exit Sub
    End If

    ReplaceInAllStoryRanges document, searchText, replacement
    ReplaceInAllShapes document, searchText, replacement
    Exit Sub

ErrorHandler:
End Sub

' Returns True when a placeholder exists anywhere in the supplied Word document.
' The placeholder argument can be provided as Name or <<Name>>.
Public Function PlaceholderExists(ByVal document As Document, ByVal placeholder As String) As Boolean
    Dim searchText As String

    On Error GoTo ErrorHandler

    If document Is Nothing Then
        PlaceholderExists = False
        Exit Function
    End If

    searchText = NormalizePlaceholder(placeholder)

    If Len(searchText) = 0 Then
        PlaceholderExists = False
        Exit Function
    End If

    PlaceholderExists = TextExistsInAllStories(document, searchText) Or TextExistsInAllShapes(document, searchText)
    Exit Function

ErrorHandler:
    PlaceholderExists = False
End Function

' Counts placeholders matching the supported <<Name>> format across the document.
Public Function GetPlaceholderCount(ByVal document As Document) As Long
    On Error GoTo ErrorHandler

    If document Is Nothing Then
        GetPlaceholderCount = 0
        Exit Function
    End If

    GetPlaceholderCount = CountPlaceholdersInAllStories(document) + CountPlaceholdersInAllShapes(document)
    Exit Function

ErrorHandler:
    GetPlaceholderCount = 0
End Function

' Validates that the supplied Word template contains at least one supported placeholder.
Public Function ValidateTemplate(ByVal document As Document) As Boolean
    On Error GoTo ErrorHandler

    ValidateTemplate = (GetPlaceholderCount(document) > 0)
    Exit Function

ErrorHandler:
    ValidateTemplate = False
End Function

Private Sub ReplaceInAllStoryRanges(ByVal document As Document, ByVal searchText As String, ByVal replacement As String)
    Dim storyRange As Range
    Dim currentRange As Range

    On Error GoTo ErrorHandler

    For Each storyRange In document.StoryRanges
        Set currentRange = storyRange

        Do While Not currentRange Is Nothing
            ReplaceTextInRange currentRange, searchText, replacement
            Set currentRange = currentRange.NextStoryRange
        Loop
    Next storyRange

ErrorHandler:
End Sub

Private Sub ReplaceTextInRange(ByVal targetRange As Range, ByVal searchText As String, ByVal replacement As String)
    On Error GoTo ErrorHandler

    With targetRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = searchText
        .Replacement.Text = replacement
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With

ErrorHandler:
End Sub

Private Function TextExistsInAllStories(ByVal document As Document, ByVal searchText As String) As Boolean
    Dim storyRange As Range
    Dim currentRange As Range

    On Error GoTo ErrorHandler

    For Each storyRange In document.StoryRanges
        Set currentRange = storyRange

        Do While Not currentRange Is Nothing
            If TextExistsInRange(currentRange, searchText) Then
                TextExistsInAllStories = True
                Exit Function
            End If

            Set currentRange = currentRange.NextStoryRange
        Loop
    Next storyRange

    TextExistsInAllStories = False
    Exit Function

ErrorHandler:
    TextExistsInAllStories = False
End Function

Private Function TextExistsInRange(ByVal targetRange As Range, ByVal searchText As String) As Boolean
    Dim searchRange As Range

    On Error GoTo ErrorHandler

    Set searchRange = targetRange.Duplicate

    With searchRange.Find
        .ClearFormatting
        .Text = searchText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        TextExistsInRange = .Execute
    End With

    Exit Function

ErrorHandler:
    TextExistsInRange = False
End Function

Private Function CountPlaceholdersInAllStories(ByVal document As Document) As Long
    Dim storyRange As Range
    Dim currentRange As Range
    Dim placeholderCount As Long

    On Error GoTo ErrorHandler

    For Each storyRange In document.StoryRanges
        Set currentRange = storyRange

        Do While Not currentRange Is Nothing
            If currentRange.StoryType <> wdTextFrameStory Then
                placeholderCount = placeholderCount + CountPlaceholdersInRange(currentRange)
            End If

            Set currentRange = currentRange.NextStoryRange
        Loop
    Next storyRange

    CountPlaceholdersInAllStories = placeholderCount
    Exit Function

ErrorHandler:
    CountPlaceholdersInAllStories = 0
End Function

Private Function CountPlaceholdersInRange(ByVal targetRange As Range) As Long
    Dim searchRange As Range
    Dim placeholderCount As Long

    On Error GoTo ErrorHandler

    Set searchRange = targetRange.Duplicate

    With searchRange.Find
        .ClearFormatting
        .Text = PLACEHOLDER_WILDCARD_PATTERN
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = True
        .MatchSoundsLike = False
        .MatchAllWordForms = False

        Do While .Execute
            placeholderCount = placeholderCount + 1
            searchRange.Collapse wdCollapseEnd
        Loop
    End With

    CountPlaceholdersInRange = placeholderCount
    Exit Function

ErrorHandler:
    CountPlaceholdersInRange = 0
End Function

Private Sub ReplaceInAllShapes(ByVal document As Document, ByVal searchText As String, ByVal replacement As String)
    Dim shape As Shape
    Dim section As Section
    Dim headerFooter As HeaderFooter

    On Error Resume Next

    For Each shape In document.Shapes
        ReplaceInShape shape, searchText, replacement
    Next shape

    For Each section In document.Sections
        For Each headerFooter In section.Headers
            ReplaceInShapesCollection headerFooter.Shapes, searchText, replacement
        Next headerFooter

        For Each headerFooter In section.Footers
            ReplaceInShapesCollection headerFooter.Shapes, searchText, replacement
        Next headerFooter
    Next section
End Sub

Private Sub ReplaceInShapesCollection(ByVal shapes As Shapes, ByVal searchText As String, ByVal replacement As String)
    Dim shape As Shape

    On Error Resume Next

    For Each shape In shapes
        ReplaceInShape shape, searchText, replacement
    Next shape
End Sub

Private Sub ReplaceInShape(ByVal shape As Shape, ByVal searchText As String, ByVal replacement As String)
    Dim groupItem As Shape

    On Error Resume Next

    If shape.Type = SHAPE_TYPE_GROUP Then
        For Each groupItem In shape.GroupItems
            ReplaceInShape groupItem, searchText, replacement
        Next groupItem
    End If

    If ShapeHasText(shape) Then
        ReplaceTextInRange shape.TextFrame.TextRange, searchText, replacement
    End If
End Sub

Private Function TextExistsInAllShapes(ByVal document As Document, ByVal searchText As String) As Boolean
    Dim shape As Shape
    Dim section As Section
    Dim headerFooter As HeaderFooter

    On Error Resume Next

    For Each shape In document.Shapes
        If TextExistsInShape(shape, searchText) Then
            TextExistsInAllShapes = True
            Exit Function
        End If
    Next shape

    For Each section In document.Sections
        For Each headerFooter In section.Headers
            If TextExistsInShapesCollection(headerFooter.Shapes, searchText) Then
                TextExistsInAllShapes = True
                Exit Function
            End If
        Next headerFooter

        For Each headerFooter In section.Footers
            If TextExistsInShapesCollection(headerFooter.Shapes, searchText) Then
                TextExistsInAllShapes = True
                Exit Function
            End If
        Next headerFooter
    Next section

    TextExistsInAllShapes = False
End Function

Private Function TextExistsInShapesCollection(ByVal shapes As Shapes, ByVal searchText As String) As Boolean
    Dim shape As Shape

    On Error Resume Next

    For Each shape In shapes
        If TextExistsInShape(shape, searchText) Then
            TextExistsInShapesCollection = True
            Exit Function
        End If
    Next shape

    TextExistsInShapesCollection = False
End Function

Private Function TextExistsInShape(ByVal shape As Shape, ByVal searchText As String) As Boolean
    Dim groupItem As Shape

    On Error Resume Next

    If shape.Type = SHAPE_TYPE_GROUP Then
        For Each groupItem In shape.GroupItems
            If TextExistsInShape(groupItem, searchText) Then
                TextExistsInShape = True
                Exit Function
            End If
        Next groupItem
    End If

    If ShapeHasText(shape) Then
        TextExistsInShape = TextExistsInRange(shape.TextFrame.TextRange, searchText)
    End If
End Function

Private Function CountPlaceholdersInAllShapes(ByVal document As Document) As Long
    Dim shape As Shape
    Dim section As Section
    Dim headerFooter As HeaderFooter
    Dim placeholderCount As Long

    On Error Resume Next

    For Each shape In document.Shapes
        placeholderCount = placeholderCount + CountPlaceholdersInShape(shape)
    Next shape

    For Each section In document.Sections
        For Each headerFooter In section.Headers
            placeholderCount = placeholderCount + CountPlaceholdersInShapesCollection(headerFooter.Shapes)
        Next headerFooter

        For Each headerFooter In section.Footers
            placeholderCount = placeholderCount + CountPlaceholdersInShapesCollection(headerFooter.Shapes)
        Next headerFooter
    Next section

    CountPlaceholdersInAllShapes = placeholderCount
End Function

Private Function CountPlaceholdersInShapesCollection(ByVal shapes As Shapes) As Long
    Dim shape As Shape
    Dim placeholderCount As Long

    On Error Resume Next

    For Each shape In shapes
        placeholderCount = placeholderCount + CountPlaceholdersInShape(shape)
    Next shape

    CountPlaceholdersInShapesCollection = placeholderCount
End Function

Private Function CountPlaceholdersInShape(ByVal shape As Shape) As Long
    Dim groupItem As Shape
    Dim placeholderCount As Long

    On Error Resume Next

    If shape.Type = SHAPE_TYPE_GROUP Then
        For Each groupItem In shape.GroupItems
            placeholderCount = placeholderCount + CountPlaceholdersInShape(groupItem)
        Next groupItem
    End If

    If ShapeHasText(shape) Then
        placeholderCount = placeholderCount + CountPlaceholdersInRange(shape.TextFrame.TextRange)
    End If

    CountPlaceholdersInShape = placeholderCount
End Function

Private Function ShapeHasText(ByVal shape As Shape) As Boolean
    On Error GoTo ErrorHandler

    ShapeHasText = (shape.TextFrame.HasText <> 0)
    Exit Function

ErrorHandler:
    ShapeHasText = False
End Function

Private Function NormalizePlaceholder(ByVal placeholder As String) As String
    placeholder = Trim$(placeholder)

    If Len(placeholder) = 0 Then
        NormalizePlaceholder = vbNullString
    ElseIf Left$(placeholder, Len(PLACEHOLDER_PREFIX)) = PLACEHOLDER_PREFIX _
            And Right$(placeholder, Len(PLACEHOLDER_SUFFIX)) = PLACEHOLDER_SUFFIX Then
        NormalizePlaceholder = placeholder
    Else
        NormalizePlaceholder = PLACEHOLDER_PREFIX & placeholder & PLACEHOLDER_SUFFIX
    End If
End Function

Private Function ValueToText(ByVal value As Variant) As String
    On Error GoTo ErrorHandler

    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then
        ValueToText = vbNullString
    Else
        ValueToText = CStr(value)
    End If

    Exit Function

ErrorHandler:
    ValueToText = vbNullString
End Function
