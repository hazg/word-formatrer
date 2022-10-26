Attribute VB_Name = "functions"


Private Function ProcessImages()

    Application.ScreenUpdating = False

    Dim objShape As Shape

    For Each objShape In ActiveDocument.Shapes
        If objShape.Type = msoPicture Then
            objShape.WrapFormat.Type = wdWrapTopBottom
        End If
    Next objShape
    
    Application.ScreenUpdating = True

End Function

Private Sub ParagraphsWalker()

    ' ”дал€ет все концы абзацев Ђ^pї, кроме следующих случаев:

    ' Ц если абзац заканчиваетс€ точкой, двоеточием, восклицательным или вопросительным знаком;
    ' Ц если следующий абзац начинаетс€ с любым количеством пробелов или табул€ций (в т.ч. нулевым) после которых идет цифра и точка (например, Ђ1.ї), или цифра и скобка (напр. Ђ1)ї), или знак минуса, дефиса или тире Ђ-ї, Ђ?ї, ЂЧї или буллет ЂХї;
    ' Ц если следующий абзац начинаетс€ с табул€ции, или с 3-х, 4-х, 5-ти, 6-ти, 7-ми, 8-ми, 9-ти, или 10-ти пробелов (то есть с аналога табул€ции).
    
    ' Application.ScreenUpdating = False
    Dim sBegin As String
    Dim sEnd As String
    Dim sKeep As String
    
    sBegin = "BEGIN"
    sEnd = "END"
    sKeep = "KEEP"
    
    Dim i As Integer
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        If (i <= ActiveDocument.Paragraphs.Count) Then
            With ActiveDocument.Paragraphs(i)
                
                Dim withDot As Boolean
                Dim nextWithTabs As Boolean
                Dim nextWithSpaces As Boolean
                
                .Range.Text = sBegin & .Range.Text
                
                withDot = TestString("[\.][^13]{1}", ActiveDocument.Paragraphs(i), True)
                
                
                If i < ActiveDocument.Paragraphs.Count Then
                    
                    ' *\d[\)\.\-\?ЧХ]
                    nextWithTabs = TestString(sBegin & "[^s ]*[!^s ]", ActiveDocument.Paragraphs(i + 1))
                    nextWithSpaces = TestString(sBegin & "[ ^t^s]@", ActiveDocument.Paragraphs(i + 1))
                    
                    ' TODO - REMOVE
                    
                    nextWithTabs = False
                    nextWithSpaces = False
                Else
                    Debug.Print ("Nothing")
                    nextWithTabs = False
                    nextWithSpaces = False
                End If
                
                Dim keep As Boolean
                
                keep = withDot Or nextWithTabs Or nextWithSpaces
                Debug.Print ("KEEP? " & keep)
                If keep Then
                    Debug.Print ("REPLACE " & ActiveDocument.Paragraphs(i).Range)
                    ReplaceString "^p", sKeep & Chr(13), False, ActiveDocument.Paragraphs(i)
                    ' i = ActiveDocument.Paragraphs.Count
                End If
                ' nextWithTabs =
                Debug.Print (i & " from " & ActiveDocument.Paragraphs.Count)
                Debug.Print ("dot: " & withDot & ", tabs: " & nextWithTabs & ", spaces: " & nextWithSpaces)
                
            End With
        End If
    Next i
    ReplaceString sBegin, "", False
    
    ' ReplaceString "^p", "", False
    ' ReplaceString "KEEP", "^13", False
    
    ' ReplaceString "[^13]@" & sEnd, "^13"
    ' ReplaceString sEnd, "", False
    
    Application.ScreenUpdating = True
    
End Sub

' ### FUNCTIONS ###


Private Function TestString(pattern As String, ByRef par As Paragraph, Optional wildcards As Boolean = True) As Boolean
    
    Debug.Print ("Search '" & pattern & "' in " & par.Range)
    With par.Range.Find
        .forward = True
        .MatchWildcards = wildcards
        TestString = .Execute(pattern)
        
        ' Debug.Print ("##### " & TestString)
    End With
    
    Selection.Collapse Direction:=wdCollapseStart
    
End Function


Private Function CorrectImagesMargins(margin As Integer)
    
    Application.ScreenUpdating = False
    
    With ActiveDocument.PageSetup
        .LeftMargin = Application.CentimetersToPoints(margin)
        .RightMargin = Application.CentimetersToPoints(margin)
        .TopMargin = Application.CentimetersToPoints(margin)
        .BottomMargin = Application.CentimetersToPoints(margin)
    End With
    
    Application.ScreenUpdating = True
    
End Function
