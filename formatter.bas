Attribute VB_Name = "formatter"
Option Explicit

' ������ ���� �������
Sub �������������_��������()

    ������_�������������_��������_��_����
    ��������_��������_�����_�������_����������
    ��������_��������_�����_������
    ����������_�������������_��������
    ��������_��������_�_������_�_�_�����_������
    �����������_�������
    ����������_������_�����_������_�����_��_�����
    ��������_������_�����
    ��������_������_���������
    ������_��_�
    �������_�������_��������
    ���������_�����
    ������_������_��_�����������
    ��������_�������_������
    ������_�������_��_����
    ������_������_�����_������_��_�����
    �����������_������_�����_���_�_��
    ����������_�����������_��������_�����_����������
    ����������_�����������_��������_�_��������
    ����������_������������_�������_�����_��������
    ���������_������_���������
    �����������_���������_�������
    �������_�_�����_��������������
    �������_�_������_����������_��������_�����
    �������_�_������������_��_������_�_�����
    �������_�_�����_�����_�_�������
    �������_�_������������
    ��������_��������_�����_�_������_�������������_��_pdf

End Sub

' 1
Sub ������_�������������_��������_��_����()
Attribute ������_�������������_��������_��_����.VB_Description = "������ ������������� �������� �� ���� ������"
Attribute ������_�������������_��������_��_����.VB_ProcData.VB_Invoke_Func = "Project.formatter.������_�������������_��������_��_����"

    ' �������� ������� � ������������� ������� �� ���� ������
    ReplaceString "([^s ])@[^s ]", "\1"

End Sub

' 2
Sub ��������_��������_�����_�������_����������()
Attribute ��������_��������_�����_�������_����������.VB_Description = "�������� �������� ����� ������� ���������� �. , ; : ! ?�"
Attribute ��������_��������_�����_�������_����������.VB_ProcData.VB_Invoke_Func = "Project.formatter.��������_��������_�����_�������_����������"

    ' ������� ������� ����� ������� ���������� �. , ; : ! ?�
    ReplaceString "([^s ])@([^s .,;:])", "\2"

End Sub

' 3
Sub ��������_��������_�����_������()
Attribute ��������_��������_�����_������.VB_Description = "������� ������� ����� ����������� � ����� ����������� �������� �( ) {} []�"
Attribute ��������_��������_�����_������.VB_ProcData.VB_Invoke_Func = "Project.formatter.��������_��������_�����_������"

    ' ������� ������� ����� ������������� ������ � ����� ����������� �������
    ReplaceString "\([^s ]@([!^s ])", "(\1"
    ReplaceString "[^s ]@\)", ")"

    ReplaceString "\{[^s ]@([!^s ])", "{\1"
    ReplaceString "[^s ]@\}", "}"

    ReplaceString "\[[^s ]@([!^s ])", "[\1"
    ReplaceString "[^s ]@\]", "]"

End Sub

' 4
Sub ����������_�������������_��������()
Attribute ����������_�������������_��������.VB_Description = "�������� ������������� ������� �� ����"
Attribute ����������_�������������_��������.VB_ProcData.VB_Invoke_Func = "Project.formatter.����������_�������������_��������"

    ' �������� ������������� ������� �� ����
    ReplaceString "[^s ]@([!^s ])", " \1"

End Sub


' 5
Sub ��������_��������_�_������_�_�_�����_������()
Attribute ��������_��������_�_������_�_�_�����_������.VB_Description = "�������� �������� � ������ � � ����� ������ (������ ������� �^p�)"
Attribute ��������_��������_�_������_�_�_�����_������.VB_ProcData.VB_Invoke_Func = "Project.formatter.��������_��������_�_������_�_�_�����_������"

    ' �������� �������� � ������ � � ����� ������ (������ ������� �^p�)
    ReplaceString "^13[ ^s]@([!^s ])", "^13\1"
    ReplaceString " " + vbCr, vbCr

End Sub

' 6
Sub �����������_�������()
Attribute �����������_�������.VB_Description = "��������� �������� ������� ""�"" �� �x�"
Attribute �����������_�������.VB_ProcData.VB_Invoke_Func = "Project.formatter.�����������_�������"

    ' ��������� �������� ������� "�" �� ���
    ReplaceString """", """", False
    ReplaceString Chr(147), "�", False
    ReplaceString Chr(148), "�", False

End Sub

' 7
Sub ����������_������_�����_������_�����_��_�����()
Attribute ����������_������_�����_������_�����_��_�����.VB_Description = "�������� ���, ��� � �.�. ������ ������ ������ ������ �� ���� ������ ������ "
Attribute ����������_������_�����_������_�����_��_�����.VB_ProcData.VB_Invoke_Func = "Project.formatter.����������_������_�����_������_�����_��_�����"

    ' �������� ���, ��� � �.�. ������ ������ ������ ������ �� ���� ������ ������
    ReplaceString "[^13]{2;}([!^13])", "^13^13\1"

End Sub

' 8
Sub ��������_������_�����()
Attribute ��������_������_�����.VB_Description = "������� ��� ������ ������"
Attribute ��������_������_�����.VB_ProcData.VB_Invoke_Func = "Project.formatter.��������_������_�����"

    ' ������� ��� ������ ������
    ReplaceString "^13@([!^13])", "^13\1"

End Sub

' 9
Sub ��������_������_���������()
Attribute ��������_������_���������.VB_Description = "�������� ������ ��������� ���"
Attribute ��������_������_���������.VB_ProcData.VB_Invoke_Func = "Project.formatter.��������_������_���������"

    ' �������� ������ ��������� ���
    ReplaceString "^-", ""

End Sub

' 10
Sub ������_��_�()
Attribute ������_��_�.VB_Description = "������ �+-� �� ���"
Attribute ������_��_�.VB_ProcData.VB_Invoke_Func = "Project.formatter.������_��_�"

    ' ������ �+-� �� ���
    ReplaceString "+-", "�", False

End Sub

' 11
Sub �������_�������_��������()
Attribute �������_�������_��������.VB_Description = "�������� �������� ����������� � ��������, ������ ���������� ��������� �� ������� � �����, ����� ������ ���������� �����������"
Attribute �������_�������_��������.VB_ProcData.VB_Invoke_Func = "Project.formatter.�������_�������_��������"

    ' �������� �������� ����������� � ��������,
    ' ������ ���������� ��������� �� ������� � �����
    ' ����� ������ ���������� �����������
    ProcessImages

End Sub

' 12
Sub ���������_�����()
Attribute ���������_�����.VB_Description = "�������� ���� �� 2 �� �� ���� ������"
Attribute ���������_�����.VB_ProcData.VB_Invoke_Func = "Project.formatter.���������_�����"

    ' �������� ���� �� 2 �� �� ���� ������
    With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
    End With

End Sub

' 13
Sub ������_������_��_�����������()
Attribute ������_������_��_�����������.VB_Description = "������ � ��������� ���� ������ �� ����� �� normal.dot, �������� ������ ������ � "
Attribute ������_������_��_�����������.VB_ProcData.VB_Invoke_Func = "Project.formatter.������_������_��_�����������"

    ' ������ � ��������� ���� ������ �� ����� ��
    ' normal.dot, �������� ������ ������ � ���������
    With ActiveDocument
        .AttachedTemplate = "Normal.dotm"
        .RemoveLockedStyles
        .UpdateStyles
    End With

End Sub

' 14
Sub ��������_�������_������()
Attribute ��������_�������_������.VB_Description = "�������� ��������� � ������ ������� ������ � ���������"
Attribute ��������_�������_������.VB_ProcData.VB_Invoke_Func = "Project.formatter.��������_�������_������"

    ' ������� ��������� � ������ ������� ������ � ���������
    With ActiveDocument
      .RemoveDocumentInformation (wdRDIAll)
      .Save
    End With

End Sub

' 15
Sub ������_�������_��_����()
Attribute ������_�������_��_����.VB_Description = "������ �� ���� ������ � ������ ������ � ������, ����������� ���������"
Attribute ������_�������_��_����.VB_ProcData.VB_Invoke_Func = "Project.formatter.������_�������_��_����"

    ' ������ ���� (-) �� ����� (�) � ������ ������ � ������, ����������� ���������
    ReplaceString "^13 - ", "�"

End Sub

' 16
Sub ������_������_�����_������_��_�����()
Attribute ������_������_�����_������_��_�����.VB_Description = "������ �� ���� ����� ������ ����� �������� ���� �����, ����� ��� ����� ������"
Attribute ������_������_�����_������_��_�����.VB_ProcData.VB_Invoke_Func = "Project.formatter.������_������_�����_������_��_�����"

    ' ������ �� ���� ����� ������ ����� �������� ���� �����, ����� ��� ����� ������
    ReplaceString "([0-9 ]{1;})[�]", "\1-"

End Sub

' 17
Sub �����������_������_�����_���_�_��()
Attribute �����������_������_�����_���_�_��.VB_Description = "����������� ������ ����� ���λ, ���λ, ���λ ���λ ���λ ���"
Attribute �����������_������_�����_���_�_��.VB_ProcData.VB_Invoke_Func = "Project.formatter.�����������_������_�����_���_�_��"

    ' ����������� ������ ����� ���λ, ���λ, ���λ ���λ ���λ ��λ

    Dim nbsp As String
    nbsp = ChrW(8239) ' ������ ������������ �������

    Dim abrs() As String
    abrs = Split("���λ,���λ,���λ,���λ,���λ,��λ", ",")

    Dim abbr As Variant

    For Each abbr In abrs
        ReplaceString abbr & "[^s ]*", abbr & nbsp, True
    Next abbr

End Sub

' 18
Sub ����������_�����������_��������_�����_����������()
Attribute ����������_�����������_��������_�����_����������.VB_Description = "���������� ����������� �������� ����� � �.�, ���� ����� ���� ���� ��������� �����. ��������, �. ���������. ���������� ���. ��������, ����. �������������, ����. ��������, ����. �������, ��. �������, ��. 14�, ��. 1�, ����. 6�."
Attribute ����������_�����������_��������_�����_����������.VB_ProcData.VB_Invoke_Func = "Project.formatter.����������_�����������_��������_�����_����������"

    ' ����������� ������ ����� � �.� ���� ����� ���� ���� ��������� �����
    ' ... ��,���,���,���,�,�,���

    Dim nbsp As String
    nbsp = ChrW(8239) ' ������ ������������ �������

    Dim abrsWithCapital() As String
    abrsWithCapital = Split("�,��,���,���,���,�,�,���", ",")

    Dim abbr As Variant

    For Each abbr In abrsWithCapital
        ReplaceString " (" & abbr & "\.)([�-�])", " \1" & nbsp & "\2", True
        ReplaceString " (" & abbr & "\.)[^s ]*([�-�])", " \1" & nbsp & "\2", True
    Next abbr

End Sub

' 19
Sub ����������_�����������_��������_�_��������()
Attribute ����������_�����������_��������_�_��������.VB_Description = "������ ��.�.������ �� ��. �. ������ � ������������ ���������; ������� �.�.� �� ������� �. �.� � ������������ ���������"
Attribute ����������_�����������_��������_�_��������.VB_ProcData.VB_Invoke_Func = "Project.formatter.����������_�����������_��������_�_��������"

    Dim nbsp As String
    nbsp = ChrW(8239) ' ������ ������������ �������

    ' ������ ��.�.������ �� ��. �. ������ � ������������ ���������
    ReplaceString "([�-�]\.)([�-�]\.)([�-�][�-�]*)", "\1" & nbsp & "\2" & nbsp & "\3", True

    ' ������� �.�.� �� ������� �. �.� � ������������ ���������
    ReplaceString "([�-�][�-�]* )([�-�]\.)([�-�]\.)", "\1" & nbsp & "\2" & nbsp & "\3", True


End Sub

' 20
Sub ����������_������������_�������_�����_��������()
Attribute ����������_������������_�������_�����_��������.VB_Description = "���������� ������������ ������� ����� ������ � ��������� �� ��� ������"
Attribute ����������_������������_�������_�����_��������.VB_ProcData.VB_Invoke_Func = "Project.formatter.����������_������������_�������_�����_��������"

    ' ���������� ������������ ������� ����� ������ � ��������� �� ��� ������
    ReplaceString "([0-9])([! 0-9])", "\1 \2", True

End Sub

' 21
Sub ���������_������_���������()
Attribute ���������_������_���������.VB_Description = "������ �*�, � * �, ���, � � �, �x�, � x � (������� ����� �� � ���������� ���) ����� ������� �� ���������� ��������� ���� ���������"
Attribute ���������_������_���������.VB_ProcData.VB_Invoke_Func = "Project.formatter.���������_������_���������"

    ' ������ �*�, � * �, ���, � � �, �x�, � x � (������� ����� �� � ���������� ���)
    ' ������� ����� ������� �� ���������� ��������� ���� ���������.

    Dim multuplies() As String
    multuplies = Split("*,x,�", ",")

    Dim multiple As Variant

    For Each multiple In multuplies
        ReplaceString "([0-9])" & multiple & "([0-9])", "\1 " & ChrW(215) & " \2", True
        ReplaceString "([0-9])[^s ]*" & multiple & "[^s ]([0-9])", "\1 " & ChrW(215) & " \2", True
    Next multiple

End Sub

' 22
Sub �����������_���������_�������()
Attribute �����������_���������_�������.VB_Description = "������� �� ����� ����� �^p� ������ ����� �����諶�, ������� ������ �^l�, ������� �������� �^m�, ��� ������� ������� �^b�"
Attribute �����������_���������_�������.VB_ProcData.VB_Invoke_Func = "Project.formatter.�����������_���������_�������"

    ' ������� �� ����� ����� �^p� ������ ����� �����諶�, ������� ������ �^l�,
    ' ������� �������� �^m�, ��� ������� ������� �^b�

    ReplaceString "�", vbCr, False
    ReplaceString "^l", vbCr, False
    ReplaceString "^m", vbCr, False
    ReplaceString "^b", vbCr, False

End Sub

' 23
Sub �������_�_�����_��������������()
Attribute �������_�_�����_��������������.VB_Description = "����� ������� �����, �������, ������ ����� � ������ �������"
Attribute �������_�_�����_��������������.VB_ProcData.VB_Invoke_Func = "Project.formatter.�������_�_�����_��������������"

    ' ����� ������� �����, �������, ������ ����� � ������ �������
    Dim table As table

    For Each table In ActiveDocument.Tables

        Debug.Print table.Style

        table.Style = "Table Normal"
        table.Select
        Selection.ClearFormatting
        Selection.Collapse Direction:=wdCollapseStart

    Next table

End Sub

' 24
Sub �������_�_������_����������_��������_�����()
Attribute �������_�_������_����������_��������_�����.VB_Description = "������ ���������� �������� ����� � �������� �� ����� ��������"
Attribute �������_�_������_����������_��������_�����.VB_ProcData.VB_Invoke_Func = "Project.formatter.�������_�_������_����������_��������_�����"

    '������ ���������� �������� ����� � �������� �� ����� ��������

    Dim table As table

    For Each table In ActiveDocument.Tables
        table.Rows.AllowBreakAcrossPages = False
    Next table

End Sub

' 25
Sub �������_�_������������_��_������_�_�����()
Attribute �������_�_������������_��_������_�_�����.VB_Description = "������������ � ������� ������� �� ������ � �����"
Attribute �������_�_������������_��_������_�_�����.VB_ProcData.VB_Invoke_Func = "Project.formatter.�������_�_������������_��_������_�_�����"

    ' ������������ � ������� ������� �� ������ � �����
    Dim table As table

    For Each table In ActiveDocument.Tables

        With table.Range
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cells.VerticalAlignment = wdCellAlignVerticalCenter
        End With

    Next table

End Sub

' 26
Sub �������_�_�����_�����_�_�������()
Attribute �������_�_�����_�����_�_�������.VB_Description = "������������ ������� �������� ����� � ������� ������"
Attribute �������_�_�����_�����_�_�������.VB_ProcData.VB_Invoke_Func = "Project.formatter.�������_�_�����_�����_�_�������"

    ' ������������ ������� �������� ����� � ������� ������
    Dim table As table

    For Each table In ActiveDocument.Tables
        table.TopPadding = 0
        table.BottomPadding = 0

        With table.Range.Cells.Borders

            .DistanceFromLeft = 0
            .DistanceFromTop = 0

        End With
    Next table

End Sub

' 27
Sub �������_�_������������()
Attribute �������_�_������������.VB_Description = "������������ ������ �� ������ ��������, ����� �� ����������, ����� ����� �� ������"
Attribute �������_�_������������.VB_ProcData.VB_Invoke_Func = "Project.formatter.�������_�_������������"

    ' ������������ ������ �� ������ ��������, ����� �� ����������, ����� ����� �� ������
    Dim table As table

    For Each table In ActiveDocument.Tables
        With table
            .AutoFitBehavior (wdAutoFitWindow)
            .AutoFitBehavior (wdAutoFitContent)
            .AutoFitBehavior (wdAutoFitWindow)
        End With
    Next table

End Sub

' 28
Sub ��������_��������_�����_�_������_�������������_��_pdf()

    ' ������� ��� ����� ������� �^p�, ����� ��������� �������:

    Dim keep As String
    keep = "0c1eff75-9cce-46c7-9965-05f1cc26dbf0"

    ' � ���� ��������� ����� ���������� � ����� ����������� �������� ��� ��������� (� �.�. �������) ����� ������� ���� ����� � ����� (��������, �1.�), ��� ����� � ������ (����. �1)�), ��� ���� ������, ������ ��� ���� �-�, �?�, ��� ��� ������ ���;
    ReplaceString "(^13[0-9]{1;}[\)\.\-\?��])", keep & "\1"

    ' � ���� ����� ������������� ������, ����������, ��������������� ��� �������������� ������;
    ReplaceString "([\.\!\?][^13])", "\1" & keep

    ' � ���� ��������� ����� ���������� � ���������, ��� � 3-�, 4-�, 5-��, 6-��, 7-��, 8-��, 9-��, ��� 10-�� �������� (�� ���� � ������� ���������).
    ReplaceString "([^13][ ]){3;}", keep & "\1"
    ReplaceString "([^13][^t]){1;}", keep & "\1"

    ReplaceString "^13", ""
    ReplaceString keep, "^13"

End Sub

Private Function ReplaceString(pattern As String, replace As String, Optional wildcards As Boolean = True, Optional par As Variant = Null)

    Selection.Collapse Direction:=wdCollapseStart

    Application.ScreenUpdating = False

    Dim f As Find

    If IsNull(par) Then
        Set f = ActiveDocument.Range.Find()
    Else
        Dim p As Paragraph
        Set p = par
        Set f = p.Range.Find()
    End If

    With f

        .Text = pattern
        .Replacement.Text = replace
        .forward = True
        .Format = False
        .Wrap = wdFindContinue
        .MatchWildcards = wildcards
        .Execute replace:=wdReplaceAll

    End With

    Application.ScreenUpdating = True

End Function

Private Function ProcessImages()

    ' Application.ScreenUpdating = False

    Dim objShape As Shape

    For Each objShape In ActiveDocument.Shapes
        If objShape.Type = msoPicture Then
            objShape.WrapFormat.Type = wdWrapSquare

        End If
    Next objShape

    Application.ScreenUpdating = True

End Function


