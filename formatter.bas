Attribute VB_Name = "formatter"
Option Explicit

' Запуск всех функций
Sub Форматировать_Документ()

    Замена_множественных_пробелов_на_один
    Удаление_пробелов_перед_знаками_препинания
    Удаление_пробелов_около_скобок
    Сокращение_множественных_пробелов
    Удаление_пробелов_в_начале_и_в_конце_абзаца
    Исправление_кавычек
    Сокращение_любого_числа_пустых_строк_до_одной
    Удаление_пустых_строк
    Удаление_мягких_переносов
    Замена_на_±
    Очистка_формата_картинок
    Коррекция_полей
    Замена_стилей_на_собственные
    Удаление_скрытых_данных
    Замена_дефисов_на_тире
    Замена_дефиса_перед_цифрой_на_минус
    Неразрывный_пробел_после_ООО_и_др
    Добавление_неразрывных_пробелов_после_сокращений
    Добавление_неразрывных_пробелов_в_инициалы
    Добавление_неразрывного_пробела_после_значений
    Коррекция_знаков_умножения
    Исправление_слипшихся_абзацев
    Таблицы_»_сброс_форматирования
    Таблицы_»_снятие_разрешения_переноса_строк
    Таблицы_»_выравнивание_по_центру_и_влево
    Таблицы_»_сброс_полей_в_ячейках
    Таблицы_»_выравнивание
    Удаление_разрывов_строк_в_тексте_скопированном_из_pdf

End Sub

' 1
Sub Замена_множественных_пробелов_на_один()
Attribute Замена_множественных_пробелов_на_один.VB_Description = "Замена множественных пробелов на один пробел"
Attribute Замена_множественных_пробелов_на_один.VB_ProcData.VB_Invoke_Func = "Project.formatter.Замена_множественных_пробелов_на_один"

    ' Заменяет двойные и множественные пробелы на один пробел
    ReplaceString "([^s ])@[^s ]", "\1"

End Sub

' 2
Sub Удаление_пробелов_перед_знаками_препинания()
Attribute Удаление_пробелов_перед_знаками_препинания.VB_Description = "Удаление пробелов перед знаками препинания «. , ; : ! ?»"
Attribute Удаление_пробелов_перед_знаками_препинания.VB_ProcData.VB_Invoke_Func = "Project.formatter.Удаление_пробелов_перед_знаками_препинания"

    ' Удаляет пробелы перед знаками препинания «. , ; : ! ?»
    ReplaceString "([^s ])@([^s .,;:])", "\2"

End Sub

' 3
Sub Удаление_пробелов_около_скобок()
Attribute Удаление_пробелов_около_скобок.VB_Description = "Удаляет пробелы после открывающей и перед закрывающей скобками «( ) {} []»"
Attribute Удаление_пробелов_около_скобок.VB_ProcData.VB_Invoke_Func = "Project.formatter.Удаление_пробелов_около_скобок"

    ' Удаляет пробелы после открывающихся скобок и перед закрывающей скобкой
    ReplaceString "\([^s ]@([!^s ])", "(\1"
    ReplaceString "[^s ]@\)", ")"

    ReplaceString "\{[^s ]@([!^s ])", "{\1"
    ReplaceString "[^s ]@\}", "}"

    ReplaceString "\[[^s ]@([!^s ])", "[\1"
    ReplaceString "[^s ]@\]", "]"

End Sub

' 4
Sub Сокращение_множественных_пробелов()
Attribute Сокращение_множественных_пробелов.VB_Description = "Заменяет множественные пробелы на один"
Attribute Сокращение_множественных_пробелов.VB_ProcData.VB_Invoke_Func = "Project.formatter.Сокращение_множественных_пробелов"

    ' Заменяет множественные пробелы на один
    ReplaceString "[^s ]@([!^s ])", " \1"

End Sub


' 5
Sub Удаление_пробелов_в_начале_и_в_конце_абзаца()
Attribute Удаление_пробелов_в_начале_и_в_конце_абзаца.VB_Description = "Удаление пробелов в начале и в конце абзаца (вокруг символа «^p»)"
Attribute Удаление_пробелов_в_начале_и_в_конце_абзаца.VB_ProcData.VB_Invoke_Func = "Project.formatter.Удаление_пробелов_в_начале_и_в_конце_абзаца"

    ' Удаление пробелов в начале и в конце абзаца (вокруг символа «^p»)
    ReplaceString "^13[ ^s]@([!^s ])", "^13\1"
    ReplaceString " " + vbCr, vbCr

End Sub

' 6
Sub Исправление_кавычек()
Attribute Исправление_кавычек.VB_Description = "Корректно заменяет кавычки ""…"" на «x»"
Attribute Исправление_кавычек.VB_ProcData.VB_Invoke_Func = "Project.formatter.Исправление_кавычек"

    ' Корректно заменяет кавычки "…" на «…»
    ReplaceString """", """", False
    ReplaceString Chr(147), "«", False
    ReplaceString Chr(148), "»", False

End Sub

' 7
Sub Сокращение_любого_числа_пустых_строк_до_одной()
Attribute Сокращение_любого_числа_пустых_строк_до_одной.VB_Description = "Заменяет две, три и т.д. подряд идущие пустые строки на одну пустую строку "
Attribute Сокращение_любого_числа_пустых_строк_до_одной.VB_ProcData.VB_Invoke_Func = "Project.formatter.Сокращение_любого_числа_пустых_строк_до_одной"

    ' Заменяет две, три и т.д. подряд идущие пустые строки на одну пустую строку
    ReplaceString "[^13]{2;}([!^13])", "^13^13\1"

End Sub

' 8
Sub Удаление_пустых_строк()
Attribute Удаление_пустых_строк.VB_Description = "Удаляет все пустые строки"
Attribute Удаление_пустых_строк.VB_ProcData.VB_Invoke_Func = "Project.formatter.Удаление_пустых_строк"

    ' Удаляет все пустые строки
    ReplaceString "^13@([!^13])", "^13\1"

End Sub

' 9
Sub Удаление_мягких_переносов()
Attribute Удаление_мягких_переносов.VB_Description = "Удаление мягких переносов «¬»"
Attribute Удаление_мягких_переносов.VB_ProcData.VB_Invoke_Func = "Project.formatter.Удаление_мягких_переносов"

    ' Удаление мягких переносов «¬»
    ReplaceString "^-", ""

End Sub

' 10
Sub Замена_на_±()
Attribute Замена_на_±.VB_Description = "Замена «+-» на «±»"
Attribute Замена_на_±.VB_ProcData.VB_Invoke_Func = "Project.formatter.Замена_на_±"

    ' Замена «+-» на «±»
    ReplaceString "+-", "±", False

End Sub

' 11
Sub Очистка_формата_картинок()
Attribute Очистка_формата_картинок.VB_Description = "Удаление привязки изображений к позициям, замена параметров обтекания на «сверху и снизу», сброс прочих параметров изображений"
Attribute Очистка_формата_картинок.VB_ProcData.VB_Invoke_Func = "Project.formatter.Очистка_формата_картинок"

    ' Удаление привязки изображений к позициям,
    ' замена параметров обтекания на «сверху и снизу»
    ' сброс прочих параметров изображений
    ProcessImages

End Sub

' 12
Sub Коррекция_полей()
Attribute Коррекция_полей.VB_Description = "Изменяет поля на 2 см со всех сторон"
Attribute Коррекция_полей.VB_ProcData.VB_Invoke_Func = "Project.formatter.Коррекция_полей"

    ' Изменяет поля на 2 см со всех сторон
    With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2)
    End With

End Sub

' 13
Sub Замена_стилей_на_собственные()
Attribute Замена_стилей_на_собственные.VB_Description = "Замена в документе всех стилей на стили из normal.dot, удаление лишних стилей в "
Attribute Замена_стилей_на_собственные.VB_ProcData.VB_Invoke_Func = "Project.formatter.Замена_стилей_на_собственные"

    ' Замена в документе всех стилей на стили из
    ' normal.dot, удаление лишних стилей в документе
    With ActiveDocument
        .AttachedTemplate = "Normal.dotm"
        .RemoveLockedStyles
        .UpdateStyles
    End With

End Sub

' 14
Sub Удаление_скрытых_данных()
Attribute Удаление_скрытых_данных.VB_Description = "Удаление авторства и прочих скрытых данных в документе"
Attribute Удаление_скрытых_данных.VB_ProcData.VB_Invoke_Func = "Project.formatter.Удаление_скрытых_данных"

    ' Очистка авторства и прочих скрытых данных в документе
    With ActiveDocument
      .RemoveDocumentInformation (wdRDIAll)
      .Save
    End With

End Sub

' 15
Sub Замена_дефисов_на_тире()
Attribute Замена_дефисов_на_тире.VB_Description = "Замена на тире дефиса в начале абзаца и дефиса, окруженного пробелами"
Attribute Замена_дефисов_на_тире.VB_ProcData.VB_Invoke_Func = "Project.formatter.Замена_дефисов_на_тире"

    ' Замена тире (-) на дефис (–) в начале абзаца и дефиса, окруженного пробелами
    ReplaceString "^13 - ", "–"

End Sub

' 16
Sub Замена_дефиса_перед_цифрой_на_минус()
Attribute Замена_дефиса_перед_цифрой_на_минус.VB_Description = "Замена на знак минус дефиса после которого идет цифра, сразу или через пробел"
Attribute Замена_дефиса_перед_цифрой_на_минус.VB_ProcData.VB_Invoke_Func = "Project.formatter.Замена_дефиса_перед_цифрой_на_минус"

    ' Замена на знак минус дефиса после которого идет цифра, сразу или через пробел
    ReplaceString "([0-9 ]{1;})[–]", "\1-"

End Sub

' 17
Sub Неразрывный_пробел_после_ООО_и_др()
Attribute Неразрывный_пробел_после_ООО_и_др.VB_Description = "Неразрывный пробел после «ООО», «ОАО», «ЗАО» «ПАО» «НАО» «АО"
Attribute Неразрывный_пробел_после_ООО_и_др.VB_ProcData.VB_Invoke_Func = "Project.formatter.Неразрывный_пробел_после_ООО_и_др"

    ' Неразрывный пробел после «ООО», «ОАО», «ЗАО» «ПАО» «НАО» «АО»

    Dim nbsp As String
    nbsp = ChrW(8239) ' Символ неразрывного пробела

    Dim abrs() As String
    abrs = Split("«ООО»,«ОАО»,«ЗАО»,«ПАО»,«НАО»,«АО»", ",")

    Dim abbr As Variant

    For Each abbr In abrs
        ReplaceString abbr & "[^s ]*", abbr & nbsp, True
    Next abbr

End Sub

' 18
Sub Добавление_неразрывных_пробелов_после_сокращений()
Attribute Добавление_неразрывных_пробелов_после_сокращений.VB_Description = "Добавление неразрывных пробелов после « г.», если после него идет заглавная буква. Например, г. Самарканд. Аналогично «ул. Широкая», «пер. Старомонетный», «дер. Вербилки», «пос. Рабочий», «с. Вятское», «д. 14», «к. 1», «стр. 6»."
Attribute Добавление_неразрывных_пробелов_после_сокращений.VB_ProcData.VB_Invoke_Func = "Project.formatter.Добавление_неразрывных_пробелов_после_сокращений"

    ' Неразрывный пробел после « г.» если после него идет заглавная буква
    ' ... ул,пер,дер,пос,с,д,стр

    Dim nbsp As String
    nbsp = ChrW(8239) ' Символ неразрывного пробела

    Dim abrsWithCapital() As String
    abrsWithCapital = Split("г,ул,пер,дер,пос,с,д,стр", ",")

    Dim abbr As Variant

    For Each abbr In abrsWithCapital
        ReplaceString " (" & abbr & "\.)([А-Я])", " \1" & nbsp & "\2", True
        ReplaceString " (" & abbr & "\.)[^s ]*([А-Я])", " \1" & nbsp & "\2", True
    Next abbr

End Sub

' 19
Sub Добавление_неразрывных_пробелов_в_инициалы()
Attribute Добавление_неразрывных_пробелов_в_инициалы.VB_Description = "Замена «А.Б.Иванов» на «А. Б. Иванов» с неразрывными пробелами; «Иванов А.Б.» на «Иванов А. Б.» с неразрывными пробелами"
Attribute Добавление_неразрывных_пробелов_в_инициалы.VB_ProcData.VB_Invoke_Func = "Project.formatter.Добавление_неразрывных_пробелов_в_инициалы"

    Dim nbsp As String
    nbsp = ChrW(8239) ' Символ неразрывного пробела

    ' Замена «А.Б.Иванов» на «А. Б. Иванов» с неразрывными пробелами
    ReplaceString "([А-Я]\.)([А-Я]\.)([А-Я][а-я]*)", "\1" & nbsp & "\2" & nbsp & "\3", True

    ' «Иванов А.Б.» на «Иванов А. Б.» с неразрывными пробелами
    ReplaceString "([А-Я][а-я]* )([А-Я]\.)([А-Я]\.)", "\1" & nbsp & "\2" & nbsp & "\3", True


End Sub

' 20
Sub Добавление_неразрывного_пробела_после_значений()
Attribute Добавление_неразрывного_пробела_после_значений.VB_Description = "Добавление неразрывного пробела между цифрой и следующей за ней буквой"
Attribute Добавление_неразрывного_пробела_после_значений.VB_ProcData.VB_Invoke_Func = "Project.formatter.Добавление_неразрывного_пробела_после_значений"

    ' Добавление неразрывного пробела между цифрой и следующей за ней буквой
    ReplaceString "([0-9])([! 0-9])", "\1 \2", True

End Sub

' 21
Sub Коррекция_знаков_умножения()
Attribute Коррекция_знаков_умножения.VB_Description = "Замена «*», « * », «х», « х », «x», « x » (русские буквы ха и английские икс) между цифрами на окруженный пробелами знак умножения"
Attribute Коррекция_знаков_умножения.VB_ProcData.VB_Invoke_Func = "Project.formatter.Коррекция_знаков_умножения"

    ' Замена «*», « * », «х», « х », «x», « x » (русские буквы ха и английские икс)
    ' стоящих между цифрами на окруженный пробелами знак умножения.

    Dim multuplies() As String
    multuplies = Split("*,x,х", ",")

    Dim multiple As Variant

    For Each multiple In multuplies
        ReplaceString "([0-9])" & multiple & "([0-9])", "\1 " & ChrW(215) & " \2", True
        ReplaceString "([0-9])[^s ]*" & multiple & "[^s ]([0-9])", "\1 " & ChrW(215) & " \2", True
    Next multiple

End Sub

' 22
Sub Исправление_слипшихся_абзацев()
Attribute Исправление_слипшихся_абзацев.VB_Description = "Переход на новый абзац «^p» вместо конца строки«¶», разрыва строки «^l», разрыва страницы «^m», или разрыва раздела «^b»"
Attribute Исправление_слипшихся_абзацев.VB_ProcData.VB_Invoke_Func = "Project.formatter.Исправление_слипшихся_абзацев"

    ' Переход на новый абзац «^p» вместо конца строки«¶», разрыва строки «^l»,
    ' разрыва страницы «^m», или разрыва раздела «^b»

    ReplaceString "¶", vbCr, False
    ReplaceString "^l", vbCr, False
    ReplaceString "^m", vbCr, False
    ReplaceString "^b", vbCr, False

End Sub

' 23
Sub Таблицы_»_сброс_форматирования()
Attribute Таблицы_»_сброс_форматирования.VB_Description = "Сброс толщины линий, заливки, высоты строк и ширины колонок"
Attribute Таблицы_»_сброс_форматирования.VB_ProcData.VB_Invoke_Func = "Project.formatter.Таблицы_»_сброс_форматирования"

    ' Сброс толщины линий, заливки, высоты строк и ширины колонок
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
Sub Таблицы_»_снятие_разрешения_переноса_строк()
Attribute Таблицы_»_снятие_разрешения_переноса_строк.VB_Description = "Снятие разрешения переноса строк в таблицах на новую страницу"
Attribute Таблицы_»_снятие_разрешения_переноса_строк.VB_ProcData.VB_Invoke_Func = "Project.formatter.Таблицы_»_снятие_разрешения_переноса_строк"

    'Снятие разрешения переноса строк в таблицах на новую страницу

    Dim table As table

    For Each table In ActiveDocument.Tables
        table.Rows.AllowBreakAcrossPages = False
    Next table

End Sub

' 25
Sub Таблицы_»_выравнивание_по_центру_и_влево()
Attribute Таблицы_»_выравнивание_по_центру_и_влево.VB_Description = "Выравнивание в ячейках таблицы по центру и влево"
Attribute Таблицы_»_выравнивание_по_центру_и_влево.VB_ProcData.VB_Invoke_Func = "Project.formatter.Таблицы_»_выравнивание_по_центру_и_влево"

    ' Выравнивание в ячейках таблицы по центру и влево
    Dim table As table

    For Each table In ActiveDocument.Tables

        With table.Range
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Cells.VerticalAlignment = wdCellAlignVerticalCenter
        End With

    Next table

End Sub

' 26
Sub Таблицы_»_сброс_полей_в_ячейках()
Attribute Таблицы_»_сброс_полей_в_ячейках.VB_Description = "Установление нулевых значений полей в ячейках таблиц"
Attribute Таблицы_»_сброс_полей_в_ячейках.VB_ProcData.VB_Invoke_Func = "Project.formatter.Таблицы_»_сброс_полей_в_ячейках"

    ' Установление нулевых значений полей в ячейках таблиц
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
Sub Таблицы_»_выравнивание()
Attribute Таблицы_»_выравнивание.VB_Description = "Выравнивание таблиц по ширине страницы, затем по содержанию, затем снова по ширине"
Attribute Таблицы_»_выравнивание.VB_ProcData.VB_Invoke_Func = "Project.formatter.Таблицы_»_выравнивание"

    ' Выравнивание таблиц по ширине страницы, затем по содержанию, затем снова по ширине
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
Sub Удаление_разрывов_строк_в_тексте_скопированном_из_pdf()

    ' Удаляет все концы абзацев «^p», кроме следующих случаев:

    Dim keep As String
    keep = "0c1eff75-9cce-46c7-9965-05f1cc26dbf0"

    ' – если следующий абзац начинается с любым количеством пробелов или табуляций (в т.ч. нулевым) после которых идет цифра и точка (например, «1.»), или цифра и скобка (напр. «1)»), или знак минуса, дефиса или тире «-», «?», «—» или буллет «•»;
    ReplaceString "(^13[0-9]{1;}[\)\.\-\?—•])", keep & "\1"

    ' – если абзац заканчивается точкой, двоеточием, восклицательным или вопросительным знаком;
    ReplaceString "([\.\!\?][^13])", "\1" & keep

    ' – если следующий абзац начинается с табуляции, или с 3-х, 4-х, 5-ти, 6-ти, 7-ми, 8-ми, 9-ти, или 10-ти пробелов (то есть с аналога табуляции).
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


