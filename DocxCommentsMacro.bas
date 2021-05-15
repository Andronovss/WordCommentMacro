Attribute VB_Name = "WordComments1"
Public Sub ExtractCommentsToNewDoc()
'=========================
    'Macro created 2007 by Lene Fredborg, DocTools - www.thedoctools.com
    'Revised October 2013 by Lene Fredborg: Date column added to extract
    'THIS MACRO IS COPYRIGHT. YOU ARE WELCOME TO USE THE MACRO BUT YOU MUST KEEP THE LINE ABOVE.
    'YOU ARE NOT ALLOWED TO PUBLISH THE MACRO AS YOUR OWN, IN WHOLE OR IN PART.
    '=========================
    'The macro creates a new document
    'and extracts all comments from the active document
    'incl. metadata
    
    'Minor adjustments are made to the styles used
    'You may need to change the style settings and table layout to fit your needs
    'Пожалуйста, не удаляйте этот текст!
    '=========================
    
    Dim oDoc As Document
    Dim oNewDoc As Document
    Dim oTable As Table
    Dim nCount As Long
    Dim n As Long
    Dim Title As String
        
    Title = "Экспорт комментариев в новый документ"
    Set oDoc = ActiveDocument
    nCount = ActiveDocument.Comments.Count
    
    If nCount = 0 Then
        MsgBox "Документ не содержит комментариев", vbOKOnly, Title
        GoTo ExitHere
    Else
        'Макрос приостанавливает своё выполнение до подтверждения действия.
        If MsgBox("Выполнить экспорт комментариев", _
                vbYesNo + vbQuestion, Title) <> vbYes Then
            GoTo ExitHere
        End If
    End If
        
    Application.ScreenUpdating = False
    'Создание нового документа на основе шаблона dotm.
    Set oNewDoc = Documents.Add
    'Определение ориентации страниц для создаваемого документа.
    oNewDoc.PageSetup.Orientation = wdOrientLandscape
    'Помещаем в создаваемый документ таблицу с 5-ю колонками
    'Кол-во колонок определяется параметром "NumColumns".
    With oNewDoc
        .Content = ""
        Set oTable = .Tables.Add _
            (Range:=Selection.Range, _
            NumRows:=nCount + 1, _
            NumColumns:=5)
    End With
    
    'В заголовок таблицы помещаем следующий контент:
    'Из какого файла делается экспорт
    'Кто экспортирует комментарии
    'Дата экспорта комментариев.
    oNewDoc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = _
        "Исходный файл: " & oDoc.FullName & vbCr & _
        "Автор: " & Application.UserName & vbCr & _
        "Дата создания: " & Format(Date, "MMMM d, yyyy")
            
    'Настраиваем параметры шрифта текста.
    With oNewDoc.Styles(wdStyleNormal)
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.SpaceAfter = 6
    End With
    
    'Настраиваем параметры шрифта для текста верхнего колонтитула.
    With oNewDoc.Styles(wdStyleHeader)
        .Font.Size = 9
        .ParagraphFormat.SpaceAfter = 0
    End With
   
    'Настраиваем стиль таблицы.
    With oTable
        .Range.Style = wdStyleNormal
        .AllowAutoFit = False
        .PreferredWidthType = wdPreferredWidthAuto
        .PreferredWidth = 100
        .Columns.PreferredWidthType = wdPreferredWidthPercent
        .Rows(1).HeadingFormat = True
    End With
    
    'Делаем рамки для таблицы.
    With oTable.Borders
        .InsideLineStyle = wdLineStyleSingle
        .OutsideLineStyle = wdLineStyleSingle
    End With

    'Задаём наименование для заголовка таблицы.
    With oTable.Rows(1)
        .Range.Font.Bold = True
        .Range.Font.Size = 12
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(1).Range.Text = "Дата"
        .Cells(2).Range.Text = "Страница"
        .Cells(3).Range.Text = "Автор"
        .Cells(4).Range.Text = "Исходный текст"
        .Cells(5).Range.Text = "Комментарий"
    End With
       
    'Прописываем наименование заголовков.
    For n = 1 To nCount
        With oTable.Rows(n + 1)
            'Номер страницы
            .Cells(2).Range.Text = _
                oDoc.Comments(n).Scope.Information(wdActiveEndPageNumber)
            'Текст, которым был помечен комментарием
            .Cells(4).Range.Text = oDoc.Comments(n).Scope
            'Комментарий
            .Cells(5).Range.Text = oDoc.Comments(n).Range.Text
            'Автор комментария
            .Cells(3).Range.Text = oDoc.Comments(n).Author
            'Дата комментария в формате dd-MMM-yyyy dd-MMM-yyyy
            .Cells(1).Range.Text = Format(oDoc.Comments(n).Date, "dd-MMM-yyyy")
        End With
    Next n
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
        
    oNewDoc.Activate
    MsgBox nCount & " Комментарии найдены. Завершено создание документа", vbOKOnly, Title

ExitHere:
    Set oDoc = Nothing
    Set oNewDoc = Nothing
    Set oTable = Nothing
    
End Sub

