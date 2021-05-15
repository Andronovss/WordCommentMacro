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
    '����������, �� �������� ���� �����!
    '=========================
    
    Dim oDoc As Document
    Dim oNewDoc As Document
    Dim oTable As Table
    Dim nCount As Long
    Dim n As Long
    Dim Title As String
        
    Title = "������� ������������ � ����� ��������"
    Set oDoc = ActiveDocument
    nCount = ActiveDocument.Comments.Count
    
    If nCount = 0 Then
        MsgBox "�������� �� �������� ������������", vbOKOnly, Title
        GoTo ExitHere
    Else
        '������ ���������������� ��� ���������� �� ������������� ��������.
        If MsgBox("��������� ������� ������������", _
                vbYesNo + vbQuestion, Title) <> vbYes Then
            GoTo ExitHere
        End If
    End If
        
    Application.ScreenUpdating = False
    '�������� ������ ��������� �� ������ ������� dotm.
    Set oNewDoc = Documents.Add
    '����������� ���������� ������� ��� ������������ ���������.
    oNewDoc.PageSetup.Orientation = wdOrientLandscape
    '�������� � ����������� �������� ������� � 5-� ���������
    '���-�� ������� ������������ ���������� "NumColumns".
    With oNewDoc
        .Content = ""
        Set oTable = .Tables.Add _
            (Range:=Selection.Range, _
            NumRows:=nCount + 1, _
            NumColumns:=5)
    End With
    
    '� ��������� ������� �������� ��������� �������:
    '�� ������ ����� �������� �������
    '��� ������������ �����������
    '���� �������� ������������.
    oNewDoc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text = _
        "�������� ����: " & oDoc.FullName & vbCr & _
        "�����: " & Application.UserName & vbCr & _
        "���� ��������: " & Format(Date, "MMMM d, yyyy")
            
    '����������� ��������� ������ ������.
    With oNewDoc.Styles(wdStyleNormal)
        .Font.Name = "Times New Roman"
        .Font.Size = 10
        .ParagraphFormat.LeftIndent = 0
        .ParagraphFormat.SpaceAfter = 6
    End With
    
    '����������� ��������� ������ ��� ������ �������� �����������.
    With oNewDoc.Styles(wdStyleHeader)
        .Font.Size = 9
        .ParagraphFormat.SpaceAfter = 0
    End With
   
    '����������� ����� �������.
    With oTable
        .Range.Style = wdStyleNormal
        .AllowAutoFit = False
        .PreferredWidthType = wdPreferredWidthAuto
        .PreferredWidth = 100
        .Columns.PreferredWidthType = wdPreferredWidthPercent
        .Rows(1).HeadingFormat = True
    End With
    
    '������ ����� ��� �������.
    With oTable.Borders
        .InsideLineStyle = wdLineStyleSingle
        .OutsideLineStyle = wdLineStyleSingle
    End With

    '����� ������������ ��� ��������� �������.
    With oTable.Rows(1)
        .Range.Font.Bold = True
        .Range.Font.Size = 12
        .Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cells(1).Range.Text = "����"
        .Cells(2).Range.Text = "��������"
        .Cells(3).Range.Text = "�����"
        .Cells(4).Range.Text = "�������� �����"
        .Cells(5).Range.Text = "�����������"
    End With
       
    '����������� ������������ ����������.
    For n = 1 To nCount
        With oTable.Rows(n + 1)
            '����� ��������
            .Cells(2).Range.Text = _
                oDoc.Comments(n).Scope.Information(wdActiveEndPageNumber)
            '�����, ������� ��� ������� ������������
            .Cells(4).Range.Text = oDoc.Comments(n).Scope
            '�����������
            .Cells(5).Range.Text = oDoc.Comments(n).Range.Text
            '����� �����������
            .Cells(3).Range.Text = oDoc.Comments(n).Author
            '���� ����������� � ������� dd-MMM-yyyy dd-MMM-yyyy
            .Cells(1).Range.Text = Format(oDoc.Comments(n).Date, "dd-MMM-yyyy")
        End With
    Next n
    
    Application.ScreenUpdating = True
    Application.ScreenRefresh
        
    oNewDoc.Activate
    MsgBox nCount & " ����������� �������. ��������� �������� ���������", vbOKOnly, Title

ExitHere:
    Set oDoc = Nothing
    Set oNewDoc = Nothing
    Set oTable = Nothing
    
End Sub

