Attribute VB_Name = "B_���o���ꗗ�쐬_20150128"
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
'�@�ߎ����x���}�N���iWord�Łj(2/3)�@B_���o���ꗗ�쐬_20150128
'Copyright(C)2014 Blue Panda
'This software is released under the MIT License.
'http://opensource.org/licenses/mit-license.php
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
Option Explicit
Dim TextStyle As Integer
Dim fRng As Range
Dim xls As Object
Dim MaxRowP As Long, DaimeiRow As Long, MokujiEndRow As Long
Dim First As String
Dim BorderEdgeTopRow As Long, ListUpperEnd As Long, HonsokuEndRow As Long
Dim FirstArticleRow As Long
Dim Hen As Integer, SyouSetsuKanMoku As Integer
Dim JyouColumn As Integer
Dim NumberOfArticle As Long
Dim LineOnOff As Integer
Dim i As Long, j As Long, n As Long

Sub ���o���ꗗ�쐬()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     ���傤�����́uSuper�@��Web�v�y�ё��@�K�uD1-Law.com�v�i���@�K �@��񑍍��f�[�^�x�[�X�j
' �@�@����o�͂��ꂽ�𕶃f�[�^�̃t�@�C�����т�e-Gov�̏𕶃f�[�^��Word�ɓ]�L�����t�@�C������
' �@�@�𕶂̌��o���ꗗ���쐬����B
'
' �y�����z
'     2015/01/28�@����Ł@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y���l�z
' �@�@����Ɏ኱�̎��Ԃ�v���܂��B
' �@�@�G���[���ŏ��������f���ꂽ�ꍇ�́A�ēx�}�N�������s���Ă��������B
'
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

    Call ���o���ꗗ�쐬_Lite
    Call �̍ك^�e
    Call �r���^�e
    
    xls.Cells(1, 11).Value = Now()
    xls.Cells(1, JyouColumn + 1).Value = xls.Cells(1, 11).Value - xls.Cells(1, 10).Value
    xls.Cells(1, JyouColumn + 1).NumberFormatLocal = "h:mm:ss"
    xls.Visible = True
    
End Sub
Sub ���o���ꗗ�쐬_������������\��()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     ���傤�����́uSuper�@��Web�v�̃f�[�^��p
' �@�@���o���ꗗ�̍쐬�Ɠ��l�̏������s���A���������̃f�[�^������͈͂Ɋ܂߁A�̍ق𒲐�����B
'
' �y�����z
'     2015/01/28�@����Ł@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y���l�z
' �@�@����Ɏ኱�̎��Ԃ�v���܂��B
' �@�@�G���[���ŏ��������f���ꂽ�ꍇ�́A�ēx�}�N�������s���Ă��������B
'
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
Dim notice As Integer
  
    Call ���o���ꗗ�쐬_Lite
    
    If TextStyle = 2 Then
        Call �̍ك��R
        Call �r�����R
    Else
    'Super�@��Web�ȊO�̏𕶃f�[�^�ƍl������ꍇ�̏���
        xls.Visible = False
        notice = MsgBox("Super�@��Web�̏𕶃f�[�^�ł͂Ȃ��Ǝv���܂��B��������������\�����Ȃ����̂ɕύX���܂���?", vbYesNo, "�m�F")
        Select Case notice
            Case vbYes
                xls.Visible = True
                Call �̍ك^�e
                Call �r���^�e
            Case vbNo
                xls.Visible = True
                Call �̍ك��R
                Call �r�����R
            Case Else
                xls.Visible = True
                Call �̍ك��R
                Call �r�����R
            End Select
    End If
    
    xls.Cells(1, 11).Value = Now()
    xls.Cells(1, JyouColumn + 2).Value = xls.Cells(1, 11).Value - xls.Cells(1, 10).Value
    xls.Cells(1, JyouColumn + 2).NumberFormatLocal = "h:mm:ss"
    xls.Visible = True
    
End Sub


Private Sub ���o���ꗗ�쐬_Lite()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     ���o���ꗗ���쐬�B
'
' �y�����z
'     2015/01/28�@����Ł@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '�𕶂̌`���𔻕ʁi1 e-Gov �A2 Super�@��Web �A3 �@��񑍍��f�[�^�x�[�X �A4 �V�X�e���j

    Dim myRange As Range
    Set myRange = ActiveDocument.Range(Start:=0, End:=1)
     
    TextStyle = 3
    If myRange.Font.Name = "�l�r �o�S�V�b�N" Then TextStyle = 1
    
    If myRange = "��" Or myRange = "��" Then TextStyle = 2
    
    If myRange = "��" And myRange.ParagraphFormat.LeftIndent = 0 _
    And myRange.ParagraphFormat.FirstLineIndent = 0 Then TextStyle = 4
    
    If myRange = "��" And myRange.ParagraphFormat.LeftIndent = 0 _
    And myRange.ParagraphFormat.FirstLineIndent = 0 Then TextStyle = 4
    
  '���p���ʂ�S�p�ɂ���B
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Text = "[)(]{1,}"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.CharacterWidth = wdWidthFullWidth
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

    Selection.WholeStory
    Selection.Copy
    Selection.Collapse Direction:=wdCollapseStart
    

    On Error Resume Next

  'Excel�̃V�[�g���쐬
    Set xls = CreateObject("Excel.Application")
   
    With xls
    
        .SheetsInNewWorkbook = 1
        .Workbooks.Add
        
        .Cells(1, 10).Value = Now()
        
        '���m�F�p�i������m�F����ۃJ���}���O���B�j
        '.Visible = True
        
        .Cells(1, 16).Select
        .ActiveSheet.Paste
        
        '�������΍�
        Excel.Application.CutCopyMode = False
        .Cells(1, 15).Select
        .Selection.Copy

        MaxRowP = 5000  '���̍s��MaxRowP���ݒ肳��Ȃ������ꍇ�̑΍�
        MaxRowP = .Cells(Rows.Count, 16).End(xlUp).Row

        On Error GoTo 0
   
      '���H�p�i�X�y�[�X�̑O�܂łŃJ�b�g�j
        'e-Gov
        If TextStyle = 1 Then
            For i = 1 To MaxRowP
                .Cells(i, 15).Value = Left(.Cells(i, 16), (InStr(.Cells(i, 16), "�@")))
                'e -Gov�̂��߂̏���
                If Left(.Cells(i, 16), 2) = "�@��" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "�@")) + 1))
                End If
                If Left(.Cells(i, 16), 3) = "�@�@��" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "�@")) + 2))
                End If
                If Left(.Cells(i, 16), 4) = "�@�@�@��" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "�@")) + 3))
                End If
                If Left(.Cells(i, 16), 5) = "�@�@�@�@��" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "�@")) + 4))
                End If
                If Left(.Cells(i, 16), 6) = "�@�@�@�@�@��" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "�@")) + 5))
                End If
                If Left(.Cells(i, 16), 7) = "�@�@�@�@�@�@��" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "�@")) + 6))
                End If
                If Left(.Cells(i, 16), 6) = "�@�@�@���@��" Then
                    .Cells(i, 15).Value = LTrim(.Cells(i, 16))
                End If
            Next
        End If
        
        'Super�@��Web
        If TextStyle = 2 Then
            For i = 1 To MaxRowP
                .Cells(i, 15).Value = Left(.Cells(i, 16), (InStr(.Cells(i, 16), "�@")))
                'Super�@��Web�̉�������\���̂��߂̏���
                If .Cells(i, 16).IndentLevel = 3 Then .Cells(i, 16).Value = " " + .Cells(i, 16).Value
            Next
        End If

        '�@��񑍍��f�[�^�x�[�X or �V�X�e��
        If TextStyle = 3 Or TextStyle = 4 Then
            For i = 1 To MaxRowP
                .Cells(i, 15).Value = Left(.Cells(i, 16), (InStr(.Cells(i, 16), "�@")))
            Next
        End If
        
        '�𕶂̍\���̊m�F�i���@�Ȃǁu�ڎ��v�u�����v�̂Ȃ��ڎ����z��j
        Hen = 0
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "��") > 1 Then
                Hen = 1
                Exit For
            End If
        Next
        
        SyouSetsuKanMoku = 1
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "��") > 1 Then
                SyouSetsuKanMoku = 2
                Exit For
            End If
        Next
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "��") > 1 Then
                SyouSetsuKanMoku = 3
                Exit For
            End If
        Next
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "��") > 1 Then
                SyouSetsuKanMoku = 4
                Exit For
            End If
        Next
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "��") > 1 Then
                SyouSetsuKanMoku = 5
                Exit For
            End If
        Next
        
        '�ڎ����͖̏������폜
        MokujiEndRow = 4
        For i = 1 To MaxRowP
            If Left(.Cells(i, 16), 2) = "����" Or Left(.Cells(i, 16), 3) = "�@����" Then
                MokujiEndRow = i
                Exit For
            End If
        Next
        
        If MokujiEndRow = 4 Then
            If Hen = 1 Then
                n = 0
                i = 1
                Do
                    If Left(.Cells(i, 15), 3) = "����" Then n = n + 1
                    i = i + 1
                    If n = 2 Then MokujiEndRow = i - 2
                Loop Until n = 2 Or i = MaxRowP
            End If
        
            If Hen = 0 Then
                n = 0
                i = 1
                Do
                    If Left(.Cells(i, 15), 3) = "����" Then n = n + 1
                    i = i + 1
                    If n = 2 Then MokujiEndRow = i - 2
                Loop Until n = 2 Or i = MaxRowP
            End If
        End If
        '�ڎ����͖̏������폜
        .Range(.Cells(1, 15), .Cells(MokujiEndRow, 15)).Clear
    
        '���o���ꗗ�̖`�������i�薼�A�������𓙁j�̏���

        For i = 1 To 3
            .Cells(i, 1).Value = .Cells(i, 16)
        Next
        
         j = 5
        JyouColumn = SyouSetsuKanMoku + Hen + 1
        
        '�@��񑍍��f�[�^�x�[�X�̏ꍇ�̐ݒ�
        If TextStyle = 3 Then
            .Cells(4, 1).Value = .Cells(4, 16)
            j = 6
        End If
        
        FirstArticleRow = j
        
    '�͖��A�𖼁A���o�����̓]�L�iSuper�@��Web�ȊO�̏ꍇ�j
    
        If TextStyle <> 2 Then

            For i = FirstArticleRow To MaxRowP

                '�ҁA�́A�߁A���A��
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 2).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 2 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 3 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 4 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 5 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
            
                '�𖼋y�ь��o��
                If InStr(.Cells(i, 15), "��") > 1 Then
                    If InStr(.Cells(i - 1, 16), "�i") = 1 Then
                        .Cells(j, JyouColumn).Value = .Cells(i, 15) + " " + .Cells(i - 1, 16)
                        j = j + 1
                    Else
                        If InStr(.Cells(i, 16), "�@�폜") > 1 Then
                            .Cells(j, JyouColumn).Value = .Cells(i, 15) + "�@�폜"
                            j = j + 1
                        Else
                            .Cells(j, JyouColumn).Value = .Cells(i, 15)
                            j = j + 1
                        End If
                    End If
                End If
            
                '���@��
                If InStr(.Cells(i, 15), "��") >= 1 Then
                    .Cells(j, 2).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
            Next
        End If
    
    '�͖��A�𖼁A���o�����̓]�L�iSuper�@��Web�̏ꍇ�j
    
        If TextStyle = 2 Then

            For i = FirstArticleRow To MaxRowP
        
                '���������iSuper�@��Web�̂ݑΉ��j
                If Left(.Cells(i, 16), 3) = " �i��" Or Left(.Cells(i, 16), 3) = " �i��" _
                Or Left(.Cells(i, 16), 3) = " �i��" Or Left(.Cells(i, 16), 3) = " �i��" Then
                    .Cells(j - 1, JyouColumn + 1).Value = LTrim(.Cells(i, 16))
                End If
                
                '�ҁA�́A�߁A���A��
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 2).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 2 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 3 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 4 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "��") > 1 Then
                    .Cells(j, 5 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
            
                '�𖼋y�ь��o��
                If InStr(.Cells(i, 15), "��") > 1 Then
                    If InStr(.Cells(i - 1, 16), "�i") = 1 Then
                        .Cells(j, JyouColumn).Value = .Cells(i, 15) + " " + .Cells(i - 1, 16)
                        j = j + 1
                    Else
                        If InStr(.Cells(i, 16), "�@�폜") > 1 Then
                            .Cells(j, JyouColumn).Value = .Cells(i, 15) + "�@�폜"
                            j = j + 1
                        Else
                            .Cells(j, JyouColumn).Value = .Cells(i, 15)
                            j = j + 1
                        End If
                    End If
                End If
            
                '���@��
                If InStr(.Cells(i, 15), "��") >= 1 Then
                    .Cells(j, 2).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
            Next
        End If
    End With
       
            
      '����Ɨp�̃Z���̃f�[�^�������i�m�F�̍ۂ͖`���ɃJ���}��������B�j
        xls.Range(xls.Cells(1, 15), xls.Cells(MaxRowP, 20)).Clear
        
        xls.Cells(1, 1).Select
        
        xls.Visible = True

End Sub

Private Sub �̍ك^�e()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     �쐬���ꂽ���o���ꗗ�̑̍ق𒲐�����B���������𕔕�������͈͂Ɋ܂߂Ȃ��B
'
' �y�����z
'     2015/01/28�@����Ł@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
    With xls
    
       '��̕��A�s�̍����A�܂�Ԃ��ݒ�A�t�H���g����
        .Range(.Cells(1, 1), .Cells(1, JyouColumn - 1)).ColumnWidth = 1.5
        .Range(.Cells(1, JyouColumn), .Cells(1, JyouColumn)).ColumnWidth = 74
        .Range(.Cells(1, 1), .Cells(j, 1)).RowHeight = 17.25
        
        .Range(.Cells(1, JyouColumn), .Cells(j, JyouColumn)).WrapText = True '���y�ь��o���̃Z����܂�Ԃ��\��

        .Range(.Cells(1, 1), .Cells(j, JyouColumn + 1)).Font.Name = "�l�r �S�V�b�N"
        .Range(.Cells(1, 1), .Cells(j, JyouColumn + 1)).Font.Size = 9
        
      '����͈͂̐ݒ�
        .ActiveSheet.PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(j - 1, JyouColumn)).Address
        
      '�w�b�_�[�A�t�b�^�[
        .ActiveSheet.PageSetup.LeftHeader = "�y���o���ꗗ�z " & .Cells(1, 1).Value
        .ActiveSheet.PageSetup.RightHeader = "&10&P / &N"
        .Cells(1, JyouColumn + 1).Value = Now()
        .Cells(1, JyouColumn + 1).NumberFormat = "yyyy/MM/dd H:mm:ss"
        .ActiveSheet.PageSetup.RightFooter = "&07" & " " & .Cells(1, JyouColumn + 1).Text
        
    End With
    
End Sub

Private Sub �̍ك��R()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     �쐬���ꂽ���o���ꗗ�̑̍ق𒲐�����B���������𕔕�������͈͂Ɋ܂߂�B
'
' �y�����z
'     2015/01/28�@����Ł@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
    With xls
    
      '��̕��A�s�̍����A�܂�Ԃ��ݒ�A�t�H���g����
        .Range(.Cells(1, 1), .Cells(1, JyouColumn - 1)).ColumnWidth = 1.5
        .Range(.Cells(1, JyouColumn), .Cells(1, JyouColumn)).ColumnWidth = 50
        .Range(.Cells(1, JyouColumn + 1), .Cells(1, JyouColumn + 1)).ColumnWidth = 70
        .Range(.Cells(1, 1), .Cells(j, 1)).RowHeight = 17.25
        
        .Range(.Cells(1, JyouColumn), .Cells(j, JyouColumn + 1)).WrapText = True '���y�ь��o���̃Z����܂�Ԃ��\��

        .Range(.Cells(1, 1), .Cells(j, JyouColumn + 1)).Font.Name = "�l�r �S�V�b�N"
        .Range(.Cells(1, 1), .Cells(j, JyouColumn)).Font.Size = 9
        .Range(.Cells(1, JyouColumn + 1), .Cells(j, JyouColumn + 1)).Font.Size = 8
        
      '����̌����i�����j
        .ActiveSheet.PageSetup.Orientation = xlLandscape
        
      '����͈͂̐ݒ�
        .ActiveSheet.PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(j - 1, JyouColumn + 1)).Address
        
      '�s�̍����̒���
        .Range(.Cells(1, 1), .Cells(j - 1, 1)).EntireRow.AutoFit
        
      '�w�b�_�[�A�t�b�^�[
        .ActiveSheet.PageSetup.LeftHeader = "�y���o���ꗗ�z " & .Cells(1, 1).Value
        .ActiveSheet.PageSetup.RightHeader = "&10&P / &N"
        .Cells(1, JyouColumn + 2).Value = Now()
        .Cells(1, JyouColumn + 2).NumberFormat = "yyyy/MM/dd H:mm:ss"
        .ActiveSheet.PageSetup.RightFooter = "&07" & " " & .Cells(1, JyouColumn + 2).Text
        
    End With
    
End Sub

Private Sub �r���^�e()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     �쐬���ꂽ���o���ꗗ�Ɍr���������B
'
' �y�����z
'     2015/01/28�@����Ł@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
    With xls

   '�͖�������ꍇ�A���o���ꗗ�̖{�������Ɍr��������
        BorderEdgeTopRow = 0
        If JyouColumn >= 3 Then
            '�Q��ڂ̏���
            For i = 1 To j
                If Left(.Cells(i, 2), 1) = "��" Then
                    BorderEdgeTopRow = i
                    Exit For
                End If
            Next
            ListUpperEnd = BorderEdgeTopRow
            
            For i = ListUpperEnd + 1 To j
                If Left(xls.Cells(i, 2), 1) = "��" Then
                    .Range(.Cells(BorderEdgeTopRow, 2), .Cells(i - 1, JyouColumn)).BorderAround _
                    LineStyle:=xlContinuous, Weight:=xlThin  'xlHairline
                    BorderEdgeTopRow = i
                End If
                If Left(.Cells(i, 2), 1) = "��" Then
                    .Range(.Cells(BorderEdgeTopRow, 2), .Cells(i - 1, JyouColumn)).BorderAround _
                    LineStyle:=xlContinuous, Weight:=xlThin  'xlHairline
                    HonsokuEndRow = i - 1
                    Exit For
                End If
            Next
            

            If JyouColumn >= 4 Then
                '�R��ڂ̏���
                LineOnOff = 0
                For i = 1 To HonsokuEndRow
                    If Left(.Cells(i, 3), 1) = "��" Then
                         .Range(xls.Cells(i, 3), .Cells(i, JyouColumn)).Borders(xlEdgeTop).LineStyle = xlContinuous
                         .Range(xls.Cells(i, 3), .Cells(i, JyouColumn)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                         LineOnOff = 1
                    End If
                    
                    If Left(.Cells(i, 2), 1) = "��" Then
                        LineOnOff = 0
                    End If
                    
                    If LineOnOff = 1 Then
                        .Cells(i, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Cells(i, 3).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                    End If
                Next
                
                If JyouColumn >= 5 Then
                    '�S��ڂ̏���
                    LineOnOff = 0
                    For i = 1 To HonsokuEndRow
                        If Left(.Cells(i, 4), 1) = "��" Then
                            .Range(.Cells(i, 4), .Cells(i, JyouColumn)).Borders(xlEdgeTop).LineStyle = xlContinuous
                            .Range(.Cells(i, 4), .Cells(i, JyouColumn)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                            LineOnOff = 1
                        End If
                    
                        If Left(.Cells(i, 2), 1) = "��" Or Left(.Cells(i, 3), 1) = "��" Then
                            LineOnOff = 0
                        End If
                        
                        If LineOnOff = 1 Then
                            .Cells(i, 4).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Cells(i, 4).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                        End If
                    Next
                    
                    If JyouColumn >= 6 Then
                        '�T��ڂ̏���
                        LineOnOff = 0
                        For i = 1 To HonsokuEndRow
                            If Left(.Cells(i, 5), 1) = "��" Then
                                .Range(.Cells(i, 5), .Cells(i, JyouColumn)).Borders(xlEdgeTop).LineStyle = xlContinuous
                                .Range(.Cells(i, 5), .Cells(i, JyouColumn)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                                LineOnOff = 1
                            End If
                            
                            If Left(.Cells(i, 2), 1) = "��" Or Left(.Cells(i, 3), 1) = "��" _
                            Or Left(.Cells(i, 4), 1) = "��" Then
                                LineOnOff = 0
                            End If
                           
                            If LineOnOff = 1 Then
                                .Cells(i, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
                                .Cells(i, 5).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                            End If
                        Next
                        
                        If JyouColumn >= 7 Then
                            '�U��ڂ̏���
                            LineOnOff = 0
                            For i = 1 To HonsokuEndRow
                                If Left(.Cells(i, 6), 1) = "��" Then
                                    .Range(.Cells(i, 6), .Cells(i, JyouColumn)).Borders(xlEdgeTop).LineStyle = xlContinuous
                                    .Range(.Cells(i, 6), .Cells(i, JyouColumn)).Borders(xlEdgeTop).Weight = xlThin  'xlHairline
                                    LineOnOff = 1
                                End If
                    
                                If Left(.Cells(i, 2), 1) = "��" Or Left(.Cells(i, 3), 1) = "��" _
                                Or Left(.Cells(i, 4), 1) = "��" Or Left(.Cells(i, 5), 1) = "��" Then
                                    LineOnOff = 0
                                End If
                                
                                If LineOnOff = 1 Then
                                    .Cells(i, 6).Borders(xlEdgeLeft).LineStyle = xlContinuous
                                    .Cells(i, 6).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                                End If
                            Next
                        End If
                    End If
                End If
            End If
            
            '���̗�
            For i = 2 To HonsokuEndRow
                If .Cells(i, JyouColumn) <> "" And .Cells(i - 1, JyouColumn) = "" Then
                    .Cells(i, JyouColumn).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Cells(i, JyouColumn).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                End If
                
                If .Cells(i, JyouColumn) <> "" Then
                    .Cells(i, JyouColumn).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Cells(i, JyouColumn).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                End If
            Next
        End If
        .Visible = True
    End With

End Sub

Private Sub �r�����R()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     �쐬���ꂽ���o���ꗗ�Ɍr���������B
'
' �y�����z
'     2015/01/28�@����Ł@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-'

    With xls
        
      '�͖�������ꍇ�A���o���ꗗ�̖{�������Ɍr��������
        BorderEdgeTopRow = 0
        If JyouColumn >= 3 Then
            '�Q��ڂ̏���
            For i = 1 To j
                If Left(.Cells(i, 2), 1) = "��" Then
                    BorderEdgeTopRow = i
                    Exit For
                End If
            Next
            ListUpperEnd = BorderEdgeTopRow
            
            For i = ListUpperEnd + 1 To j
                If Left(.Cells(i, 2), 1) = "��" Then
                    .Range(.Cells(BorderEdgeTopRow, 2), .Cells(i - 1, JyouColumn + 1)).BorderAround _
                    LineStyle:=xlContinuous, Weight:=xlThin  'xlHairline
                    BorderEdgeTopRow = i
                End If
                If Left(.Cells(i, 2), 1) = "��" Then
                    .Range(.Cells(BorderEdgeTopRow, 2), .Cells(i - 1, JyouColumn + 1)).BorderAround _
                    LineStyle:=xlContinuous, Weight:=xlThin  'xlHairline
                    HonsokuEndRow = i - 1
                    Exit For
                End If
            Next

            If JyouColumn >= 4 Then
                '�R��ڂ̏���
                LineOnOff = 0
                For i = 1 To HonsokuEndRow
                    If Left(.Cells(i, 3), 1) = "��" Then
                         .Range(.Cells(i, 3), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                         .Range(.Cells(i, 3), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                         LineOnOff = 1
                    End If
                    
                    If Left(.Cells(i, 2), 1) = "��" Then
                        LineOnOff = 0
                    End If
                    
                    If LineOnOff = 1 Then
                        .Cells(i, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Cells(i, 3).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                    End If
                Next
                
                If JyouColumn >= 5 Then
                    '�S��ڂ̏���
                    LineOnOff = 0
                    For i = 1 To HonsokuEndRow
                        If Left(.Cells(i, 4), 1) = "��" Then
                            .Range(.Cells(i, 4), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                            .Range(.Cells(i, 4), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                            LineOnOff = 1
                        End If
                    
                        If Left(.Cells(i, 2), 1) = "��" Or Left(.Cells(i, 3), 1) = "��" Then
                            LineOnOff = 0
                        End If
                        
                        If LineOnOff = 1 Then
                            .Cells(i, 4).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Cells(i, 4).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                        End If
                    Next
                    
                    If JyouColumn >= 6 Then
                        '�T��ڂ̏���
                        LineOnOff = 0
                        For i = 1 To HonsokuEndRow
                            If Left(.Cells(i, 5), 1) = "��" Then
                                .Range(.Cells(i, 5), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                                .Range(.Cells(i, 5), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                                LineOnOff = 1
                            End If
                            
                            If Left(.Cells(i, 2), 1) = "��" Or Left(.Cells(i, 3), 1) = "��" _
                            Or Left(.Cells(i, 4), 1) = "��" Then
                                LineOnOff = 0
                            End If
                           
                            If LineOnOff = 1 Then
                                .Cells(i, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
                                .Cells(i, 5).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                            End If
                        Next
                        
                        If JyouColumn >= 7 Then
                            '�U��ڂ̏���
                            LineOnOff = 0
                            For i = 1 To HonsokuEndRow
                                If Left(.Cells(i, 6), 1) = "��" Then
                                    .Range(.Cells(i, 6), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                                    .Range(.Cells(i, 6), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).Weight xlThin 'xlHairline
                                    LineOnOff = 1
                                End If
                    
                                If Left(.Cells(i, 2), 1) = "��" Or Left(.Cells(i, 3), 1) = "��" _
                                Or Left(.Cells(i, 4), 1) = "��" Or Left(.Cells(i, 5), 1) = "��" Then
                                    LineOnOff = 0
                                End If
                                
                                If LineOnOff = 1 Then
                                    .Cells(i, 6).Borders(xlEdgeLeft).LineStyle = xlContinuous
                                    .Cells(i, 6).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                                End If
                            Next
                        End If
                    End If
                End If
            End If
            
            '���̗�
            For i = 2 To HonsokuEndRow
                If .Cells(i, JyouColumn) <> "" And .Cells(i - 1, JyouColumn) = "" Then
                    .Range(.Cells(i, JyouColumn), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range(.Cells(i, JyouColumn), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                End If
                
                If .Cells(i, JyouColumn) <> "" Then
                    .Range(.Cells(i, JyouColumn), .Cells(i, JyouColumn + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Range(.Cells(i, JyouColumn), .Cells(i, JyouColumn + 1)).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                End If
            Next
            
            '�P�s�����ɏ��ɐF�t��
            For i = 2 To HonsokuEndRow
                If .Cells(i, JyouColumn) <> "" Then
                    .Range(.Cells(i, JyouColumn), .Cells(i, JyouColumn + 1)).Interior.ColorIndex = 15
                i = i + 1
                End If
            Next

        '�c���i���D�݂Łj
        '.Range(.Cells(ListUpperEnd, JyouColumn), .Cells(HonsokuEndRow, JyouColumn)).Borders(xlEdgeRight).LineStyle = xlContinuous
        '.Range(.Cells(ListUpperEnd, JyouColumn), .Cells(HonsokuEndRow, JyouColumn)).Borders(xlEdgeRight).Weight = xlThin
        End If

    End With

End Sub
