Attribute VB_Name = "C_���r�ꗗ�쐬_20150128"
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
'�@�ߎ����x���}�N���iWord�Łj(3/3)�@C_���r�ꗗ�쐬_20150128
'Copyright(C)2014 Blue Panda
'This software is released under the MIT License.
'http://opensource.org/licenses/mit-license.php
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
Option Explicit
    Dim fRng As Range
    Dim Jyou As String '�Y�������
    Dim Kou As String '�Y�����鍀
    Dim Gou As String '�Y�����鍆
    Dim Saibun1 As String '�Y�����鍆�̍ו� �C���n
    Dim Saibun2 As String '�Y�����鍆�̍ו� (1)(2)(3)
    Dim Saibun3 As String '�Y�����鍆�̍ו� (i)(ii)(iii)
   
Sub ���r�ꗗ�쐬()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     ���傤�����́uSuper�@��Web�v����o�͂��ꂽ�𕶃f�[�^���烋�r�̕t���ꂽ�����̈ꗗ���쐬����B
'
' �y�����z
'     2015/01/28�@����ō쐬
'
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y���l�z
' �@�@�������̏����ł����Ă��u�����v�Ƃ̕\�L�͂���܂���B
'�@�@ �uSuper�@��Web�v�ȊO�̃f�[�^�ł́A���r�̕\���`���̍��قɂ��A�������o�͂���܂���B
'
' �y���Ή��̓��e���z
' �@�@����Ȉʒu�̃��r�ɂ��ẮA�������\������Ȃ����ꂪ����܂��B
' �@�@���̍ו����S�i�K�̂��̂ɂ͑Ή��ł��܂���B
' �@�@�ꍀ�Ȃ琬��{���ɂ͑Ή��ł��܂���B
'
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '�ϐ��̐錾
    Dim xls As Object
    Dim myfield As Object
    Dim mycode As String
    Dim j As Integer
    Dim myRange As Range
    
  '�C���f���g����
    Call �C���f���g����

    Set myRange = ActiveDocument.Sentences(1)
    
  'Excel�̃V�[�g���쐬
    Set xls = CreateObject("Excel.Application")
   
    With xls
        .SheetsInNewWorkbook = 1
        .Workbooks.Add
    End With

  '�^�C�g�����쐬
    With xls
        .Cells(1, 1).Value = "�y���r�ꗗ�z " & myRange
        .Cells(2, 2).Value = "���r�t���̕���"
        .Cells(2, 3).Value = "�Y�������@�i���j�����̏������܂ޏꍇ������܂��B"
        .Cells(2, 4).Value = "�y�[�W"
    End With
    
    j = 3
    
  '�G���[���̏���
    On Error GoTo 0
    
  '���r�̏���]�L
    For Each myfield In ActiveDocument.Fields
        
        mycode = myfield.Code
        myfield.Select
        
      '���r
        xls.Cells(j, 2).Value = ruby(mycode)
        
      '���r�̈ʒu�i�����j
        Call �����ʒu�\��
        xls.Cells(j, 3).Value = Jyou & " " & Kou & " " & Gou & " " & Saibun1 & " " & Saibun2 & " " & Saibun3

      '���r�̈ʒu�i�y�[�W�j
        xls.Cells(j, 4).Value = Selection.Information(wdActiveEndPageNumber)
        
        j = j + 1

    Next
    
    With xls
       '��̕��A�s�̍����A�t�H���g����
        .Range(.Cells(1, 1), .Cells(1, 1)).ColumnWidth = 3
        .Range(.Cells(1, 2), .Cells(1, 2)).ColumnWidth = 25
        .Range(.Cells(1, 3), .Cells(1, 3)).ColumnWidth = 48
        .Range(.Cells(1, 4), .Cells(1, 4)).ColumnWidth = 8
        .Range(.Cells(1, 1), .Cells(j - 1, 1)).RowHeight = 30
        
        .Range(.Cells(1, 1), .Cells(1, 1)).Font.Size = 12
        .Range(.Cells(1, 1), .Cells(1, 1)).Font.Bold = True
        .Range(.Cells(2, 1), .Cells(j - 1, 4)).Font.Size = 11
        .Range(.Cells(2, 2), .Cells(2, 4)).Font.Bold = True
        
        .Range(.Cells(2, 2), .Cells(j - 1, 3)).IndentLevel = 1
        .Range(.Cells(2, 4), .Cells(j - 1, 4)).HorizontalAlignment = xlCenter
    
        With .Range(xls.Cells(2, 2), .Cells(j - 1, 4))
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
        
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin

            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
        
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
        
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Weight = xlThin
        
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlThin
        End With

        .Range(.Cells(2, 2), .Cells(2, 4)).Interior.ColorIndex = 15
        
      '�\��
        .Visible = True
        
    End With

End Sub
 
Private Sub �C���f���g����()

'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'�@�@ �C���f���g��11�|�C���g�̕��������ɕϊ�����B
'
' �y�����z
'     2015/01/28�@����ō쐬
'
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

    Dim FirstFontSize As Single, ChangedFontSize As Single
    Dim Indenting As Single, Indented As Single
    Dim i As Integer
    
    '�C���f���g����
    If ActiveDocument.Range(0, 0).Font.Size = 11 Then  '��11�|�C���g�̏ꍇ�̓C���f���g�����s�v
    Else
        FirstFontSize = ActiveDocument.Range(0, 0).Font.Size  '�����̃t�H���g�T�C�Y���擾
        ChangedFontSize = 11  '�Q�Ə𕶖{���̃t�H���g�T�C�Y��ݒ�
    '2�s�ڈȍ~�̃C���f���g�iLeftIndent�j�𒲐�
       
        For i = 1 To 10
  
            Indenting = i * FirstFontSize
            Indented = i * ChangedFontSize
            
            Set fRng = ActiveDocument.Range(Start:=0, End:=0)
            fRng.Find.ParagraphFormat.LeftIndent = Indenting
            Do
                If fRng.Find.Execute = False Then Exit Do
                fRng.ParagraphFormat.LeftIndent = Indented
                fRng.Collapse Direction:=wdCollapseEnd
            Loop
        Next i
    
    '1�s�ڂ̃C���f���g�iFirstLineIndent�j�𒲐�
       
        For i = -4 To -1
            Indenting = i * FirstFontSize
            Indented = i * ChangedFontSize
            
            Set fRng = ActiveDocument.Range(Start:=0, End:=0)
            fRng.Find.ParagraphFormat.FirstLineIndent = Indenting
            Do
                If fRng.Find.Execute = False Then Exit Do
                fRng.ParagraphFormat.FirstLineIndent = Indented
                fRng.Collapse Direction:=wdCollapseEnd
            Loop
        Next i
    End If
    
End Sub

Function ruby(codetext)

'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'�@�@ ���r�̂����������烋�r�����o���B
'
'�@�@�@����
'�@�@�@�����@���@�����i�����j
'
' �y�����z
'     2015/01/28�@����ō쐬
'
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

    Dim i, pos As Long
    Dim rubybase, rubytext, t As String
    
    pos = 0
    rubybase = ""
    rubytext = ""
    
    For i = 1 To Len(codetext)
        t = Mid(codetext, i, 1)
        If t = "(" Or t = ")" Or t = "," Then
            pos = pos + 1
        ElseIf pos = 2 Then
            rubytext = rubytext & t
        ElseIf pos = 4 Then
            rubybase = rubybase & t
        End If
    Next i
    
    ruby = rubybase & " �i" & rubytext & "�j"
     
End Function

Private Sub �����ʒu�\��()

'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'�@�@ �w��ʒu�̏�����\�����܂��B
'
'
' �y�����z
'     2015/01/28�@����ō쐬
'
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
    Dim Syozoku As String '�Y������v�f
    Dim a As Integer
    Dim i As Integer
    Dim JyouKou1 As String '�����̉��̒l
    
    '�l���N���A
    Jyou = ""
    Kou = ""
    Gou = ""
    Saibun1 = ""
    Saibun2 = ""
    Saibun3 = ""
    
    '�Y���ӏ��̈ʒu�̕\��

    Set fRng = ActiveDocument.Range(Start:=0, End:=Selection.Range.End)
    a = fRng.Paragraphs.Count
    Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a).Range.Start, End:=ActiveDocument.Paragraphs(a).Range.End) 'Selection.Range

    '���߂̗v�f
    If InStr(fRng, "�@") > 0 Then
        Syozoku = Left(fRng, (InStr((fRng), "�@") - 1))
        JyouKou1 = Syozoku
    Else
        '�w��̈ʒu�����o���̏ꍇ
        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = 0 Then
            Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a + 1).Range.Start, End:=ActiveDocument.Paragraphs(a + 1).Range.End)
            JyouKou1 = Left(fRng, (InStr((fRng), "�@") - 1))
            Jyou = JyouKou1 & "���o��"
        End If
    End If
        
        'Case�P�@�w��̈ʒu�����y�э��̏ꍇ
        If InStr(Left(fRng, (InStr((fRng), "�@") - 1)), "��") > 0 And Jyou = "" Then
                Kou = "��ꍀ"
        Else
            If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 And Jyou = "" Then
                Kou = "��" & Syozoku & "��"
            End If
             
        End If
        
        'Case�Q�@�w��̈ʒu���͖����̏ꍇ
        If fRng.ParagraphFormat.FirstLineIndent = -44 Then
                Jyou = "�͖���"
        End If
        
        'Casa�R�@�w��̈ʒu�����̏ꍇ
        If fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
            Gou = "��" & Syozoku & "��"
            i = 1
            Do Until (fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                
                If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    If InStr(Left(fRng, (InStr((fRng), "�@") - 1)), "��") > 0 Then
                        Kou = "��ꍀ"
                    Else
                        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                            JyouKou1 = Left(fRng, (InStr((fRng), "�@") - 1))
                            Kou = "��" & JyouKou1 & "��"
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If
             
        'Case�S�@�w��̈ʒu�����̍ו�1�i��ɃC���n�j�̏ꍇ
        If fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
            Saibun1 = Syozoku
            i = 1
            Do Until (fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    JyouKou1 = Left(fRng, (InStr((fRng), "�@") - 1))
                    Gou = "��" & JyouKou1 & "��"
                End If
                i = i + 1
            Loop
            
            Do Until (fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                
                If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    If InStr(Left(fRng, (InStr((fRng), "�@") - 1)), "��") > 0 Then
                        Kou = "��ꍀ"
                    Else
                        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                            JyouKou1 = Left(fRng, (InStr((fRng), "�@") - 1))
                            Kou = "��" & JyouKou1 & "��"
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If
        
        'Case�T�@�w��̈ʒu�����̍ו�2�i���(1)(2)(3)�j�̏ꍇ
        If fRng.ParagraphFormat.LeftIndent = 44 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
            Saibun2 = Syozoku
            i = 1
            '���̍ו�1
            Do Until (fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    Saibun1 = Left(fRng, (InStr((fRng), "�@") - 1))
                End If
                i = i + 1
            Loop
            '��
            i = 1
            Do Until (fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    JyouKou1 = Left(fRng, (InStr((fRng), "�@") - 1))
                    Gou = "��" & JyouKou1 & "��"
                End If
                i = i + 1
            Loop
            '��
            i = 1
            Do Until (fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                
                If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    If InStr(Left(fRng, (InStr((fRng), "�@") - 1)), "��") > 0 Then
                        Kou = "��ꍀ"
                    Else
                        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                            JyouKou1 = Left(fRng, (InStr((fRng), "�@") - 1))
                            Kou = "��" & JyouKou1 & "��"
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If
        
   
        'Case�U�@�w��̈ʒu�����̍ו�3�i���(i)(ii)(iii)�j�̏ꍇ
        If fRng.ParagraphFormat.LeftIndent = 55 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
            Saibun3 = Syozoku
            i = 1
            '���̍ו�2
            Do Until (fRng.ParagraphFormat.LeftIndent = 55 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 44 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    Saibun1 = Left(fRng, (InStr((fRng), "�@") - 1))
                End If
                i = i + 1
            Loop
            '���̍ו�1
            Do Until (fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    Saibun1 = Left(fRng, (InStr((fRng), "�@") - 1))
                End If
                i = i + 1
            Loop
            '��
            Do Until (fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    JyouKou1 = Left(fRng, (InStr((fRng), "�@") - 1))
                    Gou = "��" & JyouKou1 & "��"
                End If
                i = i + 1
            Loop
            '��
            Do Until (fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                
                If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    If InStr(Left(fRng, (InStr((fRng), "�@") - 1)), "��") > 0 Then
                        Kou = "��ꍀ"
                    Else
                        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                            JyouKou1 = Left(fRng, (InStr((fRng), "�@") - 1))
                            Kou = "��" & JyouKou1 & "��"
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If

             
    '�����̏�
    If Jyou = "" Then
        Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a).Range.Start, End:=ActiveDocument.Paragraphs(a).Range.End)
        If InStr(Left(fRng, (InStr((fRng), "�@") - 1)), "��") > 0 Then
            Jyou = JyouKou1
        Else
            i = 1
            Do Until InStr(JyouKou1, "��") > 0 Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If InStr(fRng, "�@") > 0 Then
                    JyouKou1 = Left(fRng, (InStr((fRng), "�@") - 1))
                End If
                i = i + 1
            Loop
            Jyou = JyouKou1
        End If
    End If
        
End Sub
