Attribute VB_Name = "A_�Q�Ə𕶍쐬��_20150128"
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
'�@�ߎ����x���}�N���iWord�Łj(1/3)�@A_�Q�Ə𕶍쐬��_20150128
'Copyright(C)2014 Blue Panda
'This software is released under the MIT License.
'http://opensource.org/licenses/mit-license.php
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

Sub �Q�Ə𕶍쐬()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     ���傤�����́uSuper�@��Web�v���͑��@�K�uD1-Law.com�v�i���@�K �@��񑍍��f�[�^�x�[�X�j
' �@�@����o�͂��ꂽ�𕶃f�[�^�����s���{�@�K���̗l���ɒ�������B
'
' �y�����z
'     2014/10/29�@ver.1.01 �s�v�ȃR�[�h�̍폜�A�ǐ�����̂��߂̃R�[�h�̐����A
'�@�@�@�@�@�@�@�@�@�@�@�@�@�t�b�^�[�̃o�O������
' �@�@2014/10/01�@ver.1.0
' �@�@2014/09/12�@�b��ō쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y���l�z
' �@�@�C���f���g�̏��̓_�E�����[�h�����t�@�C���̏��𗘗p���Ă��܂��B
'
' �y���Ή��̓��e���z
' �@�@�uD1-Law.com�v�̃��r�����́A���f�[�^�̎d�l�ɂ��A���ʏ����ŕ\������܂��B
' �@�@�ꕔ�����@�ߗ߂̏𕶂ł̎g�p�͑z�肵�Ă��܂���B
'
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '�ϐ��̐錾
    Dim FirstFontSize As Single, ChangedFontSize As Single
    Dim Indenting As Single, Indented As Single
    Dim fRng As Range
    Dim i As Integer
    Dim sWrds() As Variant, rWrds() As Variant
    Dim Num As Integer
    
  '������h�~�i�}�N���̍������B�����ŉ����B�j���}�N���̓�����������邽�߂����č쓮�����Ă��Ȃ��B
        'Application.ScreenUpdating = False
    
  '�w�b�_�[���̓t�b�^�[�̕ҏW��ʂɂȂ��Ă���ꍇ�ւ̑Ή�
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
 
  '[1a] �C���f���g�̒����i11�|�C���g�̃t�H���g�̃C���f���g�ɒ����j
        Call �C���f���g����110
    
  '[2a] ���p�X�y�[�X�̏����i�Q�A�����锼�p�X�y�[�X�͑S�p�ɒu�����A�P�P�Ƃ̂��͖̂����j
  
        Call ���p�X�y�[�X����_�S��
    
  '[3a] ���ԍ����̏c�����̏������s���B
  
        Call ���ԍ��c����������_�S��

  '[4a] �̍ق̒����i���s���{�@�K���̑̍فj
  
        Call �ْ̍���_�S��

  '[5a] �𖼂̏�ԍ��y�я͖������S�V�b�N�ɂ���B
  
        Call �𖼓��S�V�b�N_�S��

  '[6a] �w�b�_�[�y�уt�b�^�[�ɍ쐬�����A�@�ߖ��y�уy�[�W�ԍ���������B�i�y�[�W�ԍ��͍��̂݁j
  
        Call �w�b�_�[�t�b�^�[����_�Ж�
      
  '�����̈ʒu�ɖ߂�B
        ActiveDocument.Range(0, 0).Select

  '������h�~����
        'Application.ScreenUpdating = True
    
End Sub
Sub �Q�Ə𕶍쐬_���ʈ���p()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     �}�N���u�Q�Ə𕶍쐬�v�̃w�b�_�[�y�уt�b�^�[�𗼖ʈ���ɓK�����`�ɂ������́B
'
' �y�����z
'�@�@ 2015/01/28�@ver.1.02 �t�b�^�[�̃o�O������
'     2014/10/29�@ver.1.01 �s�v�ȃR�[�h�̍폜�A�ǐ�����̂��߂̃R�[�h�̐����A
'�@�@�@�@�@�@�@�@�@�@�@�@�@�t�b�^�[�̃o�O������
' �@�@2014/10/01�@ver.1.0�@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

    Dim FirstFontSize As Single, ChangedFontSizet As Single
    Dim Indenting As Single, Indented As Single
    Dim fRng As Range
    Dim i As Integer
    Dim sWrds() As Variant, rWrds() As Variant
    Dim Num As Integer
    
  '�w�b�_�[���̓t�b�^�[�̕ҏW��ʂɂȂ��Ă���ꍇ�ւ̑Ή�
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
  
  '[1a] �C���f���g�̒����i11�|�C���g�̃t�H���g�̃C���f���g�ɒ����j
        Call �C���f���g����110
    
  '[2a] ���p�X�y�[�X�̏����i�Q�A�����锼�p�X�y�[�X�͑S�p�ɒu�����A�P�P�Ƃ̂��͖̂����j
  
        Call ���p�X�y�[�X����_�S��
    
  '[3a] ���ԍ����̏c�����̏������s���B
  
        Call ���ԍ��c����������_�S��
 
  '[4a] �̍ق̒����i���s���{�@�K���̑̍فj
  
        Call �ْ̍���_�S��

  '[5a] �𖼂̏�ԍ��y�я͖������S�V�b�N�ɂ���B
  
        Call �𖼓��S�V�b�N_�S��

  '[6b] �w�b�_�[�y�уt�b�^�[�ɍ쐬�����A�@�ߖ��y�уy�[�W�ԍ���������B�i�y�[�W�ԍ��͍��E���c�̏��j
  
        Call �w�b�_�[�t�b�^�[����_����
    
  '�����̈ʒu�ɖ߂�B
        ActiveDocument.Range(0, 0).Select

  '������h�~����
        'Application.ScreenUpdating = True
        
End Sub

Sub �V�������H�p����()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'     ���傤�����́uSuper�@��Web�v���͑��@�K�uD1-Law.com�v�i���@�K �@��񑍍��f�[�^�x�[�X�j
'     ����o�͂��ꂽ�𕶃f�[�^���A�V���Ώƕ\���̏c�����̗l���̏��ނ̍쐬�ɓK�����`�ɉ��H����B
'
' �y�����z
'     2014/10/29�@ver.1.0�@ �V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

    Dim FirstFontSize As Single, ChangedFontSizet As Single
    Dim Indenting As Single, Indented As Single
    Dim fRng As Range
    Dim i As Integer
    Dim sWrds() As Variant, rWrds() As Variant
    Dim Num As Integer
    
  '�w�b�_�[���̓t�b�^�[�̕ҏW��ʂɂȂ��Ă���ꍇ�ւ̑Ή�
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
  
  '[1b] �C���f���g�̒����i11�|�C���g�̃t�H���g�̃C���f���g�ɒ����j
        Call �C���f���g����105
    
  '[2a] ���p�X�y�[�X�̏����i�Q�A�����锼�p�X�y�[�X�͑S�p�ɒu�����A�P�P�Ƃ̂��͖̂����j
  
        Call ���p�X�y�[�X����_�S��
    
  '[3a] ���ԍ����̏c�����̏������s���B
  
        Call ���ԍ��c����������_�S��
 
  '[6a] �w�b�_�[�y�уt�b�^�[�ɍ쐬�����A�@�ߖ��y�уy�[�W�ԍ���������B�i�y�[�W�ԍ��͍��̂݁j
  
        Call �w�b�_�[�t�b�^�[����_�Ж�
    
  '�{���̃t�H���g�̒���
        Selection.WholeStory
        With Selection
            .Font.Name = "�l�r ����"
            .Font.Size = 10.5
            .Orientation = wdTextOrientationHorizontal
        End With
    
  '�i�g�i�P�i�j
        ActiveDocument.PageSetup.TextColumns.SetCount NumColumns:=1

  '�p���ݒ�
    With ActiveDocument.PageSetup
        .TopMargin = MillimetersToPoints(10)
        .BottomMargin = MillimetersToPoints(10)
        .LeftMargin = MillimetersToPoints(20)
        .RightMargin = MillimetersToPoints(20)
        .HeaderDistance = MillimetersToPoints(5)
        .FooterDistance = MillimetersToPoints(5)
    End With

  '�\�̈ʒu
        Dim tbl As Table
        For Each tbl In ActiveDocument.Tables
            With tbl
                .Style = "�\ (�i�q)"
                .Rows.Alignment = wdAlignRowCenter
                .AutoFitBehavior (wdAutoFitContent)
                .PreferredWidthType = wdPreferredWidthPercent
                .PreferredWidth = 95
            End With
        Next
    
  '[7] ���r�����̊m�F
  
        Call ���ʃ��r�m�F
    
  '�����̈ʒu�ɖ߂�B
        ActiveDocument.Range(0, 0).Select

        
End Sub

Sub e_Gov�𕶏���_�S��_�^�e()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@ e-Gov����R�s�[���Ă����𕶂̑̍ق��A�c�����̑̍قɐ�����B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
    
    Dim aFld As Field
    Dim para As Paragraph
    Dim fRng As Range
    Dim fsnormal As Single
    Dim RE As Object, Matches As Object

  '[9a] e-Gov�̃t�H�[�}�b�g������������B

        Call e_Gov�t�H�[�}�b�g������_�S��
    
  '[10b] �I��͈͂̋󔒍s�y�ѐ����^�u�̍폜
    
        Call �󔒍s���폜_�I��͈�

  '[2a] ���p�X�y�[�X�̏����i�Q�A�����锼�p�X�y�[�X�͑S�p�ɒu�����A�P�P�Ƃ̂��͖̂����j
  
        Call ���p�X�y�[�X����_�S��

  '[3a] ���ԍ����̏c�����̏������s���B
  
        Call ���ԍ��c����������_�S��

  '[8a] ���̌`����K�p���ׂ��@�߂̎������𐄒肵�A�K�p����B

        Call ������_�S��
    
  '�����̌������u�������i���{�ꕶ��������90�x��]�j�v�ɂ���B
        Selection.WholeStory
        Selection.Orientation = wdTextOrientationHorizontalRotatedFarEast

End Sub

Sub e_Gov�𕶏���_�S��_���R()
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@e-Gov����R�s�[���Ă����𕶂̑̍ق��A�������̑̍قɐ�����B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
  
    Dim aFld As Field
    Dim para As Paragraph
    Dim fRng As Range
    Dim fsnormal As Single
    Dim RE As Object, Matches As Object
    
  '[9a] e-Gov�̃t�H�[�}�b�g������������B

        Call e_Gov�t�H�[�}�b�g������_�S��
    
  '[10b] �I��͈͂̋󔒍s�y�ѐ����^�u�̍폜
    
        Call �󔒍s���폜_�I��͈�

  '[2a] ���p�X�y�[�X�̏����i�Q�A�����锼�p�X�y�[�X�͑S�p�ɒu�����A�P�P�Ƃ̂��͖̂����j
  
        Call ���p�X�y�[�X����_�S��

  '[3b] ���ԍ����̏���
  
        Call ���ԍ�������_�S��_���R

  '[8a] ���̌`����K�p���ׂ��@�߂̎������𐄒肵�A�K�p����B

        Call ������_�S��

End Sub

Private Sub �C���f���g����110()
'[1a]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�S���z�C���f���g�̒����i11�|�C���g�̃t�H���g�̃C���f���g�ɒ����j
'�@�@�@�����傤�����́uSuper�@��Web�v���͑��@�K�uD1-Law.com�v�i���@�K �@��񑍍��f�[�^�x�[�X�j��p
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
    
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

Private Sub �C���f���g����105()
'[1b]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�S���z�C���f���g�̒����i10.5�|�C���g�̃t�H���g�̃C���f���g�ɒ����j
'�@�@�@�����傤�����́uSuper�@��Web�v���͑��@�K�uD1-Law.com�v�i���@�K �@��񑍍��f�[�^�x�[�X�j��p
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

    If ActiveDocument.Range(0, 0).Font.Size = 10.5 Then  '��10.5�|�C���g�̏ꍇ�̓C���f���g�����s�v
    Else
        FirstFontSize = ActiveDocument.Range(0, 0).Font.Size  '�����̃t�H���g�T�C�Y���擾
        ChangedFontSize = 10.5  '�Q�Ə𕶖{���̃t�H���g�T�C�Y��ݒ�
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

Private Sub ���p�X�y�[�X����_�S��()
'[2a]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�S���z���p�X�y�[�X�̏����i�Q�A�����锼�p�X�y�[�X�͑S�p�ɒu�����A�P�P�Ƃ̂��͖̂����j
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

    With ActiveDocument.Content.Find
        .Text = "  "
        .Replacement.Text = "�@"
        .MatchByte = True
        .Execute Replace:=wdReplaceAll
    End With
    
    With ActiveDocument.Content.Find
        .Text = " "
        .Replacement.Text = ""
        .MatchByte = True
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub

Private Sub ���p�X�y�[�X����_�I��͈�()
'[2b]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�I��͈́z���p�X�y�[�X�̏����i�Q�A�����锼�p�X�y�[�X�͑S�p�ɒu�����A�P�P�Ƃ̂��͖̂����j
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

    Set fRng = Selection.Range
    With fRng.Find
        .Text = "  "
        .Replacement.Text = "�@"
        .MatchByte = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
    
    Set fRng = Selection.Range
    With fRng.Find
        .Text = " "
        .Replacement.Text = ""
        .MatchByte = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub

Sub ���ԍ��c����������_�S��()
'[3a]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@ �S���ɍ��ԍ����ɏc�������̏������s���B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�ǐ�����̂��߂̃R�[�h�̐���
'
' �y���ӎ����z
' �@�@�g�p����錳������������ꍇ�A�R�[�h�̏C�����K�v�ł��B
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '���p�����Ɣ��p���ʓ��w ()[]���� �x��S�p�ɂ���B
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Text = "[0-9)([\]����]{1,}"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.CharacterWidth = wdWidthFullWidth
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

  '���ʕt�������𔼊p�ɂ��A�c�����̏������s���B
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "�i[�O-�X]{1,}�j"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine '�����L������ꍇ �� wdHorizontalInVerticalResizeLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
  '���ԍ��̂Ȃ����̏����@�u���P�@�v���u�@�@�v�@���u���Q�O�@�v�܂�
    sWrds = Array("���P�@", "���Q�@", "���R�@", "���S�@", "���T�@", "���U�@", "���V�@", "���W�@", "���X�@", "���P�O�@" _
                , "���P�P�@", "���P�Q�@", "���P�R�@", "���P�S�@", "���P�T�@", "���P�U�@", "���P�V�@", "���P�W�@", "���P�X�@", "���Q�O�@")
    rWrds = Array("�@�@", "�A�@", "�B�@", "�C�@", "�D�@", "�E�@", "�F�@", "�G�@", "�H�@", "�I�@", _
                "�J�@", "�K�@", "�L�@", "�M�@", "�N�@", "�O�@", "�P�@", "�Q�@", "�R�@", "�S�@")
    With ActiveDocument.Content.Find
        For Num = LBound(sWrds) To UBound(sWrds)
            .Text = sWrds(Num)
            .Replacement.Text = rWrds(Num)
            .Execute Replace:=wdReplaceAll
        Next Num
    End With
    
  '���ԍ��̓񌅂̐����𔼊p�ɂ��A�c�����̏������s���B
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[�O-�X]{2}�@[!�@]"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
  '�Z�p�����ŕW�L����Ă���@�ߔԍ��y�єN�����̏����i�Q�`�R���̐����𔼊p�ɂ��A�c�����̏������s���B�j
  
   '�N�i���@��񑍍��f�[�^�x�[�X�̕\�L�ɑΉ����镔���j
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[�����a��][�O-�X]{2,3}�N"  '�@���@�g�p����錳�����������ہA���̕����ɉ��M�̕K�v������
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveStart wdCharacter, 1
        fRng.MoveEnd wdCharacter, -1
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
   '��
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "�N[�O-�X]{2}��"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveStart wdCharacter, 1
        fRng.MoveEnd wdCharacter, -1
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
   '��
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[�O-�X]{2}��"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveStart wdCharacter, 1
        fRng.MoveEnd wdCharacter, -1
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
   '�@�ߔԍ�
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[�O-�X]{2,3}��"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveStart wdCharacter, 1
        fRng.MoveEnd wdCharacter, -1
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

  '�ii�j�A�iii�j�A�iiii�j�̊��ʂ𔼊p�ɂ��A�c�����̏������s���B
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "�i[ivx]{1,}�j"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

End Sub

Private Sub ���ԍ�������_�S��_���R()
'[3b]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
'�@�@ �S���ɍ��ԍ����̔��p�������s���B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
  '���p�̐����A���ʓ���S�p�ɂ���B

    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[0-9)([\]����]{1,}"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.CharacterWidth = wdWidthFullWidth
        fRng.Collapse Direction:=wdCollapseEnd
        fRng.MoveEnd unit:=wdCharacter, Count:=-1
    Loop


  '�y���̍ו�(1)(2)(3)�z�S�Ă̊��ʕt�������𔼊p�ɂ���B
    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "�i[�O-�X]{1,}�j[!�@]"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        With fRng
            .MoveEnd unit:=wdCharacter, Count:=-1
            .CharacterWidth = wdWidthHalfWidth
            .Collapse Direction:=wdCollapseEnd
        End With
    Loop
    
    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "�i[�O-�X]{1,}�j�@"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        With fRng
            .MoveEnd unit:=wdCharacter, Count:=-1
            .CharacterWidth = wdWidthHalfWidth
            .Collapse Direction:=wdCollapseEnd
            With .ParagraphFormat
                .LeftIndent = 4 * fsnormal
                .FirstLineIndent = -1 * fsnormal
            End With
        End With
    Loop
        
   '�y���̍ו�(i)(ii)(iii)�z���ʂ𔼊p�ɂ���B
    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "�i[ivx]{1,}�j"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        With fRng
            .MoveEnd unit:=wdCharacter, Count:=-1
            .CharacterWidth = wdWidthHalfWidth
            .Collapse Direction:=wdCollapseEnd
        End With
    Loop
        
    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "�i[ivx]{1,}�j�@"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        With fRng
            .MoveEnd unit:=wdCharacter, Count:=-1
            .CharacterWidth = wdWidthHalfWidth
            .Collapse Direction:=wdCollapseEnd
            With .ParagraphFormat
                .LeftIndent = 5 * fsnormal
                .FirstLineIndent = -1 * fsnormal
            End With
        End With
    Loop
        
  '���ԍ��̂Ȃ����̏����@�u���P�@�v���u�@�@�v�@���u���Q�O�@�v�܂�
    sWrds = Array("���P�@", "���Q�@", "���R�@", "���S�@", "���T�@", "���U�@", "���V�@", "���W�@", "���X�@", "���P�O�@" _
                , "���P�P�@", "���P�Q�@", "���P�R�@", "���P�S�@", "���P�T�@", "���P�U�@", "���P�V�@", "���P�W�@", "���P�X�@", "���Q�O�@")
    rWrds = Array("�@�@", "�A�@", "�B�@", "�C�@", "�D�@", "�E�@", "�F�@", "�G�@", "�H�@", "�I�@", _
                "�J�@", "�K�@", "�L�@", "�M�@", "�N�@", "�O�@", "�P�@", "�Q�@", "�R�@", "�S�@")
    With ActiveDocument.Content.Find
        For Num = LBound(sWrds) To UBound(sWrds)
            .Text = sWrds(Num)
            .Replacement.Text = rWrds(Num)
            .Execute Replace:=wdReplaceAll
        Next Num
    End With
    
  '�y���ԍ��i10�`�j�z�S�Ă̍��ԍ��̓񌅂̐����𔼊p�ɂ���B
  '���i�Q���j�̎�����
    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[�O-�X]{2}�@[!�@]"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        With fRng
            .MoveEnd wdCharacter, -2
            .CharacterWidth = wdWidthHalfWidth
            .Collapse Direction:=wdCollapseEnd
            With .ParagraphFormat
                .LeftIndent = fsnormal
                .FirstLineIndent = -1 * fsnormal
            End With
        End With
    Loop

End Sub
Private Sub �ْ̍���_�S��()
'[4a]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�S���z�p���T�C�Y�y�уt�H���g�̑̍ق𒲐�����B
'�@�@�@���u���傤�����v�̌��s���{�@�K�i���}�́j���`�S�T�C�Y�Ɉ���������̂ɋ߂��̍قƂ��Ă���B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
  
  '�{���̃t�H���g�̒���
    Selection.WholeStory
    With Selection
        .Font.Name = "�l�r ����"
        .Font.Size = 11  'ChangedFontSize�Ɠ���̒l�Ƃ���B
        .Orientation = wdTextOrientationHorizontalRotatedFarEast '�������i���{�ꕶ��������90�x��]�j
    End With
    
  '�p���T�C�Y��A4�ɕύX
    ActiveDocument.PageSetup.PaperSize = wdPaperA4
        
  '�s��
    Selection.WholeStory
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20.2
    End With

  '�p���ݒ�
    With ActiveDocument.PageSetup
        .Orientation = wdOrientLandscape '�p������������
        .TopMargin = MillimetersToPoints(23)
        .BottomMargin = MillimetersToPoints(23)
        .LeftMargin = MillimetersToPoints(25)
        .RightMargin = MillimetersToPoints(25)
        .HeaderDistance = MillimetersToPoints(12.7)
        .FooterDistance = MillimetersToPoints(12.7)
        .LinesPage = 23 '23�s�i���ʓI�ɂP�s30�����j
        .LayoutMode = wdLayoutModeLineGrid
    End With
    
  '�i�g
    With ActiveDocument.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
        .LineBetween = True
        .Width = MillimetersToPoints(117.1)
        .Spacing = MillimetersToPoints(12.7)
    End With
    
  '�\�̈ʒu
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        With tbl
            .Style = "�\ (�i�q)"
            .Rows.Alignment = wdAlignRowCenter
            .AutoFitBehavior (wdAutoFitContent)
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 91
        End With
    Next
    
  '�@�ߖ��i�P�s�ځj�̃t�H���g�T�C�Y�̒���
    If ActiveDocument.Range(0, 1) = "��" Or ActiveDocument.Range(0, 1) = "��" Then
        With ActiveDocument.Sentences(1)
            .Font.Size = 15
            .ParagraphFormat.LeftIndent = 48
            .ParagraphFormat.FirstLineIndent = -15
        End With
    Else
        With ActiveDocument.Sentences(1)
            .Font.Size = 15
            .ParagraphFormat.LeftIndent = 33
            .ParagraphFormat.FirstLineIndent = 0
        End With
    End If
    
End Sub

Private Sub �𖼓��S�V�b�N_�S��()
'[5a]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�S���z�𖼂̏�ԍ����S�V�b�N�ɂ���B
'�@�@�@�����傤�����́uSuper�@��Web�v���͑��@�K�uD1-Law.com�v�i���@�K �@��񑍍��f�[�^�x�[�X�j��p
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}���@[!�@]"
        .ParagraphFormat.LeftIndent = 11
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}����[���O�l�ܘZ������\�S��]{1,}�@[!�@]"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}����[���O�l�ܘZ������\�S��]{1,}��[���O�l�ܘZ������\�S��]{1,}�@[!�@]"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
  '�͖������S�V�b�N�ɂ���B�i�������A�ڎ��͏����B�j
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}�ҁ@[!�@]"
        .ParagraphFormat.LeftIndent = 66
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}�́@[!�@]"
        .ParagraphFormat.LeftIndent = 77
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    

    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}�͂�[���O�l�ܘZ������\�S��]{1,}�@[!�@]"
        .ParagraphFormat.LeftIndent = 77
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}�߁@[!�@]"
        .ParagraphFormat.LeftIndent = 88
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    

    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}�߂�[���O�l�ܘZ������\�S��]{1,}�@[!�@]"
        .ParagraphFormat.LeftIndent = 88
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}���@[!�@]"
        .ParagraphFormat.LeftIndent = 99
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}����[���O�l�ܘZ������\�S��]{1,}�@[!�@]"
        .ParagraphFormat.LeftIndent = 99
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}�ځ@[!�@]"
        .ParagraphFormat.LeftIndent = 110
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "��[���O�l�ܘZ������\�S��]{1,}�ڂ�[���O�l�ܘZ������\�S��]{1,}�@[!�@]"
        .ParagraphFormat.LeftIndent = 110
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .Text = "���@��"
     End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.Font.Name = "�l�r �S�V�b�N"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

End Sub

Private Sub �w�b�_�[�t�b�^�[����_�Ж�()
'[6a]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�Жʈ���p�z�w�b�_�[�y�уt�b�^�[�ɍ쐬�����A�@�ߖ��y�уy�[�W�ԍ���t�L����B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '�w�b�_�[�ɍ쐬������������B�i���ꐫ�̒S�ہA�o�[�W�����Ǘ��̈Ӑ}�j
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.WholeStory
    Selection.TypeBackspace
        With Selection
            .Font.Name = "�l�r �S�V�b�N"
            .Font.Size = 6
            .ParagraphFormat.Alignment = wdAlignParagraphRight
            .TypeText Text:=" " '�����̑O�ɉ������Ȍオ����ꍇ�͍���""���ɕ�������L������B
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "DATE  \@ ""yyyy/MM/dd H:mm:ss"" ", PreserveFormatting:=False
            .InsertAfter " " '�����̌�ɉ������Ȍオ����ꍇ�͍���""���ɕ�������L������B
        End With
    
  '�t�b�^�[�i�y�[�W�����j�ɖ@�ߖ��ƃy�[�W�ԍ���������B

    '�t�b�^�[�̕\�̍폜�i���@�K�l���ւ̑Ή��B�P���Ƀt�b�^�[���폜����ƃG���[�ƂȂ�B�j
    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryFooter
    Selection.WholeStory
    Selection.TypeBackspace
    Selection.TypeText Text:="�@"
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    Dim myRange As Range
    Set myRange = ActiveDocument.Sentences(1)
    '���傤���� Super�@��Web�p�̃t�b�^�[�i���@�K �@��񑍍��f�[�^�x�[�X�p�ɂ͓��Y�����͕s�v�j
    If ActiveDocument.Range(0, 1) = "��" Then myRange.MoveStart wdCharacter, 1
    If ActiveDocument.Range(0, 1) = "��" Then myRange.MoveStart wdCharacter, 1
     
  '�ȉ��A���傤���� Super�@��Web�A���@�K �@��񑍍��f�[�^�x�[�X���ʂ̏���
    myRange.MoveEnd wdCharacter, -1
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryFooter
    Selection.WholeStory
    Selection.TypeBackspace
   
        With Selection
            .Font.Name = "�l�r ����"
            .Font.Size = 10
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:="�i" & myRange & "�j�@"
            .Fields.Add Range:=Selection.Range, Text:="PAGE  \* DBNUM1 "
            .InsertAfter " �^ "
            .Collapse Direction:=wdCollapseEnd
            .Fields.Add Range:=Selection.Range, Text:="NUMPAGES  \* DBNUM1 ", Type:=wdFieldEmpty
        End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

End Sub

Private Sub �w�b�_�[�t�b�^�[����_����()
'[6b]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y���ʈ���p�z�w�b�_�[�y�уt�b�^�[�ɍ쐬�����A�@�ߖ��y�уy�[�W�ԍ���t�L����B
'
' �y�����z
'�@�@ 2015/01/28  ver.1.01  �o�O�̏C��
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '��y�[�W�̃w�b�_�[�ɍ쐬������������B�i���ꐫ�̒S�ہA�o�[�W�����Ǘ��̈Ӑ}�j
    ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = True
  
    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryHeader
    Selection.WholeStory
    Selection.TypeBackspace
        With Selection
            .Font.Name = "�l�r �S�V�b�N"
            .Font.Size = 6
            .ParagraphFormat.Alignment = wdAlignParagraphRight
            .TypeText Text:=" " '�����̑O�ɉ������Ȍオ����ꍇ�͍���""���ɕ�������L������B
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "DATE  \@ ""yyyy/MM/dd H:mm:ss"" ", PreserveFormatting:=False
            .InsertAfter " " '�����̌�ɉ������Ȍオ����ꍇ�͍���""���ɕ�������L������B
        End With
        
  '�t�b�^�[�̕\�̍폜�i���@�K�l���ւ̑Ή��B�P���Ƀt�b�^�[���폜����ƃG���[�ƂȂ�B�j
    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryFooter
    Selection.WholeStory
    Selection.TypeBackspace
    Selection.TypeText Text:="�@"
   
  '�����y�[�W�̃t�b�^�[�ɍ쐬������������B�i���ꐫ�̒S�ہA�o�[�W�����Ǘ��̈Ӑ}�j
    If Selection.Information(wdNumberOfPagesInDocument) <> 1 Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekEvenPagesFooter
        Selection.WholeStory
        Selection.TypeBackspace
        With Selection
            .Font.Name = "�l�r �S�V�b�N"
            .Font.Size = 6
            .ParagraphFormat.Alignment = wdAlignParagraphRight
            .TypeText Text:=" " '�����̑O�ɉ������Ȍオ����ꍇ�͍���""���ɕ�������L������B
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "DATE  \@ ""yyyy/MM/dd H:mm:ss"" ", PreserveFormatting:=False
            .InsertAfter " " '�����̌�ɉ������Ȍオ����ꍇ�͍���""���ɕ�������L������B
        End With
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If

  '�t�b�^�[�i�y�[�W�����j�ɖ@�ߖ��ƃy�[�W�ԍ���������B
    Dim myRange As Range
    Set myRange = ActiveDocument.Sentences(1)
    '���傤���� Super�@��Web�p�̃t�b�^�[�i���@�K �@��񑍍��f�[�^�x�[�X�p�ɂ͓��Y�����͕s�v�j
    If ActiveDocument.Range(0, 1) = "��" Then myRange.MoveStart wdCharacter, 1
    If ActiveDocument.Range(0, 1) = "��" Then myRange.MoveStart wdCharacter, 1
     
  '�ȉ��A���傤���� Super�@��Web�A���@�K �@��񑍍��f�[�^�x�[�X���ʂ̏���
    myRange.MoveEnd wdCharacter, -1

    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryFooter
    Selection.WholeStory
    Selection.TypeBackspace
    
        With Selection
            .Font.Name = "�l�r ����"
            .Font.Size = 10
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:="�i" & myRange & "�j�@"
            .Fields.Add Range:=Selection.Range, Text:="PAGE  \* DBNUM1 "
            .InsertAfter " �^ "
            .Collapse Direction:=wdCollapseEnd
            .Fields.Add Range:=Selection.Range, Text:="NUMPAGES  \* DBNUM1 ", Type:=wdFieldEmpty
        End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

  '�����y�[�W
  
    If Selection.Information(wdNumberOfPagesInDocument) <> 1 Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekEvenPagesHeader
        Selection.WholeStory
        Selection.TypeBackspace
    
        With Selection
            .Font.Name = "�l�r ����"
            .Font.Size = 10
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:="�i" & myRange & "�j�@"
            .Fields.Add Range:=Selection.Range, Text:="PAGE  \* DBNUM1 "
            .InsertAfter " �^ "
            .Collapse Direction:=wdCollapseEnd
            .Fields.Add Range:=Selection.Range, Text:="NUMPAGES  \* DBNUM1 ", Type:=wdFieldEmpty
        End With
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If

End Sub

Sub ���ʃ��r�m�F()
'[7]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@���r���������ʏ����ɂ��Ă���ƍl������ӏ����m�F����B�Y���ӏ���Ԏ��ɂ��邱�Ƃ��ł���B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
    Dim fRng As Range
    Dim i As Integer
    Dim rc As Integer
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[��-�I]�i[��-��]{1,}�j"
    End With
    
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.Collapse Direction:=wdCollapseEnd
        i = i + 1
    Loop
    
    If i = 0 Then
        MsgBox "���r���������ʊO�����ɂ��Ă���ƍl������ӏ��͂���܂���ł����B"
        
    Else
        rc = MsgBox("���r���������ʏ����ɂ��Ă��邱�Ƃƍl������ӏ��� " & i & " �ӏ��ł��B" & vbCrLf & "�Y���ӏ��ɐԎ��������s���܂����H", vbYesNo + vbQuestion, "�m�F")

        If rc = vbYes Then
            i = 0
            Set fRng = ActiveDocument.Range(Start:=0, End:=0)
            With fRng.Find
                .MatchWildcards = True
                .Text = "[��-�I]�i[��-��]{1,}�j"
            End With
            Do
                If fRng.Find.Execute = False Then Exit Do
                fRng.MoveStart unit:=wdCharacter, Count:=1
                fRng.Font.ColorIndex = wdRed
                fRng.Font.Bold = wdToggle
                fRng.Collapse Direction:=wdCollapseEnd
                i = i + 1
            Loop
            MsgBox "�Ԏ��������������܂����B" & vbCrLf & "�������� �c " & i & " ��", vbOKOnly, "�Ԏ���������"
        End If
    End If
End Sub

Private Sub ������_�S��()
'[8a]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�S���z���̌`����K�p���ׂ��@�߂̎������𐄒肵�A�K�p����B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '�W���̃t�H���g�T�C�Y���擾
    fsnormal = ActiveDocument.Styles(wdStyleNormal).Font.Size
    
    Set RE = CreateObject("VBScript.RegExp")
        RE.Global = False
    
    For Each para In ActiveDocument.Paragraphs
    
      '���o���̎�����
        RE.Pattern = "^�i.+�j"
        Set Matches = RE.Execute(para.Range)
        
        If Matches.Count > 0 Then
            With para.Range.ParagraphFormat
                .LeftIndent = 0
                .FirstLineIndent = fsnormal
            End With
            
        Else
          '���̎�����
            RE.Pattern = "^��[���O�l�ܘZ������\�S��]+��(��[���O�l�ܘZ������\�S��]+)*[�@�E�`]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
          '���i�P���y�ъۈ͂ݐ����j�̎�����
            RE.Pattern = "^[�P-�X�@�A�B�C�D�E�F�G�H�I�J�K�L�M�N�O�P�Q�R�S]{1}[�@�E�`]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
          '���i�Q���j�̎�����
            RE.Pattern = "^[0-9�O-�X]{2}[�@�E�`]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
          '���̎�����
            RE.Pattern = "^[���O�l�ܘZ������\�S]{1,}[�@�E�`]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 2 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '���̍ו��̎������i�P�j�C���n�c
            RE.Pattern = "^[�A-��]{1}[�@�E�`]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 3 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '���̍ו��̎������i�Q�j(1)(2)(3)�c
            RE.Pattern = "^\([0-9]{1,}\)[�@�E�`]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 4 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '���̍ו��̎������i�R�j(i)(ii)(iii)�c
            RE.Pattern = "^\([ivx]{1,}\)[�@�E�`]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 5 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '����
            RE.Pattern = "^���@��"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 3 * fsnormal
                    .FirstLineIndent = 0
                End With
            End If
            
           '���R
            RE.Pattern = "^���@�R"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 5 * fsnormal
                    .FirstLineIndent = 0
                End With
            End If
            
        End If
    Next
    Set RE = Nothing
    Set Matches = Nothing

End Sub

Private Sub ������_�I��͈�()
'[8b]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�I��͈́z���̌`����K�p���ׂ��@�߂̎������𐄒肵�A�K�p����B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '�W���̃t�H���g�T�C�Y���擾
    fsnormal = ActiveDocument.Styles(wdStyleNormal).Font.Size
    
    Set RE = CreateObject("VBScript.RegExp")
        RE.Global = False
    
    For Each para In Selection.Paragraphs
    
      '���o���̎�����
        RE.Pattern = "^�i.+�j"
        Set Matches = RE.Execute(para.Range)
        
        If Matches.Count > 0 Then
            With para.Range.ParagraphFormat
                .LeftIndent = 0
                .FirstLineIndent = fsnormal
            End With
            
        Else
          '���̎�����
            RE.Pattern = "^��[���O�l�ܘZ������\�S��]+��(��[���O�l�ܘZ������\�S��])*"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
          '���i�P���y�ъۈ͂ݐ����j�̎�����
            RE.Pattern = "^[�P-�X�@�A�B�C�D�E�F�G�H�I�J�K�L�M�N�O�P�Q�R�S]{1}�@"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
          '���̎�����
            RE.Pattern = "^[���O�l�ܘZ������\�S]{1,}�@"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 2 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '���̍ו��̎������i�P�j�C���n�c
            RE.Pattern = "^[�A-��]{1}�@"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 3 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '���̍ו��̎������i�Q�j(1)(2)(3)�c
            RE.Pattern = "^([0-9]{1,})�@"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 4 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '���̍ו��̎������i�R�j(i)(ii)(iii)�c
            RE.Pattern = "^([ivx]{1,})�@"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 5 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '����
            RE.Pattern = "^���@��"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 3 * fsnormal
                    .FirstLineIndent = 0
                End With
            End If
            
           '���R
            RE.Pattern = "^���@�R"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 5 * fsnormal
                    .FirstLineIndent = 0
                End With
            End If
            
        End If
    Next
    Set RE = Nothing
    Set Matches = Nothing

End Sub

Private Sub e_Gov�t�H�[�}�b�g������_�S��()
'[9a]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�S���ze-Gov�̃t�H�[�}�b�g������������B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '�n�C�p�[�����N�̉���
    For Each aFld In ActiveDocument.Fields
        aFld.Unlink
    Next
    
  '�t�H���g�̃t�H�[�}�b�g���N���A
    Selection.WholeStory
    Selection.ClearFormatting
    
  '�I��͈͂̕\�̈ʒu�̒���
    Dim tbl As Table
    For Each tbl In Selection.Tables
        With tbl
            .Style = "�\ (�i�q)"
            .Rows.Alignment = wdAlignRowCenter
            .AutoFitBehavior (wdAutoFitContent)
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 92
        End With
    Next

End Sub

Private Sub e_Gov�t�H�[�}�b�g������_�I��͈�()
'[9b]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�y�I��͈́ze-Gov�̃t�H�[�}�b�g������������B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-

  '�n�C�p�[�����N�̉���
    For Each aFld In Selection.Fields
        aFld.Unlink
    Next
    
  '�t�H���g�̃t�H�[�}�b�g���N���A
    
    Selection.ClearFormatting
    
  '�I��͈͂̕\�̈ʒu�̒���
    Dim tbl As Table
    For Each tbl In Selection.Tables
        With tbl
            .Style = "�\ (�i�q)"
            .Rows.Alignment = wdAlignRowCenter
            .AutoFitBehavior (wdAutoFitContent)
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 92
        End With
    Next

End Sub

Private Sub �󔒍s���폜_�I��͈�()
'[10b]
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
' �y�T�v�z
' �@�@�󔒍s�����폜����B
'
' �y�����z
' �@�@2014/10/29�@ver.1.0 �@�V�K�쐬
'
' �y�ۑ�z�\���̃C���f���g������B�\����I���̑ΏۊO�ɂł��Ȃ����H
'
'-�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�- -�-
    Dim para As Paragraph
    
  '�����^�u�����s�ɒu������B
    With Selection.Find
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Text = vbVerticalTab
        .Replacement.Text = vbCr
        .Execute Replace:=wdReplaceAll
    End With
  '�P�Ƃ̉��s�̍폜
    For Each para In Selection.Paragraphs
        With para.Range
            If .Characters.Count = 1 Then .Delete
        End With
    Next
    
  '�����̉��s�̒u��
    With Selection.Find
        .MatchWildcards = True
        .Text = "^13{2,}"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub


