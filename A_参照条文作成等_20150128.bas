Attribute VB_Name = "A_参照条文作成等_20150128"
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
'法令事務支援マクロ（Word版）(1/3)　A_参照条文作成等_20150128
'Copyright(C)2014 Blue Panda
'This software is released under the MIT License.
'http://opensource.org/licenses/mit-license.php
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

Sub 参照条文作成()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     ぎょうせいの「Super法令Web」又は第一法規「D1-Law.com」（第一法規 法情報総合データベース）
' 　　から出力された条文データを現行日本法規風の様式に調整する。
'
' 【履歴】
'     2014/10/29　ver.1.01 不要なコードの削除、可読性向上のためのコードの整理、
'　　　　　　　　　　　　　フッターのバグを解消
' 　　2014/10/01　ver.1.0
' 　　2014/09/12　暫定版作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【備考】
' 　　インデントの情報はダウンロードしたファイルの情報を利用しています。
'
' 【未対応の内容等】
' 　　「D1-Law.com」のルビ文字は、元データの仕様により、括弧書きで表示されます。
' 　　一部改正法令令の条文での使用は想定していません。
'
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  '変数の宣言
    Dim FirstFontSize As Single, ChangedFontSize As Single
    Dim Indenting As Single, Indented As Single
    Dim fRng As Range
    Dim i As Integer
    Dim sWrds() As Variant, rWrds() As Variant
    Dim Num As Integer
    
  'ちらつき防止（マクロの高速化。末尾で解除。）※マクロの動作を可視化するためあえて作動させていない。
        'Application.ScreenUpdating = False
    
  'ヘッダー又はフッターの編集画面になっている場合への対応
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
 
  '[1a] インデントの調整（11ポイントのフォントのインデントに調整）
        Call インデント調整110
    
  '[2a] 半角スペースの処理（２つ連続する半角スペースは全角に置換し、１つ単独のものは抹消）
  
        Call 半角スペース処理_全文
    
  '[3a] 項番号等の縦中横の処理を行う。
  
        Call 項番号縦中横等処理_全文

  '[4a] 体裁の調整（現行日本法規風の体裁）
  
        Call 体裁調整_全文

  '[5a] 条名の条番号及び章名等をゴシックにする。
  
        Call 条名等ゴシック_全文

  '[6a] ヘッダー及びフッターに作成日時、法令名及びページ番号を加える。（ページ番号は左のみ）
  
        Call ヘッダーフッター処理_片面
      
  '初字の位置に戻る。
        ActiveDocument.Range(0, 0).Select

  'ちらつき防止解除
        'Application.ScreenUpdating = True
    
End Sub
Sub 参照条文作成_両面印刷用()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     マクロ「参照条文作成」のヘッダー及びフッターを両面印刷に適した形にしたもの。
'
' 【履歴】
'　　 2015/01/28　ver.1.02 フッターのバグを解消
'     2014/10/29　ver.1.01 不要なコードの削除、可読性向上のためのコードの整理、
'　　　　　　　　　　　　　フッターのバグを解消
' 　　2014/10/01　ver.1.0　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

    Dim FirstFontSize As Single, ChangedFontSizet As Single
    Dim Indenting As Single, Indented As Single
    Dim fRng As Range
    Dim i As Integer
    Dim sWrds() As Variant, rWrds() As Variant
    Dim Num As Integer
    
  'ヘッダー又はフッターの編集画面になっている場合への対応
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
  
  '[1a] インデントの調整（11ポイントのフォントのインデントに調整）
        Call インデント調整110
    
  '[2a] 半角スペースの処理（２つ連続する半角スペースは全角に置換し、１つ単独のものは抹消）
  
        Call 半角スペース処理_全文
    
  '[3a] 項番号等の縦中横の処理を行う。
  
        Call 項番号縦中横等処理_全文
 
  '[4a] 体裁の調整（現行日本法規風の体裁）
  
        Call 体裁調整_全文

  '[5a] 条名の条番号及び章名等をゴシックにする。
  
        Call 条名等ゴシック_全文

  '[6b] ヘッダー及びフッターに作成日時、法令名及びページ番号を加える。（ページ番号は左右左…の順）
  
        Call ヘッダーフッター処理_両面
    
  '初字の位置に戻る。
        ActiveDocument.Range(0, 0).Select

  'ちらつき防止解除
        'Application.ScreenUpdating = True
        
End Sub

Sub 新旧等加工用処理()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     ぎょうせいの「Super法令Web」又は第一法規「D1-Law.com」（第一法規 法情報総合データベース）
'     から出力された条文データを、新旧対照表等の縦書きの様式の書類の作成に適した形に加工する。
'
' 【履歴】
'     2014/10/29　ver.1.0　 新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

    Dim FirstFontSize As Single, ChangedFontSizet As Single
    Dim Indenting As Single, Indented As Single
    Dim fRng As Range
    Dim i As Integer
    Dim sWrds() As Variant, rWrds() As Variant
    Dim Num As Integer
    
  'ヘッダー又はフッターの編集画面になっている場合への対応
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
  
  '[1b] インデントの調整（11ポイントのフォントのインデントに調整）
        Call インデント調整105
    
  '[2a] 半角スペースの処理（２つ連続する半角スペースは全角に置換し、１つ単独のものは抹消）
  
        Call 半角スペース処理_全文
    
  '[3a] 項番号等の縦中横の処理を行う。
  
        Call 項番号縦中横等処理_全文
 
  '[6a] ヘッダー及びフッターに作成日時、法令名及びページ番号を加える。（ページ番号は左のみ）
  
        Call ヘッダーフッター処理_片面
    
  '本文のフォントの調整
        Selection.WholeStory
        With Selection
            .Font.Name = "ＭＳ 明朝"
            .Font.Size = 10.5
            .Orientation = wdTextOrientationHorizontal
        End With
    
  '段組（１段）
        ActiveDocument.PageSetup.TextColumns.SetCount NumColumns:=1

  '用紙設定
    With ActiveDocument.PageSetup
        .TopMargin = MillimetersToPoints(10)
        .BottomMargin = MillimetersToPoints(10)
        .LeftMargin = MillimetersToPoints(20)
        .RightMargin = MillimetersToPoints(20)
        .HeaderDistance = MillimetersToPoints(5)
        .FooterDistance = MillimetersToPoints(5)
    End With

  '表の位置
        Dim tbl As Table
        For Each tbl In ActiveDocument.Tables
            With tbl
                .Style = "表 (格子)"
                .Rows.Alignment = wdAlignRowCenter
                .AutoFitBehavior (wdAutoFitContent)
                .PreferredWidthType = wdPreferredWidthPercent
                .PreferredWidth = 95
            End With
        Next
    
  '[7] ルビ文字の確認
  
        Call 括弧ルビ確認
    
  '初字の位置に戻る。
        ActiveDocument.Range(0, 0).Select

        
End Sub

Sub e_Gov条文処理_全文_タテ()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　 e-Govからコピーしてきた条文の体裁を、縦書きの体裁に整える。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
    
    Dim aFld As Field
    Dim para As Paragraph
    Dim fRng As Range
    Dim fsnormal As Single
    Dim RE As Object, Matches As Object

  '[9a] e-Govのフォーマットを初期化する。

        Call e_Govフォーマット初期化_全文
    
  '[10b] 選択範囲の空白行及び垂直タブの削除
    
        Call 空白行等削除_選択範囲

  '[2a] 半角スペースの処理（２つ連続する半角スペースは全角に置換し、１つ単独のものは抹消）
  
        Call 半角スペース処理_全文

  '[3a] 項番号等の縦中横の処理を行う。
  
        Call 項番号縦中横等処理_全文

  '[8a] 文の形から適用すべき法令の字下げを推定し、適用する。

        Call 字下げ_全文
    
  '文字の向きを「横書き（日本語文字を左に90度回転）」にする。
        Selection.WholeStory
        Selection.Orientation = wdTextOrientationHorizontalRotatedFarEast

End Sub

Sub e_Gov条文処理_全文_ヨコ()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　e-Govからコピーしてきた条文の体裁を、横書きの体裁に整える。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
  
    Dim aFld As Field
    Dim para As Paragraph
    Dim fRng As Range
    Dim fsnormal As Single
    Dim RE As Object, Matches As Object
    
  '[9a] e-Govのフォーマットを初期化する。

        Call e_Govフォーマット初期化_全文
    
  '[10b] 選択範囲の空白行及び垂直タブの削除
    
        Call 空白行等削除_選択範囲

  '[2a] 半角スペースの処理（２つ連続する半角スペースは全角に置換し、１つ単独のものは抹消）
  
        Call 半角スペース処理_全文

  '[3b] 項番号等の処理
  
        Call 項番号等処理_全文_ヨコ

  '[8a] 文の形から適用すべき法令の字下げを推定し、適用する。

        Call 字下げ_全文

End Sub

Private Sub インデント調整110()
'[1a]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【全文】インデントの調整（11ポイントのフォントのインデントに調整）
'　　　※ぎょうせいの「Super法令Web」又は第一法規「D1-Law.com」（第一法規 法情報総合データベース）専用
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
    
    If ActiveDocument.Range(0, 0).Font.Size = 11 Then  '→11ポイントの場合はインデント処理不要
    Else
        FirstFontSize = ActiveDocument.Range(0, 0).Font.Size  '初字のフォントサイズを取得
        ChangedFontSize = 11  '参照条文本文のフォントサイズを設定
    '2行目以降のインデント（LeftIndent）を調整
       
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
    
    '1行目のインデント（FirstLineIndent）を調整
       
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

Private Sub インデント調整105()
'[1b]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【全文】インデントの調整（10.5ポイントのフォントのインデントに調整）
'　　　※ぎょうせいの「Super法令Web」又は第一法規「D1-Law.com」（第一法規 法情報総合データベース）専用
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

    If ActiveDocument.Range(0, 0).Font.Size = 10.5 Then  '→10.5ポイントの場合はインデント処理不要
    Else
        FirstFontSize = ActiveDocument.Range(0, 0).Font.Size  '初字のフォントサイズを取得
        ChangedFontSize = 10.5  '参照条文本文のフォントサイズを設定
    '2行目以降のインデント（LeftIndent）を調整
       
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
    
    '1行目のインデント（FirstLineIndent）を調整
       
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

Private Sub 半角スペース処理_全文()
'[2a]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【全文】半角スペースの処理（２つ連続する半角スペースは全角に置換し、１つ単独のものは抹消）
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

    With ActiveDocument.Content.Find
        .Text = "  "
        .Replacement.Text = "　"
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

Private Sub 半角スペース処理_選択範囲()
'[2b]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【選択範囲】半角スペースの処理（２つ連続する半角スペースは全角に置換し、１つ単独のものは抹消）
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

    Set fRng = Selection.Range
    With fRng.Find
        .Text = "  "
        .Replacement.Text = "　"
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

Sub 項番号縦中横等処理_全文()
'[3a]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　 全文に項番号等に縦中横等の処理を行う。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　可読性向上のためのコードの整理
'
' 【留意事項】
' 　　使用される元号が加わった場合、コードの修正が必要です。
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  '半角数字と半角括弧等『 ()[]｢｣｡､ 』を全角にする。
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchFuzzy = False
        .MatchWildcards = True
        .Text = "[0-9)([\]｢｣｡､]{1,}"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.CharacterWidth = wdWidthFullWidth
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

  '括弧付き数字を半角にし、縦中横の処理を行う。
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "（[０-９]{1,}）"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine '幅を広くする場合 → wdHorizontalInVerticalResizeLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
  '項番号のない項の処理　「○１　」→「①　」　※「○２０　」まで
    sWrds = Array("○１　", "○２　", "○３　", "○４　", "○５　", "○６　", "○７　", "○８　", "○９　", "○１０　" _
                , "○１１　", "○１２　", "○１３　", "○１４　", "○１５　", "○１６　", "○１７　", "○１８　", "○１９　", "○２０　")
    rWrds = Array("①　", "②　", "③　", "④　", "⑤　", "⑥　", "⑦　", "⑧　", "⑨　", "⑩　", _
                "⑪　", "⑫　", "⑬　", "⑭　", "⑮　", "⑯　", "⑰　", "⑱　", "⑲　", "⑳　")
    With ActiveDocument.Content.Find
        For Num = LBound(sWrds) To UBound(sWrds)
            .Text = sWrds(Num)
            .Replacement.Text = rWrds(Num)
            .Execute Replace:=wdReplaceAll
        Next Num
    End With
    
  '項番号の二桁の数字を半角にし、縦中横の処理を行う。
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[０-９]{2}　[!　]"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
  '算用数字で標記されている法令番号及び年月日の処理（２～３桁の数字を半角にし、縦中横の処理を行う。）
  
   '年（※法情報総合データベースの表記に対応する部分）
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[治正和成][０-９]{2,3}年"  '　←　使用される元号が増えた際、この部分に加筆の必要性あり
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveStart wdCharacter, 1
        fRng.MoveEnd wdCharacter, -1
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
   '月
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "年[０-９]{2}月"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveStart wdCharacter, 1
        fRng.MoveEnd wdCharacter, -1
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
   '日
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "月[０-９]{2}日"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveStart wdCharacter, 1
        fRng.MoveEnd wdCharacter, -1
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
   '法令番号
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[０-９]{2,3}号"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveStart wdCharacter, 1
        fRng.MoveEnd wdCharacter, -1
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

  '（i）、（ii）、（iii）の括弧を半角にし、縦中横の処理を行う。
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "（[ivx]{1,}）"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.CharacterWidth = wdWidthHalfWidth
        fRng.HorizontalInVertical = wdHorizontalInVerticalFitInLine
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

End Sub

Private Sub 項番号等処理_全文_ヨコ()
'[3b]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'　　 全文に項番号等の半角処理を行う。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
  '半角の数字、括弧等を全角にする。

    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[0-9)([\]｢｣｡､]{1,}"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.CharacterWidth = wdWidthFullWidth
        fRng.Collapse Direction:=wdCollapseEnd
        fRng.MoveEnd unit:=wdCharacter, Count:=-1
    Loop


  '【号の細分(1)(2)(3)】全ての括弧付き数字を半角にする。
    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "（[０-９]{1,}）[!　]"
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
        .Text = "（[０-９]{1,}）　"
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
        
   '【号の細分(i)(ii)(iii)】括弧を半角にする。
    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "（[ivx]{1,}）"
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
        .Text = "（[ivx]{1,}）　"
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
        
  '項番号のない項の処理　「○１　」→「①　」　※「○２０　」まで
    sWrds = Array("○１　", "○２　", "○３　", "○４　", "○５　", "○６　", "○７　", "○８　", "○９　", "○１０　" _
                , "○１１　", "○１２　", "○１３　", "○１４　", "○１５　", "○１６　", "○１７　", "○１８　", "○１９　", "○２０　")
    rWrds = Array("①　", "②　", "③　", "④　", "⑤　", "⑥　", "⑦　", "⑧　", "⑨　", "⑩　", _
                "⑪　", "⑫　", "⑬　", "⑭　", "⑮　", "⑯　", "⑰　", "⑱　", "⑲　", "⑳　")
    With ActiveDocument.Content.Find
        For Num = LBound(sWrds) To UBound(sWrds)
            .Text = sWrds(Num)
            .Replacement.Text = rWrds(Num)
            .Execute Replace:=wdReplaceAll
        Next Num
    End With
    
  '【項番号（10～）】全ての項番号の二桁の数字を半角にする。
  '項（２桁）の字下げ
    Set fRng = ActiveDocument.Range(0, 0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[０-９]{2}　[!　]"
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
Private Sub 体裁調整_全文()
'[4a]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【全文】用紙サイズ及びフォントの体裁を調整する。
'　　　※「ぎょうせい」の現行日本法規（紙媒体）をＡ４サイズに印刷したものに近い体裁としている。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
  
  '本文のフォントの調整
    Selection.WholeStory
    With Selection
        .Font.Name = "ＭＳ 明朝"
        .Font.Size = 11  'ChangedFontSizeと同一の値とする。
        .Orientation = wdTextOrientationHorizontalRotatedFarEast '横書き（日本語文字を左に90度回転）
    End With
    
  '用紙サイズをA4に変更
    ActiveDocument.PageSetup.PaperSize = wdPaperA4
        
  '行間
    Selection.WholeStory
    With Selection.ParagraphFormat
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 20.2
    End With

  '用紙設定
    With ActiveDocument.PageSetup
        .Orientation = wdOrientLandscape '用紙を横向きに
        .TopMargin = MillimetersToPoints(23)
        .BottomMargin = MillimetersToPoints(23)
        .LeftMargin = MillimetersToPoints(25)
        .RightMargin = MillimetersToPoints(25)
        .HeaderDistance = MillimetersToPoints(12.7)
        .FooterDistance = MillimetersToPoints(12.7)
        .LinesPage = 23 '23行（結果的に１行30文字）
        .LayoutMode = wdLayoutModeLineGrid
    End With
    
  '段組
    With ActiveDocument.PageSetup.TextColumns
        .SetCount NumColumns:=2
        .EvenlySpaced = True
        .LineBetween = True
        .Width = MillimetersToPoints(117.1)
        .Spacing = MillimetersToPoints(12.7)
    End With
    
  '表の位置
    Dim tbl As Table
    For Each tbl In ActiveDocument.Tables
        With tbl
            .Style = "表 (格子)"
            .Rows.Alignment = wdAlignRowCenter
            .AutoFitBehavior (wdAutoFitContent)
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 91
        End With
    Next
    
  '法令名（１行目）のフォントサイズの調整
    If ActiveDocument.Range(0, 1) = "○" Or ActiveDocument.Range(0, 1) = "●" Then
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

Private Sub 条名等ゴシック_全文()
'[5a]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【全文】条名の条番号をゴシックにする。
'　　　※ぎょうせいの「Super法令Web」又は第一法規「D1-Law.com」（第一法規 法情報総合データベース）専用
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}条　[!　]"
        .ParagraphFormat.LeftIndent = 11
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}条の[一二三四五六七八九十百千]{1,}　[!　]"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}条の[一二三四五六七八九十百千]{1,}の[一二三四五六七八九十百千]{1,}　[!　]"
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
  '章名等をゴシックにする。（ただし、目次は除く。）
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}編　[!　]"
        .ParagraphFormat.LeftIndent = 66
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}章　[!　]"
        .ParagraphFormat.LeftIndent = 77
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    

    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}章の[一二三四五六七八九十百千]{1,}　[!　]"
        .ParagraphFormat.LeftIndent = 77
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}節　[!　]"
        .ParagraphFormat.LeftIndent = 88
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    

    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}節の[一二三四五六七八九十百千]{1,}　[!　]"
        .ParagraphFormat.LeftIndent = 88
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}款　[!　]"
        .ParagraphFormat.LeftIndent = 99
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}款の[一二三四五六七八九十百千]{1,}　[!　]"
        .ParagraphFormat.LeftIndent = 99
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}目　[!　]"
        .ParagraphFormat.LeftIndent = 110
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "第[一二三四五六七八九十百千]{1,}目の[一二三四五六七八九十百千]{1,}　[!　]"
        .ParagraphFormat.LeftIndent = 110
    End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.MoveEnd wdCharacter, -2
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .Text = "附　則"
     End With
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.Font.Name = "ＭＳ ゴシック"
        fRng.Collapse Direction:=wdCollapseEnd
    Loop

End Sub

Private Sub ヘッダーフッター処理_片面()
'[6a]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【片面印刷用】ヘッダー及びフッターに作成日時、法令名及びページ番号を付記する。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  'ヘッダーに作成日時を加える。（同一性の担保、バージョン管理の意図）
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    Selection.WholeStory
    Selection.TypeBackspace
        With Selection
            .Font.Name = "ＭＳ ゴシック"
            .Font.Size = 6
            .ParagraphFormat.Alignment = wdAlignParagraphRight
            .TypeText Text:=" " '日時の前に加えた以後がある場合は左の""内に文字列を記入する。
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "DATE  \@ ""yyyy/MM/dd H:mm:ss"" ", PreserveFormatting:=False
            .InsertAfter " " '日時の後に加えた以後がある場合は左の""内に文字列を記入する。
        End With
    
  'フッター（ページ左側）に法令名とページ番号を加える。

    'フッターの表の削除（第一法規様式への対応。単純にフッターを削除するとエラーとなる。）
    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryFooter
    Selection.WholeStory
    Selection.TypeBackspace
    Selection.TypeText Text:="　"
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    
    Dim myRange As Range
    Set myRange = ActiveDocument.Sentences(1)
    'ぎょうせい Super法令Web用のフッター（第一法規 法情報総合データベース用には当該処理は不要）
    If ActiveDocument.Range(0, 1) = "○" Then myRange.MoveStart wdCharacter, 1
    If ActiveDocument.Range(0, 1) = "●" Then myRange.MoveStart wdCharacter, 1
     
  '以下、ぎょうせい Super法令Web、第一法規 法情報総合データベース共通の処理
    myRange.MoveEnd wdCharacter, -1
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryFooter
    Selection.WholeStory
    Selection.TypeBackspace
   
        With Selection
            .Font.Name = "ＭＳ 明朝"
            .Font.Size = 10
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:="（" & myRange & "）　"
            .Fields.Add Range:=Selection.Range, Text:="PAGE  \* DBNUM1 "
            .InsertAfter " ／ "
            .Collapse Direction:=wdCollapseEnd
            .Fields.Add Range:=Selection.Range, Text:="NUMPAGES  \* DBNUM1 ", Type:=wdFieldEmpty
        End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

End Sub

Private Sub ヘッダーフッター処理_両面()
'[6b]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【両面印刷用】ヘッダー及びフッターに作成日時、法令名及びページ番号を付記する。
'
' 【履歴】
'　　 2015/01/28  ver.1.01  バグの修正
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  '奇数ページのヘッダーに作成日時を加える。（同一性の担保、バージョン管理の意図）
    ActiveDocument.PageSetup.OddAndEvenPagesHeaderFooter = True
  
    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryHeader
    Selection.WholeStory
    Selection.TypeBackspace
        With Selection
            .Font.Name = "ＭＳ ゴシック"
            .Font.Size = 6
            .ParagraphFormat.Alignment = wdAlignParagraphRight
            .TypeText Text:=" " '日時の前に加えた以後がある場合は左の""内に文字列を記入する。
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "DATE  \@ ""yyyy/MM/dd H:mm:ss"" ", PreserveFormatting:=False
            .InsertAfter " " '日時の後に加えた以後がある場合は左の""内に文字列を記入する。
        End With
        
  'フッターの表の削除（第一法規様式への対応。単純にフッターを削除するとエラーとなる。）
    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryFooter
    Selection.WholeStory
    Selection.TypeBackspace
    Selection.TypeText Text:="　"
   
  '偶数ページのフッターに作成日時を加える。（同一性の担保、バージョン管理の意図）
    If Selection.Information(wdNumberOfPagesInDocument) <> 1 Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekEvenPagesFooter
        Selection.WholeStory
        Selection.TypeBackspace
        With Selection
            .Font.Name = "ＭＳ ゴシック"
            .Font.Size = 6
            .ParagraphFormat.Alignment = wdAlignParagraphRight
            .TypeText Text:=" " '日時の前に加えた以後がある場合は左の""内に文字列を記入する。
            .Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, Text:= _
                "DATE  \@ ""yyyy/MM/dd H:mm:ss"" ", PreserveFormatting:=False
            .InsertAfter " " '日時の後に加えた以後がある場合は左の""内に文字列を記入する。
        End With
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If

  'フッター（ページ左側）に法令名とページ番号を加える。
    Dim myRange As Range
    Set myRange = ActiveDocument.Sentences(1)
    'ぎょうせい Super法令Web用のフッター（第一法規 法情報総合データベース用には当該処理は不要）
    If ActiveDocument.Range(0, 1) = "○" Then myRange.MoveStart wdCharacter, 1
    If ActiveDocument.Range(0, 1) = "●" Then myRange.MoveStart wdCharacter, 1
     
  '以下、ぎょうせい Super法令Web、第一法規 法情報総合データベース共通の処理
    myRange.MoveEnd wdCharacter, -1

    ActiveWindow.ActivePane.View.SeekView = wdSeekPrimaryFooter
    Selection.WholeStory
    Selection.TypeBackspace
    
        With Selection
            .Font.Name = "ＭＳ 明朝"
            .Font.Size = 10
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:="（" & myRange & "）　"
            .Fields.Add Range:=Selection.Range, Text:="PAGE  \* DBNUM1 "
            .InsertAfter " ／ "
            .Collapse Direction:=wdCollapseEnd
            .Fields.Add Range:=Selection.Range, Text:="NUMPAGES  \* DBNUM1 ", Type:=wdFieldEmpty
        End With
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument

  '偶数ページ
  
    If Selection.Information(wdNumberOfPagesInDocument) <> 1 Then
        ActiveWindow.ActivePane.View.SeekView = wdSeekEvenPagesHeader
        Selection.WholeStory
        Selection.TypeBackspace
    
        With Selection
            .Font.Name = "ＭＳ 明朝"
            .Font.Size = 10
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .TypeText Text:="（" & myRange & "）　"
            .Fields.Add Range:=Selection.Range, Text:="PAGE  \* DBNUM1 "
            .InsertAfter " ／ "
            .Collapse Direction:=wdCollapseEnd
            .Fields.Add Range:=Selection.Range, Text:="NUMPAGES  \* DBNUM1 ", Type:=wdFieldEmpty
        End With
        ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
    End If

End Sub

Sub 括弧ルビ確認()
'[7]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　ルビ文字を括弧書きにしていると考えられる箇所を確認する。該当箇所を赤字にすることもできる。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
    Dim fRng As Range
    Dim i As Integer
    Dim rc As Integer
    
    Set fRng = ActiveDocument.Range(Start:=0, End:=0)
    With fRng.Find
        .MatchWildcards = True
        .Text = "[一-鶴]（[ぁ-ん]{1,}）"
    End With
    
    Do
        If fRng.Find.Execute = False Then Exit Do
        fRng.Collapse Direction:=wdCollapseEnd
        i = i + 1
    Loop
    
    If i = 0 Then
        MsgBox "ルビ文字を括弧外書きにしていると考えられる箇所はありませんでした。"
        
    Else
        rc = MsgBox("ルビ文字を括弧書きにしていることと考えられる箇所は " & i & " 箇所です。" & vbCrLf & "該当箇所に赤字処理を行いますか？", vbYesNo + vbQuestion, "確認")

        If rc = vbYes Then
            i = 0
            Set fRng = ActiveDocument.Range(Start:=0, End:=0)
            With fRng.Find
                .MatchWildcards = True
                .Text = "[一-鶴]（[ぁ-ん]{1,}）"
            End With
            Do
                If fRng.Find.Execute = False Then Exit Do
                fRng.MoveStart unit:=wdCharacter, Count:=1
                fRng.Font.ColorIndex = wdRed
                fRng.Font.Bold = wdToggle
                fRng.Collapse Direction:=wdCollapseEnd
                i = i + 1
            Loop
            MsgBox "赤字処理が完了しました。" & vbCrLf & "処理件数 … " & i & " 件", vbOKOnly, "赤字処理完了"
        End If
    End If
End Sub

Private Sub 字下げ_全文()
'[8a]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【全文】文の形から適用すべき法令の字下げを推定し、適用する。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  '標準のフォントサイズを取得
    fsnormal = ActiveDocument.Styles(wdStyleNormal).Font.Size
    
    Set RE = CreateObject("VBScript.RegExp")
        RE.Global = False
    
    For Each para In ActiveDocument.Paragraphs
    
      '見出しの字下げ
        RE.Pattern = "^（.+）"
        Set Matches = RE.Execute(para.Range)
        
        If Matches.Count > 0 Then
            With para.Range.ParagraphFormat
                .LeftIndent = 0
                .FirstLineIndent = fsnormal
            End With
            
        Else
          '条の字下げ
            RE.Pattern = "^第[一二三四五六七八九十百千]+条(の[一二三四五六七八九十百千]+)*[　・～]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
          '項（１桁及び丸囲み数字）の字下げ
            RE.Pattern = "^[１-９①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]{1}[　・～]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
          '項（２桁）の字下げ
            RE.Pattern = "^[0-9０-９]{2}[　・～]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
          '号の字下げ
            RE.Pattern = "^[一二三四五六七八九十百]{1,}[　・～]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 2 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '号の細分の字下げ（１）イロハ…
            RE.Pattern = "^[ア-ン]{1}[　・～]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 3 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '号の細分の字下げ（２）(1)(2)(3)…
            RE.Pattern = "^\([0-9]{1,}\)[　・～]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 4 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '号の細分の字下げ（３）(i)(ii)(iii)…
            RE.Pattern = "^\([ivx]{1,}\)[　・～]"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 5 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '附則
            RE.Pattern = "^附　則"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 3 * fsnormal
                    .FirstLineIndent = 0
                End With
            End If
            
           '理由
            RE.Pattern = "^理　由"
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

Private Sub 字下げ_選択範囲()
'[8b]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【選択範囲】文の形から適用すべき法令の字下げを推定し、適用する。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  '標準のフォントサイズを取得
    fsnormal = ActiveDocument.Styles(wdStyleNormal).Font.Size
    
    Set RE = CreateObject("VBScript.RegExp")
        RE.Global = False
    
    For Each para In Selection.Paragraphs
    
      '見出しの字下げ
        RE.Pattern = "^（.+）"
        Set Matches = RE.Execute(para.Range)
        
        If Matches.Count > 0 Then
            With para.Range.ParagraphFormat
                .LeftIndent = 0
                .FirstLineIndent = fsnormal
            End With
            
        Else
          '条の字下げ
            RE.Pattern = "^第[一二三四五六七八九十百千]+条(の[一二三四五六七八九十百千])*"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
          '項（１桁及び丸囲み数字）の字下げ
            RE.Pattern = "^[１-９①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳]{1}　"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
          '号の字下げ
            RE.Pattern = "^[一二三四五六七八九十百]{1,}　"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 2 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '号の細分の字下げ（１）イロハ…
            RE.Pattern = "^[ア-ン]{1}　"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 3 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '号の細分の字下げ（２）(1)(2)(3)…
            RE.Pattern = "^([0-9]{1,})　"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 4 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '号の細分の字下げ（３）(i)(ii)(iii)…
            RE.Pattern = "^([ivx]{1,})　"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 5 * fsnormal
                    .FirstLineIndent = -1 * fsnormal
                End With
            End If
            
           '附則
            RE.Pattern = "^附　則"
            Set Matches = RE.Execute(para.Range)
            If Matches.Count > 0 Then
                With para.Range.ParagraphFormat
                    .LeftIndent = 3 * fsnormal
                    .FirstLineIndent = 0
                End With
            End If
            
           '理由
            RE.Pattern = "^理　由"
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

Private Sub e_Govフォーマット初期化_全文()
'[9a]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【全文】e-Govのフォーマットを初期化する。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  'ハイパーリンクの解除
    For Each aFld In ActiveDocument.Fields
        aFld.Unlink
    Next
    
  'フォントのフォーマットをクリア
    Selection.WholeStory
    Selection.ClearFormatting
    
  '選択範囲の表の位置の調整
    Dim tbl As Table
    For Each tbl In Selection.Tables
        With tbl
            .Style = "表 (格子)"
            .Rows.Alignment = wdAlignRowCenter
            .AutoFitBehavior (wdAutoFitContent)
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 92
        End With
    Next

End Sub

Private Sub e_Govフォーマット初期化_選択範囲()
'[9b]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　【選択範囲】e-Govのフォーマットを初期化する。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  'ハイパーリンクの解除
    For Each aFld In Selection.Fields
        aFld.Unlink
    Next
    
  'フォントのフォーマットをクリア
    
    Selection.ClearFormatting
    
  '選択範囲の表の位置の調整
    Dim tbl As Table
    For Each tbl In Selection.Tables
        With tbl
            .Style = "表 (格子)"
            .Rows.Alignment = wdAlignRowCenter
            .AutoFitBehavior (wdAutoFitContent)
            .PreferredWidthType = wdPreferredWidthPercent
            .PreferredWidth = 92
        End With
    Next

End Sub

Private Sub 空白行等削除_選択範囲()
'[10b]
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
' 　　空白行等を削除する。
'
' 【履歴】
' 　　2014/10/29　ver.1.0 　新規作成
'
' 【課題】表中のインデントが壊れる。表中を選択の対象外にできないか？
'
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
    Dim para As Paragraph
    
  '垂直タブを改行に置換する。
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
  '単独の改行の削除
    For Each para In Selection.Paragraphs
        With para.Range
            If .Characters.Count = 1 Then .Delete
        End With
    Next
    
  '複数の改行の置換
    With Selection.Find
        .MatchWildcards = True
        .Text = "^13{2,}"
        .Replacement.Text = "^p"
        .Execute Replace:=wdReplaceAll
    End With
    
End Sub


