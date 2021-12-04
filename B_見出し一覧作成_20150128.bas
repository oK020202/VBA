Attribute VB_Name = "B_見出し一覧作成_20150128"
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
'法令事務支援マクロ（Word版）(2/3)　B_見出し一覧作成_20150128
'Copyright(C)2014 Blue Panda
'This software is released under the MIT License.
'http://opensource.org/licenses/mit-license.php
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
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

Sub 見出し一覧作成()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     ぎょうせいの「Super法令Web」及び第一法規「D1-Law.com」（第一法規 法情報総合データベース）
' 　　から出力された条文データのファイル並びにe-Govの条文データをWordに転記したファイルから
' 　　条文の見出し一覧を作成する。
'
' 【履歴】
'     2015/01/28　試作版　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【備考】
' 　　動作に若干の時間を要します。
' 　　エラー等で処理が中断された場合は、再度マクロを実行してください。
'
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

    Call 見出し一覧作成_Lite
    Call 体裁タテ
    Call 罫線タテ
    
    xls.Cells(1, 11).Value = Now()
    xls.Cells(1, JyouColumn + 1).Value = xls.Cells(1, 11).Value - xls.Cells(1, 10).Value
    xls.Cells(1, JyouColumn + 1).NumberFormatLocal = "h:mm:ss"
    xls.Visible = True
    
End Sub
Sub 見出し一覧作成_改正履歴印刷表示()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     ぎょうせいの「Super法令Web」のデータ専用
' 　　見出し一覧の作成と同様の処理を行い、改正履歴のデータを印刷範囲に含め、体裁を調整する。
'
' 【履歴】
'     2015/01/28　試作版　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【備考】
' 　　動作に若干の時間を要します。
' 　　エラー等で処理が中断された場合は、再度マクロを実行してください。
'
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
Dim notice As Integer
  
    Call 見出し一覧作成_Lite
    
    If TextStyle = 2 Then
        Call 体裁ヨコ
        Call 罫線ヨコ
    Else
    'Super法令Web以外の条文データと考えられる場合の処理
        xls.Visible = False
        notice = MsgBox("Super法令Webの条文データではないと思われます。改正履歴を印刷表示しないものに変更しますか?", vbYesNo, "確認")
        Select Case notice
            Case vbYes
                xls.Visible = True
                Call 体裁タテ
                Call 罫線タテ
            Case vbNo
                xls.Visible = True
                Call 体裁ヨコ
                Call 罫線ヨコ
            Case Else
                xls.Visible = True
                Call 体裁ヨコ
                Call 罫線ヨコ
            End Select
    End If
    
    xls.Cells(1, 11).Value = Now()
    xls.Cells(1, JyouColumn + 2).Value = xls.Cells(1, 11).Value - xls.Cells(1, 10).Value
    xls.Cells(1, JyouColumn + 2).NumberFormatLocal = "h:mm:ss"
    xls.Visible = True
    
End Sub


Private Sub 見出し一覧作成_Lite()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     見出し一覧を作成。
'
' 【履歴】
'     2015/01/28　試作版　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  '条文の形式を判別（1 e-Gov 、2 Super法令Web 、3 法情報総合データベース 、4 システム）

    Dim myRange As Range
    Set myRange = ActiveDocument.Range(Start:=0, End:=1)
     
    TextStyle = 3
    If myRange.Font.Name = "ＭＳ Ｐゴシック" Then TextStyle = 1
    
    If myRange = "○" Or myRange = "●" Then TextStyle = 2
    
    If myRange = "○" And myRange.ParagraphFormat.LeftIndent = 0 _
    And myRange.ParagraphFormat.FirstLineIndent = 0 Then TextStyle = 4
    
    If myRange = "●" And myRange.ParagraphFormat.LeftIndent = 0 _
    And myRange.ParagraphFormat.FirstLineIndent = 0 Then TextStyle = 4
    
  '半角括弧を全角にする。
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

  'Excelのシートを作成
    Set xls = CreateObject("Excel.Application")
   
    With xls
    
        .SheetsInNewWorkbook = 1
        .Workbooks.Add
        
        .Cells(1, 10).Value = Now()
        
        '▼確認用（動作を確認する際カンマを外す。）
        '.Visible = True
        
        .Cells(1, 16).Select
        .ActiveSheet.Paste
        
        'メモリ対策
        Excel.Application.CutCopyMode = False
        .Cells(1, 15).Select
        .Selection.Copy

        MaxRowP = 5000  '次の行でMaxRowPが設定されなかった場合の対策
        MaxRowP = .Cells(Rows.Count, 16).End(xlUp).Row

        On Error GoTo 0
   
      '加工用（スペースの前まででカット）
        'e-Gov
        If TextStyle = 1 Then
            For i = 1 To MaxRowP
                .Cells(i, 15).Value = Left(.Cells(i, 16), (InStr(.Cells(i, 16), "　")))
                'e -Govのための処理
                If Left(.Cells(i, 16), 2) = "　第" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "　")) + 1))
                End If
                If Left(.Cells(i, 16), 3) = "　　第" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "　")) + 2))
                End If
                If Left(.Cells(i, 16), 4) = "　　　第" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "　")) + 3))
                End If
                If Left(.Cells(i, 16), 5) = "　　　　第" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "　")) + 4))
                End If
                If Left(.Cells(i, 16), 6) = "　　　　　第" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "　")) + 5))
                End If
                If Left(.Cells(i, 16), 7) = "　　　　　　第" Then
                    .Cells(i, 15).Value = LTrim(Left(.Cells(i, 16), (InStr(LTrim(.Cells(i, 16)), "　")) + 6))
                End If
                If Left(.Cells(i, 16), 6) = "　　　附　則" Then
                    .Cells(i, 15).Value = LTrim(.Cells(i, 16))
                End If
            Next
        End If
        
        'Super法令Web
        If TextStyle = 2 Then
            For i = 1 To MaxRowP
                .Cells(i, 15).Value = Left(.Cells(i, 16), (InStr(.Cells(i, 16), "　")))
                'Super法令Webの改正履歴表示のための処理
                If .Cells(i, 16).IndentLevel = 3 Then .Cells(i, 16).Value = " " + .Cells(i, 16).Value
            Next
        End If

        '法情報総合データベース or システム
        If TextStyle = 3 Or TextStyle = 4 Then
            For i = 1 To MaxRowP
                .Cells(i, 15).Value = Left(.Cells(i, 16), (InStr(.Cells(i, 16), "　")))
            Next
        End If
        
        '条文の構造の確認（民法など「目次」「附則」のない目次も想定）
        Hen = 0
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "編") > 1 Then
                Hen = 1
                Exit For
            End If
        Next
        
        SyouSetsuKanMoku = 1
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "章") > 1 Then
                SyouSetsuKanMoku = 2
                Exit For
            End If
        Next
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "節") > 1 Then
                SyouSetsuKanMoku = 3
                Exit For
            End If
        Next
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "款") > 1 Then
                SyouSetsuKanMoku = 4
                Exit For
            End If
        Next
        For i = 1 To MaxRowP
            If InStr(.Cells(i, 15), "目") > 1 Then
                SyouSetsuKanMoku = 5
                Exit For
            End If
        Next
        
        '目次中の章名等を削除
        MokujiEndRow = 4
        For i = 1 To MaxRowP
            If Left(.Cells(i, 16), 2) = "附則" Or Left(.Cells(i, 16), 3) = "　附則" Then
                MokujiEndRow = i
                Exit For
            End If
        Next
        
        If MokujiEndRow = 4 Then
            If Hen = 1 Then
                n = 0
                i = 1
                Do
                    If Left(.Cells(i, 15), 3) = "第一編" Then n = n + 1
                    i = i + 1
                    If n = 2 Then MokujiEndRow = i - 2
                Loop Until n = 2 Or i = MaxRowP
            End If
        
            If Hen = 0 Then
                n = 0
                i = 1
                Do
                    If Left(.Cells(i, 15), 3) = "第一章" Then n = n + 1
                    i = i + 1
                    If n = 2 Then MokujiEndRow = i - 2
                Loop Until n = 2 Or i = MaxRowP
            End If
        End If
        '目次中の章名等を削除
        .Range(.Cells(1, 15), .Cells(MokujiEndRow, 15)).Clear
    
        '見出し一覧の冒頭部分（題名、改正履歴等）の処理

        For i = 1 To 3
            .Cells(i, 1).Value = .Cells(i, 16)
        Next
        
         j = 5
        JyouColumn = SyouSetsuKanMoku + Hen + 1
        
        '法情報総合データベースの場合の設定
        If TextStyle = 3 Then
            .Cells(4, 1).Value = .Cells(4, 16)
            j = 6
        End If
        
        FirstArticleRow = j
        
    '章名、条名、見出し等の転記（Super法令Web以外の場合）
    
        If TextStyle <> 2 Then

            For i = FirstArticleRow To MaxRowP

                '編、章、節、款、目
                If InStr(.Cells(i, 15), "編") > 1 Then
                    .Cells(j, 2).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "章") > 1 Then
                    .Cells(j, 2 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "節") > 1 Then
                    .Cells(j, 3 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "款") > 1 Then
                    .Cells(j, 4 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "目") > 1 Then
                    .Cells(j, 5 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
            
                '条名及び見出し
                If InStr(.Cells(i, 15), "条") > 1 Then
                    If InStr(.Cells(i - 1, 16), "（") = 1 Then
                        .Cells(j, JyouColumn).Value = .Cells(i, 15) + " " + .Cells(i - 1, 16)
                        j = j + 1
                    Else
                        If InStr(.Cells(i, 16), "　削除") > 1 Then
                            .Cells(j, JyouColumn).Value = .Cells(i, 15) + "　削除"
                            j = j + 1
                        Else
                            .Cells(j, JyouColumn).Value = .Cells(i, 15)
                            j = j + 1
                        End If
                    End If
                End If
            
                '附　則
                If InStr(.Cells(i, 15), "附") >= 1 Then
                    .Cells(j, 2).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
            Next
        End If
    
    '章名、条名、見出し等の転記（Super法令Webの場合）
    
        If TextStyle = 2 Then

            For i = FirstArticleRow To MaxRowP
        
                '改正履歴（Super法令Webのみ対応）
                If Left(.Cells(i, 16), 3) = " （平" Or Left(.Cells(i, 16), 3) = " （昭" _
                Or Left(.Cells(i, 16), 3) = " （大" Or Left(.Cells(i, 16), 3) = " （明" Then
                    .Cells(j - 1, JyouColumn + 1).Value = LTrim(.Cells(i, 16))
                End If
                
                '編、章、節、款、目
                If InStr(.Cells(i, 15), "編") > 1 Then
                    .Cells(j, 2).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "章") > 1 Then
                    .Cells(j, 2 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "節") > 1 Then
                    .Cells(j, 3 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "款") > 1 Then
                    .Cells(j, 4 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
                If InStr(.Cells(i, 15), "目") > 1 Then
                    .Cells(j, 5 + Hen).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
            
                '条名及び見出し
                If InStr(.Cells(i, 15), "条") > 1 Then
                    If InStr(.Cells(i - 1, 16), "（") = 1 Then
                        .Cells(j, JyouColumn).Value = .Cells(i, 15) + " " + .Cells(i - 1, 16)
                        j = j + 1
                    Else
                        If InStr(.Cells(i, 16), "　削除") > 1 Then
                            .Cells(j, JyouColumn).Value = .Cells(i, 15) + "　削除"
                            j = j + 1
                        Else
                            .Cells(j, JyouColumn).Value = .Cells(i, 15)
                            j = j + 1
                        End If
                    End If
                End If
            
                '附　則
                If InStr(.Cells(i, 15), "附") >= 1 Then
                    .Cells(j, 2).Value = LTrim(.Cells(i, 16))
                    j = j + 1
                End If
            Next
        End If
    End With
       
            
      '▼作業用のセルのデータを消去（確認の際は冒頭にカンマを加える。）
        xls.Range(xls.Cells(1, 15), xls.Cells(MaxRowP, 20)).Clear
        
        xls.Cells(1, 1).Select
        
        xls.Visible = True

End Sub

Private Sub 体裁タテ()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     作成された見出し一覧の体裁を調整する。※改正履歴部分を印刷範囲に含めない。
'
' 【履歴】
'     2015/01/28　試作版　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
    With xls
    
       '列の幅、行の高さ、折り返し設定、フォント調整
        .Range(.Cells(1, 1), .Cells(1, JyouColumn - 1)).ColumnWidth = 1.5
        .Range(.Cells(1, JyouColumn), .Cells(1, JyouColumn)).ColumnWidth = 74
        .Range(.Cells(1, 1), .Cells(j, 1)).RowHeight = 17.25
        
        .Range(.Cells(1, JyouColumn), .Cells(j, JyouColumn)).WrapText = True '条及び見出しのセルを折り返し表示

        .Range(.Cells(1, 1), .Cells(j, JyouColumn + 1)).Font.Name = "ＭＳ ゴシック"
        .Range(.Cells(1, 1), .Cells(j, JyouColumn + 1)).Font.Size = 9
        
      '印刷範囲の設定
        .ActiveSheet.PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(j - 1, JyouColumn)).Address
        
      'ヘッダー、フッター
        .ActiveSheet.PageSetup.LeftHeader = "【見出し一覧】 " & .Cells(1, 1).Value
        .ActiveSheet.PageSetup.RightHeader = "&10&P / &N"
        .Cells(1, JyouColumn + 1).Value = Now()
        .Cells(1, JyouColumn + 1).NumberFormat = "yyyy/MM/dd H:mm:ss"
        .ActiveSheet.PageSetup.RightFooter = "&07" & " " & .Cells(1, JyouColumn + 1).Text
        
    End With
    
End Sub

Private Sub 体裁ヨコ()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     作成された見出し一覧の体裁を調整する。※改正履歴部分を印刷範囲に含める。
'
' 【履歴】
'     2015/01/28　試作版　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
    With xls
    
      '列の幅、行の高さ、折り返し設定、フォント調整
        .Range(.Cells(1, 1), .Cells(1, JyouColumn - 1)).ColumnWidth = 1.5
        .Range(.Cells(1, JyouColumn), .Cells(1, JyouColumn)).ColumnWidth = 50
        .Range(.Cells(1, JyouColumn + 1), .Cells(1, JyouColumn + 1)).ColumnWidth = 70
        .Range(.Cells(1, 1), .Cells(j, 1)).RowHeight = 17.25
        
        .Range(.Cells(1, JyouColumn), .Cells(j, JyouColumn + 1)).WrapText = True '条及び見出しのセルを折り返し表示

        .Range(.Cells(1, 1), .Cells(j, JyouColumn + 1)).Font.Name = "ＭＳ ゴシック"
        .Range(.Cells(1, 1), .Cells(j, JyouColumn)).Font.Size = 9
        .Range(.Cells(1, JyouColumn + 1), .Cells(j, JyouColumn + 1)).Font.Size = 8
        
      '印刷の向き（→横）
        .ActiveSheet.PageSetup.Orientation = xlLandscape
        
      '印刷範囲の設定
        .ActiveSheet.PageSetup.PrintArea = .Range(.Cells(1, 1), .Cells(j - 1, JyouColumn + 1)).Address
        
      '行の高さの調整
        .Range(.Cells(1, 1), .Cells(j - 1, 1)).EntireRow.AutoFit
        
      'ヘッダー、フッター
        .ActiveSheet.PageSetup.LeftHeader = "【見出し一覧】 " & .Cells(1, 1).Value
        .ActiveSheet.PageSetup.RightHeader = "&10&P / &N"
        .Cells(1, JyouColumn + 2).Value = Now()
        .Cells(1, JyouColumn + 2).NumberFormat = "yyyy/MM/dd H:mm:ss"
        .ActiveSheet.PageSetup.RightFooter = "&07" & " " & .Cells(1, JyouColumn + 2).Text
        
    End With
    
End Sub

Private Sub 罫線タテ()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     作成された見出し一覧に罫線を引く。
'
' 【履歴】
'     2015/01/28　試作版　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
    With xls

   '章名がある場合、見出し一覧の本則部分に罫線を引く
        BorderEdgeTopRow = 0
        If JyouColumn >= 3 Then
            '２列目の処理
            For i = 1 To j
                If Left(.Cells(i, 2), 1) = "第" Then
                    BorderEdgeTopRow = i
                    Exit For
                End If
            Next
            ListUpperEnd = BorderEdgeTopRow
            
            For i = ListUpperEnd + 1 To j
                If Left(xls.Cells(i, 2), 1) = "第" Then
                    .Range(.Cells(BorderEdgeTopRow, 2), .Cells(i - 1, JyouColumn)).BorderAround _
                    LineStyle:=xlContinuous, Weight:=xlThin  'xlHairline
                    BorderEdgeTopRow = i
                End If
                If Left(.Cells(i, 2), 1) = "附" Then
                    .Range(.Cells(BorderEdgeTopRow, 2), .Cells(i - 1, JyouColumn)).BorderAround _
                    LineStyle:=xlContinuous, Weight:=xlThin  'xlHairline
                    HonsokuEndRow = i - 1
                    Exit For
                End If
            Next
            

            If JyouColumn >= 4 Then
                '３列目の処理
                LineOnOff = 0
                For i = 1 To HonsokuEndRow
                    If Left(.Cells(i, 3), 1) = "第" Then
                         .Range(xls.Cells(i, 3), .Cells(i, JyouColumn)).Borders(xlEdgeTop).LineStyle = xlContinuous
                         .Range(xls.Cells(i, 3), .Cells(i, JyouColumn)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                         LineOnOff = 1
                    End If
                    
                    If Left(.Cells(i, 2), 1) = "第" Then
                        LineOnOff = 0
                    End If
                    
                    If LineOnOff = 1 Then
                        .Cells(i, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Cells(i, 3).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                    End If
                Next
                
                If JyouColumn >= 5 Then
                    '４列目の処理
                    LineOnOff = 0
                    For i = 1 To HonsokuEndRow
                        If Left(.Cells(i, 4), 1) = "第" Then
                            .Range(.Cells(i, 4), .Cells(i, JyouColumn)).Borders(xlEdgeTop).LineStyle = xlContinuous
                            .Range(.Cells(i, 4), .Cells(i, JyouColumn)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                            LineOnOff = 1
                        End If
                    
                        If Left(.Cells(i, 2), 1) = "第" Or Left(.Cells(i, 3), 1) = "第" Then
                            LineOnOff = 0
                        End If
                        
                        If LineOnOff = 1 Then
                            .Cells(i, 4).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Cells(i, 4).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                        End If
                    Next
                    
                    If JyouColumn >= 6 Then
                        '５列目の処理
                        LineOnOff = 0
                        For i = 1 To HonsokuEndRow
                            If Left(.Cells(i, 5), 1) = "第" Then
                                .Range(.Cells(i, 5), .Cells(i, JyouColumn)).Borders(xlEdgeTop).LineStyle = xlContinuous
                                .Range(.Cells(i, 5), .Cells(i, JyouColumn)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                                LineOnOff = 1
                            End If
                            
                            If Left(.Cells(i, 2), 1) = "第" Or Left(.Cells(i, 3), 1) = "第" _
                            Or Left(.Cells(i, 4), 1) = "第" Then
                                LineOnOff = 0
                            End If
                           
                            If LineOnOff = 1 Then
                                .Cells(i, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
                                .Cells(i, 5).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                            End If
                        Next
                        
                        If JyouColumn >= 7 Then
                            '６列目の処理
                            LineOnOff = 0
                            For i = 1 To HonsokuEndRow
                                If Left(.Cells(i, 6), 1) = "第" Then
                                    .Range(.Cells(i, 6), .Cells(i, JyouColumn)).Borders(xlEdgeTop).LineStyle = xlContinuous
                                    .Range(.Cells(i, 6), .Cells(i, JyouColumn)).Borders(xlEdgeTop).Weight = xlThin  'xlHairline
                                    LineOnOff = 1
                                End If
                    
                                If Left(.Cells(i, 2), 1) = "第" Or Left(.Cells(i, 3), 1) = "第" _
                                Or Left(.Cells(i, 4), 1) = "第" Or Left(.Cells(i, 5), 1) = "第" Then
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
            
            '条の列
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

Private Sub 罫線ヨコ()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     作成された見出し一覧に罫線を引く。
'
' 【履歴】
'     2015/01/28　試作版　新規作成
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-'

    With xls
        
      '章名がある場合、見出し一覧の本則部分に罫線を引く
        BorderEdgeTopRow = 0
        If JyouColumn >= 3 Then
            '２列目の処理
            For i = 1 To j
                If Left(.Cells(i, 2), 1) = "第" Then
                    BorderEdgeTopRow = i
                    Exit For
                End If
            Next
            ListUpperEnd = BorderEdgeTopRow
            
            For i = ListUpperEnd + 1 To j
                If Left(.Cells(i, 2), 1) = "第" Then
                    .Range(.Cells(BorderEdgeTopRow, 2), .Cells(i - 1, JyouColumn + 1)).BorderAround _
                    LineStyle:=xlContinuous, Weight:=xlThin  'xlHairline
                    BorderEdgeTopRow = i
                End If
                If Left(.Cells(i, 2), 1) = "附" Then
                    .Range(.Cells(BorderEdgeTopRow, 2), .Cells(i - 1, JyouColumn + 1)).BorderAround _
                    LineStyle:=xlContinuous, Weight:=xlThin  'xlHairline
                    HonsokuEndRow = i - 1
                    Exit For
                End If
            Next

            If JyouColumn >= 4 Then
                '３列目の処理
                LineOnOff = 0
                For i = 1 To HonsokuEndRow
                    If Left(.Cells(i, 3), 1) = "第" Then
                         .Range(.Cells(i, 3), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                         .Range(.Cells(i, 3), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                         LineOnOff = 1
                    End If
                    
                    If Left(.Cells(i, 2), 1) = "第" Then
                        LineOnOff = 0
                    End If
                    
                    If LineOnOff = 1 Then
                        .Cells(i, 3).Borders(xlEdgeLeft).LineStyle = xlContinuous
                        .Cells(i, 3).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                    End If
                Next
                
                If JyouColumn >= 5 Then
                    '４列目の処理
                    LineOnOff = 0
                    For i = 1 To HonsokuEndRow
                        If Left(.Cells(i, 4), 1) = "第" Then
                            .Range(.Cells(i, 4), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                            .Range(.Cells(i, 4), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                            LineOnOff = 1
                        End If
                    
                        If Left(.Cells(i, 2), 1) = "第" Or Left(.Cells(i, 3), 1) = "第" Then
                            LineOnOff = 0
                        End If
                        
                        If LineOnOff = 1 Then
                            .Cells(i, 4).Borders(xlEdgeLeft).LineStyle = xlContinuous
                            .Cells(i, 4).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                        End If
                    Next
                    
                    If JyouColumn >= 6 Then
                        '５列目の処理
                        LineOnOff = 0
                        For i = 1 To HonsokuEndRow
                            If Left(.Cells(i, 5), 1) = "第" Then
                                .Range(.Cells(i, 5), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                                .Range(.Cells(i, 5), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).Weight = xlThin 'xlHairline
                                LineOnOff = 1
                            End If
                            
                            If Left(.Cells(i, 2), 1) = "第" Or Left(.Cells(i, 3), 1) = "第" _
                            Or Left(.Cells(i, 4), 1) = "第" Then
                                LineOnOff = 0
                            End If
                           
                            If LineOnOff = 1 Then
                                .Cells(i, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
                                .Cells(i, 5).Borders(xlEdgeLeft).Weight = xlThin 'xlHairline
                            End If
                        Next
                        
                        If JyouColumn >= 7 Then
                            '６列目の処理
                            LineOnOff = 0
                            For i = 1 To HonsokuEndRow
                                If Left(.Cells(i, 6), 1) = "第" Then
                                    .Range(.Cells(i, 6), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                                    .Range(.Cells(i, 6), .Cells(i, JyouColumn + 1)).Borders(xlEdgeTop).Weight xlThin 'xlHairline
                                    LineOnOff = 1
                                End If
                    
                                If Left(.Cells(i, 2), 1) = "第" Or Left(.Cells(i, 3), 1) = "第" _
                                Or Left(.Cells(i, 4), 1) = "第" Or Left(.Cells(i, 5), 1) = "第" Then
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
            
            '条の列
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
            
            '１行おきに条に色付け
            For i = 2 To HonsokuEndRow
                If .Cells(i, JyouColumn) <> "" Then
                    .Range(.Cells(i, JyouColumn), .Cells(i, JyouColumn + 1)).Interior.ColorIndex = 15
                i = i + 1
                End If
            Next

        '縦線（お好みで）
        '.Range(.Cells(ListUpperEnd, JyouColumn), .Cells(HonsokuEndRow, JyouColumn)).Borders(xlEdgeRight).LineStyle = xlContinuous
        '.Range(.Cells(ListUpperEnd, JyouColumn), .Cells(HonsokuEndRow, JyouColumn)).Borders(xlEdgeRight).Weight = xlThin
        End If

    End With

End Sub
