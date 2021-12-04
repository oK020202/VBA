Attribute VB_Name = "C_ルビ一覧作成_20150128"
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
'法令事務支援マクロ（Word版）(3/3)　C_ルビ一覧作成_20150128
'Copyright(C)2014 Blue Panda
'This software is released under the MIT License.
'http://opensource.org/licenses/mit-license.php
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
Option Explicit
    Dim fRng As Range
    Dim Jyou As String '該当する条
    Dim Kou As String '該当する項
    Dim Gou As String '該当する号
    Dim Saibun1 As String '該当する号の細分 イロハ
    Dim Saibun2 As String '該当する号の細分 (1)(2)(3)
    Dim Saibun3 As String '該当する号の細分 (i)(ii)(iii)
   
Sub ルビ一覧作成()
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'     ぎょうせいの「Super法令Web」から出力された条文データからルビの付された文字の一覧を作成する。
'
' 【履歴】
'     2015/01/28　試作版作成
'
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【備考】
' 　　附則中の条項であっても「附則」との表記はありません。
'　　 「Super法令Web」以外のデータでは、ルビの表示形式の差異により、正しく出力されません。
'
' 【未対応の内容等】
' 　　特殊な位置のルビについては、正しく表示されない恐れがあります。
' 　　号の細分が４段階のものには対応できません。
' 　　一項なら成る本則には対応できません。
'
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

  '変数の宣言
    Dim xls As Object
    Dim myfield As Object
    Dim mycode As String
    Dim j As Integer
    Dim myRange As Range
    
  'インデント調整
    Call インデント調整

    Set myRange = ActiveDocument.Sentences(1)
    
  'Excelのシートを作成
    Set xls = CreateObject("Excel.Application")
   
    With xls
        .SheetsInNewWorkbook = 1
        .Workbooks.Add
    End With

  'タイトル等作成
    With xls
        .Cells(1, 1).Value = "【ルビ一覧】 " & myRange
        .Cells(2, 2).Value = "ルビ付きの文字"
        .Cells(2, 3).Value = "該当条項　（注）附則の条項を含む場合があります。"
        .Cells(2, 4).Value = "ページ"
    End With
    
    j = 3
    
  'エラー時の処理
    On Error GoTo 0
    
  'ルビの情報を転記
    For Each myfield In ActiveDocument.Fields
        
        mycode = myfield.Code
        myfield.Select
        
      'ルビ
        xls.Cells(j, 2).Value = ruby(mycode)
        
      'ルビの位置（条項）
        Call 条項位置表示
        xls.Cells(j, 3).Value = Jyou & " " & Kou & " " & Gou & " " & Saibun1 & " " & Saibun2 & " " & Saibun3

      'ルビの位置（ページ）
        xls.Cells(j, 4).Value = Selection.Information(wdActiveEndPageNumber)
        
        j = j + 1

    Next
    
    With xls
       '列の幅、行の高さ、フォント調整
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
        
      '表示
        .Visible = True
        
    End With

End Sub
 
Private Sub インデント調整()

'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'　　 インデントを11ポイントの文字相当に変換する。
'
' 【履歴】
'     2015/01/28　試作版作成
'
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

    Dim FirstFontSize As Single, ChangedFontSize As Single
    Dim Indenting As Single, Indented As Single
    Dim i As Integer
    
    'インデント調整
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

Function ruby(codetext)

'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'　　 ルビのついた文字からルビを取り出す。
'
'　　　もじ
'　　　文字　→　文字（もじ）
'
' 【履歴】
'     2015/01/28　試作版作成
'
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-

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
    
    ruby = rubybase & " （" & rubytext & "）"
     
End Function

Private Sub 条項位置表示()

'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
' 【概要】
'　　 指定位置の条項を表示します。
'
'
' 【履歴】
'     2015/01/28　試作版作成
'
'-･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･- -･-
    Dim Syozoku As String '該当する要素
    Dim a As Integer
    Dim i As Integer
    Dim JyouKou1 As String '条項の仮の値
    
    '値をクリア
    Jyou = ""
    Kou = ""
    Gou = ""
    Saibun1 = ""
    Saibun2 = ""
    Saibun3 = ""
    
    '該当箇所の位置の表示

    Set fRng = ActiveDocument.Range(Start:=0, End:=Selection.Range.End)
    a = fRng.Paragraphs.Count
    Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a).Range.Start, End:=ActiveDocument.Paragraphs(a).Range.End) 'Selection.Range

    '直近の要素
    If InStr(fRng, "　") > 0 Then
        Syozoku = Left(fRng, (InStr((fRng), "　") - 1))
        JyouKou1 = Syozoku
    Else
        '指定の位置が見出しの場合
        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = 0 Then
            Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a + 1).Range.Start, End:=ActiveDocument.Paragraphs(a + 1).Range.End)
            JyouKou1 = Left(fRng, (InStr((fRng), "　") - 1))
            Jyou = JyouKou1 & "見出し"
        End If
    End If
        
        'Case１　指定の位置が条及び項の場合
        If InStr(Left(fRng, (InStr((fRng), "　") - 1)), "条") > 0 And Jyou = "" Then
                Kou = "第一項"
        Else
            If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 And Jyou = "" Then
                Kou = "第" & Syozoku & "項"
            End If
             
        End If
        
        'Case２　指定の位置が章名等の場合
        If fRng.ParagraphFormat.FirstLineIndent = -44 Then
                Jyou = "章名等"
        End If
        
        'Casa３　指定の位置が号の場合
        If fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
            Gou = "第" & Syozoku & "号"
            i = 1
            Do Until (fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                
                If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    If InStr(Left(fRng, (InStr((fRng), "　") - 1)), "条") > 0 Then
                        Kou = "第一項"
                    Else
                        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                            JyouKou1 = Left(fRng, (InStr((fRng), "　") - 1))
                            Kou = "第" & JyouKou1 & "項"
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If
             
        'Case４　指定の位置が号の細分1（主にイロハ）の場合
        If fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
            Saibun1 = Syozoku
            i = 1
            Do Until (fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    JyouKou1 = Left(fRng, (InStr((fRng), "　") - 1))
                    Gou = "第" & JyouKou1 & "号"
                End If
                i = i + 1
            Loop
            
            Do Until (fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                
                If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    If InStr(Left(fRng, (InStr((fRng), "　") - 1)), "条") > 0 Then
                        Kou = "第一項"
                    Else
                        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                            JyouKou1 = Left(fRng, (InStr((fRng), "　") - 1))
                            Kou = "第" & JyouKou1 & "項"
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If
        
        'Case５　指定の位置が号の細分2（主に(1)(2)(3)）の場合
        If fRng.ParagraphFormat.LeftIndent = 44 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
            Saibun2 = Syozoku
            i = 1
            '号の細分1
            Do Until (fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    Saibun1 = Left(fRng, (InStr((fRng), "　") - 1))
                End If
                i = i + 1
            Loop
            '号
            i = 1
            Do Until (fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    JyouKou1 = Left(fRng, (InStr((fRng), "　") - 1))
                    Gou = "第" & JyouKou1 & "号"
                End If
                i = i + 1
            Loop
            '項
            i = 1
            Do Until (fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                
                If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    If InStr(Left(fRng, (InStr((fRng), "　") - 1)), "条") > 0 Then
                        Kou = "第一項"
                    Else
                        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                            JyouKou1 = Left(fRng, (InStr((fRng), "　") - 1))
                            Kou = "第" & JyouKou1 & "項"
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If
        
   
        'Case６　指定の位置が号の細分3（主に(i)(ii)(iii)）の場合
        If fRng.ParagraphFormat.LeftIndent = 55 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
            Saibun3 = Syozoku
            i = 1
            '号の細分2
            Do Until (fRng.ParagraphFormat.LeftIndent = 55 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 44 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    Saibun1 = Left(fRng, (InStr((fRng), "　") - 1))
                End If
                i = i + 1
            Loop
            '号の細分1
            Do Until (fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 33 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    Saibun1 = Left(fRng, (InStr((fRng), "　") - 1))
                End If
                i = i + 1
            Loop
            '号
            Do Until (fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If fRng.ParagraphFormat.LeftIndent = 22 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    JyouKou1 = Left(fRng, (InStr((fRng), "　") - 1))
                    Gou = "第" & JyouKou1 & "号"
                End If
                i = i + 1
            Loop
            '項
            Do Until (fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11) Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                
                If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                    If InStr(Left(fRng, (InStr((fRng), "　") - 1)), "条") > 0 Then
                        Kou = "第一項"
                    Else
                        If fRng.ParagraphFormat.LeftIndent = 11 And fRng.ParagraphFormat.FirstLineIndent = -11 Then
                            JyouKou1 = Left(fRng, (InStr((fRng), "　") - 1))
                            Kou = "第" & JyouKou1 & "項"
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If

             
    '所属の条名
    If Jyou = "" Then
        Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a).Range.Start, End:=ActiveDocument.Paragraphs(a).Range.End)
        If InStr(Left(fRng, (InStr((fRng), "　") - 1)), "条") > 0 Then
            Jyou = JyouKou1
        Else
            i = 1
            Do Until InStr(JyouKou1, "条") > 0 Or i = 100
                Set fRng = ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(a - i).Range.Start, End:=ActiveDocument.Paragraphs(a - i).Range.End)
                If InStr(fRng, "　") > 0 Then
                    JyouKou1 = Left(fRng, (InStr((fRng), "　") - 1))
                End If
                i = i + 1
            Loop
            Jyou = JyouKou1
        End If
    End If
        
End Sub
