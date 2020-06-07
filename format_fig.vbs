Const wdAlignParagraphCenter = 1
Const wdCollapseEnd = 0

' カレントフォルダの取得
Dim shell_obj
Dim current_dir
Set shell_obj = CreateObject( "WScript.Shell" )
current_dir = shell_obj.CurrentDirectory
Set shell_obj = Nothing

' メイン処理
main(current_dir)

Sub main(current_dir)
    Set word_obj = CreateObject("Word.Application")
    word_obj.Visible = True

    target_filename = current_dir & "\sample.docx"

    Set target_obj = word_obj.Documents.Open(target_filename)
    
    ' Call format_fig(word_obj,target_obj)
    Call format_caption(word_obj, target_obj,"図*:")

End Sub


' word_obj.Quit

Sub format_caption(word_obj,target_obj,target_str)
    target_obj.Bookmarks("\EndOfDoc").Select
    word_obj.Selection.Collapse(wdCollapseEnd)
        With word_obj.Selection.Find                     
            .text = "図*:"
            .Forward = False                 '検索方向上向き
            .Wrap = wdFindAsk                '文書の先頭/末尾まで検索したら聞く
            .Format = False              '書式にこだわらずに検索する
            .MatchCase = False           '大文字小文字区別せずに検索する  
            .MatchWholeWord = False      '(英)完全一致でなくとも検索する
            .MatchByte = False           '全角半角区別せずに検索する  
            .MatchAllWordForms = False   '(英)異なる活用形は検索しない
            .MatchSoundsLike = False     '(英)あいまいに検索しない
            .MatchFuzzy = False          '(日)あいまいに検索しない
            .MatchWildcards = True           'ワイルドカードOn
            .Execute
        '     word_obj.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
End Sub

Sub format_fig(word_obj,target_obj)
    For Each iShape In target_obj.InlineShapes
            iShape.Select
            word_obj.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Next
End Sub

Function get_current_dir(shell_obj)
    ' カレントディレクトリ取得.
    
End Function


