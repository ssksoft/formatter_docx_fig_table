' 参考：https://yutako.hateblo.jp/entry/2019/11/24/170711


Const wdAlignParagraphCenter = 1
Const wdCollapseEnd = 0
Const wdstory = 6
Const wdAlignRowCenter = 1
Const fig_name_common = "図*:"
Const table_name_common = "表*:"


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

    Set doc_obj = word_obj.Documents.Open(target_filename)
    
    Call format_fig(word_obj,doc_obj)
    Call format_table(word_obj,doc_obj)

End Sub

Sub format_caption(word_obj,doc_obj,target_str)
    word_obj.Selection.HomeKey(wdstory)
        With word_obj.Selection.Find                     
            .text = target_str
            .Forward = True                 '検索方向上向き
            ' .Wrap = wdFindAsk                '文書の先頭/末尾まで検索したら聞く
            .Format = False              '書式にこだわらずに検索する
            .MatchCase = False           '大文字小文字区別せずに検索する  
            .MatchWholeWord = False      '(英)完全一致でなくとも検索する
            .MatchByte = False           '全角半角区別せずに検索する  
            .MatchAllWordForms = False   '(英)異なる活用形は検索しない
            .MatchSoundsLike = False     '(英)あいまいに検索しない
            .MatchFuzzy = False          '(日)あいまいに検索しない
            .MatchWildcards = True           'ワイルドカードOn
            Do While .Execute
                word_obj.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            Loop
        End With
    word_obj.Selection.HomeKey(wdstory)
End Sub

Sub format_fig(word_obj,doc_obj)
    For Each iShape In doc_obj.InlineShapes
            iShape.Select
            word_obj.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Next
    Call format_caption(word_obj, doc_obj,fig_name_common)
End Sub

Sub format_table(word_obj,doc_obj)
    For Each table in doc_obj.Tables
        table.Rows.Alignment = wdAlignRowCenter
    Next
    Call format_caption(word_obj, doc_obj,table_name_common)
End Sub