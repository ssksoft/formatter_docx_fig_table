' �Q�l�Fhttps://yutako.hateblo.jp/entry/2019/11/24/170711


Const wdAlignParagraphCenter = 1
Const wdCollapseEnd = 0
Const wdstory = 6
Const wdAlignRowCenter = 1
Const fig_name_common = "�}*:"
Const table_name_common = "�\*:"


' �J�����g�t�H���_�̎擾
Dim shell_obj
Dim current_dir
Set shell_obj = CreateObject( "WScript.Shell" )
current_dir = shell_obj.CurrentDirectory
Set shell_obj = Nothing

' ���C������
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
            .Forward = True                 '�������������
            ' .Wrap = wdFindAsk                '�����̐擪/�����܂Ō��������畷��
            .Format = False              '�����ɂ�����炸�Ɍ�������
            .MatchCase = False           '�啶����������ʂ����Ɍ�������  
            .MatchWholeWord = False      '(�p)���S��v�łȂ��Ƃ���������
            .MatchByte = False           '�S�p���p��ʂ����Ɍ�������  
            .MatchAllWordForms = False   '(�p)�قȂ銈�p�`�͌������Ȃ�
            .MatchSoundsLike = False     '(�p)�����܂��Ɍ������Ȃ�
            .MatchFuzzy = False          '(��)�����܂��Ɍ������Ȃ�
            .MatchWildcards = True           '���C���h�J�[�hOn
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