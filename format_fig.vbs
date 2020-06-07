Const wdAlignParagraphCenter = 1
Const wdCollapseEnd = 0

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

    Set target_obj = word_obj.Documents.Open(target_filename)
    
    ' Call format_fig(word_obj,target_obj)
    Call format_caption(word_obj, target_obj,"�}*:")

End Sub


' word_obj.Quit

Sub format_caption(word_obj,target_obj,target_str)
    target_obj.Bookmarks("\EndOfDoc").Select
    word_obj.Selection.Collapse(wdCollapseEnd)
        With word_obj.Selection.Find                     
            .text = "�}*:"
            .Forward = False                 '�������������
            .Wrap = wdFindAsk                '�����̐擪/�����܂Ō��������畷��
            .Format = False              '�����ɂ�����炸�Ɍ�������
            .MatchCase = False           '�啶����������ʂ����Ɍ�������  
            .MatchWholeWord = False      '(�p)���S��v�łȂ��Ƃ���������
            .MatchByte = False           '�S�p���p��ʂ����Ɍ�������  
            .MatchAllWordForms = False   '(�p)�قȂ銈�p�`�͌������Ȃ�
            .MatchSoundsLike = False     '(�p)�����܂��Ɍ������Ȃ�
            .MatchFuzzy = False          '(��)�����܂��Ɍ������Ȃ�
            .MatchWildcards = True           '���C���h�J�[�hOn
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
    ' �J�����g�f�B���N�g���擾.
    
End Function


