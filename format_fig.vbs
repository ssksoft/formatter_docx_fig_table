Dim shell_obj
Dim current_dir
Const wdAlignParagraphCenter = 1

' カレントディレクトリ取得.
Set shell_obj = CreateObject( "WScript.Shell" )
current_dir = shell_obj.CurrentDirectory

Set shell_obj = Nothing

Set word_obj = CreateObject("Word.Application")
word_obj.Visible = True

target_filename = current_dir & "\sample.docx"

Set target_obj = word_obj.Documents.Open(target_filename)

For Each iShape In target_obj.InlineShapes
    iShape.Select
    word_obj.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
Next


' word_obj.Quit