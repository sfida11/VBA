' シート名称・・・(例)シート名(カウンタ)
' セル：Q8・・・シート名称を記載
' セル：Q7・・・シート数のカウンタを記載

' シート名称変更
' シート名を変更する際に使用する
' 初期設定時に使用する
Public Sub シート名変更_Click()
    ' 「[シート名称]+"("+[シートカウンタ]+")"」シート名変更
    ActiveSheet.Name = Range("Q8").Value + "(" + CStr(Range("Q7").Value) + ")"
End Sub
' シート挿入
' シートを挿入する際に使用する
' シート名称の括弧内のカウンタをインクリメントした名称として作成する
Sub シートの追加_Click()
    ' 「[シート名称]+"("+[シートカウンタ]+")"」シート追加
    Sheets(Range("Q8").Value + "(" + CStr(Range("Q7").Value) + ")").Copy After:=Sheets(Sheets.Count)
    ' [シートカウンタ]カウントアップ
    Range("Q7").Value = Range("Q7").Value + 1
End Sub
' シート削除
' シートを削除する際に使用する
' カウンタの数値が１の場合、削除出来ないメッセージをポップアップにて表示する
Sub シートの削除_Click()
    If Range("Q7").Value = 1 Then
        MsgBox "シートが無くなってしまいます" & vbCrLf & "(シートの削除は行いません)"
        Exit Sub
    End If
    ' 警告メッセージ無効化
    Application.DisplayAlerts = False
    ' 「[シート名称]+"("+[シートカウンタ]+")"」シート削除
    Sheets(Range("Q8").Value + "(" + CStr(Range("Q7").Value) + ")").Delete
    ' 警告メッセージ有効化
    Application.DisplayAlerts = True
End Sub
