Attribute VB_Name = "SpeedUp"
'高速化､構造体､動的メモリ確保､Sleepなど色々

Option Explicit

' Sleep関数を使う時
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'構造体（ユーザー定義型）
Type Person
    Name As String
    Age As Long
End Type


Sub Main()

On Error GoTo CATCH
    
    '高速化処理
    Application.ScreenUpdating = False '描画停止
    Application.EnableEvents = False 'イベント抑制
    Application.Calculation = xlCalculationManual '手動計算

    '最終行を検知して、そのサイズのメモリを動的確保
    Dim Arr() As Person
    Dim LastRowNum As Long
    LastRowNum = ThisWorkbook.Sheets("Sheet1").Range("A100").End(xlUp).Row
    ReDim Arr(LastRowNum)
    'MsgBox LBound(Arr) & vbCrLf & UBound(Arr)

    'シートのデータをメモリに格納
    Dim i As Long
    For i = 2 To UBound(Arr)
        Arr(i).Name = ThisWorkbook.Sheets("Sheet1").Cells(i, 1)
        Arr(i).Age = ThisWorkbook.Sheets("Sheet1").Cells(i, 2)
        Debug.Print Arr(i).Name & Arr(i).Age
        'DoEvents   'OSに処理を渡す時
    Next

    'MsgBox "終了"
    GoTo FINAL
CATCH:
    MsgBox "エラー終了"
FINAL:
    '高速化の解除
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

End Sub

