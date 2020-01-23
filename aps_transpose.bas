Attribute VB_Name = "aps_transpose"
Option Explicit



Sub aps_transpose()
Dim i As Integer, k As Integer, shl As Integer
Dim FinalColumn As Long, FinalRow As Long
Dim arshl As Range
Dim wsSh As Worksheet
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
ActiveSheet.Copy Before:=ActiveSheet
FinalRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row ' последн€€ заполнена€ строка
 
For i = FinalRow To 2 Step -1 'обратный цикл по строкам с низу вверх
FinalColumn = ActiveSheet.Cells(i, Columns.Count).End(xlToLeft).Column 'последн€€ права€ заполнена€ строка
shl = Application.RoundUp(((FinalColumn - 3) / 4), 0) 'колличество шлейфов из расчета шлейф кратен 4(-3 п€ть первых столбцов)
If shl > 1 Then 'условие дл€ отсева приборов без шлейфов
Range(Rows(i + 1), Rows(i + shl - 1)).Insert Shift:=xlDown 'вставка диапазона под строку i + колличество шлейфов -1
 
 If shl > 1 Then 'условие что шлейфов должно быть больше одного
 For k = 1 To (shl - 1) 'цикл вставки содержимого в шлейфе по количеству шлейфов
 Cells(i + k, 2) = "вых." & k
 Set arshl = Range(Cells(i, 4), Cells(i, 7)) 'задан диапазон с 4 по 7 столбцы
 arshl.Offset(k, 0).Value = arshl.Offset(0, 4 * k).Value 'присвоение значени€ диапазона arshl
 Next k
 End If
End If
Next i

FinalRow = ActiveSheet.Cells(Rows.Count, 10).End(xlUp).Row
FinalColumn = ActiveSheet.Cells(1, Columns.Count).End(xlToLeft).Column
Range(Cells(1, 8), Cells(FinalRow, FinalColumn)).ClearContents
On Error Resume Next
Set wsSh = Sheets("–езультат")
If wsSh Is Nothing Then
ActiveSheet.Name = "–езультат"
Else
Application.DisplayAlerts = False
Sheets("–езультат").Delete
Application.DisplayAlerts = True
ActiveSheet.Name = "–езультат"
End If
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub

