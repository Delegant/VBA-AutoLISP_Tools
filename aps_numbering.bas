Attribute VB_Name = "aps_numbering"
Option Explicit

Sub aps_numbering()
Attribute aps_numbering.VB_ProcData.VB_Invoke_Func = " \n14"

Dim Val As Integer, FirstRow As Integer, FirstColumn As Long
Dim FinalColumn As Long, FinalRow As Long
Dim wsSh As Worksheet
Dim sVal As String, pVal As String, psVal 'значение суфикса, префикса и значащей строки целиком
Dim sstr As Integer 'колличество символов в суфиксе
Dim USRange As Range
Dim i As Integer
Dim j As Integer
ActiveSheet.Copy Before:=ActiveSheet 'копируем лист перед активным
    
    On Error GoTo Errors1
    
    Set USRange = Application.InputBox _
    (prompt:="”кажите диапазон нумеруемых €чеек", Type:=8) 'запрос диапазона
    If USRange Is Nothing Then 'Exit Sub выход при отцутствии диапазона
  
 'If IsNumeric(USRange(1, 1).Value2) Then 'выход если первое значение диапазона не числовое
 'FirstValue = USRange(1, 1).Value2 'задаем значение
 
 'Else
Errors1:       MsgBox ("Ќеверное значение")
        Application.DisplayAlerts = False 'блокируем лишние сообщени€ дл€ "тихого" удалени€ листа
        ActiveSheet.Delete '"тихо" удал€ем лист
        Application.DisplayAlerts = True 'возвращаем сообщени€ системы
        Exit Sub
    End If


FinalRow = USRange.Rows.Count ' последн€€ заполнена€ строка
FinalColumn = USRange.Columns.Count 'последн€€ права€ заполнена€ строка

Application.Calculation = xlCalculationManual 'пересчет значений вручную
Application.ScreenUpdating = False 'отключение обновлени€ экрана


sVal = 0
For i = 1 To FinalRow
    For j = 1 To FinalColumn
        psVal = USRange(i, j).Value2
        sstr = suf(USRange(i, j)) 'находим длинну числового суфикса
        If sstr < 0 Then
            GoTo line1
            ElseIf sstr = 0 And sVal = 0 Then
            sVal = 1 'когда нет начального значени€
            ElseIf sstr > 0 And sVal = 0 Then
            sVal = Right(USRange(i, j).Value2, sstr) 'принимаем начальное значение
        End If
        pVal = Left(psVal, (Len(psVal) - sstr))
        USRange(i, j).Value2 = pVal + sVal 'переписываем значение
        sVal = sVal + 1
line1:
    Next j
Next i

On Error Resume Next 'при ошибке пл€шем дальше
Set wsSh = Sheets("–езультат") 'задаем переменной ссылку на объект - лист "результат"
If wsSh Is Nothing Then 'ѕроверка наличи€ листа "результат"
ActiveSheet.Name = "–езультат" 'даем текущему листу им€ "результат"
Else
Application.DisplayAlerts = False 'отключение оповещени€
Sheets("–езультат").Delete
Application.DisplayAlerts = True 'включаем оповещени€
ActiveSheet.Name = "–езультат" 'даем текущему листу им€ "результат"
End If

Application.Calculation = xlCalculationAutomatic 'пересчет значений автоматически
Application.ScreenUpdating = True 'включение обновлени€ экрана
End Sub

Private Function suf(US As Range) 'функци€ по определению размера числового суфикса
Dim USval As String
Dim i As Integer
i = 1
USval = US.Value2

Do While Not IsNumeric(USval) And Len(USval) > 0 'цикл пока значение не числовое или длина строки не ноль
 USval = Right(USval, Len(USval) - i) 'убираем один символ с права
 Loop 'повтор€ем цикл, вышли из цикла с размером суфикса (0 если его нет)
 
 Select Case US 'отсев системных знаков
 Case "<>"
 suf = -1
 Case "."
 suf = -1
 US.Value2 = "<>"
 US.Interior.Color = vbYellow
 Case Else
 suf = Len(USval)
 End Select
 
 
End Function

'Private Function pref(USRange As Range)


'End Function
