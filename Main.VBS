Option Explicit

Const AppTitle = "AJPapps - Yens to rubles calculator"
Const AppCopyright = "Линда Кайе 2009-2021. Посвящается Ариэль"

Const NA = "N/A"

'====================================================================
Public Sub Main()
  Dim Yens
  Dim Message
  Dim MainMessage
  Dim Ratios
  
  Set Ratios = GetConversionRatios(Array("JPY_RUB", "JPY_USD"))
  If Ratios Is Nothing Then Exit Sub
  
  ' Debug!
  'MsgBox "Ratio for JPY > RUB is " & Ratios("JPY_RUB") & vbCrLf & _
  '       "Ratio for JPY > USD is " & Ratios("JPY_USD") & vbCrLf & _
  '       "Ratio for JPY > BTC is " & Ratios("JPY_BTC"), _
  '       vbInformation, AppTitle
  
  Message = ""
  MainMessage = "Введите сумму в йенах. Она будет переведена " & _
                "в рубли с учётом следующего курса:" & _
                vbCrLf & vbCrLf & _
                "JPY > RUB: " & vbTab & GetSafeItem(Ratios, "JPY_RUB", NA) & vbCrLf & _
                "JPY > USD: " & vbTab & GetSafeItem(Ratios, "JPY_USD", NA)
  
  Do
    Yens = InputBox(Message & MainMessage, AppTitle)
    If Yens = "" Then 
      Exit Do
    Else
      Message = Calculate(Yens, Ratios)
      If Message <> "" Then _
        Message = Message & vbCrLf & vbCrLf
    End If
  Loop
End Sub

'====================================================================
Private Function Calculate(ByVal Yens, ByRef Ratios)
  Dim Bucks
  Dim Rubs
  
  On Error Resume Next
  
  Yens = CDbl(Yens)
  If Err.Number <> 0 Then
    Calculate = "Вы ввели неправильное число."
    Exit Function
  End If
  
  If Ratios.Exists("JPY_USD") Then
    Bucks = FormatNumber(Yens * Ratios("JPY_USD"))
  Else
    Bucks = "N/A"
  End If
  
  If Err.Number <> 0 Then
    Calculate = "В вычислениях JPY > USD произошла какая-то ошибка."
    Exit Function
  End If
  
  If Ratios.Exists("JPY_RUB") Then
    Rubs = FormatNumber(Yens * Ratios("JPY_RUB"))
  Else
    Rubs = "N/A"
  End If
  
  If Err.Number <> 0 Then
    Calculate = "В вычислениях JPY > RUB произошла какая-то ошибка."
    Exit Function
  End If
  
  Calculate = "Сумма в йенах: "    & vbTab & FormatNumber(Yens) & vbCrLf & _
              "Сумма в долларах: " & vbTab & Bucks & vbCrLf & _
              "Сумма в рублях: "   & vbTab & Rubs
End Function
