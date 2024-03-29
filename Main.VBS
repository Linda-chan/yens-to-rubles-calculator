Option Explicit

Const AppTitle = "AJPapps - Yens to rubles calculator"
Const AppCopyright = "����� ���� 2009-2021. ����������� ������"

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
  MainMessage = "������� ����� � �����. ��� ����� ���������� " & _
                "� ����� � ������ ���������� �����:" & _
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
    Calculate = "�� ����� ������������ �����."
    Exit Function
  End If
  
  If Ratios.Exists("JPY_USD") Then
    Bucks = FormatNumber(Yens * Ratios("JPY_USD"))
  Else
    Bucks = "N/A"
  End If
  
  If Err.Number <> 0 Then
    Calculate = "� ����������� JPY > USD ��������� �����-�� ������."
    Exit Function
  End If
  
  If Ratios.Exists("JPY_RUB") Then
    Rubs = FormatNumber(Yens * Ratios("JPY_RUB"))
  Else
    Rubs = "N/A"
  End If
  
  If Err.Number <> 0 Then
    Calculate = "� ����������� JPY > RUB ��������� �����-�� ������."
    Exit Function
  End If
  
  Calculate = "����� � �����: "    & vbTab & FormatNumber(Yens) & vbCrLf & _
              "����� � ��������: " & vbTab & Bucks & vbCrLf & _
              "����� � ������: "   & vbTab & Rubs
End Function
