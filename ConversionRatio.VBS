Option Explicit

Const TARGET_URL = "http://free.currconv.com/api/v7/convert?apiKey=%KEY%&q=%TARGET%&compact=ultra"

Const SXH_PROXY_SET_DEFAULT = 0
Const SXH_PROXY_SET_DIRECT = 1
Const SXH_PROXY_SET_PRECONFIG = 0
Const SXH_PROXY_SET_PROXY = 2

Const SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = &H3300
Const SXH_SERVER_CERT_IGNORE_CERT_CN_INVALID = &H1000
Const SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID = &H2000
Const SXH_SERVER_CERT_IGNORE_UNKNOWN_CA = &H100
Const SXH_SERVER_CERT_IGNORE_WRONG_USAGE = &H200

Const SXH_OPTION_ESCAPE_PERCENT_IN_URL = 1
Const SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS = 2
Const SXH_OPTION_SELECT_CLIENT_SSL_CERT = 3
Const SXH_OPTION_URL = -1
Const SXH_OPTION_URL_CODEPAGE = 0

'====================================================================
Public Function GetConversionRatios(ByRef Targets)
  Dim xmlServerHttp
  Dim TargetURL
  Dim TXT
  
  On Error Resume Next
  
  Set GetConversionRatios = Nothing
  
  TargetURL = TARGET_URL
  TargetURL = Replace(TargetURL, "%KEY%", API_KEY)
  TargetURL = Replace(TargetURL, "%TARGET%", Join(Targets, ","))
  
  Set xmlServerHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0") ' .4.0
  
  'xmlServerHttp.setProxy SXH_PROXY_SET_DEFAULT ' SXH_PROXY_SET_DIRECT
  'xmlServerHttp.setOption SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS, _
  '                        SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
  
  xmlServerHttp.open "GET", TargetURL, False
  xmlServerHttp.setRequestHeader "Pragma", "No-Cache"
  xmlServerHttp.send
  
  If Err.Number <> 0 Then
    MsgBox "Не удалось плучить курс валюты. Ошибка отправки запроса.", _
           vbCritical, AppTitle
    Exit Function
  End If
  
  ' Debug!
  'MsgBox xmlServerHttp.Status & " " & xmlServerHttp.statusText & vbCrLf & _
  '       xmlServerHttp.getAllResponseHeaders() & _
  '       xmlServerHttp.responseText
  
  ' Debug!
  'MsgBox xmlServerHttp.responseText
  'MsgBox xmlServerHttp.responseXML.xml
  
  If xmlServerHttp.Status <> 200 Then
    TXT = xmlServerHttp.responseText
    If TXT <> "" Then TXT = GetErrorMessage(TXT)
    If TXT <> "" Then TXT = vbCrLf & vbCrLf & TXT
    
    MsgBox "Не удалось плучить курс валюты." & vbCrLf & vbCrLf & _
           "Сервер говорит: " & _
           xmlServerHttp.Status & " " & xmlServerHttp.statusText & _
           TXT, vbCritical, AppTitle
    Exit Function
  End If
  
  ' Debug!
  'MsgBox TypeName(xmlServerHttp.responseBody)
  'MsgBox CStr(xmlServerHttp.responseBody)
  'MsgBox TypeName(xmlServerHttp.responseXML)
  'MsgBox TypeName(xmlServerHttp.responseXML.xml)
  'MsgBox xmlServerHttp.responseXML.xml
  'MsgBox xmlServerHttp.responseText
  
  Set GetConversionRatios = ParseRatios(xmlServerHttp.responseText)
End Function

'====================================================================
Private Function GetErrorMessage(ByVal Text)
  Dim RE 'As RegExp
  Dim MC 'As MatchCollection
  
  On Error Resume Next
  
  Set RE = New RegExp
  
  RE.Global = True
  RE.IgnoreCase = True
  RE.MultiLine = True
  
  RE.Pattern = """error"":""(.*?)"""
  
  Set MC = RE.Execute(Text)
  
  GetErrorMessage = ""
  GetErrorMessage = MC(0).SubMatches(0)
  Err.Clear
End Function

'====================================================================
Private Function ParseRatios(ByVal JSON)
  Dim RE 'As RegExp
  Dim MC 'As MatchCollection
  Dim TMP
  Dim Ratio
  
  On Error Resume Next
  
  Set ParseRatios = CreateObject("Scripting.Dictionary")
  Set RE = New RegExp
  
  RE.Global = True
  RE.IgnoreCase = True
  RE.MultiLine = True
  
  RE.Pattern = """(.*?)"":([1234567890.]*)"
  
  Set MC = RE.Execute(JSON)
  If Err.Number <> 0 Then
    MsgBox "Курс валюты получен, но JSON распарсить не удалось. #1", _
           vbCritical, AppTitle
    Exit Function
  End If
  
  ' Debug!
  'MsgBox MC.Count
  'MsgBox MC(0).SubMatches.Count
  'MsgBox MC(0).SubMatches(0)
  'MsgBox MC(0).SubMatches(1)
  
  If MC.Count = 0 Then
    MsgBox "Курс валюты получен, но JSON распарсить не удалось. #2", _
           vbCritical, AppTitle
    Exit Function
  End If
  
  For TMP = 0 To MC.Count - 1
    If MC(TMP).SubMatches.Count <> 2 Then
      MsgBox "Курс валюты получен, но JSON распарсить не удалось. #3", _
             vbCritical, AppTitle
      Exit Function
    End If
    
    Err.Clear
    Ratio = MC(TMP).SubMatches(1)
    Ratio = CDbl(Ratio)
    If Err.Number <> 0 Then
      Err.Clear
      Ratio = CDbl(Replace(Ratio, ".", ","))
    End If
    
    If Err.Number <> 0 Then
        MsgBox "Курс валюты получен, но JSON распарсить не удалось. #4", _
               vbCritical, AppTitle
      Exit Function
    End If
    
    ParseRatios.Add MC(TMP).SubMatches(0), Ratio
  Next
End Function
