Attribute VB_Name = "Module1"
Option Compare Text

Function ДательныйП(Yacheika) As String
    Application.Volatile True    ' автопересчёт формулы на листе
    Dim oHttp As Object
    Dim strURL As String
    Dim EncodeURL As String
    Dim str As Variant, N&, R&
    
    strURL = "https://cityninja.ru/morpher/"
    EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
    strURL = strURL & EncodedUrl
    
    On Error Resume Next
    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    
    If Err.Number <> 0 Then
        Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
        
    End If
    On Error GoTo 0
    If oHttp Is Nothing Then
        MsgBox "Не удалось инициализировать объект MSXML!"
        Exit Function
    End If
    
    With oHttp
    .Open "GET", strURL, False
    .SetRequestHeader "Cache-Control", "max-age=0"
    .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/48.0.2564.41 Safari/537.36 OPR/35.0.2066.10 (Edition beta)"
    .SetRequestHeader "Accept-Encoding", "UTF-8"
    .SetRequestHeader "Accept-Language", "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4"
    .send
    End With
    
    'НАЧАЛО ПАРСЕРА
    
    'КОНЕЦ ПАРСЕРА
    
    Prosklonyat = oHttp.responseText
End Function

Function GetHTTPResponse(ByVal sURL As String) As String
    Application.Volatile True    ' автопересчёт формулы на листе
    Dim oXMLHTTP As Object
    
    On Error Resume Next
     Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
   
    On Error GoTo 0
    If oXMLHTTP Is Nothing Then
        MsgBox "Не удалось инициализировать объект MSXML!"
        Exit Function
    End If
    
    With oXMLHTTP
        .Open "GET", sURL, False
        .SetRequestHeader "Accept-Encoding", "UTF-8"
        .SetRequestHeader "Accept-Language", "ru-RU,ru;q=0.8,en-US;q=0.6,en;q=0.4"
        .send
        GetHTTPResponse = .responseText
    End With
        
    Set oXMLHTTP = Nothing
End Function
Function Родительный_падеж(Yacheika As String) As String
Application.Volatile True ' автопересчёт формулы на листе

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
текст = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox текст
НачальныйТекст = "GENT"
Начало = InStr(1, текст, НачальныйТекст) + Len(НачальныйТекст) + 3
Подстрока = Mid(текст, Начало, 200)
Конец = InStr(1, Подстрока, "DATV") - 4
Родительный_падеж = Mid(текст, Начало, Конец)
End Function

Function Дательный_падеж(Yacheika As String) As String
Application.Volatile True ' автопересчёт формулы на листе

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
текст = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox текст
НачальныйТекст = "DATV"
Начало = InStr(1, текст, НачальныйТекст) + Len(НачальныйТекст) + 3
Подстрока = Mid(текст, Начало, 200)
Конец = InStr(1, Подстрока, "ACCS") - 4
Дательный_падеж = Mid(текст, Начало, Конец)
End Function

Function Винительный_падеж(Yacheika As String) As String
Application.Volatile True ' автопересчёт формулы на листе

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
текст = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox текст
НачальныйТекст = "ACCS"
Начало = InStr(1, текст, НачальныйТекст) + Len(НачальныйТекст) + 3
Подстрока = Mid(текст, Начало, 200)
Конец = InStr(1, Подстрока, "ABLT") - 4
Винительный_падеж = Mid(текст, Начало, Конец)
End Function

Function Творительный_падеж(Yacheika As String) As String
Application.Volatile True ' автопересчёт формулы на листе

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
текст = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox текст
НачальныйТекст = "ABLT"
Начало = InStr(1, текст, НачальныйТекст) + Len(НачальныйТекст) + 3
Подстрока = Mid(текст, Начало, 200)
Конец = InStr(1, Подстрока, "LOCT") - 4
Творительный_падеж = Mid(текст, Начало, Конец)
End Function

Function Предложный_падеж(Yacheika As String) As String
Application.Volatile True ' автопересчёт формулы на листе

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
текст = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox текст
НачальныйТекст = "LOCT"
Начало = InStr(1, текст, НачальныйТекст) + Len(НачальныйТекст) + 3
Подстрока = Mid(текст, Начало, 200)
Конец = InStr(1, Подстрока, "}") - 2
Предложный_падеж = Mid(текст, Начало, Конец)
End Function
