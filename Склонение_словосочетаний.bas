Attribute VB_Name = "Module1"
Option Compare Text

Function ДательныйП(Yacheika) As String
    Application.Volatile True    ' àâòîïåðåñ÷¸ò ôîðìóëû íà ëèñòå
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
        MsgBox "Íå óäàëîñü èíèöèàëèçèðîâàòü îáúåêò MSXML!"
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
    
    'ÍÀ×ÀËÎ ÏÀÐÑÅÐÀ
    
    'ÊÎÍÅÖ ÏÀÐÑÅÐÀ
    
    Prosklonyat = oHttp.responseText
End Function

Function GetHTTPResponse(ByVal sURL As String) As String
    Application.Volatile True    ' àâòîïåðåñ÷¸ò ôîðìóëû íà ëèñòå
    Dim oXMLHTTP As Object
    
    On Error Resume Next
     Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
   
    On Error GoTo 0
    If oXMLHTTP Is Nothing Then
        MsgBox "Íå óäàëîñü èíèöèàëèçèðîâàòü îáúåêò MSXML!"
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
Function Ðîäèòåëüíûé_ïàäåæ(Yacheika As String) As String
Application.Volatile True ' àâòîïåðåñ÷¸ò ôîðìóëû íà ëèñòå

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
òåêñò = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox òåêñò
Íà÷àëüíûéÒåêñò = "GENT"
Íà÷àëî = InStr(1, òåêñò, Íà÷àëüíûéÒåêñò) + Len(Íà÷àëüíûéÒåêñò) + 3
Ïîäñòðîêà = Mid(òåêñò, Íà÷àëî, 200)
Êîíåö = InStr(1, Ïîäñòðîêà, "DATV") - 4
Ðîäèòåëüíûé_ïàäåæ = Mid(òåêñò, Íà÷àëî, Êîíåö)
End Function

Function Äàòåëüíûé_ïàäåæ(Yacheika As String) As String
Application.Volatile True ' àâòîïåðåñ÷¸ò ôîðìóëû íà ëèñòå

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
òåêñò = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox òåêñò
Íà÷àëüíûéÒåêñò = "DATV"
Íà÷àëî = InStr(1, òåêñò, Íà÷àëüíûéÒåêñò) + Len(Íà÷àëüíûéÒåêñò) + 3
Ïîäñòðîêà = Mid(òåêñò, Íà÷àëî, 200)
Êîíåö = InStr(1, Ïîäñòðîêà, "ACCS") - 4
Äàòåëüíûé_ïàäåæ = Mid(òåêñò, Íà÷àëî, Êîíåö)
End Function

Function Âèíèòåëüíûé_ïàäåæ(Yacheika As String) As String
Application.Volatile True ' àâòîïåðåñ÷¸ò ôîðìóëû íà ëèñòå

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
òåêñò = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox òåêñò
Íà÷àëüíûéÒåêñò = "ACCS"
Íà÷àëî = InStr(1, òåêñò, Íà÷àëüíûéÒåêñò) + Len(Íà÷àëüíûéÒåêñò) + 3
Ïîäñòðîêà = Mid(òåêñò, Íà÷àëî, 200)
Êîíåö = InStr(1, Ïîäñòðîêà, "ABLT") - 4
Âèíèòåëüíûé_ïàäåæ = Mid(òåêñò, Íà÷àëî, Êîíåö)
End Function

Function Òâîðèòåëüíûé_ïàäåæ(Yacheika As String) As String
Application.Volatile True ' àâòîïåðåñ÷¸ò ôîðìóëû íà ëèñòå

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
òåêñò = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox òåêñò
Íà÷àëüíûéÒåêñò = "ABLT"
Íà÷àëî = InStr(1, òåêñò, Íà÷àëüíûéÒåêñò) + Len(Íà÷àëüíûéÒåêñò) + 3
Ïîäñòðîêà = Mid(òåêñò, Íà÷àëî, 200)
Êîíåö = InStr(1, Ïîäñòðîêà, "LOCT") - 4
Òâîðèòåëüíûé_ïàäåæ = Mid(òåêñò, Íà÷àëî, Êîíåö)
End Function

Function Ïðåäëîæíûé_ïàäåæ(Yacheika As String) As String
Application.Volatile True ' àâòîïåðåñ÷¸ò ôîðìóëû íà ëèñòå

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
òåêñò = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox òåêñò
Íà÷àëüíûéÒåêñò = "LOCT"
Íà÷àëî = InStr(1, òåêñò, Íà÷àëüíûéÒåêñò) + Len(Íà÷àëüíûéÒåêñò) + 3
Ïîäñòðîêà = Mid(òåêñò, Íà÷àëî, 200)
Êîíåö = InStr(1, Ïîäñòðîêà, "}") - 2
Ïðåäëîæíûé_ïàäåæ = Mid(òåêñò, Íà÷àëî, Êîíåö)
End Function
