Attribute VB_Name = "Module1"
Option Compare Text

Function ����������(Yacheika) As String
    Application.Volatile True    ' ������������ ������� �� �����
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
        MsgBox "�� ������� ���������������� ������ MSXML!"
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
    
    '������ �������
    
    '����� �������
    
    Prosklonyat = oHttp.responseText
End Function

Function GetHTTPResponse(ByVal sURL As String) As String
    Application.Volatile True    ' ������������ ������� �� �����
    Dim oXMLHTTP As Object
    
    On Error Resume Next
     Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
   
    On Error GoTo 0
    If oXMLHTTP Is Nothing Then
        MsgBox "�� ������� ���������������� ������ MSXML!"
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
Function �����������_�����(Yacheika As String) As String
Application.Volatile True ' ������������ ������� �� �����

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
����� = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox �����
�������������� = "GENT"
������ = InStr(1, �����, ��������������) + Len(��������������) + 3
��������� = Mid(�����, ������, 200)
����� = InStr(1, ���������, "DATV") - 4
�����������_����� = Mid(�����, ������, �����)
End Function

Function ���������_�����(Yacheika As String) As String
Application.Volatile True ' ������������ ������� �� �����

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
����� = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox �����
�������������� = "DATV"
������ = InStr(1, �����, ��������������) + Len(��������������) + 3
��������� = Mid(�����, ������, 200)
����� = InStr(1, ���������, "ACCS") - 4
���������_����� = Mid(�����, ������, �����)
End Function

Function �����������_�����(Yacheika As String) As String
Application.Volatile True ' ������������ ������� �� �����

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
����� = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox �����
�������������� = "ACCS"
������ = InStr(1, �����, ��������������) + Len(��������������) + 3
��������� = Mid(�����, ������, 200)
����� = InStr(1, ���������, "ABLT") - 4
�����������_����� = Mid(�����, ������, �����)
End Function

Function ������������_�����(Yacheika As String) As String
Application.Volatile True ' ������������ ������� �� �����

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
����� = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox �����
�������������� = "ABLT"
������ = InStr(1, �����, ��������������) + Len(��������������) + 3
��������� = Mid(�����, ������, 200)
����� = InStr(1, ���������, "LOCT") - 4
������������_����� = Mid(�����, ������, �����)
End Function

Function ����������_�����(Yacheika As String) As String
Application.Volatile True ' ������������ ������� �� �����

EncodedUrl = WorksheetFunction.EncodeURL(Yacheika)
����� = GetHTTPResponse("https://cityninja.ru/morpher/" + EncodedUrl)
'MsgBox �����
�������������� = "LOCT"
������ = InStr(1, �����, ��������������) + Len(��������������) + 3
��������� = Mid(�����, ������, 200)
����� = InStr(1, ���������, "}") - 2
����������_����� = Mid(�����, ������, �����)
End Function
