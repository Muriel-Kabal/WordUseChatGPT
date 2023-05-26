Sub ChatGPT()
    
    ' Set endpoint URL
    Dim endpoint As String
    Dim selectedText As String
    Dim requestBody As String
    Dim request As Object

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")

    endpoint = "http://127.0.0.1:3002/conversation"
    
    ' Create request object
    If Selection.Type = wdSelectionNormal Then
    selectedText = Selection.Text
    selectedText = Replace(selectedText, ChrW$(13), "")
    
    ' Set request properties
    request.Open "POST", endpoint, False
    request.setRequestHeader "Content-Type", "application/json"
    requestBody = "{""message"":"""& selectedText & """}"
    request.Send requestBody

    ' 解析Json
    request.ResponseText = Replace(request.ResponseText, "{", "")
    request.ResponseText = Replace(request.ResponseText, "}", "")
    Dim arr() As String
    arr = Split(request.ResponseText, ":")
    arr = Split(arr(1), ",")
    Dim ans As String
    ans = Replace(arr(0), ChrW$(34), "")
	' 换行符
	ans = Replace(ans, "\n", ChrW$(10))
    
    
    
    Selection.Text = selectedText & vbNewLine & "回答：" & ans
Else
Exit Sub
End If
End Sub