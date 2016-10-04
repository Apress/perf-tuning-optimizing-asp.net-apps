Const SERVER_NAME = "kenneth-01"
Const PATH = "/Duwamish7vb"
Dim oIE6Request
Set oIE6Request = Test.CreateRequest
Call MakeIE60GETRequest(oIE6Request)
Call SendRequest(oIE6Request)

Sub SendRequest(oRequest)
    Dim oConnection

    Set oConnection = Test.CreateConnection(SERVER_NAME, 80, False)

    If (oConnection Is Nothing) Then
        Test.Trace("Error: Unable to create connection.")
    Else
        If (oConnection.IsOpen) Then
			oRequest.Path = PATH
            Set oResponse = oConnection.Send(oRequest)
            ' check for a bad connection or request
            If (oResponse Is Nothing) Then
                Test.Trace("Error: invalid request or host not found ")
            Else
                Test.Trace("Server response:" & oResponse.ResultCode)
            End If
        Else
            Test.Trace("Connection was closed")         
        End If
        Call oConnection.Close()
    End If
End Sub

Function MakeIE60GETRequest(oRequest)
      ' function returns a request object with the appropriate header setup 
      Dim oHeaders
      If not (oRequest is nothing) Then
         ' set request line
         oRequest.Verb = "GET"
         oRequest.HTTPVersion = "HTTP/1.1"
         Set oHeaders = oRequest.Headers
         With oHeaders
            Call .RemoveAll()     
            ' set header fields
            Call .Add("Accept", "*/*")
            Call .Add("Accept-Language", "en-us")
            Call .Add("Connection", "Keep-Alive")
            Call .Add("Host", "kenneth-01")
            Call .Add("User-Agent", " Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)")
            Call .Add("Accept-Encoding", "gzip, deflate")
            Call .Add("Cookie", "(Automatic)")
         End With
      Else
         Set oRequest = Nothing
         Test.Trace("Invalid Request object")
      End If
      Set MakeIE60GETRequest = oRequest
   End Function 
