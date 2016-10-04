Option Explicit
Dim fEnableDelays
fEnableDelays = False

Sub SendRequest1()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to server"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest2()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to server"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest3()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (400)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest4()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to server"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest5()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest6()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest7()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest8()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest9()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest10()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (141)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest11()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (200)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/387.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/387.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest12()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/443.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/443.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest13()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/microsoftnetlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/microsoftnetlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest14()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/bizinternet.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/bizinternet.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest15()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (6279)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/account.aspx"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/account.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest16()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/logon.aspx"+"?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/logon.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest17()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (230)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest18()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest19()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest20()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest21()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest22()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest23()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest24()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest25()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (15483)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/logon.aspx"+"?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwxOTkzNzE0NjI0O3Q8O2w8aTwzPjs%2BO2w8"
        oRequest.Body = oRequest.Body + "dDw7bDxpPDc%2BO2k8OT47aTwxMT47PjtsPHQ8cDxwPGw8Vmlz"
        oRequest.Body = oRequest.Body + "aWJsZTs%2BO2w8bzxmPjs%2BPjs%2BOzs%2BO3Q8cDxwPGw8Vm"
        oRequest.Body = oRequest.Body + "lzaWJsZTs%2BO2w8bzxmPjs%2BPjs%2BOzs%2BO3Q8cDxwPGw8"
        oRequest.Body = oRequest.Body + "VmlzaWJsZTs%2BO2w8bzxmPjs%2BPjs%2BOzs%2BOz4%2BOz4%"
        oRequest.Body = oRequest.Body + "2BOz7wH0S4Fer2f%2Bd4b0wz2B24fFtEDg%3D%3D&LogonEmai"
        oRequest.Body = oRequest.Body + "lTextBox=john@mycom.com&LogonPasswordTextBox=john&"
        oRequest.Body = oRequest.Body + "LogonButton=Logon"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/logon.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest26()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (380)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/account.aspx"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/logon.aspx?ReturnUrl=%2fDuwamish7vb%2fsecure%2faccount.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/account.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest27()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (411)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest28()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest29()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest30()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest31()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest32()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest33()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest34()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest35()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (2424)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/account.aspx"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwxMjgwNzQxOTE7Oz67LqmSUvENu8HWw1EMqg"
        oRequest.Body = oRequest.Body + "rsHoHOuA%3D%3D&ModuleSearch%3ASearchDropDownList=T"
        oRequest.Body = oRequest.Body + "itle&ModuleSearch%3ASearchTextBox=&ModuleAccount%3"
        oRequest.Body = oRequest.Body + "AEmailTextBox=john@mycom.com&ModuleAccount%3APassw"
        oRequest.Body = oRequest.Body + "ordTextBox=&ModuleAccount%3AConfirmPasswordTextBox"
        oRequest.Body = oRequest.Body + "=&ModuleAccount%3AAcctNameTextBox=John+Done&Module"
        oRequest.Body = oRequest.Body + "Account%3AAddressTextBox=12+Forest+Road%2C+Forest+"
        oRequest.Body = oRequest.Body + "Town%2C+Woodville&ModuleAccount%3ACountryTextBox=U"
        oRequest.Body = oRequest.Body + "SA&ModuleAccount%3APhoneTextBox=%28425%29+433+3344"
        oRequest.Body = oRequest.Body + "&ModuleAccount%3AFaxTextBox=%28425%29+344+5678&Sub"
        oRequest.Body = oRequest.Body + "mitButton=Update+Account"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/account.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest36()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (390)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest37()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest38()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest39()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest40()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest41()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest42()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest43()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest44()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (7071)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/account.aspx"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwxMjgwNzQxOTE7Oz67LqmSUvENu8HWw1EMqg"
        oRequest.Body = oRequest.Body + "rsHoHOuA%3D%3D&ModuleSearch%3ASearchDropDownList=T"
        oRequest.Body = oRequest.Body + "itle&ModuleSearch%3ASearchTextBox=&ModuleAccount%3"
        oRequest.Body = oRequest.Body + "AEmailTextBox=john@mycom.com&ModuleAccount%3APassw"
        oRequest.Body = oRequest.Body + "ordTextBox=john&ModuleAccount%3AConfirmPasswordTex"
        oRequest.Body = oRequest.Body + "tBox=john&ModuleAccount%3AAcctNameTextBox=John+Don"
        oRequest.Body = oRequest.Body + "e&ModuleAccount%3AAddressTextBox=12+Forest+Road%2C"
        oRequest.Body = oRequest.Body + "+Forest+Town%2C+Woodville&ModuleAccount%3ACountryT"
        oRequest.Body = oRequest.Body + "extBox=USA&ModuleAccount%3APhoneTextBox=%28425%29+"
        oRequest.Body = oRequest.Body + "433+3344&ModuleAccount%3AFaxTextBox=%28425%29+344+"
        oRequest.Body = oRequest.Body + "5678&SubmitButton=Update+Account"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/account.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest45()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (580)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest46()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest47()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest48()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest49()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (31)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest50()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest51()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest52()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/account.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest53()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (1592)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest54()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest55()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest56()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest57()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest58()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest59()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest60()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest61()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/IMAGE_NA.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/IMAGE_NA.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest62()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/413.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/413.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest63()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/413.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/413.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest64()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/IMAGE_NA.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/IMAGE_NA.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest65()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/441.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/441.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest66()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (942)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/categories.aspx"+"?id=830"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwtNTU5MTk4NTkwO3Q8O2w8aTwxPjs%2BO2w8"
        oRequest.Body = oRequest.Body + "dDw7bDxpPDk%2BOz47bDx0PDtsPGk8MT47PjtsPHQ8O2w8aTwx"
        oRequest.Body = oRequest.Body + "Pjs%2BO2w8dDxAMDxwPHA8bDxfIUl0ZW1Db3VudDtEYXRhS2V5"
        oRequest.Body = oRequest.Body + "czs%2BO2w8aTwyPjtsPD47Pj47Pjs7Ozs7Ozs7PjtsPGk8MD47"
        oRequest.Body = oRequest.Body + "aTwxPjs%2BO2w8dDw7bDxpPDk%2BOz47bDx0PHA8cDxsPENvbW"
        oRequest.Body = oRequest.Body + "1hbmRBcmd1bWVudDs%2BO2w8NDEzfENhbm5pYmFscyBhbmQgS2"
        oRequest.Body = oRequest.Body + "luZ3N8NC45OTs%2BPjs%2BOzs%2BOz4%2BO3Q8O2w8aTw5Pjs%"
        oRequest.Body = oRequest.Body + "2BO2w8dDxwPHA8bDxDb21tYW5kQXJndW1lbnQ7PjtsPDQ0MnxC"
        oRequest.Body = oRequest.Body + "ZXlvbmQgRnJlZWRvbSBhbmQgRGlnbml0eXw3Ljk5Oz4%2BOz47"
        oRequest.Body = oRequest.Body + "Oz47Pj47Pj47Pj47Pj47Pj47Pj47Pg7OBibrPL9bKqJvsM3epa"
        oRequest.Body = oRequest.Body + "Q0gnBd&ModuleSearch%3ASearchDropDownList=Title&Mod"
        oRequest.Body = oRequest.Body + "uleSearch%3ASearchTextBox=&ModuleDailyPick%3ADaily"
        oRequest.Body = oRequest.Body + "PickList%3A_ctl0%3AAddToCartButton=Add+To+Cart"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/categories.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest67()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (561)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/shoppingcart.aspx"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=830"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/shoppingcart.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest68()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (420)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest69()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest70()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest71()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest72()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest73()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest74()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest75()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest76()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (731)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/categories.aspx"+"?id=832"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/categories.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest77()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (1392)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=832"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest78()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (11)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=832"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest79()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=832"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest80()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/categories.aspx"+"?id=833"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/categories.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest81()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (530)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest82()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest83()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest84()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest85()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest86()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (21)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest87()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest88()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest89()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (200)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/490.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/490.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest90()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (220)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/506.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/506.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest91()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (201)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/446.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/446.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest92()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/490.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/490.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest93()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/506.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/506.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest94()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (1041)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/categories.aspx"+"?id=833"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwtNTU5MTk4NTkwO3Q8O2w8aTwxPjs%2BO2w8"
        oRequest.Body = oRequest.Body + "dDw7bDxpPDk%2BOz47bDx0PDtsPGk8MT47PjtsPHQ8O2w8aTwx"
        oRequest.Body = oRequest.Body + "Pjs%2BO2w8dDxAMDxwPHA8bDxfIUl0ZW1Db3VudDtEYXRhS2V5"
        oRequest.Body = oRequest.Body + "czs%2BO2w8aTwyPjtsPD47Pj47Pjs7Ozs7Ozs7PjtsPGk8MD47"
        oRequest.Body = oRequest.Body + "aTwxPjs%2BO2w8dDw7bDxpPDk%2BOz47bDx0PHA8cDxsPENvbW"
        oRequest.Body = oRequest.Body + "1hbmRBcmd1bWVudDs%2BO2w8NDkwfFRoZSBTZXZlbiBIYWJpdH"
        oRequest.Body = oRequest.Body + "Mgb2YgSGlnaGx5IEVmZmVjdGl2ZSBQZW9wbGV8MTYuOTk7Pj47"
        oRequest.Body = oRequest.Body + "Pjs7Pjs%2BPjt0PDtsPGk8OT47PjtsPHQ8cDxwPGw8Q29tbWFu"
        oRequest.Body = oRequest.Body + "ZEFyZ3VtZW50Oz47bDw1MDZ8VGhlIE1pY3Jvc29mdCBXYXl8NC"
        oRequest.Body = oRequest.Body + "45OTs%2BPjs%2BOzs%2BOz4%2BOz4%2BOz4%2BOz4%2BOz4%2B"
        oRequest.Body = oRequest.Body + "Oz4%2BOz4ysCl2ARqJJSoNeZgzHLO0wiuDPQ%3D%3D&ModuleS"
        oRequest.Body = oRequest.Body + "earch%3ASearchDropDownList=Title&ModuleSearch%3ASe"
        oRequest.Body = oRequest.Body + "archTextBox=&ModuleDailyPick%3ADailyPickList%3A_ct"
        oRequest.Body = oRequest.Body + "l1%3AAddToCartButton=Add+To+Cart"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/categories.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest95()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (771)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/shoppingcart.aspx"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=833"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/shoppingcart.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest96()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (421)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest97()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest98()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest99()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest100()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest101()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest102()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest103()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest104()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (1783)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/categories.aspx"+"?id=837"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/categories.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest105()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (7551)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest106()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest107()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest108()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest109()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest110()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest111()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest112()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest113()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (170)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/39.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/39.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest114()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (401)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/56.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/56.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest115()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/IMAGE_NA.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/IMAGE_NA.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest116()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/17.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/17.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest117()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/19.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/19.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest118()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/20.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/20.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest119()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/21.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/21.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest120()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/24.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/24.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest121()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/27.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/27.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest122()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/29.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/29.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest123()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/31.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/31.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest124()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (90)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/33.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/33.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest125()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/34.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/34.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest126()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/35.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/35.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest127()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/36.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/36.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest128()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (101)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/37.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/37.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest129()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/39.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/39.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest130()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/40.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/40.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest131()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/42.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/42.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest132()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/45.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/45.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest133()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/41.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/41.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest134()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (130)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/46.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/46.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest135()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/48.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/48.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest136()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/49.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/49.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest137()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/50.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/50.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest138()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/52.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/52.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest139()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/53.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/53.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest140()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (131)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/55.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/55.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest141()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/56.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/56.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest142()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (160)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/61.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/61.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest143()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/63.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/63.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest144()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (150)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/64.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/64.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest145()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/68.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/68.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest146()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (170)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/72.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/72.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest147()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (51)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/71.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/71.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest148()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (150)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/74.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/74.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest149()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/75.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/75.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest150()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/77.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/77.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest151()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (130)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/79.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/79.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest152()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/81.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/81.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest153()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/82.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/82.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest154()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/83.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/83.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest155()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (131)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/87.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/87.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest156()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/88.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/88.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest157()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/90.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/90.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest158()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/91.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/91.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest159()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/94.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/94.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest160()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/95.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/95.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest161()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/97.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/97.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest162()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/98.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/98.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest163()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/99.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/99.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest164()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/100.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/100.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest165()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (111)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/104.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/104.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest166()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/105.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/105.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest167()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/106.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/106.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest168()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/107.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/107.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest169()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/108.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/108.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest170()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/109.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/109.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest171()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/111.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/111.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest172()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/110.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/110.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest173()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/115.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/115.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest174()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/112.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/112.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest175()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/116.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/116.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest176()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/119.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/119.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest177()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (51)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/122.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/122.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest178()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/120.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/120.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest179()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/125.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/125.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest180()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (200)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/127.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/127.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest181()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/129.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/129.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest182()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/131.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/131.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest183()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/132.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/132.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest184()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/136.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/136.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest185()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/137.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/137.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest186()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (401)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/138.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/138.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest187()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/140.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/140.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest188()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/141.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/141.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest189()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/139.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/139.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest190()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/142.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/142.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest191()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/143.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/143.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest192()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/145.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/145.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest193()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/146.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/146.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest194()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (51)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/148.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/148.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest195()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/150.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/150.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest196()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/151.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/151.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest197()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/152.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/152.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest198()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/153.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/153.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest199()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/154.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/154.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest200()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/155.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/155.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest201()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/158.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/158.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest202()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/160.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/160.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest203()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/156.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/156.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest204()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/162.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/162.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest205()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/163.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/163.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest206()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (61)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/166.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/166.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest207()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/168.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/168.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest208()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/164.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/164.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest209()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/170.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/170.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest210()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/171.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/171.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest211()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/172.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/172.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest212()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/174.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/174.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest213()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/175.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/175.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest214()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (341)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/178.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/178.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest215()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/179.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/179.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest216()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/183.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/183.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest217()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/184.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/184.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest218()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/181.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/181.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest219()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/185.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/185.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest220()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/187.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/187.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest221()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/189.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/189.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest222()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/190.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/190.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest223()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/336.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/336.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest224()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (51)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/485.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/485.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest225()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/498.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/498.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest226()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/495.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/495.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest227()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/499.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/499.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest228()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/510.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/510.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest229()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/525.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/525.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest230()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/526.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/526.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest231()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/530.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/530.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest232()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/531.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/531.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest233()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/532.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/532.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest234()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (3615)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/categories.aspx"+"?id=837"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwtNTU5MTk4NTkwO3Q8O2w8aTwxPjs%2BO2w8"
        oRequest.Body = oRequest.Body + "dDw7bDxpPDk%2BOz47bDx0PDtsPGk8MT47PjtsPHQ8O2w8aTwx"
        oRequest.Body = oRequest.Body + "Pjs%2BO2w8dDxAMDxwPHA8bDxfIUl0ZW1Db3VudDtEYXRhS2V5"
        oRequest.Body = oRequest.Body + "czs%2BO2w8aTwyPjtsPD47Pj47Pjs7Ozs7Ozs7PjtsPGk8MD47"
        oRequest.Body = oRequest.Body + "aTwxPjs%2BO2w8dDw7bDxpPDk%2BOz47bDx0PHA8cDxsPENvbW"
        oRequest.Body = oRequest.Body + "1hbmRBcmd1bWVudDs%2BO2w8Mzl8VGhlIFdvcmxkIEFjY29yZG"
        oRequest.Body = oRequest.Body + "luZyB0byBHYXJwfDYuOTk7Pj47Pjs7Pjs%2BPjt0PDtsPGk8OT"
        oRequest.Body = oRequest.Body + "47PjtsPHQ8cDxwPGw8Q29tbWFuZEFyZ3VtZW50Oz47bDw1NnxT"
        oRequest.Body = oRequest.Body + "dXJmYWNpbmd8OC45OTs%2BPjs%2BOzs%2BOz4%2BOz4%2BOz4%"
        oRequest.Body = oRequest.Body + "2BOz4%2BOz4%2BOz4%2BOz5CL70fTQv2r9RGlJtc8cmwZL%2Fa"
        oRequest.Body = oRequest.Body + "QQ%3D%3D&ModuleSearch%3ASearchDropDownList=Title&M"
        oRequest.Body = oRequest.Body + "oduleSearch%3ASearchTextBox=&ModuleDailyPick%3ADai"
        oRequest.Body = oRequest.Body + "lyPickList%3A_ctl0%3AAddToCartButton=Add+To+Cart"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/categories.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest235()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (591)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/shoppingcart.aspx"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/shoppingcart.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest236()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (481)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest237()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (100)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest238()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest239()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest240()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest241()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (71)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest242()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest243()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest244()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (7420)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/shoppingcart.aspx"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwxMzA0ODU2MjI5O3Q8O2w8aTwxPjs%2BO2w8"
        oRequest.Body = oRequest.Body + "dDw7bDxpPDk%2BO2k8MTE%2BO2k8MTM%2BOz47bDx0PHA8cDxs"
        oRequest.Body = oRequest.Body + "PE5hdmlnYXRlVXJsOz47bDxodHRwOi8vbG9jYWxob3N0L0R1d2"
        oRequest.Body = oRequest.Body + "FtaXNoN3ZiL3NlY3VyZS9jaGVja291dC5hc3B4Oz4%2BOz47Oz"
        oRequest.Body = oRequest.Body + "47dDw7bDxpPDE%2BOz47bDx0PEAwPHA8cDxsPFBhZ2VDb3VudD"
        oRequest.Body = oRequest.Body + "tfIUl0ZW1Db3VudDtfIURhdGFTb3VyY2VJdGVtQ291bnQ7RGF0"
        oRequest.Body = oRequest.Body + "YUtleXM7PjtsPGk8MT47aTwzPjtpPDM%2BO2w8Pjs%2BPjs%2B"
        oRequest.Body = oRequest.Body + "Ozs7Ozs7Ozs7Oz47bDxpPDA%2BOz47bDx0PDtsPGk8MT47aTwy"
        oRequest.Body = oRequest.Body + "PjtpPDM%2BOz47bDx0PDtsPGk8MD47aTwxPjtpPDI%2BO2k8Mz"
        oRequest.Body = oRequest.Body + "47PjtsPHQ8O2w8aTwzPjtpPDU%2BOz47bDx0PHA8cDxsPEVycm"
        oRequest.Body = oRequest.Body + "9yTWVzc2FnZTs%2BO2w8Q2FubmliYWxzIGFuZCBLaW5nczs%2B"
        oRequest.Body = oRequest.Body + "Pjs%2BOzs%2BO3Q8cDxwPGw8RXJyb3JNZXNzYWdlOz47bDxDYW"
        oRequest.Body = oRequest.Body + "5uaWJhbHMgYW5kIEtpbmdzOz4%2BOz47Oz47Pj47dDxwPHA8bD"
        oRequest.Body = oRequest.Body + "xUZXh0Oz47bDxDYW5uaWJhbHMgYW5kIEtpbmdzOz4%2BOz47Oz"
        oRequest.Body = oRequest.Body + "47dDxwPHA8bDxUZXh0Oz47bDwkNC45OTs%2BPjs%2BOzs%2BO3"
        oRequest.Body = oRequest.Body + "Q8cDxwPGw8VGV4dDs%2BO2w8JDQuOTk7Pj47Pjs7Pjs%2BPjt0"
        oRequest.Body = oRequest.Body + "PDtsPGk8MD47aTwxPjtpPDI%2BO2k8Mz47PjtsPHQ8O2w8aTwz"
        oRequest.Body = oRequest.Body + "PjtpPDU%2BOz47bDx0PHA8cDxsPEVycm9yTWVzc2FnZTs%2BO2"
        oRequest.Body = oRequest.Body + "w8VGhlIE1pY3Jvc29mdCBXYXk7Pj47Pjs7Pjt0PHA8cDxsPEVy"
        oRequest.Body = oRequest.Body + "cm9yTWVzc2FnZTs%2BO2w8VGhlIE1pY3Jvc29mdCBXYXk7Pj47"
        oRequest.Body = oRequest.Body + "Pjs7Pjs%2BPjt0PHA8cDxsPFRleHQ7PjtsPFRoZSBNaWNyb3Nv"
        oRequest.Body = oRequest.Body + "ZnQgV2F5Oz4%2BOz47Oz47dDxwPHA8bDxUZXh0Oz47bDwkNC45"
        oRequest.Body = oRequest.Body + "OTs%2BPjs%2BOzs%2BO3Q8cDxwPGw8VGV4dDs%2BO2w8JDQuOT"
        oRequest.Body = oRequest.Body + "k7Pj47Pjs7Pjs%2BPjt0PDtsPGk8MD47aTwxPjtpPDI%2BO2k8"
        oRequest.Body = oRequest.Body + "Mz47PjtsPHQ8O2w8aTwzPjtpPDU%2BOz47bDx0PHA8cDxsPEVy"
        oRequest.Body = oRequest.Body + "cm9yTWVzc2FnZTs%2BO2w8VGhlIFdvcmxkIEFjY29yZGluZyB0"
        oRequest.Body = oRequest.Body + "byBHYXJwOz4%2BOz47Oz47dDxwPHA8bDxFcnJvck1lc3NhZ2U7"
        oRequest.Body = oRequest.Body + "PjtsPFRoZSBXb3JsZCBBY2NvcmRpbmcgdG8gR2FycDs%2BPjs%"
        oRequest.Body = oRequest.Body + "2BOzs%2BOz4%2BO3Q8cDxwPGw8VGV4dDs%2BO2w8VGhlIFdvcm"
        oRequest.Body = oRequest.Body + "xkIEFjY29yZGluZyB0byBHYXJwOz4%2BOz47Oz47dDxwPHA8bD"
        oRequest.Body = oRequest.Body + "xUZXh0Oz47bDwkNi45OTs%2BPjs%2BOzs%2BO3Q8cDxwPGw8VG"
        oRequest.Body = oRequest.Body + "V4dDs%2BO2w8JDYuOTk7Pj47Pjs7Pjs%2BPjs%2BPjs%2BPjs%"
        oRequest.Body = oRequest.Body + "2BPjt0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47Pjs7Pjs%"
        oRequest.Body = oRequest.Body + "2BPjs%2BPjs%2BR4aIe4gx7nDLwtT7ezpERl1vb74%3D&Modul"
        oRequest.Body = oRequest.Body + "eSearch%3ASearchDropDownList=Title&ModuleSearch%3A"
        oRequest.Body = oRequest.Body + "SearchTextBox=&CartItemsDataGrid%3A_ctl2%3AQuantit"
        oRequest.Body = oRequest.Body + "yTextBox=1&CartItemsDataGrid%3A_ctl3%3AQuantityTex"
        oRequest.Body = oRequest.Body + "tBox=2&CartItemsDataGrid%3A_ctl4%3AQuantityTextBox"
        oRequest.Body = oRequest.Body + "=1&UpdateButton=Update"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/shoppingcart.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest245()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (601)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest246()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest247()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest248()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest249()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (61)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest250()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest251()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest252()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest253()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (3615)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/shoppingcart.aspx"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwxMzA0ODU2MjI5O3Q8O2w8aTwxPjs%2BO2w8"
        oRequest.Body = oRequest.Body + "dDw7bDxpPDk%2BO2k8MTE%2BO2k8MTM%2BOz47bDx0PHA8cDxs"
        oRequest.Body = oRequest.Body + "PFRleHQ7TmF2aWdhdGVVcmw7PjtsPFByb2NlZWQgdG8gQ2hlY2"
        oRequest.Body = oRequest.Body + "tvdXQNCiAgICAgICAgICAgICAgICAgICAgICAgIDtodHRwOi8v"
        oRequest.Body = oRequest.Body + "bG9jYWxob3N0L0R1d2FtaXNoN3ZiL3NlY3VyZS9jaGVja291dC"
        oRequest.Body = oRequest.Body + "5hc3B4Oz4%2BOz47Oz47dDw7bDxpPDE%2BOz47bDx0PEAwPHA8"
        oRequest.Body = oRequest.Body + "cDxsPFBhZ2VDb3VudDtfIUl0ZW1Db3VudDtfIURhdGFTb3VyY2"
        oRequest.Body = oRequest.Body + "VJdGVtQ291bnQ7RGF0YUtleXM7PjtsPGk8MT47aTwzPjtpPDM%"
        oRequest.Body = oRequest.Body + "2BO2w8Pjs%2BPjs%2BOzs7Ozs7Ozs7Oz47bDxpPDA%2BOz47bD"
        oRequest.Body = oRequest.Body + "x0PDtsPGk8MT47aTwyPjtpPDM%2BOz47bDx0PDtsPGk8MD47aT"
        oRequest.Body = oRequest.Body + "wxPjtpPDI%2BO2k8Mz47PjtsPHQ8O2w8aTwzPjtpPDU%2BOz47"
        oRequest.Body = oRequest.Body + "bDx0PHA8cDxsPEVycm9yTWVzc2FnZTs%2BO2w8Q2FubmliYWxz"
        oRequest.Body = oRequest.Body + "IGFuZCBLaW5nczs%2BPjs%2BOzs%2BO3Q8cDxwPGw8RXJyb3JN"
        oRequest.Body = oRequest.Body + "ZXNzYWdlOz47bDxDYW5uaWJhbHMgYW5kIEtpbmdzOz4%2BOz47"
        oRequest.Body = oRequest.Body + "Oz47Pj47dDxwPHA8bDxUZXh0Oz47bDxDYW5uaWJhbHMgYW5kIE"
        oRequest.Body = oRequest.Body + "tpbmdzOz4%2BOz47Oz47dDxwPHA8bDxUZXh0Oz47bDwkNC45OT"
        oRequest.Body = oRequest.Body + "s%2BPjs%2BOzs%2BO3Q8cDxwPGw8VGV4dDs%2BO2w8JDQuOTk7"
        oRequest.Body = oRequest.Body + "Pj47Pjs7Pjs%2BPjt0PDtsPGk8MD47aTwxPjtpPDI%2BO2k8Mz"
        oRequest.Body = oRequest.Body + "47PjtsPHQ8O2w8aTwzPjtpPDU%2BOz47bDx0PHA8cDxsPEVycm"
        oRequest.Body = oRequest.Body + "9yTWVzc2FnZTs%2BO2w8VGhlIE1pY3Jvc29mdCBXYXk7Pj47Pj"
        oRequest.Body = oRequest.Body + "s7Pjt0PHA8cDxsPEVycm9yTWVzc2FnZTs%2BO2w8VGhlIE1pY3"
        oRequest.Body = oRequest.Body + "Jvc29mdCBXYXk7Pj47Pjs7Pjs%2BPjt0PHA8cDxsPFRleHQ7Pj"
        oRequest.Body = oRequest.Body + "tsPFRoZSBNaWNyb3NvZnQgV2F5Oz4%2BOz47Oz47dDxwPHA8bD"
        oRequest.Body = oRequest.Body + "xUZXh0Oz47bDwkNC45OTs%2BPjs%2BOzs%2BO3Q8cDxwPGw8VG"
        oRequest.Body = oRequest.Body + "V4dDs%2BO2w8JDkuOTg7Pj47Pjs7Pjs%2BPjt0PDtsPGk8MD47"
        oRequest.Body = oRequest.Body + "aTwxPjtpPDI%2BO2k8Mz47PjtsPHQ8O2w8aTwzPjtpPDU%2BOz"
        oRequest.Body = oRequest.Body + "47bDx0PHA8cDxsPEVycm9yTWVzc2FnZTs%2BO2w8VGhlIFdvcm"
        oRequest.Body = oRequest.Body + "xkIEFjY29yZGluZyB0byBHYXJwOz4%2BOz47Oz47dDxwPHA8bD"
        oRequest.Body = oRequest.Body + "xFcnJvck1lc3NhZ2U7PjtsPFRoZSBXb3JsZCBBY2NvcmRpbmcg"
        oRequest.Body = oRequest.Body + "dG8gR2FycDs%2BPjs%2BOzs%2BOz4%2BO3Q8cDxwPGw8VGV4dD"
        oRequest.Body = oRequest.Body + "s%2BO2w8VGhlIFdvcmxkIEFjY29yZGluZyB0byBHYXJwOz4%2B"
        oRequest.Body = oRequest.Body + "Oz47Oz47dDxwPHA8bDxUZXh0Oz47bDwkNi45OTs%2BPjs%2BOz"
        oRequest.Body = oRequest.Body + "s%2BO3Q8cDxwPGw8VGV4dDs%2BO2w8JDYuOTk7Pj47Pjs7Pjs%"
        oRequest.Body = oRequest.Body + "2BPjs%2BPjs%2BPjs%2BPjt0PHA8cDxsPFRleHQ7VmlzaWJsZT"
        oRequest.Body = oRequest.Body + "s%2BO2w8WW91ciBzaG9wcGluZyBjYXJ0IGlzIGVtcHR5IC0gcG"
        oRequest.Body = oRequest.Body + "xlYXNlIGNhcnJ5IG9uIHNob3BwaW5nLg0KICAgICAgICAgICAg"
        oRequest.Body = oRequest.Body + "ICAgICAgICAgICAgICA7bzxmPjs%2BPjs%2BOzs%2BOz4%2BOz"
        oRequest.Body = oRequest.Body + "4%2BOz76wezs2hTRTIYpBI%2FjNZrsu0juiw%3D%3D&ModuleS"
        oRequest.Body = oRequest.Body + "earch%3ASearchDropDownList=Title&ModuleSearch%3ASe"
        oRequest.Body = oRequest.Body + "archTextBox=&CartItemsDataGrid%3A_ctl2%3AQuantityT"
        oRequest.Body = oRequest.Body + "extBox=1&CartItemsDataGrid%3A_ctl3%3AQuantityTextB"
        oRequest.Body = oRequest.Body + "ox=10&CartItemsDataGrid%3A_ctl4%3AQuantityTextBox="
        oRequest.Body = oRequest.Body + "1&UpdateButton=Update"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/shoppingcart.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest254()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (451)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest255()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest256()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest257()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest258()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest259()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest260()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest261()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest262()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (4407)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/checkout.aspx"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/shoppingcart.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/checkout.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest263()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (450)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest264()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest265()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (61)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest266()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest267()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest268()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest269()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/arrow.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/arrow.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest270()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest271()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest272()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (150)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/next.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/next.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest273()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/previousdisabled.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/previousdisabled.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest274()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (2494)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/checkout.aspx"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwtMTMzODAwMzQ5Nzt0PHA8bDxzdGFnZTs%2B"
        oRequest.Body = oRequest.Body + "O2w8MDs%2BPjtsPGk8MT47PjtsPHQ8O2w8aTw5PjtpPDExPjs%"
        oRequest.Body = oRequest.Body + "2BO2w8dDxwPHA8bDxWaXNpYmxlOz47bDxvPGY%2BOz4%2BOz47"
        oRequest.Body = oRequest.Body + "bDxpPDc%2BO2k8MTE%2BO2k8MTM%2BO2k8MTU%2BO2k8MTc%2B"
        oRequest.Body = oRequest.Body + "O2k8MTk%2BO2k8MjM%2BOz47bDx0PHA8cDxsPFZpc2libGU7Pj"
        oRequest.Body = oRequest.Body + "tsPG88Zj47Pj47Pjs7Pjt0PHA8cDxsPFZpc2libGU7PjtsPG88"
        oRequest.Body = oRequest.Body + "Zj47Pj47Pjs7Pjt0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj"
        oRequest.Body = oRequest.Body + "47Pjs7Pjt0PHA8cDxsPE1vZGU7PjtsPFN5c3RlbS5XZWIuVUku"
        oRequest.Body = oRequest.Body + "V2ViQ29udHJvbHMuVGV4dEJveE1vZGUsIFN5c3RlbS5XZWIsIF"
        oRequest.Body = oRequest.Body + "ZlcnNpb249MS4wLjMzMDAuMCwgQ3VsdHVyZT1uZXV0cmFsLCBQ"
        oRequest.Body = oRequest.Body + "dWJsaWNLZXlUb2tlbj1iMDNmNWY3ZjExZDUwYTNhPE11bHRpTG"
        oRequest.Body = oRequest.Body + "luZT47Pj47Pjs7Pjt0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47"
        oRequest.Body = oRequest.Body + "Pj47Pjs7Pjt0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47Pj"
        oRequest.Body = oRequest.Body + "s7Pjt0PHQ8O3A8bDxpPDA%2BO2k8MT47aTwyPjtpPDM%2BO2k8"
        oRequest.Body = oRequest.Body + "ND47aTw1PjtpPDY%2BO2k8Nz47aTw4PjtpPDk%2BOz47bDxwPD"
        oRequest.Body = oRequest.Body + "IwMDI7MjAwMj47cDwyMDAzOzIwMDM%2BO3A8MjAwNDsyMDA0Pj"
        oRequest.Body = oRequest.Body + "twPDIwMDU7MjAwNT47cDwyMDA2OzIwMDY%2BO3A8MjAwNzsyMD"
        oRequest.Body = oRequest.Body + "A3PjtwPDIwMDg7MjAwOD47cDwyMDA5OzIwMDk%2BO3A8MjAxMD"
        oRequest.Body = oRequest.Body + "syMDEwPjtwPDIwMTE7MjAxMT47Pj47Pjs7Pjs%2BPjt0PHA8cD"
        oRequest.Body = oRequest.Body + "xsPFZpc2libGU7PjtsPG88Zj47Pj47Pjs7Pjs%2BPjs%2BPjts"
        oRequest.Body = oRequest.Body + "PFByZXZpb3VzSW1hZ2VCdXR0b247TmV4dEltYWdlQnV0dG9uOz"
        oRequest.Body = oRequest.Body + "4%2Bga7ZGMlpDRoRW14aPeS8tOPRsYo%3D&ShipToNameTextB"
        oRequest.Body = oRequest.Body + "ox=John+Done&AddressTextBox=12+Forest+Road%2C+Fore"
        oRequest.Body = oRequest.Body + "st+Town%2C+Woodville&CountryTextBox=USA&PhoneNumbe"
        oRequest.Body = oRequest.Body + "rTextBox=%28425%29+433+3344&FaxTextBox=%28425%29+3"
        oRequest.Body = oRequest.Body + "44+5678&NextImageButton.x=31&NextImageButton.y=22"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/checkout.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest275()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (431)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest276()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest277()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest278()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest279()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (80)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest280()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest281()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (80)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest282()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest283()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/arrow.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/arrow.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest284()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/next.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/next.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest285()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (71)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/previous.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/previous.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest286()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (31365)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/checkout.aspx"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwtMTMzODAwMzQ5Nzt0PHA8bDxzdGFnZTs%2B"
        oRequest.Body = oRequest.Body + "O2w8MTs%2BPjtsPGk8MT47PjtsPHQ8O2w8aTw3PjtpPDk%2BO2"
        oRequest.Body = oRequest.Body + "k8MTE%2BOz47bDx0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj"
        oRequest.Body = oRequest.Body + "47PjtsPGk8MT47aTwzPjtpPDU%2BO2k8Nz47aTw5PjtpPDExPj"
        oRequest.Body = oRequest.Body + "tpPDEzPjtpPDE1PjtpPDE3PjtpPDE5PjtpPDIxPjs%2BO2w8dD"
        oRequest.Body = oRequest.Body + "xwPHA8bDxUZXh0Oz47bDxKb2huIERvbmU7Pj47Pjs7Pjt0PHA8"
        oRequest.Body = oRequest.Body + "cDxsPFZpc2libGU7PjtsPG88Zj47Pj47Pjs7Pjt0PHA8cDxsPF"
        oRequest.Body = oRequest.Body + "RleHQ7PjtsPDEyIEZvcmVzdCBSb2FkLCBGb3Jlc3QgVG93biwg"
        oRequest.Body = oRequest.Body + "V29vZHZpbGxlOz4%2BOz47Oz47dDxwPHA8bDxWaXNpYmxlOz47"
        oRequest.Body = oRequest.Body + "bDxvPGY%2BOz4%2BOz47Oz47dDxwPHA8bDxUZXh0Oz47bDxVU0"
        oRequest.Body = oRequest.Body + "E7Pj47Pjs7Pjt0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj47"
        oRequest.Body = oRequest.Body + "Pjs7Pjt0PHA8cDxsPFRleHQ7PjtsPCg0MjUpIDQzMyAzMzQ0Oz"
        oRequest.Body = oRequest.Body + "4%2BOz47Oz47dDxwPHA8bDxWaXNpYmxlOz47bDxvPGY%2BOz4%"
        oRequest.Body = oRequest.Body + "2BOz47Oz47dDxwPHA8bDxWaXNpYmxlOz47bDxvPGY%2BOz4%2B"
        oRequest.Body = oRequest.Body + "Oz47Oz47dDxwPHA8bDxUZXh0Oz47bDwoNDI1KSAzNDQgNTY3OD"
        oRequest.Body = oRequest.Body + "s%2BPjs%2BOzs%2BO3Q8cDxwPGw8VmlzaWJsZTs%2BO2w8bzxm"
        oRequest.Body = oRequest.Body + "Pjs%2BPjs%2BOzs%2BOz4%2BO3Q8cDxwPGw8VmlzaWJsZTs%2B"
        oRequest.Body = oRequest.Body + "O2w8bzx0Pjs%2BPjs%2BO2w8aTw3PjtpPDExPjtpPDEzPjtpPD"
        oRequest.Body = oRequest.Body + "E1PjtpPDE3PjtpPDE5PjtpPDIzPjs%2BO2w8dDxwPHA8bDxUZX"
        oRequest.Body = oRequest.Body + "h0O1Zpc2libGU7PjtsPCo7bzx0Pjs%2BPjs%2BOzs%2BO3Q8cD"
        oRequest.Body = oRequest.Body + "xwPGw8VGV4dDtWaXNpYmxlOz47bDwqO288dD47Pj47Pjs7Pjt0"
        oRequest.Body = oRequest.Body + "PHA8cDxsPFRleHQ7VmlzaWJsZTs%2BO2w8KjtvPHQ%2BOz4%2B"
        oRequest.Body = oRequest.Body + "Oz47Oz47dDxwPHA8bDxNb2RlOz47bDxTeXN0ZW0uV2ViLlVJLl"
        oRequest.Body = oRequest.Body + "dlYkNvbnRyb2xzLlRleHRCb3hNb2RlLCBTeXN0ZW0uV2ViLCBW"
        oRequest.Body = oRequest.Body + "ZXJzaW9uPTEuMC4zMzAwLjAsIEN1bHR1cmU9bmV1dHJhbCwgUH"
        oRequest.Body = oRequest.Body + "VibGljS2V5VG9rZW49YjAzZjVmN2YxMWQ1MGEzYTxNdWx0aUxp"
        oRequest.Body = oRequest.Body + "bmU%2BOz4%2BOz47Oz47dDxwPHA8bDxUZXh0O1Zpc2libGU7Pj"
        oRequest.Body = oRequest.Body + "tsPCo7bzx0Pjs%2BPjs%2BOzs%2BO3Q8cDxwPGw8VGV4dDtWaX"
        oRequest.Body = oRequest.Body + "NpYmxlOz47bDwqO288dD47Pj47Pjs7Pjt0PHQ8O3A8bDxpPDA%"
        oRequest.Body = oRequest.Body + "2BO2k8MT47aTwyPjtpPDM%2BO2k8ND47aTw1PjtpPDY%2BO2k8"
        oRequest.Body = oRequest.Body + "Nz47aTw4PjtpPDk%2BOz47bDxwPDIwMDI7MjAwMj47cDwyMDAz"
        oRequest.Body = oRequest.Body + "OzIwMDM%2BO3A8MjAwNDsyMDA0PjtwPDIwMDU7MjAwNT47cDwy"
        oRequest.Body = oRequest.Body + "MDA2OzIwMDY%2BO3A8MjAwNzsyMDA3PjtwPDIwMDg7MjAwOD47"
        oRequest.Body = oRequest.Body + "cDwyMDA5OzIwMDk%2BO3A8MjAxMDsyMDEwPjtwPDIwMTE7MjAx"
        oRequest.Body = oRequest.Body + "MT47Pj47Pjs7Pjs%2BPjt0PHA8cDxsPFZpc2libGU7PjtsPG88"
        oRequest.Body = oRequest.Body + "Zj47Pj47Pjs7Pjs%2BPjs%2BPjtsPFByZXZpb3VzSW1hZ2VCdX"
        oRequest.Body = oRequest.Body + "R0b247TmV4dEltYWdlQnV0dG9uOz4%2BB0pwZ3iXsoWKsZ6Fc7"
        oRequest.Body = oRequest.Body + "XAHK3h6ic%3D&CartTypeDropDownList=Card+Type+1&Name"
        oRequest.Body = oRequest.Body + "OnCardTextBox=John+Smith&CardNumberTextBox=4111111"
        oRequest.Body = oRequest.Body + "1111111111&BillingInfoTextBox=1+Microsoft+Way%0D%0"
        oRequest.Body = oRequest.Body + "ARedmond%2C+WA&ExpMonthDropDownList=7&ExpYearDropD"
        oRequest.Body = oRequest.Body + "ownListBox=2006&NextImageButton.x=27&NextImageButt"
        oRequest.Body = oRequest.Body + "on.y=10"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/checkout.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest287()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (761)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest288()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest289()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest290()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest291()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (140)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest292()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest293()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (80)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest294()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (81)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/arrow.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/arrow.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest295()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest296()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (80)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/previous.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/previous.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest297()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (80)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/confirm.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/confirm.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest298()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (2213)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/checkout.aspx"
        oRequest.Verb = "POST"
        oRequest.HTTPVersion = "HTTP/1.0"
        oRequest.EncodeBody = False
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "Content-Type", "application/x-www-form-urlencoded"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        oHeaders.Add "Content-Length", "(automatic)" 
        oRequest.Body = "__VIEWSTATE=dDwtMTMzODAwMzQ5Nzt0PHA8bDxzdGFnZTs%2B"
        oRequest.Body = oRequest.Body + "O2w8Mjs%2BPjtsPGk8MT47PjtsPHQ8O2w8aTw3PjtpPDk%2BO2"
        oRequest.Body = oRequest.Body + "k8MTE%2BOz47bDx0PHA8cDxsPFZpc2libGU7PjtsPG88Zj47Pj"
        oRequest.Body = oRequest.Body + "47PjtsPGk8MT47aTwzPjtpPDU%2BO2k8Nz47aTw5PjtpPDExPj"
        oRequest.Body = oRequest.Body + "tpPDEzPjtpPDE1PjtpPDE3PjtpPDE5PjtpPDIxPjs%2BO2w8dD"
        oRequest.Body = oRequest.Body + "xwPHA8bDxUZXh0Oz47bDxKb2huIERvbmU7Pj47Pjs7Pjt0PHA8"
        oRequest.Body = oRequest.Body + "cDxsPFRleHQ7VmlzaWJsZTs%2BO2w8KjtvPGY%2BOz4%2BOz47"
        oRequest.Body = oRequest.Body + "Oz47dDxwPHA8bDxUZXh0Oz47bDwxMiBGb3Jlc3QgUm9hZCwgRm"
        oRequest.Body = oRequest.Body + "9yZXN0IFRvd24sIFdvb2R2aWxsZTs%2BPjs%2BOzs%2BO3Q8cD"
        oRequest.Body = oRequest.Body + "xwPGw8VGV4dDtWaXNpYmxlOz47bDwqO288Zj47Pj47Pjs7Pjt0"
        oRequest.Body = oRequest.Body + "PHA8cDxsPFRleHQ7PjtsPFVTQTs%2BPjs%2BOzs%2BO3Q8cDxw"
        oRequest.Body = oRequest.Body + "PGw8VGV4dDtWaXNpYmxlOz47bDwqO288Zj47Pj47Pjs7Pjt0PH"
        oRequest.Body = oRequest.Body + "A8cDxsPFRleHQ7PjtsPCg0MjUpIDQzMyAzMzQ0Oz4%2BOz47Oz"
        oRequest.Body = oRequest.Body + "47dDxwPHA8bDxUZXh0O1Zpc2libGU7PjtsPCo7bzxmPjs%2BPj"
        oRequest.Body = oRequest.Body + "s%2BOzs%2BO3Q8cDxwPGw8VGV4dDtWaXNpYmxlOz47bDwqO288"
        oRequest.Body = oRequest.Body + "Zj47Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7PjtsPCg0MjUpIDM0NC"
        oRequest.Body = oRequest.Body + "A1Njc4Oz4%2BOz47Oz47dDxwPHA8bDxUZXh0O1Zpc2libGU7Pj"
        oRequest.Body = oRequest.Body + "tsPCo7bzxmPjs%2BPjs%2BOzs%2BOz4%2BO3Q8cDxwPGw8Vmlz"
        oRequest.Body = oRequest.Body + "aWJsZTs%2BO2w8bzxmPjs%2BPjs%2BO2w8aTwzPjtpPDU%2BO2"
        oRequest.Body = oRequest.Body + "k8Nz47aTw5PjtpPDExPjtpPDEzPjtpPDE1PjtpPDE3PjtpPDE5"
        oRequest.Body = oRequest.Body + "PjtpPDIxPjtpPDIzPjs%2BO2w8dDx0PDs7bDxpPDA%2BOz4%2B"
        oRequest.Body = oRequest.Body + "Ozs%2BO3Q8cDxwPGw8VGV4dDs%2BO2w8Sm9obiBTbWl0aDs%2B"
        oRequest.Body = oRequest.Body + "Pjs%2BOzs%2BO3Q8cDxwPGw8VGV4dDtWaXNpYmxlOz47bDwqO2"
        oRequest.Body = oRequest.Body + "88Zj47Pj47Pjs7Pjt0PHA8cDxsPFRleHQ7PjtsPDQxMTExMTEx"
        oRequest.Body = oRequest.Body + "MTExMTExMTExOz4%2BOz47Oz47dDxwPHA8bDxUZXh0O1Zpc2li"
        oRequest.Body = oRequest.Body + "bGU7PjtsPCo7bzxmPjs%2BPjs%2BOzs%2BO3Q8cDxwPGw8VGV4"
        oRequest.Body = oRequest.Body + "dDtWaXNpYmxlOz47bDwqO288Zj47Pj47Pjs7Pjt0PHA8cDxsPE"
        oRequest.Body = oRequest.Body + "1vZGU7VGV4dDs%2BO2w8U3lzdGVtLldlYi5VSS5XZWJDb250cm"
        oRequest.Body = oRequest.Body + "9scy5UZXh0Qm94TW9kZSwgU3lzdGVtLldlYiwgVmVyc2lvbj0x"
        oRequest.Body = oRequest.Body + "LjAuMzMwMC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleV"
        oRequest.Body = oRequest.Body + "Rva2VuPWIwM2Y1ZjdmMTFkNTBhM2E8TXVsdGlMaW5lPjsxIE1p"
        oRequest.Body = oRequest.Body + "Y3Jvc29mdCBXYXkNClJlZG1vbmQsIFdBOz4%2BOz47Oz47dDxw"
        oRequest.Body = oRequest.Body + "PHA8bDxUZXh0O1Zpc2libGU7PjtsPCo7bzxmPjs%2BPjs%2BOz"
        oRequest.Body = oRequest.Body + "s%2BO3Q8cDxwPGw8VGV4dDtWaXNpYmxlOz47bDwqO288Zj47Pj"
        oRequest.Body = oRequest.Body + "47Pjs7Pjt0PHQ8OztsPGk8Nj47Pj47Oz47dDx0PDtwPGw8aTww"
        oRequest.Body = oRequest.Body + "PjtpPDE%2BO2k8Mj47aTwzPjtpPDQ%2BO2k8NT47aTw2PjtpPD"
        oRequest.Body = oRequest.Body + "c%2BO2k8OD47aTw5Pjs%2BO2w8cDwyMDAyOzIwMDI%2BO3A8Mj"
        oRequest.Body = oRequest.Body + "AwMzsyMDAzPjtwPDIwMDQ7MjAwND47cDwyMDA1OzIwMDU%2BO3"
        oRequest.Body = oRequest.Body + "A8MjAwNjsyMDA2PjtwPDIwMDc7MjAwNz47cDwyMDA4OzIwMDg%"
        oRequest.Body = oRequest.Body + "2BO3A8MjAwOTsyMDA5PjtwPDIwMTA7MjAxMD47cDwyMDExOzIw"
        oRequest.Body = oRequest.Body + "MTE%2BOz4%2BO2w8aTw0Pjs%2BPjs7Pjs%2BPjt0PHA8cDxsPF"
        oRequest.Body = oRequest.Body + "Zpc2libGU7PjtsPG88dD47Pj47Pjs7Pjs%2BPjs%2BPjtsPFBy"
        oRequest.Body = oRequest.Body + "ZXZpb3VzSW1hZ2VCdXR0b247TmV4dEltYWdlQnV0dG9uOz4%2B"
        oRequest.Body = oRequest.Body + "xY94vBhcsBzy8EGQtsye03qs4X4%3D&NextImageButton.x=4"
        oRequest.Body = oRequest.Body + "8&NextImageButton.y=7"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/checkout.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest299()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (1092)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/secure/order.aspx"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/vnd.ms-powerpoint, application/vnd.ms-excel, application/msword, */*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/checkout.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Pragma", "no-cache"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/secure/order.aspx"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest300()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (1402)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/order.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest301()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/order.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest302()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (90)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/order.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest303()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/order.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest304()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/order.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest305()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/order.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest306()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (90)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/order.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest307()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (80)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/secure/order.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest308()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (3976)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/css/duwamish.css"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/css/duwamish.css"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest309()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (150)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest310()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerhome.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerhome.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest311()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (151)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannercart.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannercart.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest312()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/banneraccount.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/banneraccount.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest313()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (90)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/line.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/line.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest314()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/codebehindsource.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/codebehindsource.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest315()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (140)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/387.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/387.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest316()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (90)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/banner/bannerslice.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/banner/bannerslice.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest317()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/microsoftnetlogo.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/microsoftnetlogo.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest318()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/443.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/443.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest319()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (161)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/bizinternet.gif"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/default.aspx"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        'oHeaders.Add "Cookie", "ASP.NET_SessionId=2gvym3y2xle3qx45p1uvxt45; .ADUAUTH=3CDB06CB77DEDAEF8EE5895E16CBB3A895D858C264A3C37A96F82E7AF421117F164CD9CC652D731E300A0ACD2DD846F6C25263C166EA18758882879267A61F27"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/bizinternet.gif"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub




Sub Main()
    call SendRequest1()
    call SendRequest2()
    call SendRequest3()
    call SendRequest4()
    call SendRequest5()
    call SendRequest6()
    call SendRequest7()
    call SendRequest8()
    call SendRequest9()
    call SendRequest10()
    call SendRequest11()
    call SendRequest12()
    call SendRequest13()
    call SendRequest14()
    call SendRequest15()
    call SendRequest16()
    call SendRequest17()
    call SendRequest18()
    call SendRequest19()
    call SendRequest20()
    call SendRequest21()
    call SendRequest22()
    call SendRequest23()
    call SendRequest24()
    call SendRequest25()
    call SendRequest26()
    call SendRequest27()
    call SendRequest28()
    call SendRequest29()
    call SendRequest30()
    call SendRequest31()
    call SendRequest32()
    call SendRequest33()
    call SendRequest34()
    call SendRequest35()
    call SendRequest36()
    call SendRequest37()
    call SendRequest38()
    call SendRequest39()
    call SendRequest40()
    call SendRequest41()
    call SendRequest42()
    call SendRequest43()
    call SendRequest44()
    call SendRequest45()
    call SendRequest46()
    call SendRequest47()
    call SendRequest48()
    call SendRequest49()
    call SendRequest50()
    call SendRequest51()
    call SendRequest52()
    call SendRequest53()
    call SendRequest54()
    call SendRequest55()
    call SendRequest56()
    call SendRequest57()
    call SendRequest58()
    call SendRequest59()
    call SendRequest60()
    call SendRequest61()
    call SendRequest62()
    call SendRequest63()
    call SendRequest64()
    call SendRequest65()
    call SendRequest66()
    call SendRequest67()
    call SendRequest68()
    call SendRequest69()
    call SendRequest70()
    call SendRequest71()
    call SendRequest72()
    call SendRequest73()
    call SendRequest74()
    call SendRequest75()
    call SendRequest76()
    call SendRequest77()
    call SendRequest78()
    call SendRequest79()
    call SendRequest80()
    call SendRequest81()
    call SendRequest82()
    call SendRequest83()
    call SendRequest84()
    call SendRequest85()
    call SendRequest86()
    call SendRequest87()
    call SendRequest88()
    call SendRequest89()
    call SendRequest90()
    call SendRequest91()
    call SendRequest92()
    call SendRequest93()
    call SendRequest94()
    call SendRequest95()
    call SendRequest96()
    call SendRequest97()
    call SendRequest98()
    call SendRequest99()
    call SendRequest100()
    call SendRequest101()
    call SendRequest102()
    call SendRequest103()
    call SendRequest104()
    call SendRequest105()
    call SendRequest106()
    call SendRequest107()
    call SendRequest108()
    call SendRequest109()
    call SendRequest110()
    call SendRequest111()
    call SendRequest112()
    call SendRequest113()
    call SendRequest114()
    call SendRequest115()
    call SendRequest116()
    call SendRequest117()
    call SendRequest118()
    call SendRequest119()
    call SendRequest120()
    call SendRequest121()
    call SendRequest122()
    call SendRequest123()
    call SendRequest124()
    call SendRequest125()
    call SendRequest126()
    call SendRequest127()
    call SendRequest128()
    call SendRequest129()
    call SendRequest130()
    call SendRequest131()
    call SendRequest132()
    call SendRequest133()
    call SendRequest134()
    call SendRequest135()
    call SendRequest136()
    call SendRequest137()
    call SendRequest138()
    call SendRequest139()
    call SendRequest140()
    call SendRequest141()
    call SendRequest142()
    call SendRequest143()
    call SendRequest144()
    call SendRequest145()
    call SendRequest146()
    call SendRequest147()
    call SendRequest148()
    call SendRequest149()
    call SendRequest150()
    call SendRequest151()
    call SendRequest152()
    call SendRequest153()
    call SendRequest154()
    call SendRequest155()
    call SendRequest156()
    call SendRequest157()
    call SendRequest158()
    call SendRequest159()
    call SendRequest160()
    call SendRequest161()
    call SendRequest162()
    call SendRequest163()
    call SendRequest164()
    call SendRequest165()
    call SendRequest166()
    call SendRequest167()
    call SendRequest168()
    call SendRequest169()
    call SendRequest170()
    call SendRequest171()
    call SendRequest172()
    call SendRequest173()
    call SendRequest174()
    call SendRequest175()
    call SendRequest176()
    call SendRequest177()
    call SendRequest178()
    call SendRequest179()
    call SendRequest180()
    call SendRequest181()
    call SendRequest182()
    call SendRequest183()
    call SendRequest184()
    call SendRequest185()
    call SendRequest186()
    call SendRequest187()
    call SendRequest188()
    call SendRequest189()
    call SendRequest190()
    call SendRequest191()
    call SendRequest192()
    call SendRequest193()
    call SendRequest194()
    call SendRequest195()
    call SendRequest196()
    call SendRequest197()
    call SendRequest198()
    call SendRequest199()
    call SendRequest200()
    call SendRequest201()
    call SendRequest202()
    call SendRequest203()
    call SendRequest204()
    call SendRequest205()
    call SendRequest206()
    call SendRequest207()
    call SendRequest208()
    call SendRequest209()
    call SendRequest210()
    call SendRequest211()
    call SendRequest212()
    call SendRequest213()
    call SendRequest214()
    call SendRequest215()
    call SendRequest216()
    call SendRequest217()
    call SendRequest218()
    call SendRequest219()
    call SendRequest220()
    call SendRequest221()
    call SendRequest222()
    call SendRequest223()
    call SendRequest224()
    call SendRequest225()
    call SendRequest226()
    call SendRequest227()
    call SendRequest228()
    call SendRequest229()
    call SendRequest230()
    call SendRequest231()
    call SendRequest232()
    call SendRequest233()
    call SendRequest234()
    call SendRequest235()
    call SendRequest236()
    call SendRequest237()
    call SendRequest238()
    call SendRequest239()
    call SendRequest240()
    call SendRequest241()
    call SendRequest242()
    call SendRequest243()
    call SendRequest244()
    call SendRequest245()
    call SendRequest246()
    call SendRequest247()
    call SendRequest248()
    call SendRequest249()
    call SendRequest250()
    call SendRequest251()
    call SendRequest252()
    call SendRequest253()
    call SendRequest254()
    call SendRequest255()
    call SendRequest256()
    call SendRequest257()
    call SendRequest258()
    call SendRequest259()
    call SendRequest260()
    call SendRequest261()
    call SendRequest262()
    call SendRequest263()
    call SendRequest264()
    call SendRequest265()
    call SendRequest266()
    call SendRequest267()
    call SendRequest268()
    call SendRequest269()
    call SendRequest270()
    call SendRequest271()
    call SendRequest272()
    call SendRequest273()
    call SendRequest274()
    call SendRequest275()
    call SendRequest276()
    call SendRequest277()
    call SendRequest278()
    call SendRequest279()
    call SendRequest280()
    call SendRequest281()
    call SendRequest282()
    call SendRequest283()
    call SendRequest284()
    call SendRequest285()
    call SendRequest286()
    call SendRequest287()
    call SendRequest288()
    call SendRequest289()
    call SendRequest290()
    call SendRequest291()
    call SendRequest292()
    call SendRequest293()
    call SendRequest294()
    call SendRequest295()
    call SendRequest296()
    call SendRequest297()
    call SendRequest298()
    call SendRequest299()
    call SendRequest300()
    call SendRequest301()
    call SendRequest302()
    call SendRequest303()
    call SendRequest304()
    call SendRequest305()
    call SendRequest306()
    call SendRequest307()
    call SendRequest308()
    call SendRequest309()
    call SendRequest310()
    call SendRequest311()
    call SendRequest312()
    call SendRequest313()
    call SendRequest314()
    call SendRequest315()
    call SendRequest316()
    call SendRequest317()
    call SendRequest318()
    call SendRequest319()
End Sub
Main
