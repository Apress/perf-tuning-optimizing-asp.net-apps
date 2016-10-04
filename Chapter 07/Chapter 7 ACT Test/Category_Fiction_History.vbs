Option Explicit
Dim fEnableDelays
fEnableDelays = False

Sub SendRequest1()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
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
    If fEnableDelays = True then Test.Sleep (31)
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

Sub SendRequest3()
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

Sub SendRequest4()
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

Sub SendRequest9()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest10()
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

Sub SendRequest11()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest12()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (0)
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

Sub SendRequest13()
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

Sub SendRequest14()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (14891)
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
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

Sub SendRequest15()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (7611)
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

Sub SendRequest16()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
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

Sub SendRequest17()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
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

Sub SendRequest18()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
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

Sub SendRequest19()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
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

Sub SendRequest20()
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

Sub SendRequest21()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=837"
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

Sub SendRequest22()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (100)
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

Sub SendRequest23()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (201)
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

Sub SendRequest24()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest25()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
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

Sub SendRequest26()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (170)
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

Sub SendRequest27()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest28()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest29()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
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

Sub SendRequest30()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (100)
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

Sub SendRequest31()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (201)
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

Sub SendRequest32()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
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

Sub SendRequest33()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest34()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
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

Sub SendRequest35()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest36()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest37()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest38()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (180)
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

Sub SendRequest39()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest40()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
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

Sub SendRequest41()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (80)
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

Sub SendRequest42()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest43()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest44()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (171)
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

Sub SendRequest45()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest46()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
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

Sub SendRequest47()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (140)
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

Sub SendRequest48()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest49()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest50()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (160)
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

Sub SendRequest51()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest52()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (221)
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

Sub SendRequest53()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest54()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (180)
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

Sub SendRequest55()
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

Sub SendRequest56()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (150)
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

Sub SendRequest57()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest58()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (190)
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

Sub SendRequest59()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest60()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest61()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (161)
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

Sub SendRequest62()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest63()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest64()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (190)
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

Sub SendRequest65()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest66()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (180)
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

Sub SendRequest67()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest68()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
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

Sub SendRequest69()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (141)
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

Sub SendRequest70()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest71()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest72()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (170)
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

Sub SendRequest73()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest74()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest75()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (170)
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

Sub SendRequest76()
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

Sub SendRequest77()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (70)
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

Sub SendRequest78()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
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

Sub SendRequest79()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest80()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest81()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (31)
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

Sub SendRequest82()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest83()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
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

Sub SendRequest84()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest85()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (10)
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

Sub SendRequest86()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (160)
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

Sub SendRequest87()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
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

Sub SendRequest88()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
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

Sub SendRequest89()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
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

Sub SendRequest90()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (201)
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

Sub SendRequest91()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest92()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest93()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (140)
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

Sub SendRequest94()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest95()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest96()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (190)
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

Sub SendRequest97()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest98()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest99()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest100()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
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

Sub SendRequest101()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest102()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (41)
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

Sub SendRequest103()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest104()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
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

Sub SendRequest105()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
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

Sub SendRequest106()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest107()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (130)
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

Sub SendRequest108()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest109()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
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

Sub SendRequest110()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest111()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
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

Sub SendRequest112()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest113()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest114()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (161)
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

Sub SendRequest115()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest116()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest117()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (150)
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

Sub SendRequest118()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest119()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest120()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (40)
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

Sub SendRequest121()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (110)
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

Sub SendRequest122()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest123()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest124()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (150)
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

Sub SendRequest125()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest126()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (181)
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

Sub SendRequest127()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest128()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest129()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest130()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (140)
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

Sub SendRequest131()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest132()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (190)
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

Sub SendRequest133()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest134()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest135()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (131)
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

Sub SendRequest136()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest137()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest138()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest139()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest140()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (90)
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

Sub SendRequest141()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest142()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (20)
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

Sub SendRequest143()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
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

Sub SendRequest144()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (3064)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/categories.aspx"+"?id=838"
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

Sub SendRequest145()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (5629)
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
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

Sub SendRequest146()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
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

Sub SendRequest147()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
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

Sub SendRequest148()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
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

Sub SendRequest149()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
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

Sub SendRequest150()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
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

Sub SendRequest151()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
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

Sub SendRequest152()
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
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

Sub SendRequest153()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (120)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/291.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/291.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest154()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (401)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/228.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/228.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest155()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (80)
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
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

Sub SendRequest156()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/229.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/229.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest157()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
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
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
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

Sub SendRequest158()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/235.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/235.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest159()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/238.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/238.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest160()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (50)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/236.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/236.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest161()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/291.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/291.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest162()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (270)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/297.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/297.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest163()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (61)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/298.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/298.GIF"
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
        oRequest.Path = "/Duwamish7vb/images/books/small/300.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/300.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest165()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (290)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/308.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/308.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest166()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/309.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/309.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest167()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (170)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/314.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/314.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest168()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (90)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/315.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/315.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest169()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (41)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/313.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/313.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest170()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (440)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/327.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/327.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest171()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (200)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/331.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/331.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest172()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (61)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/334.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/334.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest173()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/335.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/335.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest174()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (80)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/341.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/341.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest175()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (60)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/342.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/342.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest176()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (140)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/345.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/345.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest177()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (401)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/353.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/353.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest178()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (400)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/364.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/364.GIF"
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
        oRequest.Path = "/Duwamish7vb/images/books/small/370.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/370.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest180()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (140)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/371.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/371.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest181()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (201)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/380.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/380.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest182()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (220)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/382.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/382.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest183()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (30)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/383.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/383.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest184()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (150)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/393.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/393.GIF"
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
        oRequest.Path = "/Duwamish7vb/images/books/small/396.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/396.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest186()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (351)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/417.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/417.GIF"
        Else
            strStatusCode = oResponse.ResultCode
        End If
        oConnection.Close
    End If
End Sub

Sub SendRequest187()
    Dim oConnection, oRequest, oResponse, oHeaders, strStatusCode
    If fEnableDelays = True then Test.Sleep (411)
    Set oConnection = Test.CreateConnection("kenneth-01", 80, false)
    If (oConnection is Nothing) Then
        Test.Trace "Error: Unable to create connection to kenneth-01"
    Else
        Set oRequest = Test.CreateRequest
        oRequest.Path = "/Duwamish7vb/images/books/small/527.GIF"
        oRequest.Verb = "GET"
        oRequest.HTTPVersion = "HTTP/1.0"
        set oHeaders = oRequest.Headers
        oHeaders.RemoveAll
        oHeaders.Add "Accept", "*/*"
        oHeaders.Add "Referer", "http://kenneth-01/Duwamish7vb/categories.aspx?id=838"
        oHeaders.Add "Accept-Language", "en-us"
        oHeaders.Add "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0; .NET CLR 1.0.3705)"
        'oHeaders.Add "Host", "kenneth-01"
        oHeaders.Add "Host", "(automatic)"
        oHeaders.Add "Cookie", "(automatic)"
        Set oResponse = oConnection.Send(oRequest)
        If (oResponse is Nothing) Then
            Test.Trace "Error: Failed to receive response for URL to " + "/Duwamish7vb/images/books/small/527.GIF"
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
End Sub
Main
