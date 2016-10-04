Const SERVER_NAME   = "kenneth-01"
Const DATA_FILE_PATH = "C:/Projects/ASP.NET/Chapter7/Code/AspNetChap7ACTTest/"
Const DATA_FILE_NAME = "RequestPaths.dat"


' Run the test
Call Main()


Sub Main
      
   Dim strURLPath, arrURLPaths(100)
   Dim lCount, lNumberOfItemsInArray
   Dim oConnection, oUser, oResponse, oRequest
      
   strURLPath = ""
   lCount = 0
   lNumberOfItemsInArray = 0
   strURLPath = ""


   Test.GetNextUser()
   
   Set oConnection = Test.CreateConnection(SERVER_NAME, 80, False)
   
   If (oConnection Is Nothing) Then
      Call Test.Trace("E: Unable to create connection." & vbCrLf)
   ElseIf (oConnection.IsOpen = False) Then
      Call Test.Trace("E: Unable to open connection." & vbCrLf)
   Else
      Set oRequest = Test.CreateRequest

      ' Populate the URL array and store the number of 
      ' items read from file.
      lNumberOfItemsInArray = GetDataFromFile( (DATA_FILE_PATH & DATA_FILE_NAME) , arrURLPaths)
      
      ' loop through each URL
      For lCount = 0 To (lNumberOfItemsInArray - 1)
         ' get an element from the array
         strURLPath = arrURLPaths(lCount) 
         ' assign Request.Path property
         oRequest.Path = strURLPath
   
         ' send the request
         Set oResponse = oConnection.Send(oRequest)
         ' check for a bad request or connection
         If (oResponse Is Nothing) Then
            Call Test.Trace("E: Invalid request or host not found." & VbCrLf)         
         Else
            Call Test.Trace("I: Requested '" & strURLPath & "'" & vbCrLf)
            Call Test.Trace("I: Server response: " & oResponse.ResultCode & VbCrLf)
         End If
      Next
   End If
End Sub



Function GetDataFromFile(strFilePath, ByRef arrData)
   ' declare local variables
   Dim oFileSys
   Dim oDataFile
   Dim strLine
   Dim lCount
   
   ' initialize variables
   strLine = ""
   lCount = 0
   
   ' create an FSO
   Set oFileSys = CreateObject("Scripting.FileSystemObject")
   ' open file for reading
   Set oDataFile = oFileSys.OpenTextFile(strFilePath, 1)

   Do While (oDataFile.AtEndOfStream <> True)   
       ' read one line at a time
      strLine = oDataFile.ReadLine
      ' assign value from each line to an array element
      ' and check for empty lines
      If (strLine <> "") Then
         arrData(lCount) = strLine
         lCount = lCount + 1
      Else
         Test.Trace("E: Found empty line in data file" & vbCrLf)
      End If
   Loop

   Call oDataFile.Close()
   ' return the number of lines read
   GetDataFromFile = lCount
End Function
