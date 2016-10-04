Option Explicit

Dim m_oProject, m_oController

Set m_oProject = OpenProject("C:\Projects\ASP.NET\Chapter7\Code\AspNetChap7ACTTest", "AspNetChap7ACTTest.act")
Set m_oController = CreateObject("ACT.Controller")

Call RunTest(m_oProject, "Category_Fiction_History", m_oController)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function to open an ACT project. 
' Returns the open project.
'
Function OpenProject(strProjectPath, strProjectFileName)
   Dim oProject

   On Error Resume Next

   Set oProject = CreateObject("ACT.Project")
   
   ' check for null object
   If (Not(IsObject(oProject))) Then
      WScript.Echo("Error creating project object")
      Call WScript.Quit()
   Else
      ' open the project
      Call oProject.Open(strProjectPath, strProjectFileName, False)
      ' check for any VB error
      If (Err.Number > 0) Then
         WScript.Echo("Error opening project")
         ' exit the script
         Call WScript.Quit()
      End If   
   End If
   Set OpenProject = oProject
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Procedure to run the specified test from the
' ACT project.
'
Sub RunTest(oProject, strTestName, oController)
	Dim oTest, bIsRunning

	bIsRunning = oController.TestIsRunning
	If bIsRunning Then
		Call oController.StopTest()
		WScript.Echo("ACT is already running a test.")
	Else
		Set oTest = oProject.Tests.Item(strTestName)
		'WScript.Echo("Starting test....")
		Call oController.StartTest(oProject, oTest, False)
	End If
End Sub


