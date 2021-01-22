testPath = "C:\Users\demo\Documents\Unified Functional Testing\MyAlphaWeb_Demo_2"
'Dim objFSO
'Set objFSO = CreateObject(“Scripting.FileSystemObject”)
'DoesFolderExist = objFSO.FolderExists(testPath)
'Set objFSO = Nothing
'If DoesFolderExist Then
Dim qtApp
Dim qtTest
Dim qtResultsOpt
Set qtApp = CreateObject("QuickTest.Application")
qtApp.Launch
qtApp.Visible = True
qtApp.Open testPath, False
Set qtTest = qtApp.Test
Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtResultsOpt.ResultsLocation = "C:\Users\demo\Documents\Unified Functional Testing\UFT\UFTWorking\res\Test1"
qtTest.Run qtResultsOpt,True
'qtTest.Run
qtTest.Close
qtApp.Quit
'Else
'End If