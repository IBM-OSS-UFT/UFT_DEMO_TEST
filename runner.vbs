'Create QTP object
Set QTP = CreateObject("QuickTest.Application")
QTP.Launch
QTP.Visible = TRUE
 
'Open QTP Test
QTP.Open "C:\Program Files\Git\Test\GUITest1", TRUE 'Set the QTP test path
 
'Set Result location
Set qtpResultsOpt = CreateObject("QuickTest.RunResultsOptions")
qtpResultsOpt.ResultsLocation = "C:\Program Files\Git\Test\GUITest1\Res1\Report"
'Run QTP test
QTP.Test.Run qtpResultsOpt
 
'Close QTP
QTP.Test.Close
QTP.Quit