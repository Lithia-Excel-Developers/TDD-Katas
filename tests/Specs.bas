Attribute VB_Name = "Specs"
Public Sub RunSpecs()
    Dim Reporter As New WorkbookReporter
    Reporter.ConnectTo SpecRunner

    Reporter.Start NumSuites:=0
    '                         ^ adjust NumSuites to match number of suites output
    '                           (used for reporting progress)
    Reporter.Output Utilities_Specs.Specs
    ' Reporter.Output Suite2

    Reporter.Done
End Sub

