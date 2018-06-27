Attribute VB_Name = "Utilities_Specs"
Option Explicit

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "Utility Functions"
    
    With Specs.It("SumTwoNumbers() should add two numbers")
    
        .Expect(SumTwoNumbers(1, 1)).ToEqual 2
        
    End With
    
End Function
