Attribute VB_Name = "Utilities_Specs"
Option Explicit

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "Utility Functions"
    
    With Specs.It("SumTwoNumbers() should add two numbers")
        
        .Expect(sumtwonumbers(1, 1)).ToEqual 2
        .Expect(sumtwonumbers(-250, 112.65)).ToEqual (112.65 - 250)
        .Expect(sumtwonumbers("1", "1")).ToEqual 2
        
    End With
    
End Function
