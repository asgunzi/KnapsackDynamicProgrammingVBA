Attribute VB_Name = "M_Generate"
Option Explicit

Sub generateNewProblem()
'It generates a new case for knapsack problem
'It uses random numbers for weights, values and maximum weight

Dim W As Long
Dim n As Long
Dim i As Long
Dim arrCase As Variant 'this array will contain the solution


W = Application.WorksheetFunction.RandBetween(10, 1000)

n = Application.WorksheetFunction.RandBetween(5, 500)
ReDim arrCase(1 To n, 1 To 3)

For i = 1 To n
    arrCase(i, 1) = i
    arrCase(i, 2) = Application.WorksheetFunction.RandBetween(100, 1000) 'Value
    arrCase(i, 3) = Application.WorksheetFunction.RandBetween(1, 100) 'Weight
    
Next i

Range("b8:d1000").ClearContents
Range("b8").Resize(n, 3) = arrCase
Range("c5") = W


End Sub
