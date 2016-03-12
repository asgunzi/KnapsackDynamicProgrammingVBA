Attribute VB_Name = "M_DynamicProg"
Option Explicit

Sub testDynamicProg()
Dim arrValues As Variant, arrWeigths As Variant, TotalWeight As Long, arrAux As Variant, arrSolution As Variant
Dim n As Long, FO As Long

n = Application.WorksheetFunction.CountA(Range("b8:b10000"))

Plan1.Activate
arrValues = Range("c8:c" & 7 + n)
arrWeigths = Range("d8:d" & 7 + n)
TotalWeight = Range("c5")

dynamicProgramming arrValues, arrWeigths, TotalWeight, arrAux, arrSolution, FO

'Paste FO
Range("I5") = FO
Range("h8:i10000").ClearContents
If n > 0 Then
    Range("h8").Resize(n, 2) = arrSolution
End If



End Sub


Private Sub dynamicProgramming(arrValues As Variant, arrWeigths As Variant, TotalWeight As Long, arrAux As Variant, arrSolution As Variant, FO As Long)
Dim W As Long, i As Long
Dim n As Long
Dim strSolution As Variant 'array to save the solution
Dim v1 As Long, v2 As Long
Dim swap As Variant, idx As Long

If Not IsArray(arrValues) Then
    Exit Sub
End If


n = UBound(arrValues, 1)



ReDim arrAux(0 To n, 0 To TotalWeight)
ReDim strSolution(0 To n, 0 To TotalWeight)
ReDim arrSolution(1 To n, 1 To 2)


For W = 0 To TotalWeight
    For i = 1 To n
        If arrWeigths(i, 1) <= W Then
            v1 = arrAux(i - 1, W)
            v2 = arrValues(i, 1) + arrAux(i - 1, W - arrWeigths(i, 1))
            
            If v1 >= v2 Then
                arrAux(i, W) = v1
                strSolution(i, W) = strSolution(i - 1, W)
            Else
                arrAux(i, W) = v2
                strSolution(i, W) = strSolution(i - 1, W - arrWeigths(i, 1)) & i & ","
            
            End If
            
            
        Else
            arrAux(i, W) = arrAux(i - 1, W)
            strSolution(i, W) = strSolution(i - 1, W)
        End If
    
    Next i
Next W


FO = arrAux(n, TotalWeight)


'Explode the solution
swap = Strings.Split(strSolution(n, TotalWeight), ",")




'Preparation of solution array
For i = 1 To n
    arrSolution(i, 1) = i
    arrSolution(i, 2) = 0
Next i


'If it is the solution, fill with 1
For i = 0 To UBound(swap)
    If swap(i) <> "" Then
        idx = Conversion.CLng(swap(i))
        arrSolution(idx, 2) = 1
    End If
Next i




End Sub

Private Function max(v1 As Long, v2 As Long) As Long

    If v1 >= v2 Then
        max = v1
    Else
        max = v2
    End If
End Function
