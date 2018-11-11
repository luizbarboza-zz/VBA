
Function Tabuada() As String

Dim Tabu As String
Dim Calc As String

For i = 1 To 9 Step 1
    For j = 1 To 9 Step 1
        Calc = "" & i & "x" & j & "=" & i * j
        Tabu = Tabu & Calc & " "
    Next
    Tabu = Tabu & vbCrLf
Next
   
   
Tabuada = Tabu
 

End Function