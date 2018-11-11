Function MaiorNota() As Double

Dim Maior As Double
Maior = 0
  
Dim Notas(10) As Double
Notas(1) = 7
Notas(2) = 5
Notas(3) = 5
Notas(4) = 7
Notas(5) = 7
Notas(6) = 8
Notas(7) = 9
Notas(8) = 9
Notas(9) = 10
Notas(10) = 7

For i = 1 To 9 Step 1
    If (Notas(i) > Maior) Then
        Maior = Notas(i)
    End If
Next

MaiorNota = Maior

  
   
End Function

 
