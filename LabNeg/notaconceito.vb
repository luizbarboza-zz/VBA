Function NotaConceito(nota As Double) As String

If (nota >= 9) Then
    NotaConceito = "A"
ElseIf (nota < 9 And nota >= 7) Then
    NotaConceito = "B"
ElseIf (nota < 7 And nota >= 5) Then
    NotaConceito = "D"
Else
    NotaConceito = "F"
End If
 

End Function
