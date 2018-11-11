Function MediaTurmaRange(Alunos As Range) As Double

Dim Soma As Double
Dim Qtd As Integer
Dim Nota As Double

Qtd = Alunos.Count

For i = 1 To Qtd Step 1
    Nota = CDbl(Alunos.Cells(i, 1).Value)
    Soma = Soma + Nota
Next

MediaTurmaRange = Soma / Qtd

End Function
