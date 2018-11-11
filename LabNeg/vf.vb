Function VFut(taxa As Double, npar As Integer, pgto As Double) As Double

Dim acc As Double
Dim pagto As Double
pagto = pgto * -1
acc = pagto

For i = 1 To npar - 1 Step 1
    acc = acc * (1 + taxa) + pagto
Next

VFut = acc
 
End Function

