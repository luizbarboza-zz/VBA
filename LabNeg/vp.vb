Function VPres(taxa As Double, npar As Integer, pgto As Double) As Double

Dim acc As Double
Dim pagto As Double
pagto = pgto * -1
acc = 0

For i = 1 To npar Step 1
    acc = acc + pagto / ((1 + taxa) ^ i)
Next

 
VPres = acc
 
End Function 