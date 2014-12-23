Attribute VB_Name = "Modulo2"
Public Sub Deseleccionar(LimiteInf, LimiteSup As Integer)
  Dim i As Integer
  For i = LimiteInf To LimiteSup Step 1
    Form1.chk(i).Value = 0
  Next
End Sub

