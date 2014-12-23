Attribute VB_Name = "Modulo1"
Public Sub Seleccionar(LimiteInf, LimiteSup As Integer)
  Dim i As Integer
  For i = LimiteInf To LimiteSup Step 1
    Form1.chk(i).Value = 1
  Next
End Sub
