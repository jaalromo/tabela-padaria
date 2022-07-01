Option Explicit
Dim codigonumero, Quantida, nome, total As Double
Dim productos(1 To 50, 1 To 4)
Dim filas, columnas As Integer

Private Sub Codigo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'numeros 96 a 105
'If esNumeroCode(KeyCode) Then
codigonumero = Codigo.Value
Quantida = Quantidade.Value
   If codigonumero = "" Or codigonumero = 0 Then codigonumero = 1
    codigonumero = codigonumero + 1
    NomeProduto.Text = ActiveWorkbook.Worksheets("Precos").Cells(codigonumero, 2)
    nome = Cells(codigonumero, 2)
    ValorTotal.Value = Format(ActiveWorkbook.Worksheets("Precos").Cells(codigonumero, 7) * Quantida, "0.00")
   'End If
End Sub

Private Sub UserForm_Initialize()
filas = 0
End Sub

Function esNumeroCode(tecla)
'numeros 96 a 105
If tecla < 96 Or tecla > 105 Then
esNumeroCode = False
Else
esNumeroCode = True
End If
End Function

Function esNumeroAscii(tecla)
'numeros 48 a 57
If tecla < 48 Or tecla > 57 Then
esNumeroAscii = False
Else
esNumeroAscii = True
End If
End Function

Private Sub botonfin_Enter()
ListBox1.AddItem Quantidade.Text & "  " & NomeProduto.Text & "  " & ValorTotal.Text
total = total + ValorTotal.Text
totalidad.Value = Format(total, "0.00")
End Sub

Private Sub botonfin_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'enter = 13, + = 107

If KeyCode = 13 Then
Codigo.Text = ""
Codigo.SetFocus
Quantidade.Text = 0
total = 0
ListBox1.Clear
totalidad.Value = 0
Codigo.Text = ""
Codigo.SetFocus
Quantidade.Text = 0
ElseIf KeyCode = 107 Then
KeyCode = 0
Quantidade.Text = 0
Codigo.Text = ""
Codigo.SetFocus
End If

End Sub

'Private Sub Codigo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'If esNumero(KeyAscii) Then
'codigonumero = Codigo.Value
'Quantida = Quantidade.Value
 '   If codigonumero = "" Or codigonumero = 0 Then codigonumero = 1
'    codigonumero = codigonumero + 1
 '   NomeProduto.Text = ActiveWorkbook.Worksheets("Precos").Cells(codigonumero, 2)
'    nome = Cells(codigonumero, 2)
'    ValorTotal.Value = Format(ActiveWorkbook.Worksheets("Precos").Cells(codigonumero, 7) * Quantida, "0.00")
 '   End If

'End Sub

Private Sub Quantidade_Change()
Quantida = Quantidade.Text
If Quantida = "" Then Quantida = 0
ValorTotal.Value = Format(ActiveWorkbook.Worksheets("Precos").Cells(codigonumero, 7) * Quantida, "0.00")
End Sub

Private Sub Quantidade_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
If KeyAscii = 43 Then '"+" = 43
KeyAscii = 0
ListBox1.AddItem Quantidade.Text & "  " & NomeProduto.Text & "  " & ValorTotal.Text
total = total + ValorTotal.Text
totalidad.Value = Format(total, "0.00")
filas = filas + 1
columnas = 1

'nombre de producto
productos(filas, columnas) = ActiveWorkbook.Worksheets("Precos").Cells(codigonumero, 2)
columnas = columnas + 1
'cantidad
productos(filas, columnas) = Quantida
columnas = columnas + 1
'total item
productos(filas, columnas) = ValorTotal.Text
columnas = columnas + 1
'lucro total
productos(filas, columnas) = ActiveWorkbook.Worksheets("Precos").Cells(codigonumero, 9) * Quantida
columnas = columnas + 1

Quantidade.Text = 0
Codigo.Text = ""
Codigo.SetFocus
MsgBox (productos(filas, 1))
End If
End Sub


+ 107
backspace 8