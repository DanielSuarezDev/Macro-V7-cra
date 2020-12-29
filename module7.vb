Sub Varrr()

lar = Sheets("VARIACION").Range("A" & Rows.Count).End(xlUp).Row


For i = 4 To lar
que = Sheets("VARIACION").Cells(i, 25)
If Not IsError(Sheets("VARIACION").Cells(i, 25)) Then
Sheets("VARIACION").Cells(i, 25) = Application.VLookup(Sheets("VARIACION").Cells(i, 4), Sheets("VALIDA CANCELADOS").Range("E:F"), 2, 0)
Else
Sheets("VARIACION").Cells(i, 25) = "NO"
End If
que = Sheets("VARIACION").Cells(i, 25)

If Not IsError(Sheets("VARIACION").Cells(i, 25)) Then
Sheets("VARIACION").Cells(i, 25) = Application.VLookup(Sheets("VARIACION").Cells(i, 4), Sheets("VALIDA CANCELADOS").Range("E:F"), 2, 0)
Else
Sheets("VARIACION").Cells(i, 25) = "NO"
End If


If Sheets("VARIACION").Cells(i, 25) <> "NO" Then
  Sheets("VARIACION").Cells(i, 25).Interior.ColorIndex = 8
  Sheets("VARIACION").Cells(i, 1).Interior.ColorIndex = 8
Else
  Sheets("VARIACION").Cells(i, 25).Interior.ColorIndex = 2
 Sheets("VARIACION").Cells(i, 1).Interior.ColorIndex = 2
End If
Next i
End Sub
Sub VALIDACION_REPORTE()

lar = Sheets("VALIDACION").Range("A" & Rows.Count).End(xlUp).Row


For i = 3 To lar
CC = Sheets("VALIDACION").Cells(i, 2)
If Not IsError(Sheets("VALIDACION").Cells(i, 5)) Then
Sheets("VALIDACION").Cells(i, 4) = Application.VLookup(Sheets("VALIDACION").Cells(i, 2), Sheets("CONTABILIZADOS").Range("A:M"), 13, 0)
Sheets("VALIDACION").Cells(i, 5) = Application.VLookup(Sheets("VALIDACION").Cells(i, 2), Sheets("CONTABILIZADOS").Range("A:A"), 1, 0)
Else
Sheets("VALIDACION").Cells(i, 4) = "NO"
Sheets("VALIDACION").Cells(i, 5) = "NO"
End If


If Trim(Sheets("VALIDACION").Cells(i, 1)) = Trim(Sheets("VALIDACION").Cells(i, 4)) And Trim(Sheets("VALIDACION").Cells(i, 2)) = Trim(Sheets("VALIDACION").Cells(i, 5)) Then
Sheets("VALIDACION").Cells(i, 3) = "CORRECTO"
Else
Sheets("VALIDACION").Cells(i, 3) = "VALIDAR"
MsgBox ("Validar Cedula " & CC), vbCritical, "Analisis Op"
End If
Next i
End Sub



