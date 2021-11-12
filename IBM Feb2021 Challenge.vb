Option Explicit 
Dim BinGrid(), vacunados() As String
Dim largo, ancho As Integer
Dim posiblesResp() As Integer
Dim filas, columnas As Integer

Sub programa()
Dim iterador As Integer

initData
For iterador = 1 To 5
    recorredor
Next
printer

End Sub

Sub initData()
Dim fila, columna As Integer

filas = Range("BinGrid").Rows.Count
columnas = Range("BinGrid").Columns.Count
ReDim BinGrid(filas + 1, columnas + 1)
ReDim vacunados(filas + 1, columnas + 1)

For fila = 1 To filas
    For columna = 1 To columnas
        BinGrid(fila, columna) = Range("BinGrid").Cells(fila, columna)
        vacunados(fila, columna) = 0
        If fila = 1 Then
            vacunados(fila - 1, columna) = 1
        End If
        If columna = 1 Then
            vacunados(fila, columna - 1) = 1
        End If
        If fila = 12 Then
            vacunados(fila + 1, columna) = 1
        End If
        If columna = 12 Then
            vacunados(fila, columna + 1) = 1
        End If
    Next
Next

End Sub

Sub recorredor()
Dim fila, columna, vacunar As Integer

For fila = 1 To filas
    For columna = 1 To columnas
        vacunar = 0
        If Mid(BinGrid(fila, columna), 1, 1) = 0 Or vacunados(fila - 1, columna) = 1 Then
            vacunar = vacunar + 1
        End If
        If Mid(BinGrid(fila, columna), 2, 1) = 0 Or vacunados(fila, columna + 1) = 1 Then
            vacunar = vacunar + 1
        End If
        If Mid(BinGrid(fila, columna), 3, 1) = 0 Or vacunados(fila + 1, columna) = 1 Then
            vacunar = vacunar + 1
        End If
        If Mid(BinGrid(fila, columna), 4, 1) = 0 Or vacunados(fila, columna - 1) = 1 Then
            vacunar = vacunar + 1
        End If
        If vacunar = 4 Then
            vacunados(fila, columna) = 1
        End If
    Next
Next

End Sub

Sub printer()
Dim fila, columna, vacunar As Integer

For fila = 1 To filas
    For columna = 1 To columnas
        Sheet2.Cells(fila, columna).Value = vacunados(fila, columna)
    Next
Next

End Sub

