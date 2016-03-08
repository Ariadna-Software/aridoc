Attribute VB_Name = "modBackup"
Option Explicit


Public Sub BACKUP_TablaIzquierda(ByRef Rs As ADODB.Recordset, ByRef CADENA As String)
Dim I As Integer
Dim nexo As String

    CADENA = ""
    nexo = ""
    For I = 0 To Rs.Fields.Count - 1
        CADENA = CADENA & nexo & Rs.Fields(I).Name
        nexo = ","
    Next I
    CADENA = "(" & CADENA & ")"
End Sub





'---------------------------------------------------
'El fichero siempre sera NF
Public Sub BACKUP_Tabla(ByRef Rs As ADODB.Recordset, ByRef Derecha As String)
Dim I As Integer
Dim nexo As String
Dim Valor As String
Dim Tipo As Integer
    Derecha = ""
    nexo = ""
    For I = 0 To Rs.Fields.Count - 1
        Tipo = Rs.Fields(I).Type
        
        If IsNull(Rs.Fields(I)) Then
            Valor = "NULL"
        Else
            
            'pruebas
            Select Case Tipo
            'TEXTO
            Case 129, 200, 201
                Valor = Rs.Fields(I)
                NombreSQL Valor    '.-----------> 23 Octubre 2003.
                Valor = "'" & Valor & "'"
            'Fecha
            Case 133
                Valor = CStr(Rs.Fields(I))
                Valor = "'" & Format(Valor, FormatoFecha) & "'"
                
            'Numero normal, sin decimales
            Case 2, 3, 16 To 19
                Valor = Rs.Fields(I)
            
            'Numero con decimales
            Case 131
                Valor = CStr(Rs.Fields(I))
                Valor = TransformaComasPuntos(Valor)
            Case Else
                Valor = "Error grave. Tipo de datos no tratado." & vbCrLf
                Valor = Valor & vbCrLf & "SQL: " & Rs.Source
                Valor = Valor & vbCrLf & "Pos: " & I
                Valor = Valor & vbCrLf & "Campo: " & Rs.Fields(I).Name
                Valor = Valor & vbCrLf & "Valor: " & Rs.Fields(I)
                MsgBox Valor, vbExclamation
                MsgBox "El programa finalizara. Avise al soporte técnico.", vbCritical
                End
            End Select
        End If
        Derecha = Derecha & nexo & Valor
        nexo = ","
    Next I
    Derecha = "(" & Derecha & ")"
End Sub
