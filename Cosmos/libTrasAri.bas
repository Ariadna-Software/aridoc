Attribute VB_Name = "libTrasAri"
Public ConnAntiguoAridoc As Connection
Public ConnNuevoAridoc As Connection
Public Const vbPermisoTotal = 2147483647

Public SeHaCancelado As Boolean

Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        Select Case Tipo
            Case "N"
                DBLet = 0
            Case "F"
                DBLet = "0:00:00"
            Case Else
                DBLet = ""
        End Select
    Else
        DBLet = vData
    End If
End Function


'Para cuando insertamos en BD, si el texto es "" y pondremos NULL
Public Function CampoANulo(ByRef T As Variant, Optional Tipo As String) As String
    Select Case Tipo
    Case "N"
        If T = 0 Then
            CampoANulo = "NULL"
        Else
            CampoANulo = TransformaComasPuntos(CStr(T))
        End If
    
    Case "F"
        If Val(T) = 0 Then
            CampoANulo = "NULL"
        Else
            CampoANulo = "'" & Format(T, FormatoFecha) & "'"
        End If
    Case Else
        If T = "" Then
            CampoANulo = "NULL"
        Else
            CampoANulo = "'" & T & "'"
        End If
    End Select
End Function


'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ",")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "." & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaComasPuntos = CADENA
End Function


Public Function espera(Segundos As Single)
    Dim T1
    T1 = Timer
    Do
    Loop Until Timer - T1 > Segundos
End Function
