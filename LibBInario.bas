Attribute VB_Name = "LibBInario"
Option Explicit

Public Const vbPermisoTotal = 2147483647

Dim i As Integer

Public Function BiarioLONG(ByRef Cad As String)
Dim TOT As Long
Dim L As Long
Dim Exp As Integer

    TOT = 0
    For i = 1 To Len(Cad)
         If Mid(Cad, i, 1) = 1 Then
            Exp = vbMaxGrupos - i
            
            L = (2 ^ (Exp))
            TOT = TOT + L
        End If
    Next i
    BiarioLONG = TOT
End Function

'Para poner el proipetario
'Est es el grupo 5  00...0010000
'es el 16 en numerico desde binario
Public Function GrupoLongBD(ByRef CodigoGrupo As Long)
Dim Cad As String
    
    For i = 1 To vbMaxGrupos
        Cad = Cad & "0"
    Next i
    
    i = vbMaxGrupos - CodigoGrupo + 1
    
    Cad = Mid(Cad, 1, i - 1) & "1" & Mid(Cad, i + 1)
  
    GrupoLongBD = BiarioLONG(Cad)
    
End Function


