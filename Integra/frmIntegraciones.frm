VERSION 5.00
Begin VB.Form frmIntegraciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Integraciones aridoc"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmIntegraciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Integrar"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1860
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5550
      TabIndex        =   1
      Top             =   1830
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   645
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "¿ Continuar con la integracion ?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Hay archivos pendientes de integrar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "frmIntegraciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumeroIntegraciones As Integer
Private NuevasCarpetasCreadas As Boolean

Private Sub Command1_Click()
    Unload Me
End Sub



'INTEGRAR
Private Sub Command2_Click()
Dim miNombre As String
Dim Errores As Boolean
Dim SubError As String

    Screen.MousePointer = vbHourglass
    Errores = False
    SubError = ""
    NuevasCarpetasCreadas = False
    Do
        ErrorLlevando = False
        miNombre = Dir(vConfig.PathArchivos & "\*." & vConfig.extensionGuia, vbArchive)
        If miNombre <> "" Then
            If Not IntegrarArchivo(miNombre) Then
                Errores = True
            Else
                If ErrorLlevando Then
                    SubError = SubError & vbCrLf & "-- " & miNombre
                End If
            End If
        End If
    Loop Until miNombre = ""
    
    'Si se han creado carpetas nuevas UPDATEAREMOS LA TABLA
    If NuevasCarpetasCreadas Then UpdatearTablaCarpetasCreadas
    
    If SubError <> "" Then MsgBox "Errores: " & vbCrLf & SubError, vbExclamation
    
    
    If Errores Then RevisarPendientes = True
    Unload Me
End Sub




Private Function IntegrarArchivo(NombreArchivo As String) As Boolean
Dim NF As Integer
Dim Linea As String
Dim Fin As Boolean
Dim Errore As Boolean
Dim Co As Integer
Dim Espacio As Long

    'Intentamos renombrarlo
    If Not CambiarNombre(NombreArchivo) Then Exit Function

    'Borramos temporal, aunque deberia estar vacia
    Conn.Execute "DELETE FROM tmpintegra where codusu = " & CodPC
    
    
    
    NF = FreeFile
    Open vConfig.PathArchivos & "\" & NombreArchivo For Input As #NF
    Fin = EOF(NF)
    Errore = False
    Co = 0
    While Not Fin
        Co = Co + 1
        Line Input #NF, Linea
        If ProcesarLinea(Linea, Co) > 0 Then
            Errore = True
            Fin = True
        Else
            Fin = EOF(NF)
        End If
        
    Wend
    Close (NF)
    
    If Errore Then Exit Function
    
    'AHora ya tengo los datos en la tabla temporal
    'Comprobaremos k estan todos los archivos
    Set miRsAux = New ADODB.Recordset
    miRsAux.Open "Select NombreArchivo from tmpIntegra where codusu =" & CodPC, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Linea = vConfig.PathArchivos & "\" & miRsAux!NombreArchivo
        Me.Label2.Caption = Linea
        Label2.Refresh
        If Dir(Linea, vbArchive) = "" Then
            MsgBox "No se encuentra el archivo: " & Linea & vbCrLf & vbCrLf & "Faltan archivos", vbExclamation
            miRsAux.Close
            Exit Function
            
        Else
            Espacio = FileLen(Linea)
           
            Linea = "UPDATE tmpintegra set tamanyo = " & Espacio & " where Nombrearchivo='" & miRsAux!NombreArchivo & "' and codusu=" & CodPC
        End If
        
        miRsAux.MoveNext
        'Ejecutamos el update
        Conn.Execute Linea
        
    Wend
    miRsAux.Close
    
    
    'Llegados aqui ya esta integrado, con lo cual solo hay hacer llamadas al aridoc
    'primero intentaremos crear las carpetas
    If Not InsertarCarpetas Then Exit Function
    
    
    'Una comprobacion mas.
    ' Todos los archvios tiene codcarpeta
    Linea = "Select codcarpeta from tmpIntegra where codusu =" & CodPC & " and codcarpeta <1"
    miRsAux.Open Linea, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRsAux.EOF Then
        MsgBox "Algunas carpetas no han sido creadas", vbExclamation
        Exit Function
    End If
    miRsAux.Close
    
    
    'AHORA LLAMAMOS A frmFTP
    'QUe ira cojiendo los archivos y llevandolos
    Linea = NombreArchivo
    NF = InStr(1, Linea, ".")
    If NF > 0 Then Linea = Mid(Linea, 1, NF - 1)
    frmvFTP.Fich = Linea
    frmvFTP.opcion = 1
    frmvFTP.Show vbModal
    IntegrarArchivo = True
    
    
    'Eliminamos la referencia
    EliminaArchivo vConfig.PathArchivos & "\" & NombreArchivo
End Function




Private Function ProcesarLinea(Lin As String, Contador As Integer) As Byte
Dim Fin As Boolean
Dim CADENA As String
Dim pos As Integer
Dim valor As String
Dim ind As Integer
Dim mImg As CImag
Dim Nombre As String
'Dim vExt As String

On Err GoTo ErrorHdle
CADENA = Lin
Fin = False
ind = 0
Set mImg = New CImag
mImg.Id = Contador
Label2.Caption = Lin
Label2.Refresh
While Not Fin
        If Mid(CADENA, 1, 1) = "|" Then 'Con esto comprobamos que los campos estan llenos
            valor = ""
            CADENA = Mid(CADENA, 2, Len(CADENA))
            ind = ind + 1
        Else
            pos = InStr(1, CADENA, "|")
           
            If pos > 0 Then
                valor = Mid(CADENA, 1, pos - 1)
                CADENA = Mid(CADENA, pos + 1, Len(CADENA))
                ind = ind + 1
                Else
                    Fin = True
                    ind = ind + 1
                    valor = CADENA
            End If
        End If
        
    Select Case ind
    Case 1: 'Nombre archivo
            mImg.NomFich = valor
    Case 2: 'Clave1
            Nombre = ProcesaLinea2(valor)
            mImg.Clave1 = Trim(Nombre)
    Case 3: 'c2
            mImg.Clave2 = valor
    Case 4: 'c3
            mImg.Clave3 = valor
    Case 5: 'fechadoc
            mImg.FechaDoc = valor
    Case 6: 'path
            mImg.NomPath = devuelvePATH(valor)   'Le cambiamos las barras
    Case Else
        valor = valor 'Sumidero para que no de error
    
    End Select
Wend

    If mImg.Agregar(CLng(CodPC)) Then ProcesarLinea = 0
' Antes de añadirlo le ponemos como fecha de digitalización la de hoy
'mImg.FechaDig = Date
'
'If mImg.Clave1 = "" Or mImg.NomFich = "" Or mImg.NomPath = "" Then
'    'Error leyendo los datos
'    ProcesarLinea = 1
'    Exit Function
'End If
'
''Comprobamos que existe el archivo
'valor = Dir(mConfig.carpetaInt & mImg.NomFich)
'vExt = ""
'If valor = "" Then
'    'NUEVA modificacion
'    'Si hemos llegado aqui, es ke no existe el archivo
'    'Ahora comprobamos si no existe el archvio con .DAT
'    If Dir(mConfig.carpetaInt & mImg.NomFich & ".DAT") = "" Then
'        'Error leyendo los datos
'        ProcesarLinea = 1
'        Exit Function
'    Else
'        vExt = ".dat"
'    End If
'End If
'
''Comprobamos si esta bien la fecha
'valor = mImg.FechaDoc
'If Not IsDate(valor) Then
'   'Error leyendo los datos
'    ProcesarLinea = 1
'    Exit Function
'End If
'
''Comprobaremos el directorio, añadiendo la carpeta inicial
'If TratarCarpeta(mImg.NomPath) = 1 Then
'    'Se ha producido un error al crear una de las carpetas
'    ProcesarLinea = 1
'    Exit Function
'End If
'
'nombre = mConfig.carpetaInt & mImg.NomFich & vExt
'mImg.NomFich = mImg.Id & ".nfi"
'mImg.Extension = NFI_Extension
'
'mImg.NomPath = mImg.NomPath & "\"
'If CompruebaCarpeta(mImg.NomPath, valor) = 3 Then  '1. con archivos  2.-vacio
'    'Se ha producido un error al crear una de las carpetas
'    ProcesarLinea = 1
'    Exit Function
'End If
'
'' luego lo asignamos a nombre, no sin antes eliminar el txt anterior
'' Puede que deberiamos de trabajar directamente sobre el nombre final
'If mImg.Agregar = 0 Then
'    If CopiaArchivo(nombre, mConfig.dirbase & mConfig.Carpeta & "\" & mImg.NomPath & mImg.NomFich) Then
'        Kill nombre
'        ProcesarLinea = 0
'    Else
'        mImg.Eliminar
'        ProcesarLinea = 1
'    End If
'    'ELSE
'    Else
'        ProcesarLinea = 1
'End If
'
Set mImg = Nothing
Exit Function
ErrorHdle:
    MsgBox "Err: " & Err.Number & vbCrLf & " Des:  " & Err.Description, vbExclamation
End Function





'-----------------------------------------------
'-----------------------------------------------
'-----------------------------------------------




Public Function ProcesaLinea2(ByRef L As String) As String
Dim I, C, l2
Dim J As Byte
l2 = ""
'Para que no tenga que hacer cada vez el select, y sabiendo que casi todo son letras y numero
'Para saber si lo tenemos que modificar
'comprobaremos que el ASC es mayor 165 para saber si hay que hacer cambios, o no
'If InStr(1, l, "CAMPA") Then Stop
For I = 1 To Len(L)
    C = Mid(L, I, 1)
    J = Asc(C)
    If J > 125 Then
        'Caracteres especiales
        Select Case J
        Case 165
            C = "Ñ"
        Case 166
            C = "ª"
        Case 167
            C = "/"
        Case 179
            C = "|"
        Case 191, 192, 193, 194, 196
            C = "-"
        Case 217, 218 ' Estas son las esquinas
            C = "-"
        End Select
    End If
    l2 = l2 & C
Next I
ProcesaLinea2 = l2
End Function


Public Function devuelvePATH(NomPath As String) As String
Dim I
Dim CADENA
Dim Cad2 As String
Cad2 = NomPath
Do
    I = InStr(1, Cad2, "/")
    If I > 0 Then
        CADENA = Mid(Cad2, 1, I - 1)
        Cad2 = CADENA & "\" & Mid(Cad2, I + 1)
    End If
Loop Until I = 0
devuelvePATH = Cad2
End Function

Private Sub Form_Load()
    Label2.Caption = ""
End Sub


Private Function CambiarNombre(ByRef Nombre As String) As Boolean
Dim C As String
Dim Cont As Integer
     On Error GoTo ECambiarNombre
     CambiarNombre = False
     C = vConfig.PathArchivos & "\" & Nombre
     Cont = InStr(1, Nombre, ".")
     Nombre = Mid(Nombre, 1, Cont) & "ira"
     Name C As vConfig.PathArchivos & "\" & Nombre
    
     CambiarNombre = True
ECambiarNombre:
    If Err.Number <> 0 Then MsgBox "Cambiar nombre: " & Err.Description, vbExclamation
End Function


Private Function InsertarCarpetas() As Boolean
Dim Cad As String
Dim RT As ADODB.Recordset
Dim Car As Ccarpetas
Dim codcarpeta As Integer

    
    InsertarCarpetas = False
    Set RT = New ADODB.Recordset
    
'    cad = "Select nombre from carpetas where codcarpeta=1"
'    miRsAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    C = miRsAux!Nombre
'    miRsAux.Close
'
    
    'Leo la carpeta RAIZ
    Set Car = New Ccarpetas
    If Car.Leer(1, False) = 1 Then
        MsgBox "Error leyendo carpeta RAIZ", vbExclamation
        Exit Function
    End If
    
    
    
    Cad = "Select carpeta from tmpintegra where  codusu = " & CodPC & " GROUP BY carpeta"
    RT.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
        Cad = RT.Fields(0)
        codcarpeta = ComprobarCarpeta(Cad, Car)
        If codcarpeta < 1 Then
            RT.Close
            Exit Function
        Else
            Cad = "UPDATE tmpIntegra SET codcarpeta =" & codcarpeta & " WHERE carpeta ='" & DevNombreSql(RT.Fields(0)) & "'"
            Conn.Execute Cad
        End If
        
        RT.MoveNext
    Wend
    RT.Close
    InsertarCarpetas = True
    
End Function


Private Function ComprobarCarpeta(Carpeta As String, ByVal C1 As Ccarpetas) As Integer
Dim I As Integer
Dim Cad1 As String
Dim Rs As ADODB.Recordset
Dim OK As Boolean
Dim C2 As Ccarpetas
Dim C As String
    ComprobarCarpeta = -1
    Cad1 = ""
   

    If Mid(Carpeta, 1, 1) = "\" Then Carpeta = Mid(Carpeta, 2)
    
    OK = False
    Set Rs = New ADODB.Recordset
    While Carpeta <> ""
        I = InStr(1, Carpeta, "\")
        If I > 0 Then
                Cad1 = Mid(Carpeta, 1, I - 1)
                Carpeta = Mid(Carpeta, I + 1)
        
        Else
            'Es la ultima
            Cad1 = Carpeta
            Carpeta = ""
        End If
        
        'Buscamos la carpeta
        C = "Select * from Carpetas where nombre='" & LCase(Cad1) & "' AND padre = " & C1.codcarpeta
        Rs.Open C, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        
            
        
        
        Set C2 = New Ccarpetas
        OK = True
        If Rs.EOF Then
            C2.Almacen = C1.Almacen
            C2.escriturag = C1.escriturag
            C2.escriturau = C1.escriturau
            C2.groupprop = C1.groupprop
            C2.lecturag = C1.lecturag
            C2.lecturau = C1.lecturau
            C2.Nombre = Cad1
            C2.padre = C1.codcarpeta
            C2.userprop = C1.userprop
            If C2.Agregar = 1 Then OK = False
            NuevasCarpetasCreadas = True
            
            
        Else
            If C2.Leer(Rs!codcarpeta, False) = 1 Then OK = False
            
        End If
        Rs.Close
        
        If Not OK Then
            'Se ha producido un error
            Exit Function
        End If
        
        'padre = C2.codcarpeta
        Set C1 = C2
        Set C2 = Nothing
    Wend
    
    ComprobarCarpeta = C1.codcarpeta
    
End Function




Private Sub UpdatearTablaCarpetasCreadas()
Dim Cad As String
    On Error GoTo EUpdatearTablaCarpetasCreadas
    
    Cad = "update actualiza set fecha=concat(curdate(), ' ' , curtime());"
    Conn.Execute Cad

    Exit Sub
EUpdatearTablaCarpetasCreadas:
    Cad = Err.Description
    Cad = "Falta tabla de lectura ""actualiza"": " & vbCrLf & vbCrLf & Cad
    MsgBox Cad, vbQuestion
End Sub
