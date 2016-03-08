VERSION 5.00
Begin VB.Form frmPruebasVerf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pruebas"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEliminaDatosNueva 
      Caption         =   "Borrar datos (Tim,car,alma.)"
      Height          =   495
      Left            =   3480
      TabIndex        =   19
      Top             =   3840
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Generar ESTR"
      Height          =   495
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Text            =   "3"
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Importar ARIDOC OLD"
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   4800
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CANCELAR"
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   4800
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ARIDOC"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Text            =   "C:\Datos\Aridoc\BDatos\BDImagen.mdb"
      Top             =   1680
      Width           =   4575
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Text            =   "root"
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Text            =   "C:\Programas\Aridoc4"
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Text            =   "raiz"
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\Datos\Aridoc\Raiz"
      Top             =   240
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Subir estructura"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "Alamacen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1680
      TabIndex        =   17
      Top             =   2280
      Width           =   840
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6240
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label3 
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "BD Aridoc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   870
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6240
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label5 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   660
   End
   Begin VB.Label Label4 
      Caption         =   "PATH aridoc.exe"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Carpeta en Destino  en Gestion Documental"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Carpeta en dico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   1365
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   6135
   End
End
Attribute VB_Name = "frmPruebasVerf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Activo As Boolean
Dim ConErrores As Boolean
Dim Fss, F

Private ParaElShell As String

Private ErrorGlobal As Boolean
Private FINALIZAR As Boolean


Private codcarpeta As Long
Private TamañoParaEspera As Long
Private NumeroArchivos As Integer


Dim ContadorSegundos As Integer
Dim varExtensiones As String

Dim vCar As Ccarpetas



Private Sub cmdEliminaDatosNueva_Click()
Dim Cad As String

    
    
    Cad = "Va a leiminar los datos de las tabasl de insercion: " & vbCrLf
    Cad = Cad & "-Timagen" & vbCrLf & "-Carpetas" & vbCrLf
    Cad = Cad & "-Extension" & vbCrLf & vbCrLf & "¿Desea continuar?"
    If MsgBox(Cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    If Not AbrirConexionNueva Then Exit Sub
    Label1.Caption = "Eliminando datos"
    Label1.Refresh
    'Elimino
    Cad = "Delete from Carpetas"
    ConnNuevoAridoc.Execute Cad
    
    Cad = "Delete from extension"
    ConnNuevoAridoc.Execute Cad
    
    Cad = "Delete from Timagen"
    ConnNuevoAridoc.Execute Cad
    
    'Cierro
    ConnNuevoAridoc.Close
    Set ConnNuevoAridoc = Nothing
    Label1.Caption = ""
End Sub

Private Sub Command1_Click()
   Text5.Text = ""
    HacerImportar False
End Sub



Private Sub HacerImportar(ImportarAridocEficiente As Boolean)
Dim I As Integer
Dim carpeta As String
Dim mipath As String

    On Error GoTo E1
   
    Dir Text1.Text   'Para ver si es real
    

    
    
    If Dir(Text5.Text, vbArchive) = "" Then
        MsgBox "base datos NO encontrada", vbExclamation
        Exit Sub
    End If
    
    I = InStrRev(Text1.Text, "\")
    carpeta = Mid(Text1.Text, I + 1)
    mipath = Mid(Text1.Text, 1, I - 1)
    
    'Esperamos antes de seguir
    If Not ImportarAridocEficiente Then
    
        If Dir(Text3.Text) <> "" Then
            MsgBox "Error carpeta aridoc", vbExclamation
            Exit Sub
        End If
    
    
    
    
    
        I = 0
        Do
            Activo = FlagActivo("iniciando")
            I = I + 1
            If I > 3 Then
                If MsgBox("El FLAG esta activo. O no se ha cerrado una instancia anterior o bien hay problemas. ¿CONTINUAR?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
                I = 0
            End If
        Loop Until Not Activo
        
        
        OpcionErrores True
        
        'Eliminamos, por si quedara residente, el archivo de intercambio
        ArchivoIntercambio True
    
    End If
    
    ErrorGlobal = False
    
    FINALIZAR = False
    TamañoParaEspera = 0
    NumeroArchivos = 0
    Set Fss = CreateObject("Scripting.FileSystemObject")

    Me.Command3.Visible = True
    
    
    If ImportarAridocEficiente Then
        ImportacionEficiente
    
    Else
        MontajeCarpetasArchivos carpeta, mipath, Text2.Text
        MsgBox "FIN"
    End If
    
    Me.Command3.Visible = False
    Label1.Caption = ""
    
    Exit Sub
E1:
    MsgBox Err.Description
End Sub

Private Sub KillFlag()
On Error Resume Next
    If Dir(Text3.Text & "\Flag.txt") <> "" Then Kill Text3.Text & "\Flag.txt"
    Err.Clear
End Sub


Private Function FlagActivo(Lugar As String) As Boolean
    If Dir(Text3.Text & "\Flag.txt") <> "" Then
        'Espera intervalo
        Label1.Caption = "FLAG. -" & Lugar

        FlagActivo = True
        espera 1
        ContadorSegundos = ContadorSegundos + 1
    Else
        Label1.Caption = Lugar
        FlagActivo = False
    End If
    
    Me.Refresh
End Function

Private Function SeHanproducidoErrores(Eliminar As Boolean) As Boolean
    

    If Eliminar Then
        'ELIMINAR
        If Dir(Text3.Text & "\ErrorShell.txt") <> "" Then Kill Text3.Text & "\ErrorShell.txt"
    
    Else
        'LEEER
        If Dir(Text3.Text & "\ErrorShell.txt") <> "" Then
            SeHanproducidoErrores = True
        Else
            SeHanproducidoErrores = False
        End If
    End If
End Function



Private Function ArchivoIntercambio(Eliminar As Boolean) As Long
Dim NF As Integer
Dim C As String

    If Eliminar Then
        If Dir(Text3.Text & "\InterShell.txt") <> "" Then Kill Text3.Text & "\InterShell.txt"

    Else
        'Leer.
        ArchivoIntercambio = -1
        C = Text3.Text & "\InterShell.txt"
        If Dir(C, vbArchive) <> "" Then
            NF = FreeFile
            Open C For Input As #NF
            Line Input #NF, C
            Close #NF
            If Val(C) > 0 Then ArchivoIntercambio = Val(C)
            Kill Text3.Text & "\InterShell.txt"
        End If
    End If
    espera 0.2
End Function



Private Sub OpcionErrores(Eliminar As Boolean)
    If Dir(Text3.Text & "\ErrorShell.txt") <> "" Then Kill Text3.Text & "\ErrorShell.txt"
End Sub


Private Function MontajeCarpetasArchivos(carpeta As String, vPath As String, CarpetaAridoc As String)
Dim FC, FCC, FCCC

    If FINALIZAR Then Exit Function

    'Montamos la carpeta
    If CrearCarpeta(carpeta, CarpetaAridoc) Then
        ErrorGlobal = True
        Exit Function
    End If
    
    Me.Refresh
    If FINALIZAR Then Exit Function
    
    
    'Si ha creado bien la carpeta, meto los archivos
    InsertarArchivos vPath & "\" & carpeta, CarpetaAridoc & "\" & carpeta
    
    
    
    If FINALIZAR Then Exit Function
    
    
    
    'Para cada subcarpeta, meto los archivos
    Set FC = Fss.getfolder(vPath & "\" & carpeta)
    Set FCC = FC.SubFolders
    For Each FCCC In FCC
        If FINALIZAR Then Exit Function
        MontajeCarpetasArchivos FCCC.Name, vPath & "\" & carpeta, CarpetaAridoc & "\" & carpeta
    Next
    Set FC = Nothing
    Set FCC = Nothing
    Set FCCC = Nothing
    
    
End Function



Private Function CrearCarpeta(Nombre As String, CarpetaAri As String) As Boolean

    ArchivoIntercambio True

    CrearCarpeta = True
    
    'Ya NO esta el flag
    ParaElShell = """" & Text3.Text & "\Aridoc.exe"" /F "
    'USUARIO
    ParaElShell = ParaElShell & Text4.Text & " """
    'nombre carpeta
    ParaElShell = ParaElShell & Nombre & """ """
    'Donde cuelga
    ParaElShell = ParaElShell & CarpetaAri & """ "
    
    Shell ParaElShell
    
    'Comprobar flag
    Do
        Activo = FlagActivo("Creando carpeta: " & Nombre)
    Loop Until Not Activo
    
   
    
    'Si ha terminado el flag
    'Vemos si hay errores
    If SeHanproducidoErrores(False) Then
        
        FINALIZAR = True
    Else
        
        
    End If
    
    
    'Si ha llegado bien codigo carpeta
    codcarpeta = ArchivoIntercambio(False)
    If codcarpeta < 1 Then
        MsgBox "Error obteniendo carpeta: " & CarpetaAri, vbExclamation
        FINALIZAR = True
    Else
        CrearCarpeta = False
    End If
    
End Function

Private Sub InsertarArchivos(CarpetaLocal As String, CarpetaAridoc As String)
Dim F1, FC
Dim IdentificacionAridoc As Long
Dim Tamanyo As Long
Dim I As Integer
Dim IDAridocAntiguo As Long

    Set F = Fss.getfolder(CarpetaLocal)
    
    
    Set FC = F.Files
    For Each F1 In FC
        If FINALIZAR Then Exit Sub
        
        If Dir(CarpetaLocal & "\" & F1.Name, vbArchive) <> "" Then
           IDAridocAntiguo = -1
           Tamanyo = FileLen(CarpetaLocal & "\" & F1.Name)
            
           IdentificacionAridoc = Insertar1Archivo(CarpetaLocal & "\" & F1.Name, CarpetaAridoc)
           Me.Refresh

           If IdentificacionAridoc > 0 Then
                'AQUI ASOCIARIAMOS LA ACCION
                I = InStr(1, F1.Name, ".")
                If I > 0 Then IDAridocAntiguo = Val(Mid(F1.Name, 1, I - 1))
            
               'Lo qu vaya aqui
               If Text5.Text <> "" Then HACERUPDATE IdentificacionAridoc, Tamanyo, IDAridocAntiguo
           Else
                MsgBox "Error con el fichero: " & F1.Name, vbExclamation
           End If
           
           NumeroArchivos = NumeroArchivos + 1
           If NumeroArchivos > 100 Then
                Label1.Caption = "Espera carga datos(Numero)"
                Label1.Refresh
                TamañoParaEspera = 0
                espera 2.5
                NumeroArchivos = 0
                Me.Refresh
                Me.SetFocus
           End If
           
           
           TamañoParaEspera = TamañoParaEspera + Tamanyo
           If TamañoParaEspera > 1500000 Then
                Label1.Caption = "Espera carga datos(Tamaño)"
                Label1.Refresh
                TamañoParaEspera = 0
                espera 3.5
           End If
        End If
        DoEvents
    Next
    Set F1 = Nothing
    Set FC = Nothing
    Set F = Nothing
End Sub



Private Function Insertar1Archivo(CL As String, CA As String) As Long
Dim T1 As Single


    T1 = Timer
    Insertar1Archivo = -1
    'Ya NO esta el flag
    ParaElShell = """" & Text3.Text & "\Aridoc.exe"" /N "
    'USUARIO
    ParaElShell = ParaElShell & Text4.Text & " """
    'nombre carpeta
    ParaElShell = ParaElShell & CL & """ """
    'Donde cuelga
    ParaElShell = ParaElShell & CA & """ "
    
    Shell ParaElShell, vbNormalFocus
    
    'Comprobar flag
    ContadorSegundos = 0
    Do
        Activo = FlagActivo("Subiendo archivo : " & CL)
        If ContadorSegundos > 30 Then KillFlag
            
    Loop Until Not Activo
    
    
    'Si ha terminado el flag
    'Vemos si hay errores
    If SeHanproducidoErrores(True) Then
        'ARHCIOV NO SUBIDP
        'Mato el fichero intercambio, por si acaso
        ArchivoIntercambio True
    Else
        'Destripamos el archivo de intercambio
        Insertar1Archivo = ArchivoIntercambio(False)
        
        
        espera 1
        
        
    End If
    
    SeHanproducidoErrores True
    
    
    
End Function



'SUBIR ARIDOC
'---------------------------------------

Private Sub Command2_Click()

    If MsgBox("Seguro que desea seguir con la importacin desde el antiguo ARIDOC?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub

    'Comprobaremos que la BD esta aquin, y que es de aridoc
    If Text5.Text = "" Then
        MsgBox "Debe introducir la BD de ARIDOC.", vbExclamation
        Exit Sub
    End If
    If Dir(Text5.Text, vbArchive) = "" Then
        MsgBox "BD aridoc mdb  NO existe", vbExclamation
        Exit Sub
    End If
    
    
    'Abrimos la antigua
    Label1.Caption = "Abriendo conexiones BD's"
    Label1.Refresh
    If AbrirConexionAntigua Then
        'Abrir conexion Nueva aridoc
        If AbrirConexionNueva Then
            'OK, conexion nueva abierta
            
            HacerImportar False
    
        Else
            
            
        End If
    Else
    
    End If
    Label1.Caption = ""
    Set ConnAntiguoAridoc = Nothing
    Set ConnNuevoAridoc = Nothing
End Sub




Public Function AbrirConexionNueva() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionNueva = False
    Set ConnNuevoAridoc = Nothing
    Set ConnNuevoAridoc = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnNuevoAridoc.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    
    
    
    'cadenaconexion
    Cad = "DSN=Aridoc;DESC= DSN;DATABASE=aridoc;;;PORT=;OPTION=;STMT=;"
    Cad = "DSN=Aridoc;"" "
    'cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=" & vUsu.CadenaConexion & ";SERVER=" & vConfig.SERVER & ";"
    'cad = cad & ";UID=" & vConfig.User
    'cad = cad & ";PWD=" & vConfig.password
    
    
    ConnNuevoAridoc.ConnectionString = Cad
    ConnNuevoAridoc.Open

    AbrirConexionNueva = True
    Exit Function
EAbrirConexion:
    MsgBox "Abrir conexión nueva." & Err.Description, vbExclamation
End Function

'
Public Function AbrirConexionAntigua() As Boolean
Dim Cad As String
On Error GoTo EAbrirConexion

    
    AbrirConexionAntigua = False
    Set ConnAntiguoAridoc = Nothing
    Set ConnAntiguoAridoc = New Connection
    'Conn.CursorLocation = adUseClient   'Si ponemos este hay opciones k no van ej select con rs!campo
    ConnAntiguoAridoc.CursorLocation = adUseServer   'Si ponemos esta alguns errores de Conn no se muestran correctamente
                        
    
    
    
    'cadenaconexion
    Cad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Text5.Text
    Cad = Cad & ";Persist Security Info=False"
    
    
    
    ConnAntiguoAridoc.ConnectionString = Cad
    ConnAntiguoAridoc.Open


    'Esta abierta
    If ProbarRegistroAntigua Then
        AbrirConexionAntigua = True
    Else
        ConnAntiguoAridoc.Close
    End If
    
    
    
    Exit Function
EAbrirConexion:
    MsgBox "AbrirConexionAntigua." & Err.Description, vbExclamation
End Function


Private Function ProbarRegistroAntigua() As Boolean
Dim RS As ADODB.Recordset

    On Error GoTo EProbarRegistroAntigua

    ProbarRegistroAntigua = False
    

    Set RS = New ADODB.Recordset
    RS.Open "select count(*) from timagen", ConnAntiguoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    ParaElShell = "NO"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then ParaElShell = ""
        End If
    End If
    RS.Close
    
    
    If ParaElShell <> "" Then
        MsgBox "VACIA", vbExclamation
        Exit Function
    End If
    
    
    'AHora muestro el mensaje para las extensiones
    RS.Open "select * from textension", ConnAntiguoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    ParaElShell = ""
    varExtensiones = "|"
    While Not RS.EOF
        varExtensiones = varExtensiones & LCase(RS!extension) & "|"
        ParaElShell = ParaElShell & RS!Cod & "    " & RS!extension & "     " & RS!Descripcion & vbCrLf
        RS.MoveNext
    Wend
    RS.Close
    
    If ParaElShell = "" Then
        MsgBox "Extensiones incorrectas"
    Else
        ParaElShell = ParaElShell & vbCrLf & vbCrLf & "Deberian coincidir con las nuevas del ARIDOC" & vbCrLf & vbCrLf & " ¿Desea continuar?"
'        If MsgBox(ParaElShell, vbQuestion + vbYesNoCancel) = vbYes Then
'            ProbarRegistroAntigua = True
'        End If
         ProbarRegistroAntigua = True
    End If
    
EProbarRegistroAntigua:
    If Err.Number <> 0 Then Err.Clear
    Set RS = Nothing
End Function



Private Sub HACERUPDATE(codigo As Long, Tamanyo As Long, AntiguoID As Long)
Dim Aux As String
Dim RS As ADODB.Recordset
Dim C As String

    
    
'Private ConnAntiguoAridoc As Connection
'Private ConnNuevoAridoc As Connection
    Aux = "Select * from TIMAGEN where id =" & AntiguoID
    Set RS = New ADODB.Recordset
    RS.Open Aux, ConnAntiguoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    'Cadena insert
    C = "INSERT INTO timagen (campo4,fecha3, importe1, importe2,codcarpeta,tamnyo,"
    C = C & "userprop, groupprop, lecturau, lecturag, escriturau, escriturag, bloqueo,"
    C = C & "codigo, codext, campo1, campo2, campo3, "
    C = C & "fecha1, fecha2,observa)"
    
    'Primeros valores
    C = C & " VALUES (NULL,NULL,NULL,NULL," & codcarpeta & ","
    'Tamanyo
    C = C & TransformaComasPuntos(CStr(Round((Tamanyo / 1024), 3)))
    'Permisos
    '2147483647
    C = C & ",0,1,0,2147483647,0,0,0,"
    
    
    ConnNuevoAridoc.Execute "DELETE from Timagen where codigo = " & codigo
    
    
    
    If RS.EOF Then
        
        'NO HA ENCONTRADO LA REFERENCIA EN LA BD
        Aux = codigo & ",-1," & codigo & ",'NULL','ERROR','" & Format(Now, "yyyy-mm-dd")
        Aux = Aux & "',NULL,NULL,NULL)"
    Else
        
    
        ' codigo, codext, campo1, campo2, campo3, "
        ' fecha1, fecha2,observa)"
        
        Aux = codigo & "," & RS!extension & ",'" & DevNombreSQL(RS!clave1) & "',"
        If IsNull(RS!clave2) Then
            Aux = Aux & "NULL"
        Else
            Aux = Aux & "'" & Trim(DevNombreSQL(RS!clave2)) & "'"
        End If
        Aux = Aux & ","
        If IsNull(RS!clave3) Then
            Aux = Aux & "NULL"
        Else
            Aux = Aux & "'" & Trim(DevNombreSQL(RS!clave3)) & "'"
        End If
        Aux = Aux & ","
        Aux = Aux & "'" & Format(RS!FechaDig, "yyyy-mm-dd") & "'"
        Aux = Aux & ","
        If IsNull(RS!fechadoc) Then
            Aux = Aux & "NULL"
        Else
            Aux = Aux & "'" & Format(RS!fechadoc, "yyyy-mm-dd") & "'"
        End If
        Aux = Aux & ","
        If IsNull(RS!des) Then
            Aux = Aux & "NULL"
        Else
            Aux = Aux & "'" & DevNombreSQL(RS!des) & "'"
        End If
        Aux = Aux & ")"
        
        
        
        
        
        
        
    End If
    Aux = C & Aux
    ConnNuevoAridoc.Execute Aux
End Sub

Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"
                    DBLet = ""
                Case "N"
                    DBLet = 0
                Case "F"
                    DBLet = "0:00:00"
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function

Private Function TransformaComasPuntos(CADENA As String) As String
    Dim I As Integer
    Do
        I = InStr(1, CADENA, ",")
        If I > 0 Then
            CADENA = Mid(CADENA, 1, I - 1) & "." & Mid(CADENA, I + 1)
        End If
        Loop Until I = 0
    TransformaComasPuntos = CADENA
End Function



Private Function DevNombreSQL(CADENA As String) As String
Dim J As Integer
Dim I As Integer
Dim Aux As String
    J = 1
    Do
        I = InStr(J, CADENA, "'")
        If I > 0 Then
            Aux = Mid(CADENA, 1, I - 1) & "\"
            CADENA = Aux & Mid(CADENA, I)
            J = I + 2
        End If
    Loop Until I = 0
    DevNombreSQL = Trim(CADENA)
End Function

Private Sub Command3_Click()
    FINALIZAR = True
End Sub


Private Sub Command4_Click()
    If MsgBox("El proceso puede llevar mucho tiempo. ¿Continuar?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    
    'LO QUE HAREMOS SERA PARECIDO a subir estructura pero MAS eficiente
    '------------------------------------------------------------------
    '------------------------------------------------------------------
    'para ello cuando llamemos a importar:
    '
    '       1: Crearemos una tabla en BDaridocOLD, donde insertaremos los archivos
    '          , de que extension son...
    '
    '       2: Verificaremos la estructura , y que todas los archivos tienen
    '           referencia en la BD, no existen duplicados etc etc
    '
    '
    '       3:  Subiremos los archivos con mput, es decir, no uno a uno, si no carpetas a la vez
    
    HacerImportar True
End Sub



Private Sub ImportacionEficiente()
Dim Donde As String
    If Not AbrirConexionAntigua Then Exit Sub
    
    
    
    If Not AbrirConexionNueva Then Exit Sub
    
    
    
    'Comprobacioines iniciales de ARIDOC nuevo.
    ' Almamacen ...
    If Not ComprobacionesInicialeNuevo Then Exit Sub
    
    'Creo tabla
    If Not CrearTablaEspecial Then Exit Sub
    
    
    'Insertar extensiones
    If Not InsertarExtensiones Then Exit Sub
    
    'Recorrer carpetas y demas insertando
    If Not RecorrerCarpetasComprobacion Then Exit Sub
        
    'Ver si todos los archivos existe, y si referencias en BD no tiene fisico
    If Not ComprobarReferenciasEfectiva Then Exit Sub
    
    If FINALIZAR Then Exit Sub
    
    
    'Ahora empieza la fiesta de verdad
    On Error GoTo 0
    On Error GoTo E12
    
    '1º Creamos todas las carpetas
    If Not CrearCarpetasNuevo Then Exit Sub
    
    'Una vez creadas las carpetas, iremos recorriendolas otra vez, pero esta vez haciendo
    'iremos copiando los rchivos 1 a 1 dios mio
    'Para ello, veremos que carpetas contienen archivos y ls iremos copiando
    HabilitarBotones False
    PasarArchivos
    If FINALIZAR Then
        MsgBox "************************" & vbCrLf & "PROCESO NO FINALIZADO BIEN", vbExclamation
    Else
        MsgBox "Proceso finalizado correctamente", vbInformation
    End If
    HabilitarBotones True
    Exit Sub
E12:
    MsgBox Donde & vbCrLf & Err.Description

End Sub


Private Sub PasarArchivos()
Dim C As String
Dim RT As ADODB.Recordset

    SeHaCancelado = False
    Set RT = New ADODB.Recordset
    C = "Select micarpeta.codigo,path from micarpeta,tablaespecial WHERE micarpeta.codigo=tablaespecial.codcarpeta GROUP BY micarpeta.codigo,path"
    RT.Open C, ConnAntiguoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set vCar = New Ccarpetas
    While Not RT.EOF
            Label1.Caption = RT!Path
            Label1.Refresh
            DoEvents
            If FINALIZAR Then
                RT.Close
                Exit Sub
            End If
            If vCar.Leer(RT!codigo, False) = 0 Then
                C = RT!Path
                frmvFTP.opcion = 1
                Set frmvFTP.CarpetaD = vCar
                frmvFTP.Origen = C
                frmvFTP.Show vbModal
                If SeHaCancelado Then FINALIZAR = True
            Else
                MsgBox "Error leyendo datos carpeta: " & RT!codigo
            End If
            RT.MoveNext
    Wend
    RT.Close
    
End Sub


Private Function CrearCarpetasNuevo() As Boolean
    Dim C As Ccarpetas
    Dim RS As ADODB.Recordset
    CrearCarpetasNuevo = False
    Set RS = New ADODB.Recordset
    RS.Open "select * from miCarpeta", ConnAntiguoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    Set C = New Ccarpetas
    While Not RS.EOF
        With C
            .Almacen = Text6.Text
            .codcarpeta = RS!codigo
            .escriturag = vbPermisoTotal
            .escriturau = vbPermisoTotal
            .groupprop = 1  'Administradores aridoc
            .lecturag = vbPermisoTotal
            .lecturau = vbPermisoTotal
            .Nombre = RS!Nombre
            .padre = RS!padre
            .userprop = 0 'root
            If .Agregar = 1 Then
                Set C = Nothing
                Exit Function
            End If
        End With
        RS.MoveNext
    Wend
    RS.Close
    CrearCarpetasNuevo = True
End Function
Private Sub BorrarTablasAuxiliares()
    On Error Resume Next

        ConnAntiguoAridoc.Execute "DROP TABLE TablaEspecial"
        ConnAntiguoAridoc.Execute "DROP TABLE miCarpeta"
        Err.Clear
End Sub


Private Function CrearTablaEspecial() As Boolean
Dim Cad As String

    On Error GoTo ECre
    CrearTablaEspecial = False
        'Las borramos
        BorrarTablasAuxiliares
    
    
        'Y las creamos
        Cad = "CREATE TABLE TablaEspecial(codigo LONG CONSTRAINT MiCampoRestringido PRIMARY KEY,Extension TEXT(3),tamayo  LONG,codcarpeta  LONG);"
        ConnAntiguoAridoc.Execute Cad
        
        Cad = "CREATE TABLE miCarpeta (codigo LONG CONSTRAINT MiCampoRestringido PRIMARY KEY,Padre INTEGER,Nombre TEXT(255),Path TEXT(255))"
        ConnAntiguoAridoc.Execute Cad
        

    CrearTablaEspecial = True
    Exit Function
ECre:
    MsgBox Err.Description, vbExclamation
End Function


'Haciendo dir, lo que haremos insertar cada archivo en la
'en la tabla temporal que hemos creado, asignandole la extension
'Si hubieran archivos sin extension incorrecta o con extension duplicada entonces
Private Function RecorrerCarpetasComprobacion() As Boolean
Dim C As String
Dim C2 As String
    RecorrerCarpetasComprobacion = False
    
    codcarpeta = InStrRev(Text1.Text, "\")
    
    If codcarpeta = 0 Then
        MsgBox "Error obteniendo carpeta en RecorrerCarpetasComprobacion", vbExclamation
        Exit Function
    End If
    C = Mid(Text1.Text, 1, codcarpeta - 1)
    C2 = Mid(Text1.Text, codcarpeta + 1)

    codcarpeta = 0
    
    ComprobarCarpetaEf C, C2, 0
    
    If Not FINALIZAR Then RecorrerCarpetasComprobacion = True
End Function

Private Function ObtenerSiguienteCodigoCarpeta() As Long
    codcarpeta = codcarpeta + 1
    ObtenerSiguienteCodigoCarpeta = codcarpeta
End Function


Private Sub ComprobarCarpetaEf(Ruta As String, carpeta As String, ByVal padre As Long)
Dim Cod As Long
Dim FC, FCC, FCCC
    
    'Para la carpeta actual, inserto en carpeta,y , inserto en archivos
    Cod = ObtenerSiguienteCodigoCarpeta
    'INSERTO LA CARPETA
    InsertarCarpeta carpeta, Cod, padre, Ruta & "\" & carpeta
    Label1.Caption = carpeta
    Label1.Refresh
                
    
    If FINALIZAR Then Exit Sub
    
    
    
    
    'Para cada subcarpeta, meto los archivos
    Set FC = Fss.getfolder(Ruta & "\" & carpeta)
    Set FCC = FC.SubFolders
    
    'Inserto los archivos de esta carpeta
    'InsertarArchivosEf FC.Path, CodCarpeta
    
    For Each FCCC In FCC
        If FINALIZAR Then Exit Sub
        ComprobarCarpetaEf Ruta & "\" & carpeta, FCCC.Name, Cod
        InsertarArchivosEf FCCC.Path, Cod
    Next
    Set FC = Nothing
    Set FCC = Nothing
    Set FCCC = Nothing

    
End Sub



Private Sub InsertarArchivosEf(CarpetaLocal As String, ByVal CCa As Long)
Dim F1, FC
Dim Tamanyo As Long
Dim I As Integer
Dim J As Integer
Dim Ex As String
Dim Cod As Long

    Set F = Fss.getfolder(CarpetaLocal)
    
    
    Set FC = F.Files
    For Each F1 In FC
        Label1.Caption = F1.Name
        Label1.Refresh
                
        If FINALIZAR Then Exit Sub
        I = InStr(1, F1.Name, ".")
        If I = 0 Then
            Kill F1.Path
            'MensajeParaFinalizar "Archivo sin extension:" & F1.Name
        Else
            Cod = Val(Mid(F1.Name, 1, I - 1))
            Ex = LCase(Mid(F1.Name, I + 1))
            If Cod = 0 Then
                MensajeParaFinalizar "Codigo archivo incorrecto:" & F1.Name & " --> Devuelve 0"
            Else
                If InStr(1, varExtensiones, "|" & Ex) = 0 Then
                    MensajeParaFinalizar "Extension NO reconcida:" & F1.Name
                Else
                    Tamanyo = FileLen(F1.Path)
                    'Ahora ya tengo el archivo, su extension, y su carpeta
                    'lo inserto en BD, en tmp
                    InsertarEnArchivoTemporal Cod, Tamanyo, Ex, CCa
                End If
            End If
        End If
        DoEvents
    Next
    Set F1 = Nothing
    Set FC = Nothing
    Set F = Nothing
End Sub


Private Sub InsertarCarpeta(Nombre As String, codigo As Long, padre As Long, mipath As String)
Dim Cad As String
    On Error Resume Next
    Cad = "INSERT INTO miCarpeta(codigo,padre,nombre,path) VALUES (" & codigo & "," & padre & ",'" & Nombre & "','" & mipath & "')"
    ConnAntiguoAridoc.Execute Cad
    If Err.Number <> 0 Then
        Cad = "ERROR: " & Err.Description & vbCrLf & Cad
        MensajeParaFinalizar Cad
        Err.Clear
    End If
End Sub


Private Sub InsertarEnArchivoTemporal(ByRef Cod As Long, ByRef tama As Long, ByRef ext As String, ByRef CodCa As Long)
Dim Cad As String
    On Error Resume Next
    Cad = "INSERT INTO tablaespecial(codigo,extension,tamayo,codcarpeta) VALUES (" & Cod & ",'"
    Cad = Cad & ext & "'," & tama & "," & codcarpeta & ")"
    ConnAntiguoAridoc.Execute Cad
    If Err.Number <> 0 Then
        Cad = "ERROR: " & Err.Description & vbCrLf & Cad
        MensajeParaFinalizar Cad
        Err.Clear
    End If

End Sub


Private Function ComprobarReferenciasEfectiva() As Boolean
Dim RS As ADODB.Recordset
Dim Cad As String

On Error GoTo EComp

    ComprobarReferenciasEfectiva = False

    Label1.Caption = "Fichero fisico --> en BD"
    Label1.Refresh
        
        
    'Borror temporal
    Cad = "DELETE FROM Temporal"
    ConnAntiguoAridoc.Execute Cad
    
    
    Cad = "Select tablaespecial.*,timagen.id,nomfich from tablaespecial left join timagen on tablaespecial.codigo= timagen.id "
    Cad = Cad & " WHERE nomfich is null"
    Set RS = New ADODB.Recordset
    RS.Open Cad, ConnAntiguoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = "INSERT INTO temporal(id,Archivo) VALUES ("
    codcarpeta = 0
    While Not RS.EOF
        Label1.Caption = "Fi-> BD: " & RS!codigo & "." & RS!extension
        Label1.Refresh
        ConnAntiguoAridoc.Execute Cad & RS!codigo & ",'" & RS!codigo & "." & RS!extension & "')"
         
        codcarpeta = codcarpeta + 1
        
        RS.MoveNext
    Wend
    RS.Close
    
    If codcarpeta > 0 Then
        frmVerError.opcion = 0
        frmVerError.Show vbModal
    
        'Existen archivos sin referencias en BD
        'Si dice continuar, deberiamos borrar los datos, para
        Cad = "Han habido " & codcarpeta & " archivo(s) de sin referencia en BD's. Deberia haber sido corregido. "
        MensajeParaFinalizar Cad
    End If
    
    If FINALIZAR Then Exit Function
    
    'Comprobamos al reves, de los que hay en BD cuales NO tienen referencia
    
    Label1.Caption = "BD---> Fichero fisico "
    Label1.Refresh
        
        
    'Borror temporal
    Cad = "DELETE FROM Temporal"
    ConnAntiguoAridoc.Execute Cad
    
    
    Cad = "Select tablaespecial.*,timagen.id,nomfich from timagen left join tablaespecial on tablaespecial.codigo= timagen.id "
    Cad = Cad & " WHERE codigo is null"
    Set RS = New ADODB.Recordset
    RS.Open Cad, ConnAntiguoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = "INSERT INTO temporal(id,Archivo) VALUES ("
    codcarpeta = 0
    While Not RS.EOF
        Label1.Caption = "Fi-> BD: " & RS!nomfich
        Label1.Refresh
        ConnAntiguoAridoc.Execute Cad & RS!id & ",'" & RS!nomfich & "')"
         
        codcarpeta = codcarpeta + 1
        
        RS.MoveNext
    Wend
    RS.Close
    
    If codcarpeta > 0 Then
        frmVerError.opcion = 1
        frmVerError.Show vbModal
    
        'Existen archivos sin referencias en BD
        'Si dice continuar, deberiamos borrar los datos, para
        Cad = "Hay " & codcarpeta & " referencias(s) en la BD's sin archivo fisico. Deberia haber sido revisado."
        MensajeParaFinalizar Cad
    End If
    
    If FINALIZAR Then Exit Function
    
    ComprobarReferenciasEfectiva = True
    
    Exit Function
EComp:
    MsgBox Err.Description, vbExclamation
End Function


Private Sub MensajeParaFinalizar(Mens As String)
    Mens = Mens & vbCrLf & vbCrLf & "¿DESEA CONTINUAR?"
    If MsgBox(Mens, vbQuestion + vbYesNoCancel) <> vbYes Then FINALIZAR = True
End Sub

'Aqui pondremos todo lo que tiene que ver con las comprobaciones inicuales
'del FTP y /o Carpeta en red.
'Si existe el almacen....
Private Function ComprobacionesInicialeNuevo() As Boolean


    ComprobacionesInicialeNuevo = False
    

    
    ComprobacionesInicialeNuevo = True

End Function


Private Function InsertarExtensiones() As Boolean
Dim RS As ADODB.Recordset
Dim C As String
Dim TieneExt As Boolean
    On Error GoTo EInsertarExtensiones

    InsertarExtensiones = False
    Set RS = New ADODB.Recordset
    C = "Select count(*) from extension"
    RS.Open C, ConnNuevoAridoc, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not RS.EOF Then
        If DBLet(RS.Fields(0), "N") > 0 Then C = ""
    End If
    RS.Close
    TieneExt = False
    If C = "" Then
        TieneExt = True
        If MsgBox("Ya existen las extensiones. Desea borrarlas?", vbQuestion + vbYesNo) = vbYes Then
            ConnNuevoAridoc.Execute "DELETE FROM EXtension"
            C = "OK"
        End If
    End If
    
    If C = "" Then
        If MsgBox("Desea continuar de igual modo?", vbQuestion + vbYesNo) = vbNo Then Exit Function
    End If
    
    If Not TieneExt Then
        C = "Select * from textension"
        RS.Open C, ConnAntiguoAridoc, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF
            C = "INSERT INTO extension (codext, descripcion, exten, Modificable, Nuevo, OfertaExe, OfertaPrint, Aparecemenu, Deshabilitada) VALUES ("
            C = C & RS!Cod & ",'" & RS!Descripcion & "','" & RS!extension & "',"
            'Modificable, Nuevo, OfertaExe, OfertaPrint, Aparecemenu, Deshabilitada
            C = C & "0,0,NULL,NULL,0,1)"
            ConnNuevoAridoc.Execute C
            RS.MoveNext
        Wend
        RS.Close
    End If
    InsertarExtensiones = True
    Exit Function
EInsertarExtensiones:
    MsgBox Err.Description
End Function

Private Sub Command5_Click()
    If Not AbrirConexionNueva Then Exit Sub
    frmInsert.Show vbModal
    ConnNuevoAridoc.Close
    Set ConnNuevoAridoc = Nothing
End Sub

Private Sub HabilitarBotones(Si As Boolean)
    Command5.Enabled = Si
    cmdEliminaDatosNueva.Enabled = Si
    Command4.Enabled = Si
End Sub

