VERSION 5.00
Begin VB.Form frmBusca2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Realizar búsqueda"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "frmBusca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8460
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   2085
      Left            =   3360
      Style           =   1  'Checkbox
      TabIndex        =   30
      Top             =   3960
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   120
      TabIndex        =   26
      Top             =   3840
      Width           =   3135
      Begin VB.OptionButton Option1 
         Caption         =   "Todo ARIDOC"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   29
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Carpeta actual"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Carpeta actual y subcarpetas"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   27
         Top             =   960
         Value           =   -1  'True
         Width           =   2655
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   25
      Top             =   5640
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   24
      Top             =   5160
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   8295
      Begin VB.TextBox txtClaves 
         Height          =   315
         Index           =   9
         Left            =   240
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Tag             =   "|T|S|||timagen|Observa|||"
         Text            =   "Text3"
         Top             =   3240
         Width           =   7815
      End
      Begin VB.TextBox txtClaves 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   2760
         TabIndex        =   8
         Tag             =   "|N|S|||timagen|Importe2|||"
         Text            =   "Text3"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtClaves 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   240
         TabIndex        =   7
         Tag             =   "|N|S|||timagen|Importe1|||"
         Text            =   "Text3"
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   6
         Left            =   5520
         TabIndex        =   6
         Tag             =   "|F|S|||timagen|fecha3|||"
         Text            =   "Text3"
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   5
         Left            =   2760
         TabIndex        =   5
         Tag             =   "|F|S|||timagen|fecha2|||"
         Text            =   "Text3"
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   4
         Tag             =   "|F|S|||timagen|fecha1|||"
         Text            =   "99/99/9999"
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   3
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   3
         Tag             =   "|T|S|||timagen|Campo4|||"
         Text            =   "Text3"
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   2
         Left            =   240
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "|T|S|||timagen|Campo3|||"
         Text            =   "Text3"
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   1
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   1
         Tag             =   "|T|S|||timagen|Campo2|||"
         Text            =   "Text3"
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtClaves 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   5520
         TabIndex        =   9
         Tag             =   "|N|S|||timagen|Tamnyo|||"
         Text            =   "Text3"
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   0
         Left            =   240
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "|T|S|||timagen|Campo1|||"
         Text            =   "Text3"
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   23
         Top             =   3000
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Importe"
         Height          =   255
         Index           =   8
         Left            =   2760
         TabIndex        =   22
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Importe"
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   6
         Left            =   5520
         TabIndex        =   20
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   19
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   3
         Left            =   4320
         TabIndex        =   17
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   1
         Left            =   4320
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Tamaño"
         Height          =   255
         Index           =   10
         Left            =   5520
         TabIndex        =   13
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Añadir final archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   480
         Width           =   1755
      End
   End
   Begin VB.Menu mnPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnSelectAll 
         Caption         =   "Selecionar todos"
      End
      Begin VB.Menu mnQuitarSeleccion 
         Caption         =   "Quitar seleccion"
      End
   End
End
Attribute VB_Name = "frmBusca2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Carpetas As String   ' La primera sera la carpeta ppal, a partir de ahi, las subcarpetas
Public TodasCarpetas As String
Public DesdeEmail As Boolean
Dim SQL As String
Dim Tabla As String

Private Sub Command1_Click(Index As Integer)
Dim OK As Boolean

    If Index = 1 Then
        DatosCopiados = ""
        Unload Me
        Exit Sub
    End If
    
    SQL = ObtenerBusqueda(Me)
    If SQL = "" Then
        MsgBox "Ponga algun campo de busqueda", vbExclamation
        Exit Sub
    End If
    
    
    
    'Llegados aqui, hay SQL, y nos falta ver
    Screen.MousePointer = vbHourglass
    BorrarTemporal2 '-> Busquedas
    If Option1(0).Value Then
        'Buscar en caparpeta actual
        OK = BusquedaCarpetaActual
    Else
    
        
        If Option1(1).Value Then
            'Carpeta  y subcarpetas
           OK = BusquedaSubcarpeta(0)
        Else
        
            'Buscar en todo el arbol de directorios
            OK = BusquedaSubcarpeta(1)
        End If
    End If
    
    If OK Then
        frmResultados2.DesdeEmail = DesdeEmail
        frmResultados2.Show vbModal
        If DatosCopiados <> "" Then Unload Me
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Ningun dato encontrado con esos datos", vbExclamation
        
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
    Me.txtClaves(0).SetFocus
End Sub

Private Sub Form_Load()
    Limpiar Me
    PonerLabels
    CargaList
    
End Sub

Private Sub PonerLabels()
Dim i As Integer
    'En funcion de la configuracion pondremos los labels
    'el texto k keramos
    Me.Label3(0).Caption = vConfig.C1
    Me.Label3(1).Caption = vConfig.C2
    Me.Label3(2).Caption = vConfig.c3
    Me.Label3(3).Caption = vConfig.c4
    Me.Label3(4).Caption = vConfig.f1
    Me.Label3(5).Caption = vConfig.f2
    Me.Label3(6).Caption = vConfig.f3
    Me.Label3(7).Caption = vConfig.imp1
    Me.Label3(8).Caption = vConfig.imp2
    Me.Label3(9).Caption = vConfig.obs
    For i = 0 To 10
        Debug.Print txtClaves(i).Tag
        txtClaves(i).Tag = Me.Label3(i).Caption & txtClaves(i).Tag
    Next i
End Sub



Private Function BusquedaCarpetaActual() As Boolean
Dim i As Long


    'Que hacemos?
    'Pues metemos los archivos con codcarpeta = CARPETA
    '
    If ModoTrabajo <> vbNorm Then
        Tabla = "hco"
    Else
        Tabla = ""
    End If

    SQL = "Select codigo,codcarpeta from timagen" & Tabla & " where " & SQL
    SQL = SQL & " AND codcarpeta = " & RecuperaValor(Me.Carpetas, 1)
    
    'Ponemos si esta seleccionadndo un tipo de archivo
    PonerTipoArchivo SQL
    
    
    'permisos
    If vUsu.codusu > 0 Then
        'Para niveles lectura
        SQL = SQL & " AND (userprop = " & vUsu.codusu
    
        'O el grupo tiene permiso
        SQL = SQL & " OR lecturag & " & vUsu.Grupo & ")"
        
    End If
    
    
    i = 0
    BusquedaCarpetaActual = False
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        InsertaBusqueda miRSAux!codigo, miRSAux!codcarpeta
        miRSAux.MoveNext
       i = i + 1
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    If i > 0 Then BusquedaCarpetaActual = True
        
End Function


Private Function BusquedaSubcarpeta(Opcion As Byte) As Boolean
Dim i As Long
Dim N As Integer
Dim cad As String
Dim Contador As Integer
Dim Aux As String

    'Dividiremos en varias la busqueda. Esto es
    'En funcion del numero de subcarpetas realizaremos n veces el proceso
    BusquedaSubcarpeta = False
    i = 1
    Contador = 0
    If Opcion = 0 Then
        'Subcarpetas
        
        Do
            N = InStr(i, Me.Carpetas, "|")
            If N > 0 Then
                Contador = Contador + 1
                i = N + 1
            End If
        Loop Until N = 0
    Else
        'TODO
        Do
            N = InStr(i, Me.TodasCarpetas, "|")
            If N > 0 Then
                Contador = Contador + 1
                i = N + 1
            End If
        Loop Until N = 0
    End If
             
    'Ya tenemos contador. Haremos de quince en quince
    i = 1
    N = 0
    
    Do
        If Opcion = 0 Then
            Aux = RecuperaValor(Me.Carpetas, CInt(i))
        Else
            Aux = RecuperaValor(Me.TodasCarpetas, CInt(i))
        End If
        i = i + 1
        If Aux <> "" Then
            If cad <> "" Then cad = cad & ","
            cad = cad & Aux
            N = N + 1
        End If
        
        If N > 15 Then
            'Lanzamos la busqueda
            If LanzarBusqueda(cad) Then BusquedaSubcarpeta = True
            cad = ""
            N = 0
        End If
    Loop Until i > Contador
    
    If N > 0 Then
        'lanzamos la busqueda
        If LanzarBusqueda(cad) Then BusquedaSubcarpeta = True
    End If
    
        
End Function

Public Function LanzarBusqueda(CADENA As String)
Dim i As Long
Dim vSQL As String

    'Que hacemos?
    'Pues metemos los archivos con codcarpeta = CARPETA
    '
    If ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt Then
        Tabla = "hco"
    Else
        Tabla = ""
    End If
    Tabla = Tabla & " as timagen "
    vSQL = "Select codigo,codcarpeta from timagen" & Tabla & " where " & SQL
    vSQL = vSQL & " AND codcarpeta in ( " & CADENA & ")"
    
    'Ponemos si esta seleccionadndo un tipo de archivo
    PonerTipoArchivo vSQL
    
    
    If vUsu.codusu > 0 Then
        'Para niveles lectura
        vSQL = vSQL & " AND (userprop = " & vUsu.codusu
    
        'O el grupo tiene permiso
        vSQL = vSQL & " OR lecturag & " & vUsu.Grupo & ")"
        
    End If
    
    i = 0
    LanzarBusqueda = False
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        InsertaBusqueda miRSAux!codigo, miRSAux!codcarpeta
        miRSAux.MoveNext
       i = i + 1
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    If i > 0 Then LanzarBusqueda = True

End Function

Private Sub CargaList()
    List1.Clear
    SQL = "Select * from extension "
    SQL = SQL & " WHERE Deshabilitada=0"
    SQL = SQL & " order by descripcion"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        List1.AddItem miRSAux!Descripcion
        List1.ItemData(List1.NewIndex) = miRSAux!codext
        List1.Selected(List1.NewIndex) = True
    
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
End Sub


Private Sub PonerTipoArchivo(mSQL As String)
Dim i As Integer
Dim cad As String

    cad = ""
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then cad = cad & "1"
    Next i
    
    If Len(cad) > 0 Then
        If Len(cad) <> List1.ListCount Then
            
            'Vale, ha seleccionado alguno de los archivos. No todos
            cad = ""
            For i = 0 To List1.ListCount - 1
                If List1.Selected(i) Then cad = cad & " OR codext = " & List1.ItemData(i)
            Next i
            
            'Le quitamos el primer OR
            cad = Trim(Mid(cad, 4))
            
            'Lo metemos en SQL
            mSQL = mSQL & " AND (" & cad & ")"
        End If
    End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnPopup
    End If
End Sub

Private Sub mnQuitarSeleccion_Click()
    Seleccionar False
End Sub

Private Sub mnSelectAll_Click()
    Seleccionar True
End Sub


Private Sub Seleccionar(Si As Boolean)
Dim i As Integer
    For i = 0 To List1.ListCount - 1
        List1.Selected(i) = Si
    Next i
End Sub

