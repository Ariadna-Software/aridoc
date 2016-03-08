VERSION 5.00
Begin VB.Form frmCarpetas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carpetas"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "frmCarpetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   5160
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   1350
      Width           =   5655
   End
   Begin VB.Frame Frame3 
      Caption         =   "Propietario"
      Height          =   1335
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   5655
      Begin VB.TextBox Text3 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   15
         Text            =   "Text2"
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   360
         Width           =   4335
      End
      Begin VB.Image imgUserGroup 
         Height          =   240
         Index           =   1
         Left            =   720
         Picture         =   "frmCarpetas.frx":030A
         Top             =   840
         Width           =   240
      End
      Begin VB.Image imgUserGroup 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmCarpetas.frx":0894
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Grupo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   880
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   440
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   12
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   10
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   5655
      Begin VB.OptionButton optEscriutra 
         Caption         =   "Propietario"
         Height          =   195
         Index           =   2
         Left            =   4200
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optEscriutra 
         Caption         =   "Grupo"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optEscriutra 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   960
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Escritura"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   675
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   5655
      Begin VB.OptionButton OptLectura 
         Caption         =   "Propietario"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptLectura 
         Caption         =   "Grupo"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton OptLectura 
         Caption         =   "Todos"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Lectura"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   660
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Ubicacion"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   4815
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmCarpetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const vbMinimoCarpetaAlmacen = 3

Private WithEvents frmU As frmUsuarios2
Attribute frmU.VB_VarHelpID = -1
Private WithEvents frmG As frmGrupos
Attribute frmG.VB_VarHelpID = -1

Public vC As Ccarpetas
Public Ubicacion As String
Public PuedeModificar As Boolean

Private ElUsuario As String

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then
        If Not DatosOk Then Exit Sub
        
        'INSERTAMOS , o modificamos
        If Label2.Tag = 0 Then
            'INSERTAMOS
            If vC.Agregar = 1 Then Exit Sub
        Else
            If vC.Modificar = 1 Then Exit Sub
        End If
        'AQUI UPDATEAMOS PARA QUE REFRESQUEN CARPETAS
        vC.ActualizaTablaActualiza
        DatosMOdificados = True
    End If
    Unload Me
End Sub

Private Sub Form_Load()
Dim B As Boolean
    CargaCombo
    Text4.Text = Ubicacion
    If vC.Nombre = "" Then
        'NUEVA CARPETA
        Label2.Caption = "NUEVA CARPETA"
        Label2.ForeColor = vbBlue
        Label2.Tag = 0
        Text1.Text = "Nueva carpeta"
        Me.optEscriutra(0).Value = True
        Me.OptLectura(0).Value = True
    Else
        'MODIFICAR PROPIEDADES
        If PuedeModificar Then
            Label2.Caption = "MODIFICAR CARPETA"
            Label2.ForeColor = vbRed
        Else
            Label2.Caption = "DATOS CARPETA"
            Label2.ForeColor = &H4000&
        End If
        Caption = "Carpeta (" & vC.codcarpeta & ")"
        Label2.Tag = 1
        Text1.Text = vC.Nombre
        Ponerpermisos
    End If
    
    PonerPropietarios
    'Icono para setuidar la carpeta
    Me.imgUserGroup(0).Visible = (vUsu.Grupo And 1)
    Me.imgUserGroup(1).Visible = (vUsu.Grupo And 1)
    'imgChangaProp.Visible = vUsu.codusu = 0
    
    Me.Command1(0).Visible = PuedeModificar
    Text1.Locked = Not PuedeModificar
    If vUsu.codusu = 0 Then
        B = True
    Else
        B = (vUsu.codusu = vC.userprop) Or (vUsu.GrupoPpal = vC.groupprop)
    End If
    Frame1.Enabled = B
    Frame2.Enabled = B
End Sub


Private Sub Ponerpermisos()
   
 'Permiso escritura usuario
    If vC.lecturag = 0 Then
        Me.OptLectura(2).Value = True
    Else
        If vC.lecturag = vbPermisoTotal Then
            Me.OptLectura(0).Value = True
        Else
            Me.OptLectura(1).Value = True
        End If
    End If
    
    'Permiso ESCRITURA
     
    If vC.escriturag = 0 Then
        Me.optEscriutra(2).Value = True
    Else
        If vC.escriturag = vbPermisoTotal Then
            Me.optEscriutra(0).Value = True
        Else
            Me.optEscriutra(1).Value = True
        End If
    End If
End Sub
Private Sub PonerPropietarios()
Dim Cad As String

    On Error GoTo EPonerPropietarios
    Set miRSAux = New adodb.Recordset
    'Usuario
    Cad = "Select nombre from usuarios WHERE codusu =" & vC.userprop
    miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then Text2.Text = DBLet(miRSAux!Nombre, "T")
    miRSAux.Close
    
    'Grupo
    Cad = "Select nomgrupo from grupos WHERE codgrupo =" & vC.groupprop
    miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then Text3.Text = DBLet(miRSAux!nomgrupo, "T")
    miRSAux.Close
    
    
    
    
    
    
    Set miRSAux = Nothing
    Exit Sub
EPonerPropietarios:
    MuestraError Err.Number, "Poner datos propietario"
End Sub


Private Sub frmG_DatoSeleccionado(CadenaSeleccion As String)
    vC.groupprop = RecuperaValor(CadenaSeleccion, 1)
    PonerPropietarios
    Screen.MousePointer = vbDefault
End Sub

Private Sub frmU_DatoSeleccionado(CadenaSeleccion As String)
Dim C As String
    Screen.MousePointer = vbHourglass
    
    C = "Select codgrupo from usuariosgrupos where codusu=" & RecuperaValor(CadenaSeleccion, 1) & " ORDER BY orden"
'    Set miRSAux = New ADODB.Recordset
'    miRSAux.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    C = ""
'    If Not miRSAux.EOF Then
'        If Not IsNull(miRSAux.Fields(0)) Then C = miRSAux.Fields(0)
'    End If
'    miRSAux.Close
'    Set miRSAux = Nothing
'    If C = "" Then
'        MsgBox "Grupo PPal para el usuario: " & CadenaSeleccion & " NO encontrado", vbExclamation
'        Exit Sub
'    End If
    
    'Llegado aqui, ponemos
    
    vC.userprop = Val(RecuperaValor(CadenaSeleccion, 1))
    'vC.groupprop = Val(C)
    PonerPropietarios
    Screen.MousePointer = vbDefault
End Sub



Private Sub imgUserGroup_Click(Index As Integer)
    If ModoTrabajo <> vbNorm Then Exit Sub
    
    If Index = 0 Then
            
        Set frmU = New frmUsuarios2
        frmU.DatosADevolverBusqueda = "0|"
        frmU.Show vbModal
        Set frmU = Nothing
    Else
    
        Set frmG = New frmGrupos
        frmG.DatosADevolverBusqueda = "0|1|"
        frmG.Show vbModal
        Set frmG = Nothing
    End If
End Sub

Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub



Private Function DatosOk() As Boolean
Dim L As Long

    DatosOk = False
    
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "" Then
        MsgBox "Nombre de carpeta en blanco", vbExclamation
        Exit Function
    End If
    
    
    
    
    '----------------------------
    'Metemos en el Objeto CARPETA
    '----------------------------
     vC.Nombre = Text1.Text
     
    'Permisos
    '-------
    
    'lectura
    If Me.OptLectura(0).Value Then
        L = vbPermisoTotal
    Else
        If Me.OptLectura(1).Value Then
'            L = GrupoLongBD(vUsu.GrupoPpal)  'ANTES
            L = GrupoLongBD(vC.groupprop)
        Else
            L = 0
        End If
    End If
    vC.lecturag = L
    
    
    'escritura
     
    If Me.optEscriutra(0).Value Then
        L = vbPermisoTotal
    Else
        If Me.optEscriutra(1).Value Then
            L = GrupoLongBD(vC.groupprop)
        Else
            L = 0
        End If
    End If
    vC.escriturag = L
        
    
    'Ahora. Si el usuario es root puede selecionar el almacen
    If Label2.Tag = 0 Then
        'NUEVO
        If vUsu.codusu = 0 Then
            If Combo1.ItemData(Combo1.ListIndex) <> vC.Almacen Then
                   CambioAlmacen
             End If
        End If
     End If
    
    DatosOk = True
End Function

Private Sub CambioAlmacen()
Dim Cad As String
    Set miRSAux = New adodb.Recordset
    Cad = "Select * from almacen where codalma=" & Combo1.ItemData(Combo1.ListIndex)
    miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then
        'Ha seleccionado otro almacen
        'El almacen contenedor
        With vC
            .version = miRSAux!version
            .Almacen = miRSAux!codalma
            .user = miRSAux!user
            .pwd = miRSAux!pwd
        End With
    End If
    miRSAux.Close
    Set miRSAux = Nothing
End Sub


Private Sub CargaCombo()
Dim Cad As String
Dim I As Integer


    Set miRSAux = New adodb.Recordset
    'ANTES. Cuando no estaba la carpeta files
    'Cad = "Select * from almacen where codalma >0 order by codalma"

    'AHORA. 30 Mayo 05
    Cad = "Select * from almacen where codalma >=" & vbMinimoCarpetaAlmacen & " order by codalma"
    
    miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Cad = "Almacen: " & miRSAux!codalma & " (" & miRSAux!pathreal & ")"
        Combo1.AddItem Cad
        Combo1.ItemData(Combo1.NewIndex) = miRSAux!codalma
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing

'    If vC.Nombre = "" Then
'
'        Cad = "NO"
'        If vC.Nombre = "" Then
'            'Si es nueva, pongo el primer almacen por defecto
'            Combo1.ListIndex = 0
'
'            If vUsu.codusu = 0 Then Cad = ""
'        End If
'        Combo1.Visible = (Cad = "")
'        If Combo1.Visible Then Combo1.Enabled = True
'    Else
'
        Cad = "NO"
        For I = 0 To Combo1.ListCount - 1
            If Combo1.ItemData(I) = vC.Almacen Then
                Combo1.ListIndex = I
                Exit For
            End If
        Next I
        If I <= Combo1.ListIndex Then
            If vUsu.codusu = 0 Then Cad = ""
        End If
        
        Combo1.Visible = Cad = ""
        If Combo1.Visible Then Combo1.Enabled = vUsu.codusu = 0
'    End If
End Sub


