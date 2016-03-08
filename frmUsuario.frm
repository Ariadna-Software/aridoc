VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuario"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   Icon            =   "frmUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   240
      MaxLength       =   40
      PasswordChar    =   "*"
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   4200
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   240
      MaxLength       =   40
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   240
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   240
      MaxLength       =   40
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   9
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   8
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   7800
      TabIndex        =   13
      Top             =   360
      Width           =   615
      Begin VB.CommandButton cdmGrupo 
         Height          =   375
         Index           =   4
         Left            =   120
         Picture         =   "frmUsuario.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton cdmGrupo 
         Height          =   375
         Index           =   3
         Left            =   120
         Picture         =   "frmUsuario.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton cdmGrupo 
         Height          =   375
         Index           =   2
         Left            =   120
         Picture         =   "frmUsuario.frx":050E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cdmGrupo 
         Height          =   375
         Index           =   1
         Left            =   120
         Picture         =   "frmUsuario.frx":0610
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cdmGrupo 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "frmUsuario.frx":0712
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   4560
      TabIndex        =   7
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5741
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   240
      MaxLength       =   40
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   22
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "SERVIDOR"
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4080
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Direccion-email"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   20
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Grupos"
      Height          =   255
      Left            =   4560
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Login"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   11
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmG As frmGrupos
Attribute frmG.VB_VarHelpID = -1

Public vU As Cusuarios
'Private CambioEnGrupo As Boolean

Private Sub cdmGrupo_Click(Index As Integer)
    If Index < 3 Then
        'Nuevo , quitar, modificar
        If Index = 0 Then
            Set frmG = New frmGrupos
            frmG.DatosADevolverBusqueda = "0|1|"
            frmG.Show vbModal
            Set frmG = Nothing
        Else
            If MsgBox("Seguro que desa quitar al usuario, del grupo " & ListView1.SelectedItem.Text & "?", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
        
    Else
        
        'Subir bajar
        If ListView1.ListItems.Count < 2 Then Exit Sub
        
        If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    
        If Index = 3 Then
            'Es el primero
            If ListView1.SelectedItem.Index = 1 Then Exit Sub
            
            'Cambiamos el de arriba por este
            CambiaNodo ListView1.SelectedItem.Index, -1
        Else
            'Bajar
            If ListView1.SelectedItem.Index = ListView1.ListItems.Count Then Exit Sub
            CambiaNodo ListView1.SelectedItem.Index, 1
        End If
    End If
    
    
    
    
    
End Sub

Private Sub CambiaNodo(Origen As Integer, Incremento As Integer)
Dim Cad1 As String
Dim Cad2 As String
Dim J As Integer


    Cad1 = ListView1.ListItems(Origen).Text
    Cad2 = ListView1.ListItems(Origen).Tag
    J = Origen + Incremento
    '
    ListView1.ListItems(Origen).Text = ListView1.ListItems(J).Text
    ListView1.ListItems(Origen).Tag = ListView1.ListItems(J).Tag
    '
    ListView1.ListItems(J).Text = Cad1
    ListView1.ListItems(J).Tag = Cad2
    
    Set ListView1.SelectedItem = ListView1.ListItems(J)
End Sub


Private Sub cmdAceptar_Click(Index As Integer)
Dim Nuevo As Boolean
Dim OK As Byte
    If Index = 0 Then
        If Not DatosOk Then Exit Sub
    
        Nuevo = False
        If vU Is Nothing Then
            'Nuevo
            Nuevo = True
            Set vU = New Cusuarios
        End If
        
        vU.login = Text1(1).Text
        vU.Nombre = Text1(2).Text
    
        vU.e_dir = Text1(3).Text
        vU.e_server = Text1(4).Text
        vU.e_login = Text1(5).Text
        vU.e_pwd = Text1(6).Text
    
    
    
        'Si es nuevo asigno el mismo password k el login pero enminusculas
        If Nuevo Then vU.Password = LCase(Text1(1).Text)
    
        If Nuevo Then
            OK = vU.Agregar
        Else
            OK = vU.Modificar
        End If
    
        If OK = 1 Then
            If Nuevo Then Set vU = Nothing
            Exit Sub
        End If
        
        If Nuevo Then Text1(0).Text = vU.codusu
        
        VolcarExtensiones
        'Si hay modifiaciones cambiamos la variable para k
        'refresque el formulario anterior
        DatosMOdificados = True
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    
    If vU Is Nothing Then
        'Nuevo
        Limpiar Me
        Text1(0).Tag = ""
    Else
        Text1(0).Tag = "MOD"
        'Ponemos campos
        Text1(0).Text = vU.codusu
        Text1(1).Text = vU.login
        Text1(2).Text = vU.Nombre
        
        'Mail
        Text1(3).Text = vU.e_dir
        Text1(4).Text = vU.e_server
        Text1(5).Text = vU.e_login
        Text1(6).Text = vU.e_pwd
        
    End If
    ListView1.SmallIcons = Admin.ImgUsersPCs
    CargarGrupos
    Me.Frame1.Enabled = vUsu.Nivel < 2   'Solo root y administradores
End Sub


Private Sub CargarGrupos()
Dim Cad As String
Dim Itm As ListItem
    Set miRSAux = New ADODB.Recordset
    ListView1.ListItems.Clear
    Cad = "Select usuariosgrupos.codgrupo,nomgrupo from usuariosgrupos,grupos "
    Cad = Cad & " WHERE usuariosgrupos.codgrupo=grupos.codgrupo AND usuariosgrupos.codusu = "
    If vU Is Nothing Then
        Cad = Cad & "-1"
    Else
        Cad = Cad & vU.codusu
    End If
    Cad = Cad & " ORDER By usuariosgrupos.Orden"
    miRSAux.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Set Itm = ListView1.ListItems.Add(, "C" & miRSAux!codgrupo)
        Itm.Text = miRSAux!nomgrupo
        Itm.SmallIcon = 1
        Itm.Tag = miRSAux!codgrupo
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
End Sub


Private Function DatosOk() As Boolean
    DatosOk = False
    
    If Text1(1).Text = "" Or Text1(2).Text = "" Then
        MsgBox "Debe rellenar login/nombre del usuario.", vbExclamation
        Exit Function
    End If
    
    
    If ListView1.ListItems.Count = 0 Then
        MsgBox "Debe asignar al menos un grupo", vbExclamation
        Exit Function
    End If
        
    If Text1(0).Tag = "" Then
        Me.Tag = DevuelveDesdeBD("Nombre", "usuarios", "login", Text1(1).Text, "T")
        If Me.Tag <> "" Then
            MsgBox "El login pertence al usuario: " & Me.Tag, vbExclamation
            Me.Tag = ""
            Exit Function
        End If
        Me.Tag = ""
    End If
    DatosOk = True
    
End Function

'Esto tal vez deberia estar en el objeto USUARIO
Private Sub VolcarExtensiones()
Dim I As Integer
Dim Cad As String
Dim Orden As Integer

    Cad = "Delete from usuariosgrupos where codusu =" & vU.codusu
    Conn.Execute Cad
    
    Orden = 1
    For I = 1 To ListView1.ListItems.Count
        Cad = "INSERT INTO usuariosgrupos (codusu, codgrupo, orden) VALUES ("
        Cad = Cad & vU.codusu & "," & ListView1.ListItems(I).Tag & "," & Orden & ")"
        Conn.Execute Cad
        Orden = Orden + 1
    Next I
    
End Sub

Private Sub frmG_DatoSeleccionado(CadenaSeleccion As String)
Dim Cad As String
Dim I As Integer
Dim Itm As ListItem

    Cad = Val(RecuperaValor(CadenaSeleccion, 1))
    If Val(Cad) <> 0 Then
        For I = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(I).Tag = Cad Then
                MsgBox "El usuario ya pertenece al grupo: " & ListView1.ListItems(I).Text, vbExclamation
                Exit For
            End If
        Next I
        
        If I > ListView1.ListItems.Count Then
            'Se ha salido sin encontrarlo
            Set Itm = ListView1.ListItems.Add(, "C" & Cad)
            Itm.Tag = Cad
            Cad = RecuperaValor(CadenaSeleccion, 2)
            Itm.Text = Cad
            Itm.SmallIcon = 1
            'CambioEnGrupo = True
        End If
    End If
End Sub



Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

