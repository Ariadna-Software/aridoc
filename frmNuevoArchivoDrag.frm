VERSION 5.00
Begin VB.Form frmNuevoArchivoDrag 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Agregar documentos en ARIDOC"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   Icon            =   "frmNuevoArchivoDrag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAridoc 
      Height          =   4935
      Left            =   120
      TabIndex        =   44
      Top             =   600
      Width           =   8535
      Begin VB.CommandButton Command1 
         Caption         =   "+"
         Height          =   255
         Left            =   1800
         TabIndex        =   49
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Text            =   "Text4"
         Top             =   480
         Width           =   8055
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   240
         TabIndex        =   45
         Top             =   1200
         Width           =   8055
      End
      Begin VB.Label Label4 
         Caption         =   "Archivos a insertar"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   960
         Width           =   3855
      End
      Begin VB.Label Label2 
         Caption         =   "CARPETA ARIDOC"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdFIN 
      Default         =   -1  'True
      Height          =   375
      Left            =   8040
      Picture         =   "frmNuevoArchivoDrag.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "INSERTAR"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdAnt 
      Height          =   375
      Left            =   7440
      Picture         =   "frmNuevoArchivoDrag.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Datos"
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdSig 
      Height          =   375
      Left            =   8040
      Picture         =   "frmNuevoArchivoDrag.frx":1296
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Carpeta ARIDOC"
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Guardar usr/pwd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame FrameDatos 
      Height          =   4935
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   8535
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   4200
         TabIndex        =   35
         Top             =   4080
         Width           =   4215
         Begin VB.OptionButton optEscriutra 
            Caption         =   "Propietario"
            Height          =   195
            Index           =   2
            Left            =   2880
            TabIndex        =   38
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton optEscriutra 
            Caption         =   "Grupo"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   37
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optEscriutra 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Caption         =   "Escritura"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   4080
         Width           =   3975
         Begin VB.OptionButton OptLectura 
            Caption         =   "Propietario"
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   33
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton OptLectura 
            Caption         =   "Grupo"
            Height          =   255
            Index           =   1
            Left            =   1440
            TabIndex        =   32
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton OptLectura 
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   31
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Lectura"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   0
            Width           =   660
         End
      End
      Begin VB.Frame Frame1 
         Height          =   3735
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   8295
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   0
            Left            =   240
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "Text3"
            Top             =   480
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   2760
            MaxLength       =   15
            TabIndex        =   17
            Text            =   "Text3"
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtClaves 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   6720
            TabIndex        =   12
            Text            =   "Text3"
            Top             =   1920
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   1
            Left            =   4320
            MaxLength       =   50
            TabIndex        =   4
            Text            =   "Text3"
            Top             =   480
            Width           =   3735
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   2
            Left            =   240
            MaxLength       =   50
            TabIndex        =   5
            Text            =   "Text3"
            Top             =   1200
            Width           =   3975
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   3
            Left            =   4320
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "Text3"
            Top             =   1200
            Width           =   3735
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   4
            Left            =   240
            TabIndex        =   7
            Text            =   "99/99/9999"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   5
            Left            =   1440
            TabIndex        =   8
            Text            =   "Text3"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox txtClaves 
            Height          =   285
            Index           =   6
            Left            =   2640
            TabIndex        =   9
            Text            =   "Text3"
            Top             =   1920
            Width           =   1095
         End
         Begin VB.TextBox txtClaves 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   3840
            TabIndex        =   10
            Text            =   "Text3"
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtClaves 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   5280
            TabIndex        =   11
            Text            =   "Text3"
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtClaves 
            Height          =   885
            Index           =   9
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Text            =   "frmNuevoArchivoDrag.frx":1820
            Top             =   2640
            Width           =   7815
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   6
            Left            =   2640
            Picture         =   "frmNuevoArchivoDrag.frx":1826
            Top             =   1680
            Width           =   240
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   5
            Left            =   1440
            Picture         =   "frmNuevoArchivoDrag.frx":1928
            Top             =   1680
            Width           =   240
         End
         Begin VB.Image Image3 
            Height          =   240
            Index           =   4
            Left            =   240
            Picture         =   "frmNuevoArchivoDrag.frx":1A2A
            Top             =   1680
            Width           =   240
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
            TabIndex        =   29
            Top             =   480
            Width           =   1755
         End
         Begin VB.Label Label3 
            Caption         =   "Tamaño"
            Height          =   255
            Index           =   10
            Left            =   6720
            TabIndex        =   28
            Top             =   1680
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   1
            Left            =   4320
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   25
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   3
            Left            =   4320
            TabIndex        =   24
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   23
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   22
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   6
            Left            =   2880
            TabIndex        =   21
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Importe"
            Height          =   255
            Index           =   7
            Left            =   3840
            TabIndex        =   20
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Importe"
            Height          =   255
            Index           =   8
            Left            =   5280
            TabIndex        =   19
            Top             =   1680
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Label3"
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   18
            Top             =   2400
            Width           =   3255
         End
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdSALIR 
      Cancel          =   -1  'True
      Caption         =   "Command2"
      Height          =   255
      Left            =   7920
      TabIndex        =   50
      Top             =   2880
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   40
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmNuevoArchivoDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Dim Primera As Boolean
Dim CamposBD As Boolean
Dim CadenaExtensiones As String


Private Sub MemorizarUsuario(Leer As Boolean)
Dim N As String
    On Error GoTo EMemo
    N = App.Path & "\memous.dat"
    If Leer Then
        If Dir(N, vbArchive) <> "" Then
              Check1.Value = 1
              AccionesFichero True, N
        Else
            Check1.Value = 0
        End If
        
        
        If Check1.Value = 1 Then
            'Dice memorizar el usuario
            'Comprobamos que el usuario es el que es
            LeerUsuario
            PonerTamanyos (vUsu Is Nothing)
        End If
    Else
        If Check1.Value = 0 Then
            If Dir(N, vbArchive) <> "" Then Kill N
        Else
            'Guaradar
            AccionesFichero False, N
        End If
    End If
    
    
    Exit Sub
EMemo:
    MuestraError Err.Number, "Leyendo fichero datos guardados"
End Sub


Private Sub Check1_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
Dim i As Integer
    frmPregunta.Opcion = 20
    frmPregunta.origenDestino = 1
    DatosCopiados = ""
    frmPregunta.Show vbModal
    If DatosCopiados <> "" Then
        i = InStr(1, DatosCopiados, "·")
        If i = 0 Then
            Text4.Text = ""
        Else
            Text4.Text = Mid(DatosCopiados, 1, i - 1)
            
            Text4.Tag = Mid(DatosCopiados, i + 1)
            'Le quito el ultimo "|"
            Text4.Tag = Mid(Text4.Tag, 1, Len(Text4.Tag) - 1)
            
            cmdFIN.SetFocus
        End If
    End If
End Sub

Private Sub cmdAnt_Click()
    CamposBD = True
    PonerTamanyos (vUsu Is Nothing)
End Sub

Private Sub cmdFIN_Click()
Dim B As Boolean
Dim C As Ccarpetas
    'INSERTAR EN BD
    'Y llevar archivos
    If Not CamposBDOK Then Exit Sub
    
    
    
    
    Set C = New Ccarpetas
    B = False
    If C.Leer(Text4.Tag, False) = 1 Then
        C.Nombre = "Error obteniendo carpeta ARIDOC."
    Else
        If C.userprop = vUsu.codusu Or (C.escriturag And vUsu.Grupo) Or vUsu.codusu = 0 Then
            B = True
        Else
            C.Nombre = "No tiene permisos de escritura en la carpeta."
        End If
    End If
    If Not B Then
        MsgBox C.Nombre, vbExclamation
        Set C = Nothing
        Exit Sub
    End If
    
    
    'Si llega aqui es que va a empezar a insertar archivos
    'ASin que avisamos y punto
    If MsgBox("Va a insertar en la gestion documental. El proceso puede llevar tiempo. " & vbCrLf & vbCrLf & "Desea continuar?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    HacerInsercionArchivos C
    Unload Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSig_Click()
    CamposBD = False
    PonerTamanyos (vUsu Is Nothing)
    Command1.SetFocus
End Sub

Private Sub Form_Activate()


    If Primera Then
        Primera = False
        
        If Me.Height > 5000 Then
            
            If listaimpresion.Count = 1 Then
                txtClaves(0).SetFocus
                
            Else
                
                Text3.SetFocus
            End If
            
        Else
            Text1.SetFocus
        End If
        
    End If
    Screen.MousePointer = vbDefault
End Sub



Private Sub PonerCampos()
    If listaimpresion.Count = 1 Then
        '¡SOLO HAY UNO
        txtClaves(10).Visible = True
        Label3(10).Visible = True
        txtClaves(0).Visible = True
        
        txtClaves(0).Text = CStr(DevuelveNombreFichero(listaimpresion.Item(1)))
        On Error Resume Next
        txtClaves(10).Text = Round(FileLen(listaimpresion.Item(1)) / 1024, 3)
        
    Else
        Text3.TabIndex = 3
        Text3.Visible = True
    End If
    txtClaves(4).Text = Format(Now, "dd/mm/yyyy")
    txtClaves(5).Text = Format(Now, "dd/mm/yyyy")
End Sub




Private Sub Form_Load()
    Primera = True
    CamposBD = True
    
    'Memorizo las extensiones que tengo ene sta variable
    CadenaExtensiones = DatosCopiados
    DatosCopiados = ""
    
    Limpiar Me
    PonerLabels
    CargarArchivos
    MemorizarUsuario True
    PonerTamanyos (vUsu Is Nothing)
    PonerCampos
End Sub



Private Sub PonerLabels()
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
  
End Sub


Private Sub AccionesFichero(Lectura As Boolean, F As String)
Dim NE As Integer
Dim C As String

    On Error GoTo EF
    
    NE = FreeFile
    If Lectura Then
        Open F For Input As #NE
        'Primera linea aleatorio
        Line Input #NE, C
        'Segnda usuario
        Line Input #NE, C
        Text1.Text = LeeLinea(C)
        '3a linea aleatorio
        Line Input #NE, C
        '4º password
        Line Input #NE, C
        Text2.Text = LeeLinea(C)
    Else
        Open F For Output As #NE
        C = EscribeLinea(Text1.Text)
        Print #NE, C
        C = EscribeLinea(Text2.Text)
        Print #NE, C
        C = EscribeLinea("D@BYZICEDUM")
        Print #NE, C
        
    End If
    Close #NE
EF:
    On Error Resume Next
    Close #NE
    Err.Clear
End Sub


Private Function LeeLinea(CADENA As String) As String

Dim i As Integer
    
    
    For i = 1 To Len(CADENA)   'Empezamos en el uno
        If (i Mod 3) = 0 Then LeeLinea = LeeLinea & Mid(CADENA, i, 1)
    Next i
End Function

Private Function EscribeLinea(CADENA As String) As String
Dim i As Integer
Dim MiValor As Integer


        For i = 1 To Len(CADENA)
            MiValor = 122 - 41 'Primero y ultimo caracter normal
            MiValor = Int((MiValor * Rnd) + 41)
            EscribeLinea = EscribeLinea & Chr(MiValor)
        Next i
        EscribeLinea = EscribeLinea & "DBZ" & vbCrLf
        
        For i = 1 To Len(CADENA)
            MiValor = 122 - 41 'Primero y ultimo caracter normal
            MiValor = Int((MiValor * Rnd) + 41)
            EscribeLinea = EscribeLinea & Chr(MiValor) & Chr(MiValor + 1)
            EscribeLinea = EscribeLinea & Mid(CADENA, i, 1)
            
        Next i
        
            
        
End Function

Private Sub Form_Unload(Cancel As Integer)
    MemorizarUsuario False
End Sub


Private Sub LeerUsuario()
Dim Password As String
Dim cad As String
    Caption = "Agregar documentos en ARIDOC"

    If Text1.Text = "" Or Text2.Text = "" Then
        Exit Sub
    End If
    
    
    Password = "Password"
    cad = DevuelveDesdeBD("codusu", "usuarios", "login", Text1.Text, "T", Password)
    
    If cad = "" Then
        MsgBox "Usuario / Password  incorrectos", vbExclamation
        Text2.Text = ""
        Exit Sub
    End If
    
    If Password <> Text2.Text Then
        Text2.Text = ""
        MsgBox "Usuario / password incorrecto", vbExclamation
        Exit Sub
    End If
    
    Set vUsu = New Cusuarios
    vUsu.Leer CInt(cad)
    Caption = Caption & " ( " & vUsu.Nombre & ")"
End Sub


Private Sub PonerTamanyos(Pequenyo As Boolean)

    If Pequenyo Then
        Me.Height = 1000
        FrameDatos.Visible = False
        
    Else
        
        FrameDatos.Visible = CamposBD
        FrameAridoc.Visible = Not CamposBD
        Me.Height = 6075
    End If
    Me.cmdAnt.Visible = Not Pequenyo And Not CamposBD
    Me.cmdFIN.Visible = Not Pequenyo And Not CamposBD
    Me.cmdSig.Visible = Not Pequenyo And CamposBD
    
End Sub

Private Sub CambioCampoUsuario()
    Set vUsu = Nothing
    LeerUsuario
    PonerTamanyos (vUsu Is Nothing)
End Sub



Private Sub frmC_Selec(vFecha As Date)
    DatosCopiados = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image3_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtClaves(Index).Text <> "" Then
        If IsDate(txtClaves(Index).Text) Then frmC.Fecha = CDate(txtClaves(Index).Text)
    End If
    DatosCopiados = ""
    frmC.Show vbModal
    Set frmC = Nothing
    If DatosCopiados <> "" Then
        txtClaves(Index).Text = DatosCopiados
        DatosCopiados = ""
    End If
End Sub



Private Sub optEscriutra_KeyPress(Index As Integer, KeyAscii As Integer)
KEYpress KeyAscii
End Sub



Private Sub OptLectura_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text1_GotFocus()
    Text1.Tag = Text1.Text
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KEYpress KeyAscii
End Sub

Private Sub Text1_LostFocus()
    If Text1.Text <> Text1.Tag Then CambioCampoUsuario
    Text1.Tag = ""
End Sub

Private Sub Text2_GotFocus()
    Text2.Tag = Text2.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
 KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus()
    If Text2.Text <> Text1.Tag Then CambioCampoUsuario
    Text2.Tag = ""
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtClaves_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub CargarArchivos()
Dim L As Integer
    For L = 1 To listaimpresion.Count
          List1.AddItem CStr(listaimpresion.Item(L))
    Next L
End Sub


Private Function CamposBDOK() As Boolean
    CamposBDOK = False
    If listaimpresion.Count = 1 Then
        If txtClaves(0).Text = "" Then
            MsgBox "Ponga valor para " & Label3(0).Caption, vbExclamation
            Exit Function
        End If
    End If
    If txtClaves(4).Text = "" Or txtClaves(4).Text = "" Then
        MsgBox "Campos " & Label3(4).Caption & " y " & Label3(5).Caption & " son obligados", vbExclamation
        Exit Function
    End If
    If Text4.Text = "" Then
        MsgBox "Carpeta ARIDOC obligada", vbExclamation
        Exit Function
    End If
    CamposBDOK = True
End Function




'-----------------------------------------------------------------------
' INSERTAR ARCHIVOS

Private Sub HacerInsercionArchivos(C As Ccarpetas)
Dim i As Long
Dim mImag As cTimagen
Dim FSS, F
Dim cad As String


    
    
    
    Set mImag = New cTimagen
    mImag.codcarpeta = C.codcarpeta
    'Ajusto los valores comunes
    mImag.campo2 = txtClaves(1).Text
    mImag.campo3 = txtClaves(2).Text
    mImag.campo4 = txtClaves(3).Text
    'Fechas
    mImag.fecha1 = txtClaves(4).Text
    mImag.fecha2 = LetDB(txtClaves(5).Text, "F")
    mImag.fecha3 = LetDB(txtClaves(6).Text, "F")
    
    'Importes
    mImag.importe1 = LetDB(txtClaves(7).Text, "N")
    mImag.importe2 = LetDB(txtClaves(8).Text, "N")
    
    'Observaciones
    mImag.observa = txtClaves(9).Text
    
    'Porpietario
    mImag.groupprop = vUsu.GrupoPpal
    mImag.userprop = vUsu.codusu
    
    
    'Permisos
    If Me.OptLectura(0).Value Then
        i = vbPermisoTotal
    Else
        If Me.OptLectura(1).Value Then
            i = GrupoLongBD(vUsu.GrupoPpal)
        Else
            i = 0
        End If
    End If
    mImag.lecturag = i
    
    
    'escritura
     
    If Me.optEscriutra(0).Value Then
       i = vbPermisoTotal
    Else
        If Me.optEscriutra(1).Value Then
            i = GrupoLongBD(vUsu.GrupoPpal)
        Else
            i = 0
        End If
    End If
    mImag.escriturag = i
    
    
    
    Set FSS = CreateObject("Scripting.FileSystemObject")
    
    
    
    
    
    For i = 1 To listaimpresion.Count
        Set F = FSS.GetFile(listaimpresion.Item(i))
        If txtClaves(6).Text = "" Then
            'No ha puesto fecha. Pondremos la fecha de creacion del archivo
            mImag.fecha3 = Format(F.DateCreated, "dd/mm/yyyy")
        End If
        cad = DevuelveNombreFichero(F.Name)
        
        If listaimpresion.Count > 1 Then
            mImag.campo1 = Trim(cad & " " & Text3.Text)
        Else
            If txtClaves(0).Text = "" Then txtClaves(0).Text = cad
            mImag.campo1 = txtClaves(0).Text
        End If
        mImag.tamnyo = Round(F.Size / 1024, 3)
        
        
        
        
        InsertarArchivo F.Path, mImag, C
        Set F = Nothing
    Next i
End Sub

Private Function DevuelveNombreFichero(ByVal C As String) As String
Dim J As Integer
        'quitaos el punto
        J = InStrRev(C, ".")
        If J > 0 Then C = Mid(C, 1, J - 1)
        'Quitamos la \
        J = InStrRev(C, "\")
        If J > 0 Then C = Mid(C, J + 1)
        
        DevuelveNombreFichero = C
End Function

Private Sub InsertarArchivo(NomArchivo As String, mI As cTimagen, ByRef vCar As Ccarpetas)
Dim i As Integer
Dim Aux As String
Dim mError As String
Dim J As Integer 'logitud de la extension
    
    'Comprobamos la extension aunque no deberia dar FALLOS
    i = InStrRev(NomArchivo, ".")
    If i = 0 Then
        'FALLO
        mError = "No se encontro la extension para el archivo: "
    Else
        Aux = LCase(Mid(NomArchivo, i + 1))
        i = InStr(1, CadenaExtensiones, Aux & ":")
        If i = 0 Then
            'No tratamos la extension
            mError = "Extension no tratada para el archivo: "
        Else
            J = Len(Aux)
            Aux = Mid(CadenaExtensiones, i + J + 1) 'Para los :
            i = InStr(1, Aux, "|")
            If i = 0 Then
                mError = "No se puede encontrar la extension"
            Else
                Aux = Mid(Aux, 1, i - 1)
            End If
        End If
    End If
    If mError <> "" Then
        MsgBox mError & vbCrLf & " .- " & NomArchivo, vbExclamation
        Exit Sub
    End If
    mI.codext = CInt(Aux) 'llevara el codigo d extension a  tratar
 

    
    
    
    
    
    
    
    If mI.Agregar(objRevision.LlevaHcoRevision, False) = 0 Then
            
        Set frmMovimientoArchivo.vDestino = vCar
        frmMovimientoArchivo.Opcion = 1
        frmMovimientoArchivo.Origen = NomArchivo
        frmMovimientoArchivo.Destino = CStr(mI.codigo)
        frmMovimientoArchivo.Show vbModal
            
            
            
            
        'Y si se producen errores entonces borramos el Imgag
        If DatosCopiados <> "" Then
        'Error llevando datos
            mI.Eliminar
        End If
            
    End If
    
    
    
    
End Sub



