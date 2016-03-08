VERSION 5.00
Begin VB.Form frmInicio 
   BorderStyle     =   0  'None
   Caption         =   "Identificacion Aridoc"
   ClientHeight    =   4470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3921.053
   ScaleMode       =   0  'User
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   4800
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargando ...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   4800
      TabIndex        =   2
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   0
      Top             =   2640
      Width           =   2055
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim T1 As Single
Dim PrimeraVez As Boolean


Private Sub Validar()
Dim Rc As Byte
Dim cad As String
Dim Password As String

    'Veremos login password. Comprobamos k codusu tiene
    Text1.Text = Trim(Text1.Text)
    Text2.Text = Trim(Text2.Text)
    If Text1.Text = "" Or Text2.Text = "" Then Exit Sub
    
    Password = "Password"
    cad = DevuelveDesdeBD("codusu", "usuarios", "login", Text1.Text, "T", Password)
    
    If cad = "" Then
        MsgBox "Usuario / Password  incorrectos", vbExclamation
        Exit Sub
    End If
    
    If Password <> Text2.Text Then
        MsgBox "Usuario / password incorrecto", vbExclamation
        Exit Sub
    End If
    
    'A partir de codusu de arriba leemos los datos del usuario
    'Usuario
    Set vUsu = New Cusuarios
    

    
    Rc = vUsu.Leer(Val(cad))

    'Segun sea la lectura
    If Rc = 1 Then
        'Error en Usuario
        MsgBox "Usuario INCORRECTO", vbExclamation
        

       Set vUsu = Nothing
       Exit Sub
    Else
        If Rc = 2 Then
            'Lo gestionamos
            '----------------
            'Vemos el equipo
            GestionarEquipo
            
            'NUEVO PC o sin configurar extensiones
            '------------------------------------
            frmConfigExtensiones.NuevoEquipo = True
            frmConfigExtensiones.Show vbModal
            End
        End If
    End If


    Text1.Visible = False
    Text2.Visible = False
    Label1(0).Visible = False
    Label1(1).Visible = False
    Label1(2).Visible = True
    Me.Refresh
    espera 0.2
    

    'CargarImagenes extensiones
    If vUsu.CargaIconosExtensiones Then
        'Hay k copiar los iconos
        frmMovimientoArchivo.Opcion = 0
        frmMovimientoArchivo.Show vbModal

    End If
    
    
    'Elimino tabla procesos
    Conn.Execute "DELETE FROM Procesos where codequipo =" & vUsu.PC
    
    'Elimino tmpfich
    Conn.Execute "DELETE FROM tmpfich where codequipo =" & vUsu.PC
    
    
    
    
    'Otras acciones
    '---------------
    DoEvents
    
    'Refrescamos
    Label1(2).Caption = Label1(2).Caption & ".."
    Me.Refresh
    espera 0.2
    
    
    
    
    'MENSAJE DE PRIVACIDAD DE LOS DATOS
    '--------------------------------
    'cad = "RECUERDE LA OBLIGACION QUE USTED TIENE DE MANTENER LA PRIVACIDAD DE LOS DATOS A LOS QUE ESTA AUTORIZADO"
    'cad = cad & " Y LAS CONSECUENCIAS DEL INCUMPLIMIENTO DE DICHA PRIVACIDAD SEGUN REAL DECRETO 994/1999 DEL 11 DE JUNIO"
    'cad = cad & " SOBRE MEDIDAS DE SEGURIDAD Y PROTECCION DE DATOS"
    cad = ""
    If vConfig.LeyProtDatos2 <> "" Then cad = vbCrLf & vConfig.LeyProtDatos2
    cad = vConfig.LeyProtDatos1 & cad
    If cad <> "" Then MsgBox cad, vbInformation
    
    

    
    Load Admin
    Label1(2).Caption = Label1(2).Caption & ".."
    Me.Refresh
    espera 0.2
    
        
    Admin.Show
    
    'Si el ultimo usuario ha cambiado lo guardo
    If Text1.Text <> Text1.Tag Then UltimoUsuario False
    
    
    Unload Me
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        
        If Not ComprobacionesPrevias Then
            End
            Exit Sub
        End If
    
    
    
        '------------------------------------------------
        ' Este trozo esta igual en la libreria de SHELL
    
        'Abrimos el DSN
        If Not AbrirConexion(True) Then
            End
            Exit Sub
        End If
        
        
        
        'Leemos el objeto Confiuracion
        Set vConfig = New CConfiguracion
        If vConfig.Leer(1) = 1 Then
            End
            Exit Sub
        End If
    
    
            
        'VEo si llevamos revsion documental o no
        Set objRevision = New HcoRevisiones
        objRevision.GuardoLasLecturas = True 'para que guarde la linea por cada una
        
        
    
        '---------------------------------------------------
        '----------------------------------------------------
                
        
        Do
        
        Loop Until Timer - T1 > 1
        
        PonerVisibles True
        If Text1.Text <> "" Then Text2.SetFocus
    End If
End Sub

Private Sub PonerVisibles(Si As Boolean)
   
    Text1.Visible = Si
    Text2.Visible = Si
    Label1(0).Visible = Si
    Label1(1).Visible = Si

End Sub

Private Sub Form_Load()

    CargaEntrada
    T1 = Timer
    Text2.Text = ""
    UltimoUsuario True
    PrimeraVez = True
    PonerVisibles False
End Sub
Private Sub CargaEntrada()
    If Dir(App.Path & "\entrada.dat", vbArchive) = "" Then
        MsgBox "Falta archivos configuracion entrada", vbExclamation
        End
    End If
    Me.Picture = LoadPicture(App.Path & "\entrada.dat")
End Sub


Private Sub Text1_GotFocus()
    With Text1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            KEYpress KeyAscii
            Validar
        End If
    End If
    
End Sub

Private Sub Text2_GotFocus()
    With Text2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    Else
        If KeyAscii = 13 Then
            KEYpress KeyAscii
            Validar
        End If
    End If
End Sub



Private Sub UltimoUsuario(Leer As Boolean)
Dim NF As Integer
Dim cad

    NF = FreeFile
    If Leer Then
        cad = ""
        If Dir(App.Path & "\ultusu.dat", vbArchive) <> "" Then
            Open App.Path & "\ultusu.dat" For Input As NF
            Line Input #NF, cad
            Close #NF
        End If
        Me.Text1.Tag = cad
        Me.Text1.Text = cad
    Else
        Open App.Path & "\ultusu.dat" For Output As NF
        Print #NF, Text1.Text
        Close #NF
    End If
End Sub
