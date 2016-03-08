VERSION 5.00
Begin VB.Form frmConfigPersonal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracion personal"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   Icon            =   "frmConfigPersonal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameMAil 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4215
      Left            =   240
      TabIndex        =   49
      Top             =   960
      Width           =   6855
      Begin VB.CheckBox Check3 
         Caption         =   "Ver pass"
         Height          =   195
         Left            =   4440
         TabIndex        =   61
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pasar a HCO"
         Height          =   735
         Left            =   2640
         TabIndex        =   57
         Top             =   3360
         Width           =   3855
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   2880
            TabIndex        =   60
            Text            =   "Text5"
            Top             =   330
            Width           =   495
         End
         Begin VB.OptionButton optPasarHco 
            Caption         =   "Meses"
            Height          =   195
            Index           =   1
            Left            =   1680
            TabIndex        =   59
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optPasarHco 
            Caption         =   "Manual"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   58
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   45
         Text            =   "Text4"
         Top             =   600
         Width           =   5055
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Comprobar al incio"
         Height          =   255
         Left            =   360
         TabIndex        =   55
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   48
         Text            =   "Text4"
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   47
         Text            =   "Text4"
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   46
         Text            =   "Text4"
         Top             =   1080
         Width           =   5055
      End
      Begin VB.Label Label5 
         Caption         =   "e-mail"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   56
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "MAIL - INTERNO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   1
         Left            =   0
         TabIndex        =   54
         Top             =   2880
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Password"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   53
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   52
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "E-MAIL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   51
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Servidor"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   50
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Frame FrameAridoc 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   240
      TabIndex        =   27
      Top             =   840
      Width           =   6975
      Begin VB.CheckBox Check1 
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   345
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   12
         Top             =   735
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2400
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   13
         Top             =   1095
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   14
         Top             =   1455
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   15
         Top             =   1815
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1800
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   16
         Top             =   2175
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   2160
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   17
         Top             =   2535
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   2400
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   2520
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   18
         Top             =   2895
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2880
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   19
         Top             =   3255
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   2400
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   3240
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   20
         Top             =   3615
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   2400
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   3600
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   21
         Top             =   3975
         Width           =   255
      End
      Begin VB.TextBox txtAncho 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   2400
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   3960
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Vista"
         Height          =   855
         Left            =   3360
         TabIndex        =   29
         Top             =   240
         Width           =   3375
         Begin VB.OptionButton optVista 
            Caption         =   "Iconos"
            Height          =   255
            Index           =   0
            Left            =   600
            TabIndex        =   31
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optVista 
            Caption         =   "Detalles"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5760
         TabIndex        =   28
         Text            =   "Text2"
         Top             =   1605
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   44
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   43
         Top             =   735
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   42
         Top             =   1095
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   41
         Top             =   1455
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   40
         Top             =   1815
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   39
         Top             =   2175
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   38
         Top             =   2535
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   37
         Top             =   2895
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   36
         Top             =   3255
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   35
         Top             =   3615
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   34
         Top             =   3975
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Ancho columna carpetas(%)"
         Height          =   195
         Left            =   3480
         TabIndex        =   33
         Top             =   1680
         Width           =   1980
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   6000
      TabIndex        =   26
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   25
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      MaxLength       =   40
      TabIndex        =   22
      Text            =   "Text3"
      Top             =   360
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Login"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label3 
      Caption         =   "Nombre"
      Height          =   195
      Left            =   2160
      TabIndex        =   23
      Top             =   120
      Width           =   1980
   End
End
Attribute VB_Name = "frmConfigPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte

Private Sub Check1_Click(Index As Integer)
    If Check1(Index).Value = 1 Then
        txtAncho(Index).Enabled = True
    Else
        txtAncho(Index).Enabled = False
    End If
End Sub

Private Sub Check3_Click()


    If Check3.Value Then
        Text4(2).PasswordChar = ""
    Else
        Text4(2).PasswordChar = "*"
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    If Opcion = 0 Then
    
        If DatosOk Then
            vUsu.preferencias.Modificar vUsu.codusu, False
            DatosMOdificados = True
            Unload Me
        End If
       
    Else
        If DatosOKMAil Then
            vUsu.ModificarDatosMail
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Limpiar Me
    'Pongo los valores en funcion de la configuracion personal
    If Opcion = 0 Then
        PonerTextoslabels
        FrameMAil.Visible = False
        Caption = "Configuración Vista Archivos"
    Else
        FrameMAil.Visible = True
        Caption = "Configuración MAIL"
    End If
    Text3.Locked = Opcion = 1
    PonerDatos
End Sub

Private Sub PonerDatos()
Dim I As Integer
    
    Text1.Text = vUsu.login
    Text3.Text = vUsu.Nombre
    If Opcion = 0 Then
        'Configuracion VISTA
            For I = 1 To Check1.Count - 1
                Check1(I).Value = 0
                txtAncho(I).Enabled = False
            Next I
            
            'Ponemos la preferencias
            With vUsu.preferencias
            
                Me.txtAncho(0).Text = .C1
                
                If .C2 > 0 Then
                    Check1(1).Value = 1
                    txtAncho(1).Text = .C2
                End If
                
                If .c3 > 0 Then
                    Check1(2).Value = 1
                    txtAncho(2).Text = .c3
                End If
                
                If .c4 > 0 Then
                    Check1(3).Value = 1
                    txtAncho(3).Text = .c4
                End If
                
                If .f1 > 0 Then
                    Check1(4).Value = 1
                    txtAncho(4).Text = .f1
                End If
                
                If .f2 > 0 Then
                    Check1(5).Value = 1
                    txtAncho(5).Text = .f2
                End If
                
                If .f3 > 0 Then
                    Check1(6).Value = 1
                    txtAncho(6).Text = .f3
                End If
                
                If .imp1 > 0 Then
                    Check1(7).Value = 1
                    txtAncho(7).Text = .imp1
                End If
                
                If .imp2 > 0 Then
                    Check1(8).Value = 1
                    txtAncho(8).Text = .imp2
                End If
                
                If .obs > 0 Then
                    Check1(9).Value = 1
                    txtAncho(9).Text = .obs
                End If
                
                If .tamayo > 0 Then
                    Check1(10).Value = 1
                    txtAncho(10).Text = .tamayo
                End If
                
                
                
                For I = 1 To Check1.Count - 1
                    If txtAncho(I).Text <> "" Then txtAncho(I).Enabled = True
                Next I
                
                'Ponemos la vista y el ancho
                Text2.Text = .Ancho
                
                If .Vista = lvwReport Then
                    optVista(1).Value = True
                Else
                    optVista(0).Value = True
                End If
                
            End With
    Else
    
        'Configuracion MAIL
        Text4(3).Text = vUsu.e_dir
        Text4(0).Text = vUsu.e_server
        Text4(1).Text = vUsu.e_login
        Text4(2).Text = vUsu.e_pwd
        
        If vUsu.preferencias.mailInicio Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If
        
        If vUsu.preferencias.mailPasarHCO = 0 Then
            optPasarHco(0).Value = True
        Else
            optPasarHco(1).Value = True
            
            Text5.Text = vUsu.preferencias.mailPasarHCO Mod 100
        
        End If
    End If
    
End Sub

Private Sub PonerTextoslabels()
    With vConfig
        Me.Label1(0).Caption = .C1
        Me.Label1(1).Caption = .C2
        Me.Label1(2).Caption = .c3
        Me.Label1(3).Caption = .c4
        Me.Label1(4).Caption = .f1
        Me.Label1(5).Caption = .f2
        Me.Label1(6).Caption = .f3
        Me.Label1(7).Caption = .imp1
        Me.Label1(8).Caption = .imp2
        Me.Label1(9).Caption = .obs
        Me.Label1(10).Caption = "Tamaño"
    End With
End Sub

Private Function DatosOk() As Boolean
Dim I As Integer
Dim J As Integer
    DatosOk = False
    
    
    'Comprobamos que esten marcados con valor
    For I = 0 To Check1.Count - 1
        If Check1(I).Value = 1 Then
            If txtAncho(I).Text = "" Then
                MsgBox "Campo " & Me.Label1(I).Caption & " esta marcado y debe tener valor.", vbExclamation
                Exit Function
            End If
        End If
    Next I
    
    'Comprobamos k los valores no superan los 10000
    For I = 0 To Check1.Count - 1
        If Check1(I).Value = 1 Then
            If txtAncho(I).Text = "" Then
                If Not IsNumeric(txtAncho(I).Text) Then
                    MsgBox "Campo " & Me.Label1(I).Caption & " debe ser numérico", vbExclamation
                    Exit Function
                Else
                    If Val(txtAncho(I).Text) > 10000 Then
                        MsgBox "Campo " & Me.Label1(I).Caption & " con valor excesivo", vbExclamation
                        
                        Exit Function
                    End If
                End If
            End If
        End If
    Next I
    
    
    If Val(txtAncho(0).Text) = 0 Then
        MsgBox "Campo: " & Label1(0).Caption & " debe tener valor mayor que cero", vbExclamation
        Exit Function
    End If
    If Text2.Text = "" Then
        MsgBox "Campo Ancho debe ser numérico", vbExclamation
        Exit Function
    End If
    If Not IsNumeric(Text2.Text) Then
        MsgBox "Campo ancho debe ser numerico", vbExclamation
        Exit Function
    End If
    
    If Val(Text2.Text) < 20 Then
        Text2.Text = "20"
    Else
        If Val(Text2.Text) > 80 Then
            Text2.Text = 80
        Else
            Text2.Text = Int(Text2.Text)
        End If
    End If
    
    'LLegados aqui asignamos las variables
     For J = 0 To Check1.Count - 1
        If Me.Check1(J).Value Then
            I = Val(Me.txtAncho(J).Text)
        Else
            I = 0
        End If
        AsignarAncho J + 1, I
    Next J
        
    vUsu.preferencias.Ancho = Val(Text2.Text)
    If Me.optVista(0).Value Then
        vUsu.preferencias.Vista = lvwIcon
    Else
        vUsu.preferencias.Vista = lvwReport
    End If
    
    DatosOk = True
End Function


Private Sub AsignarAncho(Campo As Integer, Ancho As Integer)
    With vUsu.preferencias
    Select Case Campo
    Case 1
        .C1 = Ancho
    Case 2
        .C2 = Ancho
    Case 3
        .c3 = Ancho
    Case 4
        .c4 = Ancho
    Case 5
        .f1 = Ancho
    Case 6
        .f2 = Ancho
    Case 7
        .f3 = Ancho
    Case 8
        .imp1 = Ancho
    Case 9
        .imp2 = Ancho
    Case 10
        .obs = Ancho
    Case 11
        .tamayo = Ancho
    End Select
        
    End With
End Sub

Private Sub Text2_GotFocus()
    With Text2
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = Trim(Text2.Text)
    If Text2.Text <> "" Then
        If Not IsNumeric(Text2.Text) Then
            MsgBox "Campo numérico", vbExclamation
            Text2.Text = ""
            Text2.SetFocus
        Else
            If Val(Text2.Text) > 100 Then
                MsgBox "Es un valor porcentual.", vbExclamation
                Text2.Text = ""
                Text2.SetFocus
            End If
        End If
    End If
End Sub

Private Sub txtAncho_GotFocus(Index As Integer)
    With txtAncho(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub txtAncho_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtAncho_LostFocus(Index As Integer)
    txtAncho(Index).Text = Trim(txtAncho(Index).Text)
    If txtAncho(Index).Text <> "" Then
        If Not IsNumeric(txtAncho(Index).Text) Then
            MsgBox "Campo numérico", vbExclamation
            txtAncho(Index).Text = ""
            txtAncho(Index).SetFocus
        End If
    End If
End Sub

Private Function DatosOKMAil() As Boolean
    DatosOKMAil = False
    If optPasarHco(1).Value Then
        If Text5.Text = "" Then
            MsgBox "Ha puesto pasar a hco por meses pero no ha indicado la cantidad", vbExclamation
            Exit Function
        End If
        If Not IsNumeric(Text5.Text) Then
            MsgBox "Campo numerico", vbExclamation
            Text5.Text = ""
            Text5.SetFocus
            Exit Function
        End If
        
        If Val(Text5.Text) < 1 Or Val(Text5.Text) > 12 Then
            MsgBox "El intervalo debe ser entre 1 y 12 meses", vbExclamation
            Text5.SetFocus
            Exit Function
        End If
        
        
        
    End If
    
        vUsu.e_dir = Text4(3).Text
        vUsu.e_server = Text4(0).Text
        vUsu.e_login = Text4(1).Text
        vUsu.e_pwd = Text4(2).Text
        
        vUsu.preferencias.mailInicio = Abs(Check2.Value)
            
        If optPasarHco(0).Value Then
            vUsu.preferencias.mailPasarHCO = 0
        Else
            vUsu.preferencias.mailPasarHCO = Val(Text5.Text)
        End If
    DatosOKMAil = True
    
    
    
End Function
