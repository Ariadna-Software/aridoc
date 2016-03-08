VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   Icon            =   "frmVarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrRevisiones 
      Height          =   6495
      Left            =   0
      TabIndex        =   74
      Top             =   0
      Width           =   8895
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3480
         Top             =   5880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVarios.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVarios.frx":6B6C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVarios.frx":D3CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVarios.frx":13C30
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVarios.frx":1A492
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVarios.frx":20CF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVarios.frx":27556
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmVarios.frx":2DDB8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   7
         Left            =   7560
         TabIndex        =   75
         Top             =   6000
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView3 
         Height          =   5415
         Left            =   120
         TabIndex        =   76
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Accion"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Usuario"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "PC"
            Object.Width           =   3598
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Cambios"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame FrameVolca 
      Height          =   2655
      Left            =   120
      TabIndex        =   65
      Top             =   0
      Width           =   6975
      Begin VB.TextBox txtVolcar 
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Text            =   "Text1"
         Top             =   1560
         Width           =   6615
      End
      Begin VB.TextBox txtVolcar 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   68
         Text            =   "Text1"
         Top             =   840
         Width           =   6615
      End
      Begin VB.CommandButton cmdVolca 
         Caption         =   "Volcar"
         Height          =   375
         Left            =   4560
         TabIndex        =   67
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   6
         Left            =   5760
         TabIndex        =   66
         Top             =   2040
         Width           =   975
      End
      Begin VB.Image imgGetFolder 
         Height          =   240
         Left            =   960
         Picture         =   "frmVarios.frx":3461A
         Top             =   1320
         Width           =   240
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   255
         Left            =   240
         TabIndex        =   73
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Label Label12 
         Caption         =   "Volcar estructura a disco."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   72
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label6 
         Caption         =   "Destino"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   71
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Carpeta ORIGEN"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   69
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame frConfiguracion 
      Height          =   6375
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8415
      Begin VB.TextBox txtClaves 
         Height          =   1005
         Index           =   11
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   63
         Tag             =   "|N|S|||timagen|Importe1|||"
         Text            =   "frmVarios.frx":3471C
         Top             =   4440
         Width           =   7935
      End
      Begin VB.TextBox txtClaves 
         Height          =   1005
         Index           =   10
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   61
         Tag             =   "|N|S|||timagen|Importe1|||"
         Text            =   "frmVarios.frx":34722
         Top             =   3120
         Width           =   7935
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   7080
         TabIndex        =   33
         Top             =   5760
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   6000
         TabIndex        =   32
         Top             =   5760
         Width           =   975
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   9
         Left            =   4200
         TabIndex        =   30
         Tag             =   "|N|S|||timagen|Tamnyo|||"
         Text            =   "Text3"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   28
         Tag             =   "|N|S|||timagen|Importe1|||"
         Text            =   "Text3"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Tag             =   "|N|S|||timagen|Importe1|||"
         Text            =   "Text3"
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   6
         Left            =   4200
         TabIndex        =   24
         Tag             =   "|F|S|||timagen|fecha3|||"
         Text            =   "Text3"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   22
         Tag             =   "|F|S|||timagen|fecha2|||"
         Text            =   "Text3"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Tag             =   "|F|S|||timagen|fecha1|||"
         Text            =   "99/99/9999"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   3
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   18
         Tag             =   "|T|S|||timagen|Campo4|||"
         Text            =   "Text3"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   2
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   16
         Tag             =   "|T|S|||timagen|Campo3|||"
         Text            =   "Text3"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   1
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   14
         Tag             =   "|T|S|||timagen|Campo2|||"
         Text            =   "Text3"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtClaves 
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "|T|S|||timagen|Campo1|||"
         Text            =   "Text3"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Ley proteccion de datos (II)"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   64
         Top             =   4200
         Width           =   2130
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   8160
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label3 
         Caption         =   "Ley proteccion de datos (I)"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   62
         Top             =   2880
         Width           =   1890
      End
      Begin VB.Line Line2 
         X1              =   120
         X2              =   8160
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8160
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label3 
         Caption         =   "Observaciones"
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   31
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Importe"
         Height          =   195
         Index           =   8
         Left            =   2160
         TabIndex        =   29
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label3 
         Caption         =   "Importe"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha3"
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   25
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha2"
         Height          =   195
         Index           =   5
         Left            =   2160
         TabIndex        =   23
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha1"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Campo4"
         Height          =   195
         Index           =   3
         Left            =   6240
         TabIndex        =   19
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Campo3"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Campo2"
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   15
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label3 
         Caption         =   "Campo1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame FrDatosAlmacen 
      Height          =   3855
      Left            =   720
      TabIndex        =   41
      Top             =   480
      Width           =   5775
      Begin VB.TextBox txtalma 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   42
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtalma 
         Height          =   285
         Index           =   1
         Left            =   3600
         TabIndex        =   43
         Text            =   "Text1"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtalma 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox txtalma 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   45
         Text            =   "Text1"
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtalma 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtalma 
         Height          =   285
         Index           =   5
         Left            =   3720
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   2280
         Width           =   1575
      End
      Begin VB.CommandButton cmdAlma 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   48
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmdAlma 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   49
         Top             =   3120
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   55
         Top             =   495
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Version"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   54
         Top             =   495
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Servidor"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   53
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Path"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   52
         Top             =   1695
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "User"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   51
         Top             =   2295
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Password"
         Height          =   255
         Index           =   5
         Left            =   2880
         TabIndex        =   50
         Top             =   2295
         Width           =   855
      End
   End
   Begin VB.Frame FrameAlamcen 
      Height          =   4695
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdUsuario 
         Height          =   375
         Index           =   3
         Left            =   3840
         Picture         =   "frmVarios.frx":34728
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Verificar"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUsuario 
         Height          =   375
         Index           =   0
         Left            =   2160
         Picture         =   "frmVarios.frx":3512A
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Nuevo"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUsuario 
         Height          =   375
         Index           =   1
         Left            =   2640
         Picture         =   "frmVarios.frx":3522C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Modificar"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdUsuario 
         Height          =   375
         Index           =   2
         Left            =   3120
         Picture         =   "frmVarios.frx":3532E
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Eliminar"
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   2
         Left            =   6120
         TabIndex        =   37
         Top             =   4200
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3255
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Servidor"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Path"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Label Label4 
         Caption         =   "Carpeta almacen de datos"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame FrCambioPwd 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   10
         Top             =   2880
         Width           =   975
      End
      Begin VB.CommandButton cmdPwd 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   2880
         Width           =   975
      End
      Begin VB.TextBox txtpwd 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2040
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtpwd 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2040
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtpwd 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   2040
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Reescriba el pwd nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nuevo"
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
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Actual"
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
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3720
         Picture         =   "frmVarios.frx":35430
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame FrameErrores 
      Height          =   5415
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton cmdEliminarEntradaErronea 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   4800
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Index           =   5
         Left            =   7680
         TabIndex        =   59
         Top             =   4800
         Width           =   975
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   4335
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Campo1"
            Object.Width           =   4586
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha1"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Carpeta"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   1920
         Picture         =   "frmVarios.frx":35872
         Top             =   4920
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1440
         Picture         =   "frmVarios.frx":359BC
         Top             =   4920
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte

    '0.-
    '1.-
    
    '2.-  Carpetas alamcen


    '5.- Archivos erroneos en las carpetas
    
    '6.- Volcar estructura sobre un soporte
    
    '7.- Ver revisiones
    
Dim i As Integer
Dim cad As String




Private Sub cmdAlma_Click(Index As Integer)
    If Index = 0 Then
    
        If Not datosAlmaOK Then Exit Sub
        
        'Si llega aqui, o insertaremos o modificaremos
        If Not InsertarModificarAlmacen Then Exit Sub
        
        CargaCarpetasAlmacen
    End If
    PonerDatosCarpetaAlmacen False, True
End Sub

Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdEliminarEntradaErronea_Click()
Dim C As String
    
    C = ""
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Checked Then
            C = "OK"
            Exit For
        End If
    Next i
    
    If C = "" Then
        MsgBox "Seleccione alguna entrada para eliminar", vbInformation
        Exit Sub
    End If
    
    
    C = "Seguro que desea eliminar de la BD los registros seleccionados?"
    If MsgBox(C, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    C = "DELETE from TIMAGEN where codigo ="
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).Checked Then
            Conn.Execute C & ListView2.ListItems(i).Text
        End If
    Next i
    
    DatosCopiados = "ELI"
    Unload Me
End Sub

Private Sub cmdPwd_Click()
    For i = 0 To 2
        Me.txtpwd(i).Text = Trim(Me.txtpwd(i).Text)
        If Me.txtpwd(i).Text = "" Then
            MsgBox "Campos obligatorios", vbExclamation
            Exit Sub
        End If
    Next i
    
    
    If Me.txtpwd(0).Text <> vUsu.Password Then
        MsgBox "Password actual incorrecto", vbExclamation
        Exit Sub
    End If
    
    
    If txtpwd(1).Text <> txtpwd(2) Then
        MsgBox "Campos para el nuevo password NO coinciden", vbExclamation
        Exit Sub
    End If
    
    
    vUsu.Password = txtpwd(1).Text
    If vUsu.Modificar = 0 Then
        MsgBox "Cambio realizado", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmdUsuario_Click(Index As Integer)

    If Index > 0 Then
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        If Val(ListView1.SelectedItem.Text) = 0 Then
            MsgBox "El almacen 0 no puede modificarse por programa", vbExclamation
            Exit Sub
        End If
    End If
    
    Select Case Index
    Case 0, 1
            PonerDatosCarpetaAlmacen True, Index = 0
            If Index = 0 Then txtalma(0).SetFocus
    
     
    Case 2
        DatosCopiados = "Desea eliminar el almacen " & ListView1.SelectedItem.Text & " ?"
        If MsgBox(DatosCopiados, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        
        'Compruebo k no esta referenciado por ninguna carpeta
        DatosCopiados = DevuelveDesdeBD("codcarpeta", "carpetas", "almacen", ListView1.SelectedItem.Text, "N")
        If DatosCopiados <> "" Then
            DatosCopiados = "Existen carpetas en ese almacen. " & DatosCopiados
            MsgBox DatosCopiados, vbExclamation
            Exit Sub
        End If
        
        
        'Llegados aqui lo eliminamos
        DatosCopiados = "MAL"
        VerificarAlma 1
        If DatosCopiados <> "" Then
            DatosCopiados = "No deberia eliminar el almacen del sistema. Desea continuar?"
            If MsgBox(DatosCopiados, vbQuestion + vbYesNo) = vbNo Then Exit Sub
        End If
        
        DatosCopiados = "Delete from almacen where codalma = " & ListView1.SelectedItem.Text
        Conn.Execute DatosCopiados
        
        'refrescamos
        CargaCarpetasAlmacen
        
    Case 3
        DatosCopiados = "MAL"
        VerificarAlma 0
        If DatosCopiados = "" Then MsgBox "Almacen OK", vbExclamation
            
    End Select
    
    
End Sub


Private Sub PonerDatosCarpetaAlmacen(Visible As Boolean, Nuevo As Boolean)
    Me.FrDatosAlmacen.Visible = Visible
    Me.FrameAlamcen.Enabled = Not Visible
    cmdCancelar(2).Enabled = Not Visible
    
    If Visible Then
        
        If Nuevo Then
            Limpiar Me
            'NUEVO
            'pongo el codigo y punto
                        
        Else
            With ListView1.SelectedItem
                txtalma(0).Text = .Text
                txtalma(1).Text = RecuperaValor(.Tag, 1)
                txtalma(2).Text = .SubItems(1)
                txtalma(3).Text = .SubItems(2)
                txtalma(4).Text = RecuperaValor(.Tag, 2)
                txtalma(5).Text = RecuperaValor(.Tag, 3)
                
            End With
        End If
        txtalma(0).Enabled = Nuevo
        txtalma(1).Enabled = Nuevo
        txtalma(2).Enabled = Nuevo
        txtalma(3).Enabled = Nuevo
    End If

End Sub



Private Sub cmdVolca_Click()
    txtVolcar(1).Text = Trim(txtVolcar(1).Text)
    If txtVolcar(1).Text = "" Then Exit Sub
    
    If Dir(txtVolcar(1).Text, vbDirectory) = "" Then
        MsgBox "No es una ruta válida", vbExclamation
        Exit Sub
    End If
    
    cad = txtVolcar(1).Text & "\" & NombreMSDOS(Admin.TreeView1.SelectedItem.Text)
    If Dir(cad, vbDirectory) <> "" Then
        MsgBox "Ya existe la carpeta: " & cad, vbExclamation
        Exit Sub
    End If
    cad = "Seguro que desea a llevar la estructura:" & vbCrLf & txtVolcar(0).Text & vbCrLf & vbCrLf & "sobre la carpeta :     " & txtVolcar(1).Text & "?"
    If MsgBox(cad, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    
    Label7.Caption = "Comienzo del proceso"
    HacerVolcarEstructura
    Label7.Caption = "Finalizando proceso"
    Me.Refresh
    espera 1
    Label7.Caption = ""
    
End Sub

Private Sub Command1_Click()

    If MsgBox("Seguro que desea guardar la configuración?", vbQuestion + vbYesNoCancel) = vbYes Then
        For i = 0 To 9
            txtClaves(i).Text = Trim(txtClaves(i).Text)
            If txtClaves(i).Text = "" Then txtClaves(i).Text = Label3(i).Caption
        Next i
        PonerDatosConfiguracion False
        
        If vConfig.Modificar = 0 Then
            MsgBox "Debe reiniciar la apliacion para que los cambios tengan efecto", vbExclamation
            Unload Me
        End If
        
    End If
End Sub



    '0.-Cambio password
    

Private Sub Form_Load()
Dim H As Integer
Dim W As Integer
    
    Limpiar Me
    Me.FrCambioPwd.Visible = False
    Me.frConfiguracion.Visible = False
    Me.FrameAlamcen.Visible = False
    FrDatosAlmacen.Visible = False
    FrameErrores.Visible = False
    FrameVolca.Visible = False
    FrRevisiones.Visible = False
    Select Case Opcion
    Case 0
        Label2.Caption = vUsu.Nombre
        Caption = "Cambio de CLAVE"
        FrCambioPwd.Visible = True
        H = FrCambioPwd.Height
        W = FrCambioPwd.Width
        
    Case 1
        PonerDatosConfiguracion True
        frConfiguracion.Visible = True
        H = frConfiguracion.Height + 120
        W = frConfiguracion.Width
        Caption = "Configuracion"
        Command1.Visible = vUsu.Nivel < 2
    Case 2
        H = Me.FrameAlamcen.Height
        W = Me.FrameAlamcen.Width
        Me.FrameAlamcen.Visible = True
        CargaCarpetasAlmacen
        Caption = "Carpetas almacén"
    Case 5
        H = Me.FrameErrores.Height
        W = Me.FrameErrores.Width
        Me.FrameErrores.Visible = True
        CargaArchivosErrores
        Caption = "Errores verificacion"
    
    
    Case 6
        H = Me.FrameVolca.Height + 150
        W = Me.FrameVolca.Width + 90
        Me.FrameVolca.Visible = True
        Caption = "Volcar estructura y archivos"
        Label7.Caption = ""
        Me.txtVolcar(0).Text = Admin.TreeView1.SelectedItem.FullPath
        DatosCopiados = ""
    
    Case 7
        'REVISIONES
        H = Me.FrRevisiones.Height + 150
        W = Me.FrRevisiones.Width + 90
        Me.FrRevisiones.Visible = True
        Caption = "Accesos documento"
        
        cargarRevisiones
    End Select
    Me.cmdCancelar(Opcion).Cancel = True
    Me.cmdCancelar(Opcion).Default = True
    
    Me.Width = W + 120
    Me.Height = H + 360
End Sub

Private Sub Image2_Click(Index As Integer)
    For i = 1 To ListView2.ListItems.Count
        ListView2.ListItems(i).Checked = Index = 0
    Next i
End Sub

Private Sub imgGetFolder_Click()
    cad = GetFolder("Carpeta destino")
    If cad <> "" Then txtVolcar(1).Text = cad

End Sub

Private Sub ListView3_DblClick()
    If ListView3.SelectedItem Is Nothing Then Exit Sub
    
    i = CInt(ListView3.SelectedItem.SmallIcon) - 1   'Pq le sumo uno para el icono
    cad = i & "|"
    If InStr(1, "1|5|6|7|", cad) > 0 Then
        With ListView3.SelectedItem
            If .Tag <> "" Then
        
                cad = "Fecha: " & .SubItems(1) & vbCrLf
                cad = cad & "Usuario: " & .SubItems(2) & vbCrLf
                cad = cad & "Equipo: " & .SubItems(3) & vbCrLf
                cad = cad & "------------------------------" & vbCrLf & vbCrLf & .Tag
                MsgBox cad, vbInformation
                
            End If
        End With
    End If
End Sub

Private Sub txtalma_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub

Private Sub txtpwd_GotFocus(Index As Integer)
    With txtpwd(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtpwd_KeyPress(Index As Integer, KeyAscii As Integer)
    KEYpress KeyAscii
End Sub


Private Sub PonerDatosConfiguracion(Leer As Boolean)
    With vConfig
        If Leer Then
    
                txtClaves(0).Text = .C1
                txtClaves(1).Text = .C2
                txtClaves(2).Text = .c3
                txtClaves(3).Text = .c4
                txtClaves(4).Text = .f1
                txtClaves(5).Text = .f2
                txtClaves(6).Text = .f3
                txtClaves(7).Text = .imp1
                txtClaves(8).Text = .imp2
                txtClaves(9).Text = .obs
                txtClaves(10).Text = .LeyProtDatos1
                txtClaves(11).Text = .LeyProtDatos2
        Else
            
                .C1 = txtClaves(0).Text
                .C2 = txtClaves(1).Text
                .c3 = txtClaves(2).Text
                .c4 = txtClaves(3).Text
                .f1 = txtClaves(4).Text
                .f2 = txtClaves(5).Text
                .f3 = txtClaves(6).Text
                .imp1 = txtClaves(7).Text
                .imp2 = txtClaves(8).Text
                .obs = txtClaves(9).Text
                .LeyProtDatos1 = txtClaves(10).Text
                .LeyProtDatos2 = txtClaves(11).Text
        
        End If
    End With
    
End Sub

Private Sub CargaCarpetasAlmacen()
Dim Itm As ListItem

    ListView1.ListItems.Clear
    DatosCopiados = "Select * from almacen order by codalma"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open DatosCopiados, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Set Itm = ListView1.ListItems.Add(, "C" & CStr(miRSAux!codalma))
        Itm.Text = miRSAux!codalma
        Itm.SubItems(1) = miRSAux!SRV
        Itm.SubItems(2) = miRSAux!pathreal
        
        Itm.Tag = miRSAux!version & "|" & DBLet(miRSAux!user, "T") & "|" & DBLet(miRSAux!pwd, "T") & "|"
        'Siguiente
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
End Sub

Private Function datosAlmaOK() As Boolean
    datosAlmaOK = False
    
    For i = 0 To 5
        Me.txtalma(i).Text = Trim(txtalma(i).Text)
        If i < 4 Then
            If txtalma(i).Text = "" Then
                MsgBox "Campo " & Label5(i).Caption & " requerido", vbExclamation
                Exit Function
            End If
        End If
    Next i
    
    If txtalma(0).Enabled Then
        txtalma(0).Tag = Val(txtalma(0).Text)
        DatosCopiados = DevuelveDesdeBD("codalma", "almacen", "codalma", txtalma(0).Tag, "N")
        If DatosCopiados <> "" Then
            MsgBox "Ya existe el almacen: " & txtalma(0).Text & "- " & txtalma(0).Tag, vbExclamation
            Exit Function
        End If
    End If
        
    datosAlmaOK = True
    
End Function



Private Function InsertarModificarAlmacen() As Boolean
    On Error GoTo EInsertarModificarAlmacen
    InsertarModificarAlmacen = False
    If txtalma(0).Enabled Then
        'NUEVO
        DatosCopiados = "INSERT INTO almacen (codalma, version,  SRV,pathreal, user, pwd) VALUES ("
        DatosCopiados = DatosCopiados & txtalma(0).Text & "," & txtalma(1).Text & ",'"
        DatosCopiados = DatosCopiados & txtalma(2).Text & "','" & DevNombreSql(txtalma(3).Text) & "','"
        DatosCopiados = DatosCopiados & txtalma(4).Text & "','" & txtalma(5).Text & "')"
    Else
        'MODIFICAR
        DatosCopiados = "UPDATE almacen SET user='" & txtalma(4).Text & "', pwd='" & txtalma(5).Text & "'"
        DatosCopiados = DatosCopiados & " WHERE codalma = " & txtalma(0).Text
    End If
    
    
    Conn.Execute DatosCopiados
    InsertarModificarAlmacen = True
    Exit Function
EInsertarModificarAlmacen:
    MuestraError Err.Number, "InsertarModificarAlmacen"
End Function

'0.- Solo verifica
'2.-  verifica y dir
Private Sub VerificarAlma(miOpcion As Byte)
    Me.FrameAlamcen.Tag = ListView1.SelectedItem.SubItems(1) & "|" & ListView1.SelectedItem.SubItems(2) & "|"
    Me.FrameAlamcen.Tag = Me.FrameAlamcen.Tag & ListView1.SelectedItem.Tag
      
   
    frmMovimientoArchivo.Opcion = 10 + miOpcion
    frmMovimientoArchivo.Origen = FrameAlamcen.Tag
    frmMovimientoArchivo.Show vbModal

End Sub

'En la verificacion de ficheros en una carpeta por parte del root
'si se han producido errores se muestra esto
Private Sub CargaArchivosErrores()
Dim It As ListItem
    Set miRSAux = New ADODB.Recordset
    
    Set ListView2.SmallIcons = Admin.ImageList2
    ListView2.ListItems.Clear
    'La columna
    ListView2.ColumnHeaders(2).Text = vConfig.C1
    ListView2.ColumnHeaders(3).Text = vConfig.f1
    
    DatosCopiados = "select codigo,campo1,fecha1,carpetas.codcarpeta,nombre,codext from tmpfich,timagen,carpetas"
    DatosCopiados = DatosCopiados & " Where codusu = " & vUsu.codusu
    DatosCopiados = DatosCopiados & " And codequipo = " & vUsu.PC
    DatosCopiados = DatosCopiados & " AND tmpfich.imagen =codigo AND carpetas.codcarpeta=timagen.codcarpeta"
    miRSAux.Open DatosCopiados, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    DatosCopiados = ""
    While Not miRSAux.EOF
        Set It = ListView2.ListItems.Add(, "C" & miRSAux!codigo)
        It.Text = miRSAux!codigo
        It.SubItems(1) = miRSAux!campo1
        It.SubItems(2) = Format(miRSAux!fecha1, "dd/mm/yyyy")
        It.SubItems(3) = miRSAux!Nombre
        It.SmallIcon = miRSAux!codext + 1
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
End Sub

Private Function BorrarFicheroErrores() As Boolean
    BorrarFicheroErrores = False
    On Error GoTo EBorrarFicheroErrores
    If Dir(App.Path & "\volcaest.err", vbArchive) <> "" Then Kill App.Path & "\volcaest.err"
    BorrarFicheroErrores = True
    Exit Function
EBorrarFicheroErrores:
    MuestraError Err.Number, Err.Description
End Function
Private Sub HacerVolcarEstructura()
Dim N As Node
Dim miNF As Integer  'Fichero para los errores del volcado
Dim F As Date

    On Error GoTo EHacerVolcarEstructura
    
    If Not BorrarFicheroErrores Then Exit Sub
    'Pongo el nodo que voy a empezar a volcar
    Set N = Admin.TreeView1.SelectedItem
    
    F = Now
    
    Me.cmdVolca.Enabled = False
    cmdCancelar(6).Enabled = False
    Me.Refresh
    espera 0.2
    Screen.MousePointer = vbHourglass
    Set miRSAux = New ADODB.Recordset
   
    Set listaimpresion = New Collection
    ProcesarCarpeta txtVolcar(1).Text, N
    Set N = Nothing
    If listaimpresion.Count > 0 Then
        'Se han producido errores
        If listaimpresion.Count > 5 Then
            'VOLCARLOS SOBRE FICHEROS
            miNF = FreeFile
            Open App.Path & "\volcaest.err" For Output As #miNF
            
            Print #miNF, "Proceso inciado: " & Format(F, "Long date")
            Print #miNF, "Proceso finalizado: " & Format(Now, "Long date")
            Print #miNF, "Carpeta incial: " & Admin.TreeView1.SelectedItem.FullPath
            Print #miNF, "": Print #miNF, "": Print #miNF, "": Print #miNF, "":
            Print #miNF, "Errores: "
            For i = 1 To listaimpresion.Count
                Label7.Caption = i & " de " & listaimpresion.Count
                Label7.Refresh
                Print #miNF, listaimpresion.Item(i)
            Next i
            Close #miNF
            LanzaNotePad "notepad.exe " & App.Path & "\volcaest.err"
        Else
            cad = "ERRORES." & vbCrLf & vbCrLf
            For i = 1 To listaimpresion.Count
                cad = cad & listaimpresion.Item(i) & vbCrLf
            Next i
            MsgBox cad, vbExclamation
        End If
    Else
        MsgBox "Proceso finalizado", vbInformation
    End If
    Set listaimpresion = Nothing
    
    
EHacerVolcarEstructura:
    If Err.Number <> 0 Then MuestraError Err.Number
'    Me.cmdVolca.Enabled = True
    cmdCancelar(6).Enabled = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub LanzaNotePad(CADENA As String)
    On Error Resume Next
    Shell CADENA, vbNormalFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

'NombreArchivo = ListView1.SelectedItem.Text
'            Do
'                i = InStr(1, NombreArchivo, " ")
'                If i > 0 Then NombreArchivo = Mid(NombreArchivo, 1, i - 1) & Mid(NombreArchivo, i + 1)
'            Loop Until i = 0
'            i = 0
'            Do
'                Cad = App.Path & "\temp\" & NombreArchivo
'                If i > 0 Then Cad = Cad & "(" & i & ")"
'                Cad = Cad & "." & vE.Extension
'                i = i + 1
'            Loop Until Dir(Cad, vbArchive) = "" Or i > 100
'            If i > 100 Then
'                MsgBox "Error obteniendo nombre fichero(100)", vbExclamation
'                Exit Sub
'            End If


Private Function NombreMSDOS(ByVal Origen As String) As String
Dim K As Integer
Dim Kk As Integer
Dim Caracteres As String
Dim Ch As String

    'Aqui haremos cambios
    '
    'No puede contener:
    '                    \ /  :  *  ?  "  <  >  |
    Caracteres = "\/:*?""<>|"
    For K = 1 To 9
        Do
            Ch = Mid(Caracteres, K, 1)
            Kk = InStr(1, Origen, Ch)
            If Kk > 0 Then Origen = Mid(Origen, 1, Kk - 1) & Mid(Origen, Kk + 1)
        Loop Until Kk = 0
    Next K
    NombreMSDOS = Origen
End Function


Private Sub ProcesarCarpeta(ByVal camino As String, vNo As Node)
Dim CO As Ccarpetas
Dim DES As Ccarpetas
Dim Naux As Node
Dim Fin As Boolean
    Label7.Caption = "Carpeta: " & vNo.Text
    Label7.Refresh
    Screen.MousePointer = vbHourglass
    Set CO = New Ccarpetas
    If CO.Leer(Mid(vNo.Key, 2), False) = 1 Then
        MsgBox "Error leyendo carpeta: " & vNo.Text & " - " & vNo.Key, vbExclamation
        Exit Sub
    End If
    
    'Creamos carpeta
    MkDir camino & "\" & NombreMSDOS(vNo.Text)
    
    'Copio los archivos
    
    cad = "Delete from timagenhco where codequipo =" & vUsu.PC
    Conn.Execute cad
    
    cad = "Select " & vUsu.PC & ",codigo,campo1,codext"
    cad = cad & " from timagen"
    cad = cad & " WHERE "
    'Carpeta
    cad = cad & " codcarpeta = " & Mid(vNo.Key, 2)
    'Es el usuario propietario
    If vUsu.codusu > 0 Then
        cad = cad & " AND (userprop = " & vUsu.codusu
        
        'O el grupo tiene permiso
        cad = cad & " OR (lecturag & " & vUsu.Grupo & "))"
    End If
    
    cad = "INSERT INTO timagenhco(codequipo,codigo,campo1,codext)  " & cad
    Conn.Execute cad
    cad = "Select timagenhco.*,exten from timagenhco,extension where"
    cad = cad & " timagenhco.codext=extension.codext"
    cad = cad & " AND codequipo=" & vUsu.PC
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    If Not miRSAux.EOF Then
    
        '---------------------
        'Tiene datos
        
        Set DES = New Ccarpetas
        DES.version = 1
        DES.pathreal = camino & "\" & NombreMSDOS(vNo.Text)
        
        frmMovimientoArchivo.Opcion = 17
        Set frmMovimientoArchivo.vDestino = DES
        Set frmMovimientoArchivo.vOrigen = CO
        frmMovimientoArchivo.Show vbModal
        
        Label7.Caption = "Siguiente"
        Me.Refresh
        DoEvents
        espera 1
        
    End If
    Set CO = Nothing
    miRSAux.Close
    
    Label7.Caption = "................ "
    Me.Refresh
    espera 1
    DoEvents
    Label7.Caption = "Obteniendo subcarpetas ... "
    Me.Refresh
    'Para cad hijo llamare al proceso
    If vNo.Children = 0 Then Exit Sub  'No tiene hijos
        
        
    Set Naux = vNo.Child
    Fin = False
    Do
        ProcesarCarpeta camino & "\" & NombreMSDOS(vNo.Text), Naux
        If Naux.Next Is Nothing Then
            Fin = True
        Else

                Set Naux = Naux.Next
  
        End If
    Loop Until Fin
End Sub


Private Sub cargarRevisiones()
Dim Rs As ADODB.Recordset
Dim It As ListItem

    Set ListView3.SmallIcons = Me.ImageList1
    
    cad = "SELECT revision.*, usuarios.Nombre, equipos.descripcion"
    cad = cad & " FROM (revision LEFT JOIN usuarios ON revision.usuario = usuarios.codusu) LEFT JOIN equipos ON revision.pc = equipos.codequipo"
    cad = cad & " WHERE revision.id = " & listacod.Item(1) & " ORDER BY fecha"
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        Set It = ListView3.ListItems.Add()
        It.Text = objRevision.DevuelveTextoRev(CInt(Rs!accion))
        It.SubItems(1) = Format(Rs!Fecha, "dd/mm/yyyy  hh:mm")
        It.SubItems(2) = DBLet(Rs!Nombre, "T")
        It.SubItems(3) = DBLet(Rs!Descripcion, "T")
        It.SmallIcon = CInt(Rs!accion) + 1
        It.Tag = DBLet(Rs!Cambios)
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    Screen.MousePointer = vbDefault
End Sub
