VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Admin 
   BackColor       =   &H8000000A&
   Caption         =   "Administrador de Documentos"
   ClientHeight    =   7155
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7155
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameSeparador 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      DragMode        =   1  'Automatic
      Height          =   3015
      Left            =   2520
      MousePointer    =   9  'Size W E
      TabIndex        =   7
      Top             =   3960
      Width           =   45
   End
   Begin VB.Timer TimerTree 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   480
      Top             =   6720
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8520
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nueva carpeta"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "borrar carpeta"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "copiar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "cortar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "lista"
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "iconos"
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "nueva imagen"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "propiedades"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "eliminar archivo"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "buscar"
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "MAIL"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   285
         Left            =   5715
         TabIndex        =   2
         Text            =   "d"
         Top             =   50
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Carpeta"
         Height          =   375
         Left            =   5280
         TabIndex        =   3
         Top             =   0
         Width           =   495
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      DragIcon        =   "Form2.frx":030A
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   1140
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   9340
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   0
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5355
      Left            =   2640
      TabIndex        =   6
      Top             =   1140
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9446
      View            =   2
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   8520
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":074C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2456
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":86FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":910E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":E900
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":110B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1198C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":12266
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":12B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1341A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1957C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":199D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":19AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":19BFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":19D0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1A026
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1FC48
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2543A
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":25E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":2C0E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":318D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":370CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":373E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":376FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgUsersPCs 
      Left            =   8520
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":3C830
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":3EFE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":45288
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4539A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":454AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":45A46
            Key             =   "find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":45B58
            Key             =   "abierto"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":45E72
            Key             =   "cerrado"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4618C
            Key             =   "importar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4B97E
            Key             =   "cortar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4BEC0
            Key             =   "pegar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4C402
            Key             =   "carpeta"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4C514
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4C626
            Key             =   "lista"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4C738
            Key             =   "iconos"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4C84A
            Key             =   "prop"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4C95C
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4CA6E
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4CB80
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4CC92
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4D0E4
            Key             =   "tienemail"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":4E876
            Key             =   "v_cerrado"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":550D8
            Key             =   "v_abierto"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":5B93A
            Key             =   "falta"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   8520
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageListMAIL 
      Left            =   8520
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   29
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":6219C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":625EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":62A40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":62E92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":632E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":63736
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":63B88
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":63FDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":6442C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":6A6C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":6B0D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":71372
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":77BD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":7E436
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":84C98
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":8B4FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":91D5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":985BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":98A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":98E62
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":992B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":99706
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":99B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":99FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":9FBCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":A0A1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":A0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":A1052
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":A136C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H004D2C1D&
      Caption         =   "AriDoc: Gestión documental"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   2460
      TabIndex        =   4
      Top             =   435
      Width           =   9435
   End
   Begin VB.Image Image1 
      Height          =   570
      Left            =   120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label Label2 
      BackColor       =   &H004D2C1D&
      Height          =   690
      Left            =   0
      TabIndex        =   5
      Top             =   435
      Width           =   2535
   End
   Begin VB.Menu mnArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnNueva2 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnNuevoInsertar 
         Caption         =   "Insertar"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnImportar 
         Caption         =   "&Modificar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnImprimir 
         Caption         =   "Im&primir"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnbarra29 
         Caption         =   "-"
      End
      Begin VB.Menu mnCambiarImpresora 
         Caption         =   "Cambiar impresora"
      End
      Begin VB.Menu mnbarra_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnedicion 
      Caption         =   "&Edición"
      Begin VB.Menu mncortar2 
         Caption         =   "&Cortar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnPegar 
         Caption         =   "&Pegar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnbarra6 
         Caption         =   "-"
      End
      Begin VB.Menu mnselectodo 
         Caption         =   "&Seleccionar todo"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnquitarsel 
         Caption         =   "&Quitar selección"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnportipo 
         Caption         =   "S&eleccionar por tipo archivo"
      End
      Begin VB.Menu mnbarra19 
         Caption         =   "-"
      End
      Begin VB.Menu mnRefrescar 
         Caption         =   "Actualizar"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnMensajes 
      Caption         =   "Mensajes"
      Begin VB.Menu mnNuevoMensaje 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnMensajesCol 
         Caption         =   "Enviados / Recibidos"
      End
      Begin VB.Menu mnBarramail 
         Caption         =   "-"
      End
      Begin VB.Menu mnConfigMAIL 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnTiposMensaje 
         Caption         =   "Tipos mensaje"
      End
   End
   Begin VB.Menu mnCopiaS 
      Caption         =   "&Utilidades"
      Begin VB.Menu mnComprCarpeta 
         Caption         =   "&Comprimir archivos"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnPaso 
         Caption         =   "&Pasar archivos a Histórico"
      End
      Begin VB.Menu mnCopiaSeg 
         Caption         =   "C&opia Seguridad"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnVolcarEstructura 
         Caption         =   "Volcar estructura"
      End
      Begin VB.Menu mnintegrar 
         Caption         =   "&Realizar Integracion"
      End
      Begin VB.Menu nInsertarMasiva 
         Caption         =   "Inserción masiva de datos"
         Begin VB.Menu mnInsertarmasivaCarpeta 
            Caption         =   "Carpetas / subcarpetas"
         End
         Begin VB.Menu mnImportarVarios 
            Caption         =   "&Importar por Lotes"
         End
      End
      Begin VB.Menu mnbarra20 
         Caption         =   "-"
      End
      Begin VB.Menu mnImportes 
         Caption         =   "Importes"
         Begin VB.Menu mnSumaArchivos 
            Caption         =   "Archivos seleccionados"
         End
         Begin VB.Menu mnSumaCarpetaActual 
            Caption         =   "Carpeta actual"
         End
         Begin VB.Menu mnSumasSubcarpetas 
            Caption         =   "Carpeta y subcarpetas"
         End
      End
      Begin VB.Menu mnbarra11 
         Caption         =   "-"
      End
      Begin VB.Menu mnCerrarHistorico 
         Caption         =   "Cerrar Historico"
         Visible         =   0   'False
      End
      Begin VB.Menu mnRestaurar 
         Caption         =   "Recuperar Historico"
      End
   End
   Begin VB.Menu mnMenuCopiaSeguridad 
      Caption         =   "Copia seguridad"
      Begin VB.Menu mnRecuperaDesdeBackUp 
         Caption         =   "Recuperar archivos Backup"
      End
   End
   Begin VB.Menu mnConfig 
      Caption         =   "&Configuración"
      Begin VB.Menu mnPreferencias 
         Caption         =   "Preferencias"
      End
      Begin VB.Menu mnTipoArchivos 
         Caption         =   "Tipos archivos"
      End
      Begin VB.Menu mnPlantillas 
         Caption         =   "Plantillas"
      End
      Begin VB.Menu mnbarra105 
         Caption         =   "-"
      End
      Begin VB.Menu mnCambioClave 
         Caption         =   "Cambiar password"
      End
      Begin VB.Menu mnbarra2101 
         Caption         =   "-"
      End
      Begin VB.Menu mnHerramientasadmon 
         Caption         =   "Herramientas administrativas"
         Begin VB.Menu mnParametros 
            Caption         =   "Parametros"
         End
         Begin VB.Menu mnbarra1 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnConfiguraExtenAridoc 
            Caption         =   "Configurar extensiones ARIDOC"
         End
         Begin VB.Menu mnVerificarAridoc 
            Caption         =   "Verificar ARIDOC"
         End
         Begin VB.Menu mnAlmacenDatos 
            Caption         =   "Almacén de datos"
         End
         Begin VB.Menu mnAdmonPlantillas 
            Caption         =   "Administrador plantillas"
         End
         Begin VB.Menu mnbarra103 
            Caption         =   "-"
         End
         Begin VB.Menu mnAdmonUsers 
            Caption         =   "Adminstrador de usuarios"
         End
         Begin VB.Menu mnGrupos 
            Caption         =   "Administrador de grupos"
         End
         Begin VB.Menu mnGestionEquipos 
            Caption         =   "Gestion equipos"
         End
      End
   End
   Begin VB.Menu mnAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnIndice 
         Caption         =   "&Indice"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnbarra122 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnAcerca 
         Caption         =   "&Acerca de ....."
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnMenuTree 
      Caption         =   "menutree"
      Visible         =   0   'False
      Begin VB.Menu Nuevo_Dire 
         Caption         =   "Crear Carpeta"
      End
      Begin VB.Menu Borrar_dir 
         Caption         =   "Borrar Carpeta"
      End
      Begin VB.Menu mnbarra101 
         Caption         =   "-"
      End
      Begin VB.Menu mnexpandirnodo 
         Caption         =   "Expandir nodo"
      End
      Begin VB.Menu mncontraernodo 
         Caption         =   "Contraer nodo"
      End
      Begin VB.Menu mnbarra100 
         Caption         =   "-"
      End
      Begin VB.Menu mnPropiedadesCarpeta 
         Caption         =   "Propiedades"
      End
      Begin VB.Menu mnPropAvanzadas 
         Caption         =   "Avanzadas"
      End
      Begin VB.Menu mnBarraVerificar 
         Caption         =   "-"
      End
      Begin VB.Menu mnVerficarCarpeta1 
         Caption         =   "Verificar"
         Begin VB.Menu mnVerifyCarpetaActual 
            Caption         =   "Carpeta actual"
         End
         Begin VB.Menu mnVeryfCarpSub 
            Caption         =   "Carpeta y subcarpetas"
         End
      End
   End
   Begin VB.Menu mnList 
      Caption         =   "menulist"
      Visible         =   0   'False
      Begin VB.Menu mncopiar 
         Caption         =   "pegar"
      End
      Begin VB.Menu mncortar 
         Caption         =   "cortar"
      End
      Begin VB.Menu mnbarra5 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevoArch 
         Caption         =   "Insertar"
      End
      Begin VB.Menu mnModicarArch 
         Caption         =   "Modificar"
      End
      Begin VB.Menu mnNuevoEditando 
         Caption         =   "Nuevo"
         Begin VB.Menu mnNuevoArchivoEditando 
            Caption         =   "Archivo"
         End
         Begin VB.Menu mnBarraNuevo 
            Caption         =   "-"
         End
         Begin VB.Menu mnNuevoN1 
            Caption         =   "n1"
            Index           =   1
         End
         Begin VB.Menu mnNuevoN1 
            Caption         =   "n2"
            Index           =   2
         End
         Begin VB.Menu mnNuevoN1 
            Caption         =   "n3"
            Index           =   3
         End
         Begin VB.Menu mnNuevoN1 
            Caption         =   "n4"
            Index           =   4
         End
         Begin VB.Menu mnNuevoN1 
            Caption         =   "n5"
            Index           =   5
         End
         Begin VB.Menu mnNuevoN1 
            Caption         =   "n6"
            Index           =   6
         End
         Begin VB.Menu mnbarra15 
            Caption         =   "-"
         End
         Begin VB.Menu mnInsertarDesdePlantilla 
            Caption         =   "Plantilla"
         End
      End
      Begin VB.Menu barra10 
         Caption         =   "-"
      End
      Begin VB.Menu mneliminar 
         Caption         =   "Eliminar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnbarra4 
         Caption         =   "-"
      End
      Begin VB.Menu mnEnviarPorMail 
         Caption         =   "Enviar por mail"
      End
      Begin VB.Menu mnbarra10 
         Caption         =   "-"
      End
      Begin VB.Menu mnselectall 
         Caption         =   "Seleccionar Todos"
      End
      Begin VB.Menu mnDeselectAll 
         Caption         =   "Quitar selección"
      End
      Begin VB.Menu mnBarra3 
         Caption         =   "-"
      End
      Begin VB.Menu mnPropiedades 
         Caption         =   "Propiedades"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnChkPropBarra 
         Caption         =   "-"
      End
      Begin VB.Menu mnChkPropItem 
         Caption         =   "Cambiar propietario"
      End
   End
   Begin VB.Menu mnArbolBK 
      Caption         =   "Backup"
      Visible         =   0   'False
      Begin VB.Menu mnRestaurarCarpeta 
         Caption         =   "Restaurar carpeta"
      End
   End
   Begin VB.Menu mnListBackUp 
      Caption         =   "Backup"
      Visible         =   0   'False
      Begin VB.Menu mnRestarurarArchivo 
         Caption         =   "Restaurar archivo"
      End
   End
End
Attribute VB_Name = "Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim Car As Ccarpetas
Private WithEvents frmP As frmPregunta
Attribute frmP.VB_VarHelpID = -1
Dim vOpcion As Byte


Dim NodoSeleccionado As Node
Dim NodoOrigen As Node


Dim CadenaCarpetas As String
Dim AnchoListview As Integer

'esto lo utilizamos para cortar y pegar
Dim Cortar11 As String
Dim pegar11 As String

Dim PrimeraVez As Boolean
Dim Base As Single
Dim OrderAscendente As Boolean
Dim ListviewSHIFTPresionado As Integer
Private CarpetasAbiertas As String
Dim T2 As Single

Dim ImpresoraAntesEntrar As String


'-------------------------
'-------------------------
'-------------------------
'-------------------------
'PRUEBA DE CARGA DEL ARBOL
'-------------------------
'-------------------------
'-------------------------
'-------------------------
'
Private Function INSERTAR_NODO(ByRef RSS As Recordset, SubNivel As Integer) As Integer
Dim XNodo As Node

On Error GoTo EIns_Nodo

    
    

    INSERTAR_NODO = -1
    If RSS!padre = 0 Then
        'NODO RAIZ
        Set XNodo = TreeView1.Nodes.Add(, tvwChild, "C" & RSS!codcarpeta)
    Else
    
        'NODO HIJO
        Set XNodo = TreeView1.Nodes.Add("C" & RSS!padre, tvwChild, "C" & RSS!codcarpeta)
    End If
    
    XNodo.Text = RSS!Nombre
    'En el tag metemos la seguriad
    XNodo.Tag = RSS!escriturau & "|" & RSS!escriturag & "|"
    
    
    'XNODO.Image = "cerrado"
    'XNODO.ExpandedImage = "abierto"
    If InStr(1, CarpetasAbiertas, "|" & XNodo.Key & "|") > 0 Then XNodo.Expanded = True
    'XNODO.Expanded = True
    CadenaCarpetas = CadenaCarpetas & Mid(XNodo.Key, 2) & "|"
    
    
    XNodo.Image = "v_cerrado"
    XNodo.ExpandedImage = "v_abierto"
'    If SubNivel > 4 Then
'        If Not XNodo.Expanded Then
'            XNodo.Image = "falta"
'            XNodo.ExpandedImage = "falta"
'        End If
'    Else
    If RSS!hijos > 0 Then INSERTAR_NODO = XNodo.Index
'    End If
Exit Function
EIns_Nodo:
    Cortar11 = "ERROR GRAVE." & vbCrLf & vbCrLf
    Cortar11 = Cortar11 & Err.Description & vbCrLf & vbCrLf
    Cortar11 = Cortar11 & RSS!codcarpeta & " " & DBLet(RSS!Nombre, "T")
   ' MsgBox Cortar11, vbCritical
    Cortar11 = Cortar11 & vbCrLf & vbCrLf
    Cortar11 = Cortar11 & "Verifique ARIDOC. Si persiste avise a soporte técnico"
    Cortar11 = Cortar11 & vbCrLf & vbCrLf & vbCrLf & "¿FINALIZAR?"
    If MsgBox(Cortar11, vbCritical + vbYesNo) = vbYes Then
        Conn.Close
        End
    End If
End Function





'Private Function INSERTAR_NODO_OLD(ByRef RSS As Recordset) As Integer
'Dim XNODO As Node
'
'On Error GoTo EIns_Nodo
'    INSERTAR_NODO_OLD = -1
'    If RSS!padre = 0 Then
'        'NODO RAIZ
'        Set XNODO = TreeView1.Nodes.Add(, tvwChild, "C" & RSS!codcarpeta)
'    Else
'        'NODO HIJO
'        Set XNODO = TreeView1.Nodes.Add("C" & RSS!padre, tvwChild, "C" & RSS!codcarpeta)
'    End If
'
'    XNODO.Text = RSS!Nombre
'    'En el tag metemos la seguriad
'    XNODO.Tag = RSS!escriturau & "|" & RSS!escriturag & "|"
'
'
'    'XNODO.Image = "cerrado"
'    'XNODO.ExpandedImage = "abierto"
'    XNODO.Image = "v_cerrado"
'    XNODO.ExpandedImage = "v_abierto"
'    If InStr(1, CarpetasAbiertas, "|" & XNODO.Key & "|") > 0 Then XNODO.Expanded = True
'
'
'    CadenaCarpetas = CadenaCarpetas & Mid(XNODO.Key, 2) & "|"
'
'
'    INSERTAR_NODO_OLD = XNODO.Index
'Exit Function
'EIns_Nodo:
'    'Stop
'End Function
'
'



Private Sub CargaArbol()
Dim cad As String
Dim Rs As ADODB.Recordset
Dim Nod As Node
Dim C As Integer
Dim i As Integer
Dim Contador2 As Integer


    If Not PrimeraVez Then
        If ModoTrabajo = vbNorm Then
            If TreeView1.Nodes.Count > 1 Then GuardarCarpetasAbiertas
        End If
    End If
    TreeView1.Nodes.Clear
    
    cad = " from carpetas"
    If ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt Then cad = cad & "hco"
    'Es el usuario propietario
    If vUsu.codusu > 0 Then
        cad = cad & " WHERE "
        cad = cad & "userprop = " & vUsu.codusu
    
        'O el grupo tiene permiso
        cad = cad & " OR (lecturag & " & vUsu.Grupo & ")"
        
    End If

    If ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt Then
        If vUsu.codusu = 0 Then
            cad = cad & " WHERE "
        Else
            cad = cad & " AND "
        End If
        cad = cad & "codequipo = " & vUsu.PC
    End If
    
    
    'Ordenado por padre
    cad = cad & " ORDER BY Padre,nombre"
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open "select * " & cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    If Rs.EOF Then
        MsgBox "ERROR GRAVE cargando árbol de directorios(Situacion: 1)", vbCritical
        End
    End If
    CadenaCarpetas = "|"
    
    If Rs!padre <> 0 Then
        MsgBox "Error en primer NODO. Padre != 0", vbExclamation
        End
    End If
    C = 0
    i = 0
    While i = 0
        INSERTAR_NODO Rs, 1
        Rs.MoveNext
        If Rs.EOF Then
            i = 1
        Else
            If Rs!padre <> 0 Then i = 1
        End If
        C = C + 1
    Wend
    
    'Cargo el segundo nivel
    Contador2 = TreeView1.Nodes.Count
    C = 0
    For i = 1 To Contador2
        cad = Mid(TreeView1.Nodes(i).Key, 2)
        Rs.MoveFirst
        Rs.Find " padre = " & cad, , adSearchForward, 1
        While Not Rs.EOF
            C = C + 1
            If Rs!padre = cad Then
                INSERTAR_NODO Rs, 2
            Else
                Rs.MoveLast
                
            End If
            Rs.MoveNext
        Wend
    Next i
       
    If C > 0 Then
                If Not PrimeraVez Then Label3.Caption = "     c   a   r   g   a   n   d   o  "
                'Cargo el tercer nivel
                C = Contador2 + 1
                Contador2 = TreeView1.Nodes.Count
                For i = C To Contador2
                    cad = Mid(TreeView1.Nodes(i).Key, 2)
                    Rs.MoveFirst
                    Rs.Find " padre = " & cad, , adSearchForward, 1
                    While Not Rs.EOF
                        C = C + 1
                        If Rs!padre = cad Then
                            INSERTAR_NODO Rs, 2
                        Else
                            Rs.MoveLast
                        End If
                        Rs.MoveNext
                    Wend
                Next i
                
               If C > 0 Then
                    'TERECER NIVEL
                    C = Contador2 + 1
                    Contador2 = TreeView1.Nodes.Count
                    For i = C To Contador2
                        cad = Mid(TreeView1.Nodes(i).Key, 2)
                        Rs.MoveFirst
                        Rs.Find " padre = " & cad, , adSearchForward, 1
                        While Not Rs.EOF
                            C = C + 1
                            If Rs!padre = cad Then
                                INSERTAR_NODO Rs, 2
                            Else
                                Rs.MoveLast
                            End If
                            Rs.MoveNext
                        Wend
                    Next i
                    
                    'CUARTO NIVEL
                    If C > 0 Then
                        C = Contador2 + 1
                        Contador2 = TreeView1.Nodes.Count
                        For i = C To Contador2
                            cad = Mid(TreeView1.Nodes(i).Key, 2)
                            Rs.MoveFirst
                            Rs.Find " padre = " & cad, , adSearchForward, 1
                            While Not Rs.EOF
                                C = C + 1
                                If Rs!padre = cad Then
                                    INSERTAR_NODO Rs, 2
                                Else
                                    Rs.MoveLast
                                End If
                                Rs.MoveNext
                            Wend
                        Next i
                    
                            
                    'QUINTO NIVEL
                        If C > 0 Then
                            C = Contador2 + 1
                            Contador2 = TreeView1.Nodes.Count
                            For i = C To Contador2
                                cad = Mid(TreeView1.Nodes(i).Key, 2)
                                Rs.MoveFirst
                                Rs.Find " padre = " & cad, , adSearchForward, 1
                                While Not Rs.EOF
                                    C = C + 1
                                    If Rs!padre = cad Then
                                        INSERTAR_NODO Rs, 2
                                    Else
                                        Rs.MoveLast
                                    End If
                                    Rs.MoveNext
                                Wend
                            Next i
                                
                            
                            
                            
                            T2 = Timer
                            
                            
                            C = Contador2 + 1
                            Contador2 = TreeView1.Nodes.Count
                            If Contador2 >= C Then
                                For i = C To Contador2
                                    
                                    CargaArbolRecursivo Mid(TreeView1.Nodes(i).Key, 2), Rs, 5
                                  
                                Next i
                            End If
                        
                        
                        End If '5º nivel
                    End If '4ºnivel
                End If '3 nivel
    End If
    
    
        
    Rs.Close
    If Not PrimeraVez Then Label3.Caption = " AriDoc: Gestión documental"
    If TreeView1.Nodes.Count > 2 Then TreeView1.Nodes(3).EnsureVisible
   
End Sub






Private Sub CargaArbolRecursivo(CarpePadre As String, ByRef rs1 As ADODB.Recordset, ByVal Nivel As Integer)
Dim C As Integer
Dim i As Integer
Dim CADENA As String
Dim Fin As Boolean
 
    'Este esta puesto para cuando es el arranque, que si le cuesta leer que no
    'bloquee el equipo
    If (TreeView1.Nodes.Count Mod 30) = 0 Then DoEvents


    CADENA = ""
    C = 0
    rs1.MoveFirst
    rs1.Find " padre = " & CarpePadre, , adSearchForward, 1
    Fin = rs1.EOF
    While Not Fin
        If rs1!padre = CarpePadre Then
        
            i = INSERTAR_NODO(rs1, Nivel)
            If i > 0 Then
                CADENA = CADENA & rs1!codcarpeta & "|"
                C = C + 1
            End If
            rs1.MoveNext
            If rs1.EOF Then Fin = True
        Else
            Fin = True
        End If
        
        If Timer - T2 > 1 Then
            If PrimeraVez Then
                frmInicio.Label1(2).Visible = Not frmInicio.Label1(2).Visible
                frmInicio.Label1(2).Refresh
            
            Else
                If Label3.Caption = "" Then
                    Label3.Caption = "     c   a   r   g   a   n   d   o  "
                Else
                    Label3.Caption = ""
                End If
                Label3.Refresh
            End If
            T2 = Timer
        End If
    Wend

    If C > 0 Then
        For i = 1 To C
            CargaArbolRecursivo (RecuperaValor(CADENA, i)), rs1, Nivel + 1
        Next i
    End If

End Sub







Private Sub CargaArbolRecursivoOLD(Nodo As Node, ByRef rs1 As ADODB.Recordset)
Dim C As Integer
Dim i As Integer
Dim CADENA As String

    'FALTA###
    'Este esta puesto para cuando es el arranque, que si le cuesta leer que no
    'bloquee el equipo
    DoEvents
    Me.Refresh
    rs1.MoveFirst

    CADENA = ""
    C = 0
    Cortar11 = Mid(Nodo.Key, 2)
    While Not rs1.EOF
        If Cortar11 = rs1!padre Then
            'FALTA
            'i = INSERTAR_NODO(rs1)
            If i > 0 Then
                CADENA = CADENA & i & "|"
                C = C + 1
            End If
        End If
        rs1.MoveNext
    Wend
    If C > 0 Then
        For i = 1 To C
            CargaArbolRecursivoOLD TreeView1.Nodes(Val(RecuperaValor(CADENA, i))), rs1
        Next i
    End If

End Sub


Private Sub HacerNuevaCarga()

    
End Sub



Private Sub Borrar_dir_Click()
Dim HayKRefrescar As Boolean

    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    
    
    
    'Deberiamos comprobar si tiene permisos sobre nodo padre
    If vUsu.codusu > 0 Then
        If Not TreeView1.SelectedItem.Parent Is Nothing Then
            Set Car = New Ccarpetas
            Cortar11 = ""
            If Car.Leer(CInt(Mid(TreeView1.SelectedItem.Parent.Key, 2)), False) = 1 Then
                MsgBox "Error leyendo carpeta :" & TreeView1.SelectedItem.Parent.Text, vbExclamation
            Else
                If (Car.escriturag And vUsu.Grupo) Or (Car.userprop = vUsu.codusu) Then Cortar11 = "OK"
            End If
            If Cortar11 = "" Then
                MsgBox "No tiene permiso sobre la carpeta contenedora ", vbExclamation
                Exit Sub
            Else
                Cortar11 = ""
            End If
         End If
    End If
    
    
    Screen.MousePointer = vbHourglass
    Set NodoOrigen = TreeView1.SelectedItem
    HayKRefrescar = LeerBDRefresco
    If EliminarCarpeta(TreeView1.SelectedItem) Then
        If HayKRefrescar Then
            RecargarDatos
        Else
            'QUE HACEMOS
            'ELiminamos el nodo directamente
            TreeView1.Nodes.Remove TreeView1.SelectedItem.Index
            
            If TreeView1.SelectedItem Is Nothing Then TreeView1.SelectedItem = TreeView1.Nodes(1)
            
                MostrarArchivos TreeView1.SelectedItem.Key
                Text1.Text = TreeView1.SelectedItem.FullPath
          
                
        End If
    End If

    Screen.MousePointer = vbDefault
    
End Sub



Private Sub Form_Activate()
Dim B As Byte
'La primera vez tendremos que realizar la integracion
If PrimeraVez Then
    PrimeraVez = False

    'Ahora hcemos la pregunta
  '  If ExisteCarpetaEnBD Then FicheroVerificcion False, B
    If B > 0 Then
        If B = 3 Then
            MsgBox "Fichero de comprobación de carpetas no existe, o nunca se ha realizado la comprobación", vbExclamation
        Else
            If B = 2 Then MsgBox "Desde la ultima vez que se realizó la comprobación hace mas de dos meses. Deberia volverla a hacer cuanto antes", vbExclamation
        End If
    End If
    
    Screen.MousePointer = vbHourglass
    
    HabilitarBotones False
    
    
    'Leo los tipos de mensajes
    PonerArrayTiposMensaje
    
    'ESTA VACIO LOS TIPOS DE MENSAJE
    If TotalTipos = 0 Then
        Me.mnMensajes.Enabled = False
        Me.Toolbar1.Buttons(22).Enabled = False
    Else
        If vUsu.preferencias.mailInicio Then
            CompruebaMail
            If Toolbar1.Buttons(22).Image = "tienemail" Then MsgBox vbCrLf & vbCrLf & "Tiene usted información pendiente de revisar." & vbCrLf & vbCrLf, vbExclamation

        Else
            Toolbar1.Buttons(22).Image = "mail"
        End If
    
    End If
    'Comprobar fecha SISTEMA MYSQL
    CompruebaFechaMYSQL
    
    Integracion
 
 
 
 
End If
Screen.MousePointer = vbDefault
End Sub




Private Sub Ponerpermisos()
Dim B As Boolean


    'Solo administrador root
    Me.mnAlmacenDatos.Enabled = vUsu.codusu = 0

    
    'SOLO ADMINISTRADORES
    B = vUsu.Nivel < 2
    Me.mnAdmonUsers.Enabled = B
    Me.mnGestionEquipos.Enabled = B
    Me.mnGrupos.Enabled = B
    Me.mnConfiguraExtenAridoc.Enabled = B
    
    'ADmon y avanzados
    B = vUsu.Nivel < 5
    Me.mnVerificarAridoc.Enabled = B
    mnTipoArchivos.Enabled = B
    Me.mnPlantillas.Enabled = B
    
End Sub


Private Sub Form_Resize()
Dim X, Y As Integer
Dim V ''


If WindowState = 1 Then Exit Sub ' ha pulsado minimizar
X = Me.Width
Y = Me.Height
If X < 5990 Then Me.Width = 5990
If Y < 4100 Then Me.Height = 4100
Image1.Left = TreeView1.Left
Label3.Left = Image1.Width + 15
Text1.Width = Me.Width - Text1.Left - 250
X = Me.Height - Base

TreeView1.Height = X
ListView1.Height = X
Y = Me.Width - 200
'Antes
'Y = Y / 3

'AHora va por porcentajes
Y = ((vUsu.preferencias.Ancho / 100) * Y)


TreeView1.Left = 30
TreeView1.Width = Y - 30

'Separador
Me.FrameSeparador.Left = Y + 15
Me.FrameSeparador.Top = TreeView1.Top
Me.FrameSeparador.Height = Me.TreeView1.Height

ListView1.Left = Y + 60
AnchoListview = Me.Width - 200 - Y - 30
ListView1.Width = AnchoListview
V = ListView1.Left + ListView1.Width - Label3.Left
Label3.Width = V
End Sub





Private Sub frmP_DatoSeleccionado(OpcionSeleccionada As Byte)
    vOpcion = OpcionSeleccionada
End Sub





Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ColumnHeader)
Dim L As Long
   'ListView1.SortKey = ColumnHeader.Index - 1
    If ListView1.Tag = ColumnHeader.Tag Then
        'Es el mismo, cambiamos ascendente a descendene
        OrderAscendente = Not OrderAscendente
    Else
        OrderAscendente = True
        ListView1.Tag = ColumnHeader.Tag
        
    End If
    
    Screen.MousePointer = vbHourglass
    
    pegar11 = ParaElORderBY(Val(ColumnHeader.Tag))
    vUsu.preferencias.ORDERBY = pegar11
    
    If ListView1.SelectedItem Is Nothing Then
        DatosCopiados = ""
    Else
        DatosCopiados = ListView1.SelectedItem.Key
    End If
    
    MostrarColumnas
    'CargarImagenEncabezado Val(ColumnHeader.Index)
    
    If Not TreeView1.SelectedItem Is Nothing Then
        'MostrarArchivos
        TreeView1_NodeClick TreeView1.SelectedItem
        If DatosCopiados <> "" Then
            Set ListView1.SelectedItem = Nothing
            For L = 1 To ListView1.ListItems.Count
                ListView1.ListItems(L).Selected = False
                If ListView1.ListItems(L).Key = DatosCopiados Then
                    Set ListView1.SelectedItem = ListView1.ListItems(L)
                    ListView1.SelectedItem.EnsureVisible
                    Exit For
                End If
            Next L
        End If
    End If
    Cortar11 = ""
    pegar11 = ""
    DatosCopiados = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListView1_DblClick()

Dim cad As String
Dim i As Integer
Dim vE As Cextensionpc
Dim NombreArchivo As String

    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    If ModoTrabajo = vbBackup Then
        mnpropiedades_Click
        Exit Sub
    End If
  
    
    If ListviewSHIFTPresionado And vbShiftMask Then
        'PROPIEDADES
         mnpropiedades_Click
        
    Else
        If ListviewSHIFTPresionado And vbCtrlMask Then
                
                'MODIFICAR
                 mnimportar_Click
                
                
        Else
            'Normal. Es decir, ver
            Set vE = New Cextensionpc
            i = 0
            If vE.Leer(ListView1.SelectedItem.SmallIcon - 1, vUsu.PC) = 1 Then
                i = 1
            Else
                If vE.pathexe = "" And ModoTrabajo <> vbBackup Then
                    i = 1
                    MsgBox "La extension no tiene PATH asociado", vbExclamation
                End If
            End If
                
            If i = 1 Then
                Set vE = Nothing
                Exit Sub
            End If
            
            
            
            
            Set Car = New Ccarpetas
            
            'Leemos la carpeta
            If Car.Leer(CInt(Mid(TreeView1.SelectedItem.Key, 2)), (ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt)) = 1 Then
                MsgBox "Error grave leyendo datos carpeta", vbExclamation
                Set Car = Nothing
                Exit Sub
            End If
            'Quitamos los espacios en blanco
            ''''NombreArchivo = ListView1.SelectedItem.Tag 'El tag siempre lleva el campo 1
            
            'NombreArchivo = ListView1.SelectedItem.Text
            i = DevuelveNombreFichero(ListView1.SelectedItem.Text, vE.Extension, NombreArchivo, False)
            If i > 100 Then
                MsgBox "Error obteniendo nombre fichero", vbExclamation
                Exit Sub
            End If
            
            If ModoTrabajo = vbBackup Then
                Exit Sub
            End If
            
            
            '----------------------------------------------------------------------------------------------
            AbirFichero True, Car, Val(Mid(ListView1.SelectedItem.Key, 2)), NombreArchivo, vE, 0, False, 0
            '----------------------------------------------------------------------------------------------
            
            Set Car = Nothing
            Set vE = Nothing
        End If
    End If
    
End Sub


Public Sub AbirFichero(Lectura As Boolean, ByRef Ca As Ccarpetas, Cod As Long, Destino As String, ByRef CEx As Cextensionpc, PosicionListview As Integer, EsOpcionArchivoNuevo As Boolean, Plantilla As Integer)
On Error GoTo EA
Dim cad As String
Dim PonerAtributoSoloLectura As Boolean
Dim TamanyoOriginal As Long
Dim FechaModificacion As Date
Dim T1 As Single
Dim FS, F   'File system
        
    
        If EsOpcionArchivoNuevo Then
            cad = Cod
            If Plantilla <= 0 Then
                Cod = CEx.codext
            Else
                Cod = Plantilla
            End If
            If Not TraerFicheroFisico(Ca, Destino, Cod) Then Exit Sub
            
            'Reestablecemos los valores
            Cod = Val(cad)
            Ca.Leer Ca.codcarpeta, (ModoTrabajo = 1)
            
            Screen.MousePointer = vbHourglass
            espera 0.5
        Else
            
            
            cad = Cod
            If ModoTrabajo = vbHistAnt Then
                cad = cad & "." & CEx.Extension
            End If
            If Not TraerFicheroFisico(Ca, Destino, cad) Then Exit Sub
        End If
        espera 0.2
        If Dir(Destino, vbArchive) = "" Then
            MsgBox "Se ha producido un error trayendo los datos", vbExclamation
            Exit Sub
        End If
        
        Set FS = CreateObject("Scripting.FileSystemObject")
        Set F = FS.GetFile(Destino)
        
        
        'Si la extension no dejara modificar
        ' tipo pdf o html, pondiramos Lectura=true
        PonerAtributoSoloLectura = True
        If CEx.ArchivosModificables Then
            If Not Lectura Then PonerAtributoSoloLectura = False
        End If
        'Protegemos para escritura
        If PonerAtributoSoloLectura Then SetAttr Destino, vbReadOnly
                
        '-------------------------------------------------------
        cad = """" & CEx.pathexe & """ """ & F.shortpath & """"
        If PonerAtributoSoloLectura Then
            TamanyoOriginal = Shell(cad, vbNormalFocus)
            InsertarEnProcesosAbiertos TamanyoOriginal, Destino
            espera 0.6
            objRevision.InsertaRevision Cod, 4, vUsu, ""
            
        Else
            '----------------
            'Se puede modificar
            
            'Ejecutamos hasta k no modifique no salimos
    
            FechaModificacion = F.DateLastModified
            TamanyoOriginal = F.Size
    
            'Lanzamos el visor con modificar
            Caption = "MODIFICANDO FICHERO:       " & F.shortpath
            Me.Enabled = False
            Me.Refresh
            'Abrimos
            
            espera 0.3
            
            T1 = Timer
            LanzaArchivoModificar cad
            If Timer - T1 < 0.2 Then
                
                cad = "Puede que haya una referencia en memoria de la aplicación: " & vbCrLf & CEx.pathexe
                cad = cad & vbCrLf & vbCrLf & "Esto impide bloquear el archivo de forma exclusiva."
                MsgBox cad, vbExclamation
                
            End If
            'Al cerrar reestablecemos
            '--------------------------
            Me.Enabled = True
            PonerCaption
            Me.Refresh
            cad = ""
    
            'Vuelvo a leer el fichero
            Set F = FS.GetFile(Destino)
            If F.Size <> TamanyoOriginal Then
                cad = "Cambiado"
            Else
                TamanyoOriginal = DateDiff("s", FechaModificacion, F.DateLastModified)
                If TamanyoOriginal > 1 Then cad = "cambiado"
            End If
            If cad <> "" Then
                Screen.MousePointer = vbHourglass
                'El fichero ha cambiado, Tengo k volverlo aponer en su lugar
                'Para eso es como si lo crearamos nuevo y lo llevaramos al servidor
                'Lo modificamos
                objRevision.InsertaRevision Cod, 3, vUsu, ""
                
                'Llevamos el fichero
                Set frmMovimientoArchivo.vDestino = Ca
                frmMovimientoArchivo.Opcion = 1
                frmMovimientoArchivo.Origen = Destino
                frmMovimientoArchivo.Destino = CStr(Cod)
                frmMovimientoArchivo.Show vbModal
            
                'Ahora updateo el tamañoç
                cad = CStr(Round((F.Size / 1024), 3))
                cad = "UPDATE Timagen SET tamnyo = " & TransformaComasPuntos(cad) & " where codigo =" & Cod
                Conn.Execute cad


                                
                If ListView1.View = lvwReport Then
                    If PosicionListview > 0 Then
                        For TamanyoOriginal = 1 To ListView1.ColumnHeaders.Count
                            
                            'EL 11 es el tamño
                            If ListView1.ColumnHeaders(TamanyoOriginal).Tag = 11 Then
                                cad = CStr(Round((F.Size / 1024), 3))
                                ListView1.ListItems(PosicionListview).SubItems(TamanyoOriginal - 1) = cad
                            End If
                        Next TamanyoOriginal
                    End If
                End If
            

            Else
                'Como no pasa nada
                Kill Destino
                If EsOpcionArchivoNuevo Then
                    'Es nuevo archivo Y NO ha modificado. Elimino la entrada y el item, ya que el archivo no ha llegado
                    'a subir
                    cad = "DELETE from timagen WHERE codigo =" & Cod
                    Conn.Execute cad
                    
                    ListView1.ListItems.Remove PosicionListview
                    ListView1.Refresh
                End If
            End If
            Me.Show
            Me.SetFocus
            
        End If
        ListView1.Drag vbCancel
        Screen.MousePointer = vbHourglass
        If vConfig.RevisaTareasAPI Then VerProcesosMuertos
    
EA:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Abrir/Modificar fichero"
    End If
    Set FS = Nothing
    Set F = Nothing
    ListviewSHIFTPresionado = 0
    Screen.MousePointer = vbDefault
End Sub


Private Sub ListView1_DragDrop(Source As Control, X As Single, Y As Single)
    ListView1.Drag vbCancel
    If Source.Name = Me.FrameSeparador.Name Then
        'Esta moviendo el separador
        MueveSeparador X, Y, True
    End If
End Sub




Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim Nom As String
Dim Color As Long

    Toolbar1.Buttons(15).Enabled = True
    Toolbar1.Buttons(16).Enabled = True
    Toolbar1.Buttons(17).Enabled = True
    Toolbar1.Buttons(5).Enabled = True
    mncortar2.Enabled = False
    mncortar.Enabled = True
    mnEliminar.Enabled = True
    mnPropiedades.Enabled = True
    
End Sub



Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown

CtrlDown = (Shift And vbCtrlMask) > 0
ListviewSHIFTPresionado = Shift
If KeyCode = 46 Then Eliminar

If CtrlDown Then

    If KeyCode = vbKeyC Or KeyCode = vbKeyX Then
        ' se ha pulsado ctrl + c
        If mncortar.Enabled = True Then HacerCortar
    Else
        If (KeyCode = vbKeyV) Then
        ' se ha pulsado ctrl + v
            If mncopiar.Enabled Then HacerPegar
        End If
    End If
    Else 'ctrlDown
        
        If KeyCode = 13 Then
            If Not ListView1.SelectedItem Is Nothing Then VerPropiedades ListView1.SelectedItem.Key, False
        End If
    End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
    ListviewSHIFTPresionado = 0
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'se ha pulsado el raton sobre el list

 If Button = 2 Then _
        'boton derecha. Asociar el popup
        If ModoTrabajo = vbBackup Then
            PopupMenu Me.mnListBackUp
        Else
            PopupMenu mnList
        End If
    Else
    If Button = vbLeftButton Then Set NodoOrigen = NodoSeleccionado
 End If
End Sub

Private Sub mnAcerca_Click()
    frmAbout.Show vbModal
End Sub


Private Sub mnAdmonPlantillas_Click()
    frmColPlantillas.Show vbModal
End Sub

Private Sub mnAdmonUsers_Click()
    'frmListaUserPcs.opcion = 0
    'frmListaUserPcs.Show vbModal
    frmUsuarios2.Show vbModal
End Sub

Private Sub mnAlmacenDatos_Click()
    frmVarios.Opcion = 2
    frmVarios.Show vbModal
End Sub

Private Sub mnBuscar_Click()

    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    Cortar11 = ""
    'Fiajar carpetas subcarpetas
    DatosCopiados = ""
    If TreeView1.SelectedItem <> TreeView1.Nodes(1) Then
        Cortar11 = CarpetasSubcarpetas(TreeView1.SelectedItem)
    Else
        Cortar11 = Mid(CadenaCarpetas, 2)
    End If
    frmBusca2.DesdeEmail = False
    frmBusca2.Carpetas = Cortar11
    frmBusca2.TodasCarpetas = CadenaCarpetas
    frmBusca2.Show vbModal
    ListView1.Drag vbCancel
    Cortar11 = ""
    Screen.MousePointer = vbHourglass
    If DatosCopiados <> "" Then
        
        PonerValordevueloBusqueda
        
    Else
        If Not TreeView1.SelectedItem Is Nothing Then MostrarArchivos TreeView1.SelectedItem.Key
        
    End If
    Screen.MousePointer = vbDefault
    DatosCopiados = ""
End Sub





Private Sub PonerValordevueloBusqueda()
Dim i As Integer
Dim Valor As String
Dim OK As Boolean

    'ponemos la carpeta
    Valor = RecuperaValor(DatosCopiados, 1)
    OK = False
    Set TreeView1.SelectedItem = Nothing
    For i = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Key = Valor Then
            OK = True
            Set TreeView1.SelectedItem = TreeView1.Nodes(i)
            TreeView1.SelectedItem.EnsureVisible
            TreeView1.SelectedItem.Expanded = True
            MostrarArchivos TreeView1.SelectedItem.Key
            Exit For
        End If
    Next i
    Me.Refresh
    If Not OK Then
        MsgBox "Carpeta de busqueda (" & Valor & ") NO encontrada", vbExclamation
        Exit Sub
    End If
    
    
    'Ahora busco el archivo
    OK = False
    Valor = RecuperaValor(DatosCopiados, 2)
    Me.Refresh
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Selected = False
        If ListView1.ListItems(i).Key = Valor Then
            OK = True
            Set ListView1.SelectedItem = ListView1.ListItems(i)
            ListView1.SelectedItem.EnsureVisible
        End If
    Next i
    If Not OK Then MsgBox "Archivo de busqueda (" & Valor & ") NO encontrado", vbExclamation
    
End Sub





Private Sub mnCambiarImpresora_Click()
        cd1.CancelError = True
        On Error GoTo ErrHandler
        ' Presentar el cuadro de diálogo Imprimir
        cd1.ShowPrinter
ErrHandler:
    Err.Clear
End Sub

Private Sub mnCambioClave_Click()
    frmVarios.Opcion = 0
    frmVarios.Show vbModal
End Sub

Private Sub mnCerrarHistorico_Click()
' cerramos historico

    If MsgBox("Cerrar histórico?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass
    ModoTrabajo = vbNorm
    
    mnRestaurar.Visible = True
    mnCerrarHistorico.Visible = False
    mnHerramientasadmon.Enabled = True
    ListView1.ListItems.Clear

    'Trozo antiguo de carga del arbol
    CargaArbol
    If TreeView1.Nodes.Count > 1 Then TreeView1.Nodes(2).EnsureVisible
    If TreeView1.Nodes.Count > 0 Then
        Set NodoSeleccionado = TreeView1.Nodes(1)
        Text1.Text = NodoSeleccionado.FullPath
    Else
        Text1.Text = ""
    End If
    'Ponemos not enabled a mncopiaseg y mn realizar histor
    Colores
    mnCopiaSeg.Enabled = True
    mnPaso.Enabled = True
    mnintegrar.Enabled = True
    Screen.MousePointer = vbDefault
End Sub



Private Sub mnChkPropItem_Click()
Dim i As Long
    If ModoTrabajo <> vbNorm Then
        Mensajes1 15
        Exit Sub
    End If
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
        

    'Vamos a cambiar los archivos de propietario
    If ListView1.ListItems.Count = 0 Then
        MsgBox "No tiene archivos", vbExclamation
        Exit Sub
    End If
    
    Cortar11 = ""
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected Then Cortar11 = Cortar11 & "1"
    Next i
    
    
    If Cortar11 = "" Then
        MsgBox "Ningun archivo seleccionado", vbExclamation
        Exit Sub
    End If
    
    vOpcion = 127
    Set frmP = New frmPregunta
    frmP.Opcion = 11
    frmP.origenDestino = Len(Cortar11)
    frmP.Show vbModal
    If vOpcion = 127 Then Exit Sub
    
        'LLegados aqui actualizaremos la BD
    Screen.MousePointer = vbHourglass
    Cortar11 = "Select codgrupo from usuariosgrupos where "
    Cortar11 = Cortar11 & " codusu=" & vOpcion
    Cortar11 = Cortar11 & " ORDER BY orden"
    
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open Cortar11, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cortar11 = ""
    If Not miRSAux.EOF Then
        If Not IsNull(miRSAux.Fields(0)) Then Cortar11 = miRSAux.Fields(0)
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    If Cortar11 = "" Then
        MsgBox "Grupo PPal para el usuario: " & vOpcion & " NO encontrado", vbExclamation
        Exit Sub
    End If

    pegar11 = vOpcion & " / " & Cortar11
    Cortar11 = "UPDATE timagen Set userprop = " & vOpcion & " , groupprop = " & Cortar11
    Cortar11 = Cortar11 & " WHERE codigo = "
    
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected Then
            Conn.Execute Cortar11 & Mid(ListView1.ListItems(i).Key, 2)
            'Si lleva hco revisiones entonces ponemos
            If objRevision.LlevaHcoRevision Then objRevision.InsertaRevision CLng(Mid(ListView1.ListItems(i).Key, 2)), 5, vUsu, "Cambiar prop y grupo: " & vOpcion & " / " & pegar11
        End If
    Next i
    pegar11 = ""
    Cortar11 = ""
    
    'Cargararchivos
    MostrarArchivos TreeView1.SelectedItem.Key
End Sub

Private Sub mnConfigMAIL_Click()
    frmConfigPersonal.Opcion = 1
    frmConfigPersonal.Show vbModal
End Sub

Private Sub mnConfiguraExtenAridoc_Click()
    frmColExten.Show vbModal
End Sub

Private Sub mncontraernodo_Click()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    ExpandeNodo TreeView1.SelectedItem, False
    Screen.MousePointer = vbDefault
End Sub

Private Sub mncopiar_Click()
    HacerPegar
End Sub



Private Sub mncortar_Click()
HacerCortar
End Sub

Private Sub mncortar2_Click()
' cortar desde menu
HacerCortar
End Sub

Private Sub mnDeselectAll_Click()
Dim i  As Integer

For i = 1 To ListView1.ListItems.Count
    ListView1.ListItems(i).Selected = False
    Next i
mncortar = False
mncortar2 = False
Toolbar1.Buttons(5).Enabled = False
End Sub


Private Sub mnEliminar_Click()
    Eliminar
End Sub




Private Sub mnEnviarPorMail_Click()
Dim i As Integer

    If ModoTrabajo = vbBackup Then
        MsgBox "Opcion no disponible en modo recuperar backup", vbExclamation
        Exit Sub
    End If

    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Cortar11 = ""
    'Fiajar carpetas subcarpetas
    DatosCopiados = ""
    If TreeView1.SelectedItem <> TreeView1.Nodes(1) Then
        Cortar11 = CarpetasSubcarpetas(TreeView1.SelectedItem)
    Else
        Cortar11 = Mid(CadenaCarpetas, 2)
    End If
    pegar11 = ""
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected Then pegar11 = pegar11 & ListView1.ListItems(i).Key & "|"
    Next i
    
    frmMensaje.Carpetas = Cortar11
    frmMensaje.TodasCarpetas = CadenaCarpetas
    frmMensaje.ImagenAEnviar = TreeView1.SelectedItem.Key & "|" & pegar11
    frmMensaje.Opcion = 0
    frmMensaje.Show vbModal
    Cortar11 = ""
    pegar11 = ""
End Sub

Private Sub mnexpandirnodo_Click()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    Screen.MousePointer = vbHourglass
    ExpandeNodo TreeView1.SelectedItem, True
    Screen.MousePointer = vbDefault
End Sub

Private Sub ExpandeNodo(ByRef No As Node, Expandir As Boolean)
Dim N As Node

    No.Expanded = Expandir
    If Not No.Child Is Nothing Then
        Set N = No.Child
        Do
            ExpandeNodo N, Expandir
            If Not N.Next Is Nothing Then
                Set N = N.Next
            Else
                Set N = Nothing
            End If
        Loop Until (N Is Nothing)
    End If
End Sub


Private Sub mnGestionEquipos_Click()
    frmListaUserPcs2.Opcion = 1
    frmListaUserPcs2.Show vbModal
End Sub


Private Sub mnGrupos_Click()
    frmGrupos.Show vbModal
End Sub

'---------------------------------
'---------------------------------
'---------------------------------
'  Es apretar el boton modificar M O D I F I C A R
'---------------------------------
'---------------------------------

Private Sub mnimportar_Click()
    ModificarNuevoArchivo True, 0
    'ComprobarRefrescar
End Sub


Private Sub ModificarNuevoArchivo(DesdeElListView As Boolean, vPlantilla As Integer)
Dim Img As cTimagen
Dim vE As Cextensionpc
Dim NombreArchivo As String
Dim i As Integer
Dim cad As String

    If ModoTrabajo <> vbNorm Then
        Mensajes1 13
        Exit Sub
    End If

    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    
    ListView1.Drag vbCancel
    
    Set vE = New Cextensionpc
    i = 0
    If vE.Leer(ListView1.SelectedItem.SmallIcon - 1, vUsu.PC) = 1 Then
        i = 1
    Else
        If vE.pathexe = "" Then
            MsgBox "La extensión no tiene PATH asociado", vbExclamation
            i = 1
        End If
    End If
    If i = 1 Then
        Set vE = Nothing
        Exit Sub
    End If
    
    If Not vE.ArchivosModificables Then
        MsgBox "Este tipo de archivo no es modificable", vbExclamation
        Set vE = Nothing
        Exit Sub
    End If
    
    'Leemos la carpeta
    If DesdeElListView Then
        Set Car = New Ccarpetas
        If Car.Leer(CInt(Mid(TreeView1.SelectedItem.Key, 2)), (ModoTrabajo = 1)) = 1 Then
            MsgBox "Error grave leyendo datos carpeta", vbExclamation
            Set Car = Nothing
            Exit Sub
        End If
        
    Else
        'Opcion de generar un nuevo archvio con el boton de la derecha
        ' El objeto CAR ya esta definido en el otro sitio. Lo unico k le pondre el cod carpeta
        Car.codcarpeta = CInt(Mid(TreeView1.SelectedItem.Key, 2))
    End If
    
    'Leemos la carpeta
    Set Img = New cTimagen
    
    Screen.MousePointer = vbHourglass
    espera 1.5
    
    If Img.Leer(Val(Mid(ListView1.SelectedItem.Key, 2)), objRevision.LlevaHcoRevision) = 0 Then
    
        'Si tiene los permisos
        i = 0
        If vUsu.codusu = 0 Then
            i = 1
        Else
            
            If Img.userprop = vUsu.codusu Or (Img.escriturag And vUsu.Grupo) Then i = 1
        End If
        
        If i = 1 Then
            'Quitamos los espacios en blanco
            'NombreArchivo = ListView1.SelectedItem.Tag 'El tag siempre lleva el campo 1
            NombreArchivo = ListView1.SelectedItem.Text
            Do
                i = InStr(1, NombreArchivo, " ")
                If i > 0 Then NombreArchivo = Mid(NombreArchivo, 1, i - 1) & Mid(NombreArchivo, i + 1)
            Loop Until i = 0
            i = 0
            Do
                cad = App.Path & "\temp\" & NombreArchivo
                If i > 0 Then cad = cad & "(" & i & ")"
                cad = cad & "." & vE.Extension
                i = i + 1
            Loop Until Dir(cad, vbArchive) = "" Or i > 100
            If i > 100 Then
                MsgBox "Error obteniendo nombre fichero(100)", vbExclamation
                Exit Sub
            End If
            
            AbirFichero False, Car, Val(Mid(ListView1.SelectedItem.Key, 2)), cad, vE, ListView1.SelectedItem.Index, Not DesdeElListView, vPlantilla
            
            
                
        Else
            MsgBox "No tiene permisos.", vbExclamation
        End If
            
    
    End If
    
    Set Img = Nothing
    Set vE = Nothing
    Set Car = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnImportarVarios_Click()
    InsertarMultiple 0
End Sub

Private Sub mnImprimir_Click()
Dim J As Integer
Dim i As Byte



            BorrarTemporal1
            
            i = 0
            For J = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(J).Selected Then
                    InsertaTemporal (CLng(Mid(ListView1.ListItems(J).Key, 2)))
                    i = 1
                End If
            Next J
                    
            If i = 0 Then
                MsgBox "Seleccione algun archivo para imprimir", vbExclamation
                Exit Sub
            End If
            
            
            Set Car = New Ccarpetas
            
            'Leemos la carpeta
            If Car.Leer(CInt(Mid(TreeView1.SelectedItem.Key, 2)), ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt) = 1 Then
                MsgBox "Error grave leyendo datos carpeta", vbExclamation
                Set Car = Nothing
                Exit Sub
            End If
            
            
            ImprimirDesdeTablaTemporal Me, (ModoTrabajo = vbHistAnt Or ModoTrabajo = vbHistNue)
            
            Set Car = Nothing
End Sub

            


Private Sub mnInsertarDesdePlantilla_Click()
    If ModoTrabajo <> vbNorm Then
        Mensajes1 5
        Exit Sub
    End If
    DatosCopiados = ""
    frmPlantilla.Opcion = 0
    frmPlantilla.Show vbModal
    If DatosCopiados <> "" Then
        NuevoEditando CInt(DatosCopiados) + 200
    End If
End Sub

Private Sub mnInsertarmasivaCarpeta_Click()
    InsertarMultiple 1
End Sub

Private Sub mnintegrar_Click()
    If ModoTrabajo <> vbNorm Then
        Mensajes1 5
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Integracion
    'por si acaso hay archivos nuevos en la carpeta mandamos cargar otra vez
    ComprobarRefrescar
    Screen.MousePointer = vbDefault
End Sub





Private Sub mnMensajesCol_Click()
Dim Todo As Boolean
    If ModoTrabajo = vbBackup Then
        MsgBox "Opcion no disponible recuperando historico", vbExclamation
        Exit Sub
    End If

    Toolbar1.Buttons(22).Image = "mail"
    
    Screen.MousePointer = vbHourglass
    Cortar11 = ""
    'Fiajar carpetas subcarpetas
    Todo = True
    If Not TreeView1.SelectedItem Is Nothing Then
        If TreeView1.SelectedItem <> TreeView1.Nodes(1) Then Todo = False
    End If
    If Not Todo Then
        Cortar11 = CarpetasSubcarpetas(TreeView1.SelectedItem)
    Else
        Cortar11 = Mid(CadenaCarpetas, 2)
    End If
    frmColMail2.Carpetas1 = Cortar11
    frmColMail2.TodasCarpetas1 = CadenaCarpetas
    frmColMail2.Show vbModal
End Sub

Private Sub mnModicarArch_Click()
    'importar desde POPUP
    mnimportar_Click

End Sub

Private Sub mnNueva2_Click()

'    Dim I As Integer
'
'    Set listacod = New Collection
'
'    For I = 1 To 100
'        listacod.Add "c00" & (I * 10), CStr(I + 2)
'    Next I
'
'    For I = 1 To listacod.Count
'        If (I Mod 10) = 0 Then
'            Stop
'            Debug.Print "Indice: " & I & " " & listacod.Item(I)
'            Debug.Print "Key: " & I & " " & listacod.Item(CStr(I))
'            listacod.Remove CStr(I)
'        End If
'    Next I

   
    NuevoEditando -1
    
    'AQUI pondremos lo de
    ComprobarRefrescar
End Sub

'
Private Sub mnNuevoArch_Click()

    'Nuevo arhcivo
    Insertar
    ComprobarRefrescar
End Sub




Private Sub NuevoEditando(Tipo As Integer)
Dim C As Long

  If NodoSeleccionado Is Nothing Then Exit Sub
  If ModoTrabajo <> vbNorm Then
        Mensajes1 (5)
        Exit Sub
  End If
  
  If Not NodoseleccionadoConsulta Then Exit Sub
'  If NodoSeleccionado.Parent Is Nothing Then
'        MsgBox "No se pueden insertar archivos en la carpeta Raiz", vbInformation
'        Exit Sub
'  End If
        

    
    'Comprobamos si el usuarios tiene permiso de
    Set Car = New Ccarpetas
    
    If Car.Leer(CInt(Mid(NodoSeleccionado.Key, 2)), (ModoTrabajo = 1)) = 0 Then
        
        'OK. VEMOS EL Permiso
        If Car.userprop = vUsu.codusu Or (Car.escriturag And vUsu.Grupo) Or vUsu.codusu = 0 Then
            
            Set frmNuevoArchivo.Mc = Car
            'Sera 100 + el codigo del neuvo archivo
            If Tipo < 0 Then
                Tipo = 100
            Else
                If Tipo < 200 Then Tipo = 100 + Tipo
            End If
            frmNuevoArchivo.Opcion = Tipo
            frmNuevoArchivo.Carpeta = Text1.Text
            DatosMOdificados = False
            frmNuevoArchivo.Show vbModal
            If DatosMOdificados Then
                Screen.MousePointer = vbHourglass
                MostrarArchivos TreeView1.SelectedItem.Key
                DatosCopiados = "C" & DatosCopiados
                'Situaremos el nodo en donde toca
                For C = 1 To ListView1.ListItems.Count
                    If DatosCopiados = ListView1.ListItems(C).Key Then
                        Exit For
                    End If
                Next
                If C <= ListView1.ListItems.Count Then
                    'Fale, A encontrado el nodo
                    ListView1.ListItems(C).EnsureVisible
                    Set ListView1.SelectedItem = ListView1.ListItems(C)
                    
                    'Simulamos el modificar
                    'La carpeta origen sera la del codigo carpeta 1
                    'ya que 0.- Iconos
                    '       1.- Archivos en blanco
                    '       2.- Plantillas predefinidas
                    
                    Set Car = Nothing
                    Set Car = New Ccarpetas
                    Set miRSAux = New ADODB.Recordset
                    If Tipo < 200 Then
                        miRSAux.Open "Select * from almacen where codalma= 1", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    Else
                        miRSAux.Open "Select * from almacen where codalma= 2", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    End If
                    If miRSAux.EOF Then
                        MsgBox "Error leyendo carpeta. Subtipo: " & Tipo, vbExclamation
                        Cortar11 = "NO"
                    Else
                        
                        Car.Almacen = miRSAux!codalma
                        Car.SRV = miRSAux!SRV
                        Car.version = miRSAux!version
                        Car.user = miRSAux!user
                        Car.pwd = miRSAux!pwd
                        Car.pathreal = miRSAux!pathreal
                        Cortar11 = ""
                    End If
                    
                    miRSAux.Close
                    Set miRSAux = Nothing
                    
                    If Cortar11 = "" Then
                        'Lanzamos abrir archivo
                        If Tipo < 200 Then
                            Tipo = 0
                        Else
                            Tipo = Tipo - 200
                        End If
                        ModificarNuevoArchivo False, Tipo
                    End If
                    
                    
                Else
                    MsgBox "No se ha encontrado el archivo nuevo.", vbExclamation
                End If
                Screen.MousePointer = vbDefault
            End If
        Else
            MsgBox "No tiene permiso ", vbExclamation
        End If
    Else
        MsgBox "Error leyendo carpeta : " & NodoSeleccionado.Text
    End If
  Screen.MousePointer = vbDefault

            
            
            
End Sub

Private Sub mnNuevoArchivoEditando_Click()
     NuevoEditando -1
     ComprobarRefrescar
End Sub


Private Sub mnNuevoInsertar_Click()
    'Nuevo
    ListView1.Drag vbCancel
    Insertar
    ComprobarRefrescar
End Sub

Private Sub mnNuevoMensaje_Click()


    Screen.MousePointer = vbHourglass
    Cortar11 = ""
    'Fiajar carpetas subcarpetas
    DatosCopiados = ""
    If TreeView1.SelectedItem <> TreeView1.Nodes(1) Then
        Cortar11 = CarpetasSubcarpetas(TreeView1.SelectedItem)
    Else
        Cortar11 = Mid(CadenaCarpetas, 2)
    End If
    frmMensaje.Carpetas = Cortar11
    frmMensaje.TodasCarpetas = CadenaCarpetas
    frmMensaje.ImagenAEnviar = ""
    frmMensaje.Opcion = 0
    frmMensaje.Show vbModal
    Cortar11 = ""
End Sub




Private Sub mnNuevoN1_Click(Index As Integer)
    NuevoEditando CInt(mnNuevoN1(Index).Tag)
End Sub

Private Sub mnParametros_Click()
    frmVarios.Opcion = 1
    frmVarios.Show vbModal
End Sub

Private Sub mnPaso_Click()
    'Pasar archivos a cho
    If vUsu.Nivel > 2 Then
        MsgBox "No tiene permisos asignados", vbExclamation
        Exit Sub
    End If
    
    If ModoTrabajo <> vbNorm Then
        Mensajes1 15
        Exit Sub
    End If
    
    frmHco.Opcion = 0
    frmHco.Show vbModal
    
End Sub

Private Sub mnPegar_Click()
    If ModoTrabajo <> vbNorm Then
        Mensajes1 15
        Exit Sub
    End If
    'pegar
    HacerPegar
End Sub

Private Sub mnPlantillas_Click()
    If ModoTrabajo <> vbNorm Then
        Mensajes1 15
        Exit Sub
    End If
    frmPlantilla.Opcion = 1
    frmPlantilla.Show vbModal
End Sub

Private Sub mnportipo_Click()
' vamos a realizar un seleccion por tipo de archivo
Dim exte As String
Dim Aux As String
Dim i As Integer
Dim sel As Boolean


    If ListView1.ListItems.Count = 0 Then
        MsgBox "La carpeta no contiene ningún archivo.", vbInformation
        Exit Sub
    End If
    exte = "Escriba la extension que desea buscar (gif,txt,...)" & vbCrLf & vbCrLf
    exte = exte & "Carpeta : " & TreeView1.SelectedItem.FullPath
    Aux = InputBox(exte, "Buscar por tipo de archivo")
    If Aux = "" Then Exit Sub
    Aux = LCase(Aux)
    If Len(Aux) > 3 Then
        MsgBox "La extensión debe de tener menos de 3 carácteres.", vbInformation
        Exit Sub
        End If
    Aux = Mid(Aux, Len(Aux) - 2, Len(Aux))
    
    exte = DevuelveDesdeBD("codext", "extension", "exten", Aux, "T")
    If exte = "" Then
        MsgBox "La extensión """ & Aux & """   NO es reconocida por el programa.", vbInformation
        Exit Sub
    End If
       
    'La extension aux es el icono: sel
    Aux = Val(exte) + 1
    
    'Llegados aqui, quitamos la selección de todos los archivos
    mnDeselectAll_Click
    
    
    sel = False
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).SmallIcon = Aux Then
            ListView1.ListItems(i).Selected = True
            sel = True
            End If
    Next i
    If sel Then
            ListView1.SetFocus
            mncortar = True
            mncortar2 = True
            Toolbar1.Buttons(5).Enabled = True
    End If
        
    Me.Refresh
End Sub

Private Sub mnPreferencias_Click()
    DatosMOdificados = False
    frmConfigPersonal.Opcion = 0
    frmConfigPersonal.Show vbModal
    
    If DatosMOdificados Then
        Screen.MousePointer = vbHourglass
        vUsu.preferencias.Leer vUsu.codusu   'para volver a grabar el vselect, si hiciera falta
        PonerPreferenciasPersonales
        ListView1.View = vUsu.preferencias.Vista
        Form_Resize
        MostrarColumnas
        MostrarArchivos TreeView1.SelectedItem.Key
      End If
End Sub

Private Sub mnPropAvanzadas_Click()
'Dim CAr As Ccarpetas
Dim Cpadre As Ccarpetas
Dim TienePermiso As Boolean
Dim i As Integer


    If TreeView1.SelectedItem Is Nothing Then Exit Sub

    
    'Comprobamos si el usuarios tiene permiso de
    Set Car = New Ccarpetas
    
    'Leemos la carpeta
    If Car.Leer(CInt(Mid(TreeView1.SelectedItem.Key, 2)), (ModoTrabajo = 1)) = 1 Then
        MsgBox "Error grave leyendo datos carpeta", vbExclamation
        Set Car = Nothing
        Exit Sub
    End If
    
    
    '----------------------------------------------
    If TreeView1.SelectedItem.Parent Is Nothing Then
        'NODO RAIZ
        'Solo el prompietario tiene permiso
        TienePermiso = (Car.userprop = vUsu.codusu)
        If vUsu.codusu = 0 Then TienePermiso = True
    Else
        Set Cpadre = New Ccarpetas
        If Cpadre.Leer(CInt(Mid(TreeView1.SelectedItem.Parent.Key, 2)), (ModoTrabajo = 1)) = 1 Then
            MsgBox "Error leyendo datos carpeta contenedora", vbExclamation
            i = 1
        Else
            i = 0
            'los permisos de escritura, modificacion los coje de la carpeta padre
            If vUsu.codusu = 0 Then
                TienePermiso = True
            Else
                'Si es el propeietario de la carpeta entonces puede cambiar el nombre
                If Car.userprop = vUsu.codusu Then
                    TienePermiso = True
                Else
                    TienePermiso = (Cpadre.userprop = vUsu.codusu) Or (Cpadre.escriturag And vUsu.Grupo)
                End If
            End If
        End If
        Set Cpadre = Nothing
        If i = 1 Then Exit Sub
    End If
    
       
        
        'OK. VEMOS EL Permiso
    
        
            'FALE, tiene permisos para la creacion de la carpeta
            'Ahora abrimos el forumlario
            Set frmCarpetas.vC = Car
            
            i = InStr(1, Text1.Text, Car.Nombre)
            If i > 0 Then
                frmCarpetas.Ubicacion = Mid(Text1.Text, 1, i - 1)
            Else
                frmCarpetas.Ubicacion = ""
            End If
            If ModoTrabajo <> vbNorm Then TienePermiso = False
            frmCarpetas.PuedeModificar = TienePermiso
            DatosMOdificados = False
            frmCarpetas.Show vbModal
            If DatosMOdificados Then
                Car.Leer Car.codcarpeta, (ModoTrabajo = 1)
                TreeView1.SelectedItem.Text = Car.Nombre
                Car.ActualizaTablaActualiza
            End If
End Sub

Private Sub mnpropiedades_Click()
    ' se ha pulsado en propiedades
    ' es como si hicieras doble click en listview
    Dim i  As Integer
    Dim Con As Integer
    Screen.MousePointer = vbHourglass
    
    On Error GoTo EPropiedades
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    
    Con = 0
    For i = 1 To ListView1.ListItems.Count
       If ListView1.ListItems(i).Selected = True Then Con = Con + 1
    Next i
    
    If Con > 1 Then
            'Ver espacio
            VerEspacio 0, False
        Else
            VerPropiedades ListView1.SelectedItem.Key, False
    End If
    ListviewSHIFTPresionado = 0
    Me.Refresh
EPropiedades:
        If Err.Number <> 0 Then
    
            If Err.Number = 53 Then
               
            Else
                Cortar11 = Err.Description
            End If
            MsgBox Cortar11, vbExclamation
            Err.Clear
            Cortar11 = ""
        End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnPropiedadesCarpeta_Click()
    VerEspacio 1, False
End Sub

Private Sub mnquitarsel_Click()
    'quitar selección
    mnDeselectAll_Click
End Sub

Private Sub mnRecuperaDesdeBackUp_Click()
    'RECUPERA DESDE UN BACKUP
    If ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt Then
        MsgBox "Esta en HCO. Cierre este proceso primero", vbExclamation
        Exit Sub
    End If
    
    
    
    Set TreeView1.SelectedItem = Nothing
    Set ListView1.SelectedItem = Nothing
    
            
    If ModoTrabajo = vbBackup Then
        Screen.MousePointer = vbHourglass
        Me.mnRecuperaDesdeBackUp.Caption = "Recuperar archivos backup"
        Me.mnMensajes.Enabled = True
        Me.mnCopiaS.Enabled = True
        mnConfig.Enabled = True
        
        Label3.BackColor = &H4D2C1D
        Label3.Caption = "Aridoc: Gestión documental"
        
        'Esta opcion es para reeestablecer el modo normal
        ListView1.ListItems.Clear
        Conn.Close
        espera 1
        If Not AbrirConexion(True) Then
            MsgBox "ERROR GRAVE abriendo conexion BD aridoc. La aplicación finlizará", vbCritical
            End
        End If
        RefrescarPonerNodo
            
        
        ModoTrabajo = vbNorm
    Else
        
        
        DatosMOdificados = False
        frmRecuperaBackup.Opcion = 0
        frmRecuperaBackup.Show vbModal
        If DatosMOdificados Then
            Screen.MousePointer = vbHourglass
            ListView1.ListItems.Clear
            ModoTrabajo = vbBackup
            Label3.BackColor = &HFF0000  'Azulito
            Me.mnRecuperaDesdeBackUp.Caption = "Finalizar restarurar backup"
            Me.mnMensajes.Enabled = False
            Me.mnCopiaS.Enabled = False  'UTILIDADES
            mnConfig.Enabled = False  'Configuracion
            Screen.MousePointer = vbHourglass
            RefrescarPonerNodo
        Else
            'Ha habido un error en el proceso intermedio
            'Deberiamos o reiciniar, o como minimo abri la conexion otra vez
            AbrirConexion True
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub




Private Sub mnRefrescar_Click()

    'AUQI AQUI AQUI
    If ModoTrabajo <> vbNorm Then Exit Sub
    Screen.MousePointer = vbHourglass
    RefrescarPonerNodo
    Screen.MousePointer = vbDefault
End Sub


Private Sub RefrescarPonerNodo()
Dim i As String
Dim J As String
Dim L As Integer

    On Error GoTo EREFRE
    DatosCopiados = ""
    Set NodoOrigen = Nothing
    Set NodoSeleccionado = Nothing

    If Not TreeView1.SelectedItem Is Nothing Then i = TreeView1.SelectedItem.Key
    If Not ListView1.SelectedItem Is Nothing Then J = ListView1.SelectedItem.Key
    Text1.Text = "............  Cargando estructura carpetas"
    Text1.Refresh
    CargaArbol
    TreeView1.SelectedItem = TreeView1.Nodes(1)
    If i <> "" Then
        For L = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(L).Key = i Then
                TreeView1.SelectedItem = TreeView1.Nodes(L)
                Exit For
            End If
        Next L
    
    End If
    Text1.Text = "............  Cargando archivos"
    Text1.Refresh
    ListView1.SelectedItem = Nothing
    ListView1.ListItems.Clear
    espera 0.3
    MostrarArchivos TreeView1.SelectedItem.Key
    If Not ListView1.SelectedItem Is Nothing Then
        ListView1.SelectedItem.Selected = False
        ListView1.SelectedItem = Nothing
    End If
    
    For L = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(L).Key = J Then
            ListView1.SelectedItem = ListView1.ListItems(L)
            Exit For
        End If
    Next L
    
    
    Text1.Text = TreeView1.SelectedItem.FullPath
    Me.Refresh
    Exit Sub
EREFRE:
    'No controlamos el error pq el nucleo importante(CARGAARBOL) ya tiene el suyo
    'luego si da error es volviendo a situar los nodos, y eso implicaria que
    'si da error es k ha sido mificado o borrado
    Err.Clear
End Sub


Private Sub mnRestarurarArchivo_Click()
Dim i As Integer

    'Restauramos el archivo
    If ModoTrabajo <> vbBackup Then Exit Sub
    
    Cortar11 = ""
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected Then
            Cortar11 = "N"
            Exit For
        End If
    Next i
    If Cortar11 = "" Then
        MsgBox "Seleccione los archivos a restaurar", vbExclamation
        Exit Sub
    End If
    
    
    pegar11 = "codigo, codext, codcarpeta, campo1, campo2, campo3, campo4,"
    pegar11 = pegar11 & "fecha1, fecha2, fecha3, importe1, importe2, observa, tamnyo, userprop, groupprop, lecturau, lecturag, escriturau, escriturag, bloqueo "
            
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected Then
            
            Cortar11 = "INSERT INTO timagenhco (codequipo, " & pegar11 & ") "
            Cortar11 = Cortar11 & " SELECT 0," & pegar11 & " FROM timagen"
            Cortar11 = Cortar11 & " where codigo =" & Mid(ListView1.ListItems(i).Key, 2)
            On Error Resume Next
            Conn.Execute Cortar11
            If Err.Number <> 0 Then
                If Conn.Errors(0).NativeError = 1062 Then
                    MsgBox "Ya esta pendiente de recuperar el archivo: " & ListView1.ListItems(i)
                    Err.Clear
                Else
                    MuestraError Err.Number
                End If
            End If
        End If
    Next i
    DatosCopiados = TreeView1.SelectedItem.FullPath & "|" & TreeView1.SelectedItem.Key & "|" & TreeView1.SelectedItem.Parent.Key & "|"
    frmRecuperaBackup.Opcion = 1
    frmRecuperaBackup.Show vbModal
    
End Sub


Private Sub mnrestaurar_Click()
' Vamos a recuperara de algun soporte
' un historico realizado

Dim i
 
    frmHco.Opcion = 1
    frmHco.Show vbModal
 
    If DatosCopiados = "" Then Exit Sub
    
     Screen.MousePointer = vbHourglass
     If DatosCopiados = "ANT" Then
        ModoTrabajo = vbHistAnt
     Else
        ModoTrabajo = vbHistNue
     End If
     mnRestaurar.Visible = False
     mnCerrarHistorico.Visible = True
     mnHerramientasadmon.Enabled = False
     ListView1.ListItems.Clear
     CargaArbol
    



    If TreeView1.Nodes.Count > 0 Then
        Set NodoSeleccionado = TreeView1.Nodes(1)
        Text1.Text = NodoSeleccionado.FullPath
    Else
        Text1.Text = ""
    End If

    'Expandimos todos los nodos
    If TreeView1.Nodes.Count > 1 Then
        TreeView1.Nodes(2).EnsureVisible
    End If
    Colores
    Me.Refresh
    Screen.MousePointer = vbDefault
End Sub


Private Sub Colores()
    'If ModoTrabajo = vbHist Then
    If ModoTrabajo <> vbNorm Then
        Label3.BackColor = &H80&
    Else
        Label3.BackColor = &H4D2C1D
    End If
End Sub

Private Sub mnRestaurarCarpeta_Click()
    'Rstaurar carpeta
    If ModoTrabajo <> vbBackup Then Exit Sub
    
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    
    
    
    
End Sub

Private Sub mnSelectAll_Click()
Dim i  As Integer

    Set ListView1.SelectedItem = Nothing
    
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).Selected = True
    Next i
    If i > 1 Then
        mncortar = True
        mncortar2 = True
        mnPropiedades.Enabled = True
        Toolbar1.Buttons(5).Enabled = True
        Toolbar1.Buttons(16).Enabled = True
     End If
    TreeView1.SetFocus
    ListView1.SetFocus
End Sub

Private Sub mnselectodo_Click()
    'select all
     mnSelectAll_Click
End Sub

Private Sub mnSumaArchivos_Click()
    VerEspacio 0, 1
End Sub

Private Sub mnSumaCarpetaActual_Click()
      VerEspacio 0, 2
End Sub

Private Sub mnSumasSubcarpetas_Click()
      VerEspacio 0, 3
End Sub

Private Sub mnTipoArchivos_Click()
    frmConfigExtensiones.NuevoEquipo = False
    frmConfigExtensiones.Show vbModal
End Sub





Private Sub mnTiposMensaje_Click()
    frmTiposMensajes.Show vbModal
End Sub

Private Sub mnVerificarAridoc_Click()
    If ModoTrabajo <> vbNorm Then
        Mensajes1 14
        Exit Sub
    End If
    frmVerificacion.Show vbModal
End Sub

Private Sub mnVerifyCarpetaActual_Click()

    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    Cortar11 = Mid(TreeView1.SelectedItem.Key, 2) & "|"
    HacerVerificacion
    
    
End Sub

Private Sub mnVeryfCarpSub_Click()
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    Cortar11 = CarpetasSubcarpetas(TreeView1.SelectedItem)
    HacerVerificacion
End Sub

Private Sub HacerVerificacion()
    If ModoTrabajo <> vbNorm Then
        Mensajes1 15
        Exit Sub
    End If
    'Ahora llamo al formulario poniendo le la cadena cortarr11 en origen
    Screen.MousePointer = vbHourglass
    DatosCopiados = ""
    frmMovimientoArchivo.Opcion = 15
    frmMovimientoArchivo.Origen = Cortar11
    frmMovimientoArchivo.Show vbModal
    If Not listacod Is Nothing Then
        If listacod.Count = 0 Then
            MsgBox "Verificación finalizada con éxito", vbExclamation
        Else
            frmVarios.Opcion = 5
            frmVarios.Show vbModal
            If DatosCopiados <> "" Then MostrarArchivos TreeView1.SelectedItem.Key
        End If
    End If
    Set listacod = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Sub mnVolcarEstructura_Click()
    
    
    If ModoTrabajo <> vbNorm Then
        Mensajes1 15
        Exit Sub
    End If


    
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    
    If vUsu.codusu <> 0 Then
        If TreeView1.SelectedItem.Parent Is Nothing Then
            MsgBox "Imposible realizar esta opción desde la carpeta raiz.", vbExclamation
            Exit Sub
        End If
    End If
    
    frmVarios.Opcion = 6
    frmVarios.Show vbModal

End Sub

Private Sub Nuevo_Dire_Click()
Dim NuevaCar As Ccarpetas
Dim CADENA As String
Dim i As Integer
Dim HayKRefrescar As Boolean
Dim XNodo As Node

    If ModoTrabajo <> vbNorm Then
        Mensajes1 11
        Exit Sub
    End If

    'NODO SELECCIONADO
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    
    CADENA = Mid(TreeView1.SelectedItem.Key, 2)
    
    'Comprobamos si el usuarios tiene permiso de
    Set Car = New Ccarpetas
    
    If Car.Leer(CInt(CADENA), (ModoTrabajo = 1)) = 0 Then
        
        'OK. VEMOS EL Permiso
        If Car.userprop = vUsu.codusu Or (Car.escriturag And vUsu.Grupo) Or vUsu.codusu = 0 Then
            'FALE, tiene permisos para la creacion de la carpeta
            Set NuevaCar = New Ccarpetas
            'asignamos algunos valores
            With NuevaCar
                'El almacen contenedor
                .version = Car.version
                .Almacen = Car.Almacen
                .user = Car.user
                .pwd = Car.pwd
                .groupprop = Car.groupprop
                .userprop = vUsu.codusu
                .padre = Car.codcarpeta
                .codcarpeta = -1
            End With
            
            'Ahora abrimos el forumlario
            HayKRefrescar = LeerBDRefresco
            Set frmCarpetas.vC = NuevaCar
            frmCarpetas.Ubicacion = Text1.Text
            frmCarpetas.PuedeModificar = True
            DatosMOdificados = False
            frmCarpetas.Show vbModal
            If DatosMOdificados Then
                Screen.MousePointer = vbHourglass
                If HayKRefrescar Then
                    CargaArbol
                        Cortar11 = ""
                        CADENA = "C" & NuevaCar.codcarpeta
                        For i = 1 To TreeView1.Nodes.Count
                            If TreeView1.Nodes(i).Key = CADENA Then
                                ListView1.ListItems.Clear
                                TreeView1.Nodes(i).EnsureVisible
                                Set TreeView1.SelectedItem = TreeView1.Nodes(i)
                                Cortar11 = "NO"
                                Exit For
                            End If
                        Next i
                    
                    
                Else
                    Set XNodo = TreeView1.Nodes.Add(TreeView1.SelectedItem.Key, tvwChild, "C" & NuevaCar.codcarpeta)
                    XNodo.Text = NuevaCar.Nombre
                    Me.Text1.Text = Me.Text1 & "\" & NuevaCar.Nombre
                    XNodo.Image = "v_cerrado"
                    XNodo.ExpandedImage = "v_abierto"
                    XNodo.Tag = NuevaCar.escriturau & "|" & NuevaCar.escriturag & "|"
                    i = XNodo.Index
                    
                End If
         
                ListView1.ListItems.Clear
                'No se ha encontrado el nodo
                If Cortar11 = "" Then
                    Set TreeView1.SelectedItem = TreeView1.Nodes(i)
                End If
                    
                
                Set NodoSeleccionado = TreeView1.SelectedItem
                Screen.MousePointer = vbDefault
            End If
        Else
            MsgBox "No tiene permisos", vbExclamation
        End If
        
        
    Else
        MsgBox "No se han podido leer los datos de la carpeta: " & TreeView1.SelectedItem.Text & "(" & TreeView1.SelectedItem.Key & ")", vbExclamation
    End If

Errohdle:
    
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then
        CADENA = "Error al crear la carpeta." & vbCrLf & vbCrLf
        CADENA = CADENA & "Compruebe que la ruta es correcta o que " & vbCrLf
        CADENA = CADENA & " NO exista una carpeta con ese nombre" & vbCrLf & vbCrLf
        CADENA = CADENA & Err.Number & " - " & Err.Description
        MsgBox CADENA, vbCritical
    End If
    Set Car = Nothing
End Sub

Private Sub Form_Load()
'Load
Dim i As Long
Dim SQL As String
Dim Rs As Recordset
Dim men As String

Screen.MousePointer = vbHourglass

ModoTrabajo = 0
Top = -4
Left = -4
Width = Screen.Width
Height = Screen.Height
AnchoListview = ListView1.Width
CarpetasAbiertas = ""
ListviewSHIFTPresionado = 0
    
    
    
'Si el usuario es root podra verificar ......
OrderAscendente = (vUsu.codusu = 0)
mnBarraVerificar.Visible = OrderAscendente
mnVerficarCarpeta1.Visible = OrderAscendente
mnBarraVerificar = OrderAscendente
    
Me.mnChkPropBarra.Visible = OrderAscendente
Me.mnChkPropItem.Visible = OrderAscendente
    
Me.mnMenuCopiaSeguridad.Visible = OrderAscendente

    
'Cargamos la imagen de fondo del listview
ListView1.Picture = LoadPicture(App.Path & "\fondo.dat")


PrimeraVez = True
OrderAscendente = True
ListView1.Tag = ""

CargarListviews

If ImageList3.ListImages.Count > 0 Then
    ListView1.Icons = ImageList3
    ListView1.SmallIcons = ImageList2
End If

'Cargamos el imglist4 que nos servira para el formulario de imagen
'Probablemente luego lo quitaremos
CargaList4

'Ahora vamos con el treeview
'primero Cargamos el Listview1
'CargaList1

'Set TreeView1.ImageList = ImageList1
 Set TreeView1.ImageList = Me.ImgUsersPCs
'Asociamos con los IMGLIST
'Toolbar1.ImageList = ImageList1

Toolbar1.ImageList = ImgUsersPCs

' cargamos las imagenes de los botones
Toolbar1.Buttons(1).Image = "carpeta"
Toolbar1.Buttons(2).Image = "delete"
Toolbar1.Buttons(5).Image = "cortar"
Toolbar1.Buttons(5).Enabled = False
Toolbar1.Buttons(6).Image = "pegar"
Toolbar1.Buttons(6).Enabled = False
Toolbar1.Buttons(10).Image = "lista"
Toolbar1.Buttons(11).Image = "iconos"
Toolbar1.Buttons(14).Image = "nuevo"
Toolbar1.Buttons(14).Enabled = False
Toolbar1.Buttons(15).Image = "importar"
Toolbar1.Buttons(15).Enabled = False
Toolbar1.Buttons(16).Image = "prop"
Toolbar1.Buttons(16).Enabled = False
Toolbar1.Buttons(17).Image = "eliminar"
Toolbar1.Buttons(17).Enabled = False
Toolbar1.Buttons(19).Image = "find"
Toolbar1.Buttons(21).Image = "imprimir"
mncopiar.Enabled = False
mncortar = False
mncortar2 = False
mnPegar = False
mnNueva2 = False
Me.mnNuevoInsertar.Enabled = False
mnImportar = False




' sirve para calcular despues el width
Base = 1290
Base = Base + 550 '550 es lo k mide de alto la imagen de ariadna



CargaArbol




If TreeView1.Nodes.Count > 0 Then
    Set NodoSeleccionado = TreeView1.Nodes(1)
    Text1.Text = NodoSeleccionado.FullPath
Else
    Text1.Text = ""
End If

'Expandimos todos los nodos
If TreeView1.Nodes.Count > 1 Then
    TreeView1.Nodes(2).EnsureVisible
End If

'Preferencias personales
PonerPreferenciasPersonales
ListView1.View = vUsu.preferencias.Vista


Set ListView1.ColumnHeaderIcons = Me.ImgUsersPCs
MostrarColumnas

PonerCaption

Ponerpermisos

PonerMenuNuevoDocumentos


Me.mnEnviarPorMail.Visible = (vUsu.e_server <> "")
mnbarra4.Visible = (vUsu.e_server <> "")

CargaLogoPequeño


LeerPonerImpresora True


Exit Sub
ErrHdle:
    men = " Error número:  " & Err.Number & vbCrLf
    men = " Descripción : " & vbCrLf & "          " & Err.Description
    MsgBox men, vbCritical
    End
End Sub


Private Sub LeerPonerImpresora(Leer As Boolean)
Dim P As Printer
    If Leer Then
        ImpresoraAntesEntrar = Printer.DeviceName
        
    Else
        If Printer.DeviceName = ImpresoraAntesEntrar Then Exit Sub
        For Each P In Printers
            If P.DeviceName = ImpresoraAntesEntrar Then
                ' La define como predeterminada del sistema.
                Set Printer = P
                ' Sale del bucle.
                Exit For
            End If
        Next
   End If
End Sub


Private Sub CargaLogoPequeño()
    On Error Resume Next
    'El image1 para el Logo
    'El tamaño es fijo, de 1800x570
    'Leva el screcht a TRUE para que adapte el logo si fuera mas grande
    Image1.Picture = LoadPicture(App.Path & "\MiniLog2.dat")
    If Err.Number <> 0 Then Err.Clear
End Sub

'Private Sub Recursivo(ByVal nod As Node, ByVal Camino As String)
'Dim nx As Node
'Dim aux As String
'
'    Set nx = nod.FirstSibling
'    While nx <> nod.LastSibling
'        aux = Camino & nx.Text & "\"
'        If CargaArbol(nx.Key, aux) Then
'            Recursivo nx.Child.FirstSibling, aux
'        End If
'        Set nx = nx.Next
'    Wend
'    If Inicio.Pgbar.Value < 475 Then Inicio.Pgbar.Value = Inicio.Pgbar.Value + 1
'    If nx = nod.LastSibling Then
'        aux = Camino & nx.Text & "\"
'        If CargaArbol(nx.Key, aux) Then
'            Recursivo nx.Child.FirstSibling, aux
'        End If
'      End If
'    Set nx = Nothing
'End Sub





'Private Function CargaArbol(ByVal codpadre As String, ByVal Camino As String) As Boolean
'Dim xNod As Node
'Dim miNombre As String
'Dim aux, aux2 As String
'
'
'CargaArbol = False
'miNombre = Dir(Camino, vbDirectory)
'Do While miNombre <> ""
'
'   If miNombre <> "." And miNombre <> ".." Then
'        On Error GoTo Final  'ESTO NO ESTABA
'
'      If (GetAttr(Camino & miNombre) And vbDirectory) = vbDirectory Then
'         ' hacemos la accion
'         puntero = puntero + 1
'         InsertarNodo codpadre, miNombre, puntero
'         aux = Camino & miNombre & "\"
'         CargaArbol = True
'      End If
'
'    End If
'   miNombre = Dir ' Obtiene siguiente entrada.
'   Loop
'   Exit Function
'Final:
'
'End Function


'Private Sub InsertarNodo(ByVal clavepadre As String, nombre As String, apuntador As Long)
'Dim xNod As Node
'         Set xNod = TreeView1.Nodes.Add(clavepadre, tvwChild)
'         xNod.Key = vec(apuntador)
'         xNod.Text = nombre
'         xNod.Image = "cerrado"
'         xNod.ExpandedImage = "abierto"
'         Set xNod = Nothing
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    GuardarPreferencias
    LeerPonerImpresora False
End Sub

Private Sub GuardarPreferencias()
Dim i As Integer
    On Error GoTo EGuardarPreferencias
    
    If ListView1.View = lvwReport Then
        vUsu.preferencias.Vista = lvwReport
        'Para cada columna guardamos su ancho
        For i = 1 To ListView1.ColumnHeaders.Count
            AsignarAnchoColumnas CInt(ListView1.ColumnHeaders(i).Tag), CInt(Abs(ListView1.ColumnHeaders(i).Width))
        Next i
        
        
    Else
        vUsu.preferencias.Vista = lvwIcon
    End If
    
    
    vUsu.preferencias.Modificar vUsu.codusu, False
    
    Exit Sub
EGuardarPreferencias:
    MuestraError Err.Number, "Guardar preferencias"
End Sub

Private Sub AsignarAnchoColumnas(Columna As Integer, Ancho As Integer)

    
   Select Case Columna
   Case 1
        vUsu.preferencias.C1 = Ancho
        
   Case 2
        vUsu.preferencias.C2 = Ancho
   
    Case 3
        vUsu.preferencias.c3 = Ancho
   
    Case 4
        vUsu.preferencias.c4 = Ancho
   
    'fechas
    Case 5
        vUsu.preferencias.f1 = Ancho
    Case 6
        vUsu.preferencias.f2 = Ancho
    Case 7
        vUsu.preferencias.f3 = Ancho
        
        
        
    'importes
    Case 8
        vUsu.preferencias.imp1 = Ancho
    Case 9
        vUsu.preferencias.imp2 = Ancho
    
    Case 10
        vUsu.preferencias.obs = Ancho
        
    Case 11
        vUsu.preferencias.tamayo = Ancho
    End Select
End Sub



Private Sub mnSalir_Click()
Visible = False
Unload Me
End Sub



Private Sub TimerTree_Timer()
 
    TimerTree.Enabled = False
    If TreeView1.SelectedItem Is Nothing Then Exit Sub
    Set NodoOrigen = TreeView1.SelectedItem
    TreeView1.Drag vbBeginDrag
End Sub

'El timer lo utilizamos para que word desbloquee el archivo y
'de ese modo copiarlo

'Hoy 25 de marzo me doy cuenta que aqui no entra
''''''Private Sub Timer1_Timer()
''''''Dim rc, i, f
''''''Dim f1, f2
''''''    Timer1.Enabled = False ' Deshabilitamos el timer denuevo
''''''    Timer1.Enabled = True
''''''    Set fss = CreateObject("Scripting.FileSystemObject")
''''''    Set f = fss.GetFile(NombreDoc)
''''''    f1 = f.DateCreated + 0.00003 'Le sumamos 2 segundos el tiempo entre crearlo y darle nombre
''''''    f2 = f.DateLastModified
''''''    Set f = Nothing
''''''    Set fss = Nothing
''''''    If f1 < f2 Then 'La fecha de creacion <> ultima modificacion
''''''    'El archivo ha sido modificado y guardado
''''''    'hay que insertarlo en la aplicacion y que word lo
''''''    'desbloquee hemos puesto un timer
''''''        rc = DevuelveExtension("doc")
''''''        If rc = 0 Then Exit Sub
''''''        HaSidoCancelado = True ' Le ponemos el valor por defecto
''''''        Set Img = New CImag
''''''        Img.Siguiente
''''''        Img.NomPath = CarpetaW
''''''        Img.NomFich = Img.Id & ".doc"
''''''        Img.Extension = rc
''''''        FileCopy NombreDoc, Img.NomPath & Img.NomFich
''''''        Set frmImagen1 = New frmImagen
''''''        Set frmImagen1.mImg = Img
''''''        frmImagen1.modificar = False
''''''        frmImagen1.Show vbModal
''''''        Set frmImagen1 = Nothing
''''''        If Not HaSidoCancelado Then
''''''            Screen.MousePointer = vbHourglass
''''''            MostrarArchivos (inicial & NodoSeleccionado.FullPath & "\")
''''''            Screen.MousePointer = vbDefault
''''''            Else
''''''                ' ha sido cancelado. Habra que borrar el archivo
''''''                Kill Img.NomPath & Img.Id & ".*"
''''''        End If
''''''    End If
''''''End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer
Dim It As ListItem
Dim B As Boolean

On Error GoTo ErrorTool




'
'
'Dim C As String
'For i = 4001 To 8000
'
'    C = "INSERT INTO timagen (codigo, codext, codcarpeta, campo1, campo2, campo3, campo4, fecha1, fecha2, fecha3, importe1, importe2, observa, tamnyo, userprop, groupprop, lecturau, lecturag, escriturau, escriturag, bloqueo) VALUES ("
'    C = C & i & ",11,14, 'campo: " & i & " ', '0', '0', NULL, '2000-01-01', NULL, NULL, NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0)"

'
'    Conn.Execute C
'Next i
'
'Stop



Select Case Button.Index
Case 1 'Nueva carpeta
        Nuevo_Dire_Click
Case 2 ' Borra carpeta
        Borrar_dir_Click
                
Case 5 ' cortar
       HacerCortar
Case 6 'pegar
       HacerPegar
Case 10 ' la vista es en modo lvwReport
        If NodoSeleccionado Is Nothing Then Exit Sub
        vUsu.preferencias.Vista = lvwReport
        MostrarColumnas
        ListView1.View = vUsu.preferencias.Vista
        ListView1.HideColumnHeaders = False
        i = 0
        If Not ListView1.SelectedItem Is Nothing Then i = ListView1.SelectedItem.Index
        MostrarArchivos (NodoSeleccionado.Key)
         If i > 0 Then Set ListView1.SelectedItem = ListView1.ListItems(i)
Case 11 ' la vista es en modo lvwicon
        If NodoSeleccionado Is Nothing Then Exit Sub
        'AnchoColumna 3
        vUsu.preferencias.Vista = lvwIcon
        ListView1.View = vUsu.preferencias.Vista
        ListView1.HideColumnHeaders = True
        i = 0
        If Not ListView1.SelectedItem Is Nothing Then i = ListView1.SelectedItem.Index
        MostrarArchivos (NodoSeleccionado.Key)
        If i > 0 Then Set ListView1.SelectedItem = ListView1.ListItems(i)
Case 14 ' nueva imagen
        Insertar
Case 15 ' Importar:  ES MODIFICAR
        mnimportar_Click
Case 16 ' propiedades
        'VerPropiedades
        mnpropiedades_Click
Case 17 ' elimiar
        Eliminar
Case 19 ' buscar
         mnBuscar_Click
            
Case 21 ' Imprimir

        'If ImprimirArchivos(B) = False Then Mensajes1 (9)
        mnImprimir_Click
Case 22
    'MAIL
    mnMensajesCol_Click
End Select
ListView1.SetFocus
Exit Sub
ErrorTool:
    If Err.Number <> 0 Then MostrarError Err.Number
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
End Sub



Private Sub TreeView1_Collapse(ByVal Node As MSComctlLib.Node)
    'Se ha colapsado el nodo
    If NodoSeleccionado Is Nothing Then Exit Sub
    If NodoSeleccionado <> Node Then TreeView1_NodeClick Node
    
    
End Sub

Private Sub TreeView1_DragDrop(Source As Control, X As Single, Y As Single)
Dim L  As Long

    If Source.Name = Me.FrameSeparador.Name Then
        'Esta moviendo el separador
        MueveSeparador X, Y, False
        Exit Sub
    End If


    'PARA COPIAR MOVER CARPETAS  CARPETAS CARPETAS
    '---------------------------------------------
    'presionado = False
    ListView1.Drag vbEndDrag
    
    If TreeView1.DropHighlight Is Nothing Then
        Set TreeView1.DropHighlight = Nothing
        Exit Sub
    End If
           
    Screen.MousePointer = vbHourglass
    ' comprobamos si lo podemos insertar y si es asi lo borramos de
    ' alli y lo metemos aqui  la carpeta no debe incluir subdirectorios
    
    Set NodoSeleccionado = TreeView1.DropHighlight
    If RealizarMoverCarpetas() = 1 Then
    
        'nada
    Else

    End If

    Set TreeView1.DropHighlight = Nothing
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub TreeView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
        Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
        If Not TreeView1.DropHighlight Is Nothing Then TreeView1.DropHighlight.Expanded = True
End Sub

Private Sub TreeView1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0

If KeyCode = 46 Then
    ' Hay que borrar la carpeta
    Borrar_dir_Click
    Exit Sub
    End If
If CtrlDown And KeyCode = vbKeyV Then
    ' se ha pulsado ctrl + v
        If mncopiar.Enabled Then HacerPegar
    End If
End Sub

Private Sub TreeView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Debug.Print "Mouse down" & NodoSeleccionado.Text
Toolbar1.Buttons(15).Enabled = False
Toolbar1.Buttons(16).Enabled = False
mnEliminar.Enabled = False
mnPropiedades.Enabled = False
Toolbar1.Buttons(5).Enabled = False
mncortar2 = False
ListView1.Drag vbCancel
If Button = 2 Then
    If ModoTrabajo = vbBackup Then
        'PopupMenu Me.mnArbolBK
    Else
        PopupMenu mnMenuTree
    End If
Else
    Me.TimerTree.Enabled = True
End If

End Sub


Private Sub TreeView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ListView1.Drag vbCancel
    TimerTree.Enabled = False
End Sub



Private Sub HabilitarBotones(Si As Boolean)
    Me.mnNuevoInsertar.Enabled = Si
    mnNueva2 = Si
    Me.mnNuevoArch.Enabled = Si
    mnImportar.Enabled = Si
    Toolbar1.Buttons(14).Enabled = Si
    Toolbar1.Buttons(15).Enabled = Si
    mnNuevoEditando.Enabled = Si
End Sub


Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
 
    TimerTree.Enabled = False
    'TieneArchivos = False
    ListView1.ListItems.Clear
    
    'TreeView1.Enabled = False
    If Node.Parent Is Nothing Then
  
    
    Else
       
        'ListView1.HideColumnHeaders = False
        MostrarArchivos (Node.Key)

    
    End If
    HabilitarBotones Not (Node.Parent Is Nothing)
    'TreeView1.Enabled = True
    'Set TreeView1.SelectedItem = Node
    
    Set TreeView1.DropHighlight = Nothing
    
    Set NodoSeleccionado = Node
    Set NodoOrigen = Node
    Text1.Text = Node.FullPath
    TreeView1.SetFocus

End Sub



Private Sub TreeView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim CADENA As String
'presionado = False
ListView1.Drag vbEndDrag
If TreeView1.DropHighlight Is Nothing Then
      Set TreeView1.DropHighlight = Nothing
      Exit Sub
   Else
    ' comprobamos si lo podemos insertar y si es asi lo borramos de
    ' alli y lo metemos aqui  la carpeta no debe incluir subdirectorios
    If CompruebaCarpeta(TreeView1.DropHighlight.Key) Then
        Set NodoSeleccionado = TreeView1.DropHighlight
        Set TreeView1.SelectedItem = TreeView1.DropHighlight
        Text1.Text = NodoSeleccionado.FullPath
        If RealizarMover(0) = 0 Then
            Set NodoSeleccionado = NodoOrigen
            Set TreeView1.SelectedItem = NodoSeleccionado
            Text1.Text = NodoSeleccionado.FullPath
        Else
            MostrarArchivos (NodoSeleccionado.Key)
        End If

    End If
End If
Set TreeView1.DropHighlight = Nothing
End Sub

Private Sub TreeView1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    Set TreeView1.DropHighlight = TreeView1.HitTest(X, Y)
    If Not TreeView1.DropHighlight Is Nothing Then TreeView1.DropHighlight.Expanded = True
End Sub



Private Sub MostrarColumnas()


ListView1.ColumnHeaders.Clear


'Clave1 ...
With vUsu.preferencias
    If .C1 > 0 Then CrearColumna vConfig.C1, .C1, 1
    
    If .C2 > 0 Then CrearColumna vConfig.C2, .C2, 2
    
    If .c3 > 0 Then CrearColumna vConfig.c3, .c3, 3
    
    If .c4 > 0 Then CrearColumna vConfig.c4, .c4, 4
    
    If .f1 > 0 Then CrearColumna vConfig.f1, .f1, 5
    
    If .f2 > 0 Then CrearColumna vConfig.f2, .f2, 6
    
    If .f3 > 0 Then CrearColumna vConfig.f2, .f3, 7
    
    If .imp1 > 0 Then CrearColumna vConfig.imp1, .imp1, 8, True
    
    If .imp2 > 0 Then CrearColumna vConfig.imp2, .imp2, 9, True
    
    If .obs > 0 Then CrearColumna vConfig.obs, .obs, 10
    
    If .tamayo > 0 Then CrearColumna "Tamaño", .tamayo, 11, True
    End With
    Cortar11 = ""
End Sub


Private Sub CrearColumna(Nombre As String, Ancho As Integer, Campo As Integer, Optional Derecha As Boolean)
Dim clmX As ColumnHeader
    Set clmX = ListView1.ColumnHeaders.Add()
    clmX.Text = Nombre
    clmX.Width = Ancho
    clmX.Tag = Campo
    If vUsu.preferencias.ORDERBY <> "" Then
        pegar11 = ParaElORderBY(Campo)
        If vUsu.preferencias.ORDERBY = pegar11 Then
            'ListView1.Tag = clmX.Index
            CargarImagenEncabezado clmX.Index
        End If
    End If
    If Derecha Then clmX.Alignment = lvwColumnRight

End Sub



Private Function ParaElORderBY(Campo As Integer) As String
    
    Cortar11 = ""
    Select Case Campo
    Case 1
        Cortar11 = "campo1"
    Case 2
        Cortar11 = "campo2"
    Case 3
        Cortar11 = "campo3"
    Case 4
        Cortar11 = "campo4"
    Case 5
        Cortar11 = "fecha1"
    Case 6
        Cortar11 = "fecha2"
    Case 7
        Cortar11 = "fecha3"
    Case 8
        Cortar11 = "importe1"
    Case 9
        Cortar11 = "importe2"
    Case 10
        Cortar11 = "observa"
    Case 11
        Cortar11 = "tamnyo"
    End Select
    ParaElORderBY = Cortar11
End Function




'Carpeta = string ¿pq? pq es un node.key es decir "C" & codcarpeta
Private Sub MostrarArchivos(Carpeta As String)
Dim Rs As Recordset
Dim ItmX As ListItem
Dim cad As String

Dim NomArch As String

Screen.MousePointer = vbHourglass
    
    

    ListView1.ListItems.Clear
   
    cad = "Select " & vUsu.preferencias.vSelect
    cad = cad & ", escriturag,codigo,codext"
'    If InStr(1, cad, "campo1") > 0 Then
'        X = 4
'    Else
'        cad = cad & ",campo1"
'        X = 5
'    End If
    cad = cad & " from timagen"
    If ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt Then cad = cad & "hco"
    cad = cad & " WHERE "
    'Carpeta
    cad = cad & " codcarpeta = " & Mid(Carpeta, 2)
    'Es el usuario propietario
    If vUsu.codusu > 0 Then
        cad = cad & " AND (userprop = " & vUsu.codusu
        
        'O el grupo tiene permiso
        cad = cad & " OR (lecturag & " & vUsu.Grupo & "))"
    End If

    
    If ModoTrabajo = vbHistNue Or ModoTrabajo = vbHistAnt Then cad = cad & " AND codequipo = " & vUsu.PC

    'Ordenado por padre
    If ListView1.View = lvwReport Then
        'Vemos el order by
        If vUsu.preferencias.ORDERBY = "" Then
            pegar11 = ParaElORderBY(Val(ListView1.ColumnHeaders(1).Tag))
            vUsu.preferencias.ORDERBY = pegar11
            pegar11 = ""
            
        End If
        cad = cad & " ORDER BY " & vUsu.preferencias.ORDERBY
        If Not OrderAscendente Then cad = cad & " DESC"
    Else
        cad = cad & " ORDER BY campo1"
    End If
    
    
    
    Set Rs = New ADODB.Recordset
    Rs.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText

    'Si es de iconos
    If ListView1.View = lvwReport Then
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add(, "C" & Rs!codigo)
            
        
            PonerDatosLineaListviewReport ItmX, Rs
            'ItmX.SmallIcon = Rs!codext + 1
            PonerSmallIcon ItmX, Rs!codext + 1
            Rs.MoveNext
        Wend
        
    Else
        'SOlo ICONOS
        While Not Rs.EOF
            Set ItmX = ListView1.ListItems.Add(, "C" & Rs!codigo)
            ItmX.Text = DBLet(Rs.Fields(0))
            
            'ItmX.SmallIcon = Rs!codext + 1
            'ItmX.Icon = ItmX.SmallIcon
            
            PonerIcon ItmX, Rs!codext + 1
            
            'No hace falta que sea cmpo1, campo1 siempre es visible
            ItmX.Tag = Rs!escriturag
            
            Rs.MoveNext
        Wend
    End If
    
    Rs.Close
    Set Rs = Nothing
    ListView1.Refresh


Screen.MousePointer = vbDefault
End Sub


Private Sub PonerIcon(ByRef ItemD As ListItem, ByRef codigo)
    On Error Resume Next
    ItemD.Icon = codigo
    ItemD.SmallIcon = codigo
    If Err.Number <> 0 Then
        ItemD.Icon = 1
        Err.Clear
    End If
        
End Sub

Private Sub PonerSmallIcon(ByRef ItemD As ListItem, ByRef codigo)
    On Error Resume Next
    ItemD.SmallIcon = codigo
    If Err.Number <> 0 Then
        ItemD.SmallIcon = 1
        Err.Clear
    End If
End Sub


Private Sub PonerDatosLineaListviewReport(ByRef i As ListItem, ByRef RSS As ADODB.Recordset)
Dim J As Integer
    i.Text = RSS.Fields(0)
    For J = 1 To RSS.Fields.Count - 4  '- 3 pk los dos utlimos son perimos escirura y codigo
         i.SubItems(J) = DBLet(RSS.Fields(J))
    Next J
End Sub



Private Function RealizarMover(NoHacerPregunta As Byte) As Byte
'Dim C1 As String
'Dim C2 As String
Dim i As Long
Dim CO As Long
Dim cOri As Ccarpetas

'copiamos el archivo origen en destino
' a nivel de archivos
RealizarMover = 0
If ModoTrabajo <> vbNorm Then
    Mensajes1 (6)
    Exit Function
End If

On Error GoTo errorhndl
    If NoHacerPregunta = 0 Then
        If NodoOrigen.FullPath = NodoSeleccionado.FullPath Then
            MsgBox "No se puede realizar MOVER archivos sobre una misma carpeta", vbInformation
            Exit Function
        End If
    End If

        'Hacemos la pregunta
        If NoHacerPregunta = 0 Then
            vOpcion = 0
            Set frmP = New frmPregunta
            frmP.Opcion = 1
            frmP.origenDestino = NodoOrigen.FullPath & "|" & NodoSeleccionado.FullPath & "|"
            frmP.Show vbModal
            If vOpcion = 0 Then Exit Function
        Else
            vOpcion = NoHacerPregunta
        End If
        
        Set cOri = New Ccarpetas
        
        If NoHacerPregunta = 0 Then Cortar11 = NodoOrigen.Key
            'Mover de drag and drop
        If cOri.Leer(CInt(Mid(Cortar11, 2)), (ModoTrabajo = 1)) = 1 Then
            Set cOri = Nothing
            Exit Function
        End If
                
        If vOpcion = 2 Then
            CO = 1
            'Es mover, luego tenemos k comprobar si tiene permisos sobre la carpeta origen
            'Ya k es hacer un borrar
            
            If vUsu.codusu = 0 Then
                CO = 0
            Else
                If cOri.userprop = vUsu.codusu Or (cOri.escriturag And vUsu.Grupo) Then CO = 0
            End If
        
            If CO = 1 Then
                Set cOri = Nothing
                MsgBox "No tiene permisos", vbExclamation
                Exit Function
            End If
        End If
        
        
        
        
        Set Car = New Ccarpetas
        If NoHacerPregunta = 0 Then pegar11 = NodoSeleccionado.Key
        If Car.Leer(CInt(Mid(pegar11, 2)), (ModoTrabajo = 1)) = 1 Then
            Set Car = Nothing
            Set cOri = Nothing
            Exit Function
        End If
        CO = 1
        If vUsu.codusu = 0 Then
            CO = 0
        Else
            If (Car.userprop = vUsu.codusu Or ((Car.escriturag And vUsu.Grupo))) Then CO = 0
        End If
            
        If CO = 1 Then
            MsgBox "No tiene permisos sobre la carpeta destino", vbExclamation
            Set Car = Nothing
            Set cOri = Nothing
            Exit Function
        End If
               
        If NoHacerPregunta = 0 Then
            'Borramos tmpFich1
            BorrarTemporal1
            CO = 0
            For i = 1 To ListView1.ListItems.Count
                If ListView1.ListItems(i).Selected Then
                     'Si tiene permisos lo añadimos
                    InsertaTemporal CLng(Mid(ListView1.ListItems(i).Key, 2))
                    CO = CO + 1
                End If
            Next i
            
            If CO = 0 Then Exit Function
        End If
        
        
        With frmMovimientoArchivo
            .Opcion = 3 + vOpcion
            .Destino = Mid(pegar11, 2)
            
            'Origen destino
                Set .vDestino = Car
                Set .vOrigen = cOri
            
            .Show vbModal
        End With
        RealizarMover = 1

Set TreeView1.SelectedItem = NodoSeleccionado
Screen.MousePointer = vbDefault
Exit Function
errorhndl:
    MsgBox "Error :  " & Err.Number & " -- Descripción: " & Err.Description, vbCritical + vbOKOnly
    Screen.MousePointer = vbDefault
End Function




Private Function RealizarMoverCarpetas() As Byte
Dim cad As String
Dim i As Integer
Dim CO As Integer
Dim cOri As Ccarpetas
'Empiparemos
' origen: KEY padre, key hijo
' origen: KEY padre, carpeta
Dim miOrigen As String
Dim miDestino As String


'copiamos el archivo origen en destino
' a nivel de archivos
RealizarMoverCarpetas = 0
If ModoTrabajo <> vbNorm Then
    Mensajes1 (6)
    Exit Function
End If

On Error GoTo errorhndl
    If NodoOrigen.FullPath = NodoSeleccionado.FullPath Then
        MsgBox "No se puede realizar MOVER archivos sobre una misma carpeta", vbInformation
        Set TreeView1.SelectedItem = NodoOrigen
        Exit Function
        
    End If

    If NodoOrigen.Parent Is Nothing Then
        MsgBox "No se puede copiar / mover  carpetas de primer nivel", vbExclamation
        Set TreeView1.SelectedItem = NodoOrigen
        Exit Function
    End If
    
    
    'Intenta copiar el nodo sobre sel padre del nodo, es decir, intenta copiar el hijo del padre... el mismo
    If NodoOrigen.Parent.Key = NodoSeleccionado.Key Then
        MsgBox "Es la misma carpeta", vbExclamation
        Exit Function
    End If
    
    
    'Vemos si intenta copiar una carpeta dentro de su subcarpeta
    i = InStr(1, NodoSeleccionado.FullPath & "\", NodoOrigen.FullPath & "\")
    If i > 0 Then
        MsgBox "No pude copiar/mover carpetas dentro de sus subcarpetas", vbExclamation
        Exit Function
    End If
    
    'Leemos la carpeta
    'Donde va
    Set Car = New Ccarpetas
    If Car.Leer(CInt(Mid(NodoSeleccionado.Key, 2)), (ModoTrabajo = 1)) = 1 Then GoTo errorhndl
    
    If vUsu.codusu > 0 Then
        If Not ((Car.userprop = vUsu.codusu) Or CBool((Car.escriturag And vUsu.Grupo))) Then
            MsgBox "No tiene permiso sobre la carpeta destino", vbExclamation
            GoTo errorhndl
        End If
    End If
    
        'Comprobaremos si ya existe una carpeta con ese nombre
        Set miRSAux = New ADODB.Recordset
        cad = "Select * from carpetas where padre="
        cad = cad & Mid(NodoSeleccionado.Key, 2)
        
'
'        If NodoSeleccionado.Parent Is Nothing Then
'            Cad = Cad & "0"
'        Else
'            Cad = Cad & Mid(NodoSeleccionado.Parent, 2)
'        End If
        cad = cad & " AND nombre = """ & NodoOrigen.Text & """"
        miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        CO = 0
        If Not miRSAux.EOF Then
            If Not IsNull(miRSAux.Fields(0)) Then CO = 1
        End If
        miRSAux.Close
        Set miRSAux = Nothing
        
        If CO = 1 Then
            MsgBox "Ya existe la carpeta: " & NodoOrigen.Text, vbExclamation
            GoTo errorhndl
        End If
        
        
                
        
        
        
        'Hacemos la pregunta
        vOpcion = 0
        Set frmP = New frmPregunta
        frmP.Opcion = 2
        frmP.origenDestino = NodoOrigen.FullPath & "|" & NodoSeleccionado.FullPath & "|"
        frmP.Show vbModal
        Set frmP = Nothing
        If vOpcion = 0 Then GoTo errorhndl
       
        
        
        Set cOri = New Ccarpetas
        If cOri.Leer(CInt(Mid(NodoOrigen.Key, 2)), (ModoTrabajo = 1)) = 1 Then
            Set cOri = Nothing
            GoTo errorhndl
        End If
                
                
        'MOVER MOVER
        If vOpcion = 2 Then
            CO = 1
            'Es mover, luego tenemos k comprobar si tiene permisos sobre la carpeta origen
            'Ya k es hacer un borrar
            
                
            If cOri.userprop = vUsu.codusu Or (cOri.escriturag And vUsu.Grupo) Then CO = 0
            If vUsu.codusu = 0 Then CO = 0
        
            If CO = 1 Then
                MsgBox "No tiene permisos sobre la carpeta origen.", vbExclamation
                GoTo errorhndl
            End If
            
            
            'Comprobaremos k no existen carpetas no mostradas para este
            'usuario
            If ComprobarCarpetasOcultasDentroCarpeta(cOri) = 1 Then
                MsgBox "La carpeta contiene archivos / carpetas ocultos. No se puede mover", vbExclamation
                GoTo errorhndl
            End If
            
        End If
        
                
        

        'Borramos tmpFich1
        BorrarTemporal1
        
       
        
        
        CO = 0
        'Si es copiar , solo copiaremos las que esten visibles
        
        If vOpcion = 2 Then
            'Mover
            'Lo unico  haremos sera updatear el padre de la carpeta origen, al nuevo padre
            cad = "UPDATE carpetas set padre= " & Mid(NodoSeleccionado.Key, 2)
            cad = cad & " WHERE codcarpeta =" & Mid(NodoOrigen.Key, 2)
            Conn.Execute cad
            RealizarMoverCarpetas = 1
            
            'MEto en cad, para k luego lo refresque
            If NodoOrigen.Parent Is Nothing Then
                cad = NodoOrigen.Key
            Else
                cad = NodoOrigen.Parent.Key
            End If
            cad = cad & "|" & NodoSeleccionado.Key & "|"
            RecargarDatos
            MostrarNodosDespuesMover cad
            GoTo errorhndl
        Else
            'Copiar
            cad = NodoSeleccionado.Key
            CopiarCarpetas
            RecargarDatos
            
            'AQUI ES DONDE DEBEMOS guardar en variables para el refresco
            For i = 1 To TreeView1.Nodes.Count
                If TreeView1.Nodes(i).Key = NodoOrigen.Key Then
                    TreeView1.Nodes(i).EnsureVisible
                Else
                    'Este es el destno,hay que abrir sus subnodos
                    If TreeView1.Nodes(i).Key = cad Then
                        Set NodoSeleccionado = Nothing
                        Set NodoSeleccionado = TreeView1.Nodes(i)
                        TreeView1.Nodes(i).EnsureVisible
                    End If
                End If
            Next i
            'El nodo destino, ahora lo abrimos
            Set NodoOrigen = Nothing
            If NodoSeleccionado.Children > 0 Then NodoSeleccionado.Child.EnsureVisible
            
            
        End If
        
        If CO = 0 Then GoTo errorhndl
        
        
        Set TreeView1.SelectedItem = NodoSeleccionado
        


errorhndl:
    If Err.Number <> 0 Then
        MsgBox "Error :  " & Err.Number & " -- Descripción: " & Err.Description, vbCritical + vbOKOnly
        
    End If
    Set Car = Nothing
    Set cOri = Nothing
    Screen.MousePointer = vbDefault
End Function


Private Sub MostrarNodosDespuesMover(ByRef CarpetasInvolucradas As String)
Dim i As Integer
Dim B As Byte
    
    
    For i = 1 To TreeView1.Nodes.Count
        If InStr(1, CarpetasInvolucradas, TreeView1.Nodes(i).Key & "|") > 0 Then
            
            B = B + 1
            TreeView1.Nodes(i).Expanded = True
            TreeView1.Nodes(i).EnsureVisible
            If TreeView1.Nodes(i).Children > 0 Then TreeView1.Nodes(i).FirstSibling.EnsureVisible
            
        End If
        If B > 1 Then Exit For
        
    Next i
        
    

End Sub


Private Sub RecargarDatos()
Dim i As Integer

    CargaArbol
    
    
    
    
    ListView1.ListItems.Clear
    If NodoOrigen Is Nothing Then
       i = 1
    Else
        For i = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(i).Key = NodoOrigen.Key Then Exit For
        Next i
        If i > TreeView1.Nodes.Count Then i = 1
    End If

    Set TreeView1.SelectedItem = TreeView1.Nodes(i)
    TreeView1.SelectedItem.EnsureVisible
    TreeView1.SelectedItem.Expanded = True
    If TreeView1.SelectedItem.Children > 0 Then TreeView1.SelectedItem.FirstSibling.EnsureVisible
    TreeView1_NodeClick TreeView1.Nodes(i)
    
    
    'MostrarArchivos TreeView1.SelectedItem.Key
    'Set NodoOrigen = TreeView1.SelectedItem
    'Set NodoSeleccionado = TreeView1.SelectedItem
    
End Sub

Private Function ImprimirArchivos(ByRef Mostrar As Boolean) As Boolean
'Dim NuevoDoc As Word.Document
Dim Cont As Integer
Dim sel As Boolean ' para decir si hay archivos seleccionados

    
    
Exit Function
ErrHan1:
        MsgBox "Error: " & Err.Number & vbCrLf & "Descripción: " & Err.Description, vbCritical
End Function

Private Function NodoseleccionadoConsulta() As Boolean
  On Error GoTo ENodoseleccionadoConsulta
    NodoseleccionadoConsulta = False
  If NodoSeleccionado.Parent Is Nothing Then
        MsgBox "No se pueden insertar archivos en la carpeta Raiz", vbInformation
        Exit Function
  End If

    NodoseleccionadoConsulta = True
Exit Function
ENodoseleccionadoConsulta:
    MuestraError Err.Number, "Nodo seleccionado Consulta"
End Function

Private Sub Insertar()

  If NodoSeleccionado Is Nothing Then Exit Sub
  If ModoTrabajo <> vbNorm Then
        Mensajes1 (5)
        Exit Sub
  End If
        
        
        
        
    If Not NodoseleccionadoConsulta Then Exit Sub
    
    
    
    
    'Comprobamos si el usuarios tiene permiso de
    Set Car = New Ccarpetas
    
    If Car.Leer(CInt(Mid(NodoSeleccionado.Key, 2)), (ModoTrabajo = 1)) = 0 Then
        
        'OK. VEMOS EL Permiso
        If Car.userprop = vUsu.codusu Or (Car.escriturag And vUsu.Grupo) Or vUsu.codusu = 0 Then
            
            Set frmNuevoArchivo.Mc = Car
            frmNuevoArchivo.Opcion = 0
            frmNuevoArchivo.Carpeta = Text1.Text
            DatosMOdificados = False
            frmNuevoArchivo.Show vbModal
            If DatosMOdificados Then
                Screen.MousePointer = vbHourglass
                MostrarArchivos TreeView1.SelectedItem.Key
                TreeView1.Drag vbCancel
                ListView1.Drag vbCancel
                Screen.MousePointer = vbDefault
            End If
        Else
            MsgBox "No tiene permiso ", vbExclamation
        End If
    Else
        MsgBox "Error leyendo carpeta : " & NodoSeleccionado.Text
    End If
  Screen.MousePointer = vbDefault

End Sub


Private Sub Eliminar()
Dim Rc As Byte
Dim i As Long
Dim Cont As Long
Dim Mens As String
Dim Img As cTimagen
Dim SinPer As Long
Dim TienePermiso As Boolean

    On Error GoTo errorhndl
    If ModoTrabajo <> vbNorm Then
        Mensajes1 (2)
        Exit Sub
    End If
    
    
    'Leemos la carpeta
    Set Car = New Ccarpetas
    If Car.Leer(CInt(Mid(TreeView1.SelectedItem.Key, 2)), (ModoTrabajo = 1)) = 1 Then
        MsgBox "Error grave leyendo datos carpeta", vbExclamation
        Set Car = Nothing
        Exit Sub
    End If

    Cont = 0
    Mens = ""
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected Then
            Cont = Cont + 1
            Mens = Mens & " - " & ListView1.ListItems(i).Text & vbCrLf
        End If
    Next i
    If Cont = 0 Then
        Set Car = Nothing
        Exit Sub
    End If
    
    'PERMISOS
    If vUsu.codusu > 0 Then
        i = 1
        If (Car.userprop <> vUsu.codusu) Then
            If (Car.escriturag And vUsu.Grupo) Then i = 0
        Else
            i = 0
        End If
        If i = 1 Then
            'No tienen permisos de escritura sobre la carpeta
            MsgBox "No tiene permiso de escritura sobre la carpeta", vbExclamation
            Set Car = Nothing
            Exit Sub
        End If
    
    End If
    
    
    
    If Cont = 1 Then
        Mens = "Seguro que desea eliminar el archivo " & ListView1.SelectedItem.Text & "?"
        Else
            Mens = " Seguro que desea eliminar los siguientes archivos : ?" & vbCrLf & vbCrLf & Mens
        End If

    Rc = MsgBox(Mens, vbQuestion + vbYesNoCancel + vbDefaultButton2)
    Screen.MousePointer = vbHourglass
    If Rc = vbYes Then
         BorrarTemporal1
         Mens = "INSERT INTO tmpfich(codusu,codequipo,imagen) VALUES (" & vUsu.codusu & "," & vUsu.PC & ","
         Set Img = New cTimagen
         SinPer = 0 'Sin permiso
         For i = 1 To ListView1.ListItems.Count
             If ListView1.ListItems(i).Selected Then
                '-----------------------------------
                
                
                
                If Img.Leer(Val(Mid(ListView1.ListItems(i).Key, 2)), objRevision.LlevaHcoRevision) = 0 Then
                   TienePermiso = False
                   If vUsu.codusu = 0 Then
                        TienePermiso = True
                    Else
                        If Img.userprop = vUsu.codusu Or (Img.escriturag And vUsu.Grupo) Then TienePermiso = True
                    End If
                
                   If TienePermiso Then
                        'Si tiene permiso.  Lo metemos en la tabla para k elimine
                        Conn.Execute Mens & Mid(ListView1.ListItems(i).Key, 2) & ")"
                        
                    Else
                        SinPer = SinPer + 1
                    End If
                End If
             End If
         Next i
         
         Set Img = Nothing
         Mens = "NO"  'Para ver si seguimos adelante con el proceso
         If SinPer > 0 Then
            'No tiene permisos sobre alguno de los archivos
            If SinPer = Cont Then
                'No tienen permisos sobre NINGUNO de los archivos
                If Cont = 1 Then
                    MsgBox "No tiene permisos sobre el archivo a eliminar", vbExclamation
                Else
                    MsgBox "No tiene permisos sobre ninguno de los archivos a eliminar", vbExclamation
                End If
            Else
                Mens = "No tiene permiso sobre alguno de los archivos" & vbCrLf & "¿Desea continuar igualmente?"
                If MsgBox(Mens, vbQuestion + vbYesNoCancel) = vbYes Then Mens = ""
            End If
         Else
            'Todos se pueden borrar
            Mens = ""
         End If
         
        If Mens = "" Then
             'Ahora lanzaremos movimientos de archivos a piñon
            Set frmMovimientoArchivo.vDestino = Car
            frmMovimientoArchivo.Opcion = 8
            frmMovimientoArchivo.Show vbModal
            MostrarArchivos TreeView1.SelectedItem.Key
            
            
            
        End If
    End If ' del rc=vbyes
    Screen.MousePointer = vbDefault
Exit Sub
errorhndl:
    'Por si quieres hacer algo con el error
    
    Screen.MousePointer = vbDefault
End Sub




Private Sub HacerCortar()
Dim i As Long
Dim HanCortado As Boolean
        If ModoTrabajo <> vbNorm Then
            Mensajes1 (3)
            Exit Sub
            End If
        Cortar11 = TreeView1.SelectedItem.Key
         ' aqui tenemos el path inicial
        'Set lista = New Collection
        Set listacod = New Collection
        HanCortado = False
        
        
        'Nuevo
        For i = 1 To ListView1.ListItems.Count
            If ListView1.ListItems(i).Selected Then
                listacod.Add Mid(ListView1.ListItems(i).Key, 2)
                ListView1.ListItems(i).Ghosted = True
                HanCortado = True
             End If
        Next i
        
        Toolbar1.Buttons(6).Enabled = HanCortado
        mncopiar.Enabled = HanCortado
        mnPegar = HanCortado
        Toolbar1.Buttons(5).Enabled = Not HanCortado
        mncortar2 = Not HanCortado
        mncortar.Enabled = Not HanCortado
        Me.Refresh
End Sub


Private Sub HacerPegar()
Dim i As Integer


        If ModoTrabajo <> vbNorm Then
            Mensajes1 (4)
            Exit Sub
        End If
        pegar11 = TreeView1.SelectedItem.Key

         ' aqui tenemos el path final
        Toolbar1.Buttons(5).Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        mncopiar.Enabled = False
        mncortar.Enabled = False
        mncortar2 = False
        mnPegar = False
        If pegar11 = Cortar11 Then
            MsgBox "No se puede realizar PEGAR sobre una misma carpeta", vbInformation
            Exit Sub
        End If

        Screen.MousePointer = vbHourglass
        BorrarTemporal1
        
        For i = 1 To listacod.Count
            InsertaTemporal Val(listacod(i))
        Next i
                  
        'Hacemos lo mismo que en mover
        'pero no haremos pregunta
        If RealizarMover(2) = 0 Then MostrarArchivos TreeView1.SelectedItem.Key
        
End Sub


Public Sub VerPropiedades(CadenaCodigo As String, LecturaSolo As Boolean)
Dim vImg As cTimagen
Dim SoloLeer As Boolean

    
    Screen.MousePointer = vbHourglass ' luego en el form.load lo ponemos a normal
    Set vImg = New cTimagen
    
    If vImg.Leer(CLng(Mid(CadenaCodigo, 2)), objRevision.LlevaHcoRevision) = 0 Then
        SoloLeer = True
        If Not LecturaSolo Then
             If vImg.userprop = vUsu.codusu Or (vImg.escriturag And vUsu.Grupo) Or vUsu.codusu = 0 Then SoloLeer = False
        End If
        If ModoTrabajo <> vbNorm Then
            LecturaSolo = True
            SoloLeer = True
        End If
        If SoloLeer Then
            frmNuevoArchivo.Opcion = 3
        Else
            frmNuevoArchivo.Opcion = 2
        End If
        Set frmNuevoArchivo.mImag = vImg
        frmNuevoArchivo.Carpeta = Text1.Text
        DatosMOdificados = False
        frmNuevoArchivo.Show vbModal
        If Not LecturaSolo Then
            If DatosMOdificados Then
                If ListView1.View = lvwIcon Then
                    ListView1.SelectedItem.Text = vImg.campo1
                Else
                    pegar11 = "Select " & vUsu.preferencias.vSelect
                    pegar11 = pegar11 & ", escriturag,codigo,codext"
                    pegar11 = pegar11 & " from timagen WHERE codigo = " & vImg.codigo
                    Set miRSAux = New ADODB.Recordset
                    miRSAux.Open pegar11, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    If miRSAux.EOF Then
                        MsgBox "Error recargando datos: " & vImg.codigo, vbExclamation
                    Else
                        PonerDatosLineaListviewReport ListView1.SelectedItem, miRSAux
                    End If
                    Set miRSAux = Nothing
                    pegar11 = ""
                End If
            End If
        End If
    End If
    
    Set vImg = Nothing
    Screen.MousePointer = vbDefault
End Sub






Private Function HazIntegracion() As Byte
'Dim fso2
'Dim a
'Dim EnError As String
'Dim linea As String
'Dim rc As Byte
'
'
'
'        rc = DevuelveExtension("iux")  'La extension de los archivos sera iux
'        If rc = 0 Then
'            Screen.MousePointer = vbDefault
'            Exit Function ' NO hay un visor txt
'        End If
'        EnError = ""
'
'        Set fso2 = New FileSystemObject
'        Set a = fso2.OpenTextFile(CarpetaIntegracion & ArchivoIntegracion)
'        While Not a.AtEndOfStream
'            linea = a.ReadLine
'            linea = Trim(linea)
'            If linea <> "" Then
'                If ProcesarLinea(linea, rc) = 1 Then _
'                        EnError = EnError & linea & vbCrLf
'            End If
'        Wend
'        a.Close
'
'        If EnError = "" Then
'            HazIntegracion = 0
'            Exit Function
'            End If
'        If Not fso2.FolderExists(App.Path & "\ErrInt") Then fso2.CreateFolder (App.Path & "\ErrInt")
'        ' NO ha ido bien
'        FileCopy CarpetaIntegracion & ArchivoIntegracion, App.Path & "\ErrInt\" & ArchivoIntegracion
'        Kill CarpetaIntegracion & ArchivoIntegracion
'        Set a = fso2.CreateTextFile(App.Path & "\ErrInt\2" & ArchivoIntegracion)
'        a.WriteLine ("#############################################")
'        a.WriteLine ("###                                       ###")
'        a.WriteLine ("###                                       ###")
'        a.WriteLine ("###  Error en la Integración de ficheros  ###")
'        a.WriteLine ("###                                       ###")
'        a.WriteLine ("###                                       ###")
'        a.WriteLine ("#############################################")
'        a.WriteBlankLines (2)
'        a.WriteLine (" Fecha : " & Date)
'        a.WriteLine ".-Las siguientes lineas no han podido ser integradas"
'        a.WriteBlankLines (1)
'        a.WriteLine (EnError)
'        a.WriteBlankLines (3)
'        a.WriteLine ("Consulte tambien el archivo ")
'        a.WriteLine ("    " & App.Path & "\ErrInt\" & ArchivoIntegracion)
'        a.Close
'        HazIntegracion = 1
End Function


'
'                                          ' cod de extension
'Private Function ProcesarLinea(Lin As String, rc As Byte) As Byte
'Dim Fin As Boolean
'Dim cadena As String
'Dim pos As Integer
'Dim valor As String
'Dim ind As Integer
'Dim mImg As CImag
'Dim nombre As String
'Dim vExt As String
'
'On Err GoTo ErrorHdle
'cadena = Lin
'Fin = False
'ind = 0
'Set mImg = New CImag
'mImg.Siguiente
'While Not Fin
'    If Mid(cadena, 1, 1) = "|" Then 'Con esto comprobamos que los campos estan llenos
'        valor = ""
'        cadena = Mid(cadena, 2, Len(cadena))
'        ind = ind + 1
'    Else
'        pos = InStr(1, cadena, "|")
'
'        If pos > 0 Then
'            valor = Mid(cadena, 1, pos - 1)
'            cadena = Mid(cadena, pos + 1, Len(cadena))
'            ind = ind + 1
'            Else
'                Fin = True
'                ind = ind + 1
'                valor = cadena
'        End If
'    End If
'
'Select Case ind
'Case 1: 'Nombre archivo
'        mImg.NomFich = valor
'Case 2: 'Clave1
'        nombre = ProcesaLinea2(valor)
'        nombre = Trim(nombre)
'        mImg.Clave1 = nombre
'Case 3: 'c2
'        mImg.Clave2 = valor
'Case 4: 'c3
'        mImg.Clave3 = valor
'Case 5: 'fechadoc
'        mImg.FechaDoc = valor
'Case 6: 'path
'        mImg.NomPath = devuelvePATH(valor)   'Le cambiamos las barras
'Case Else
'    valor = valor 'Sumidero para que no de error
'
'End Select
'Wend
'
'ProcesarLinea = 0
'' Antes de añadirlo le ponemos como fecha de digitalización la de hoy
'mImg.FechaDig = Date
'
'If mImg.Clave1 = "" Or mImg.NomFich = "" Or mImg.NomPath = "" Then
'    'Error leyendo los datos
'    ProcesarLinea = 1
'    Exit Function
'    End If
'
''Comprobamos que existe el archivo
'valor = Dir(CarpetaIntegracion & mImg.NomFich)
'vExt = ""
'If valor = "" Then
'    'Error leyendo los datos
'    'Nueva modificacion. Por si acaso vienen con .dat
'    If Dir(CarpetaIntegracion & mImg.NomFich & ".dat") = "" Then
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
'    End If
'
''Comprobaremos el directorio
'If TratarCarpeta(mImg.NomPath) = 1 Then
'    'Se ha producido un error al crear una de las carpetas
'    ProcesarLinea = 1
'    Exit Function
'    End If
'
'nombre = CarpetaIntegracion & mImg.NomFich & vExt
'mImg.NomFich = mImg.ID & ".iux"
'mImg.Extension = rc
'
'mImg.NomPath = mImg.NomPath & "\"
'If CompruebaCarpeta(mImg.NomPath, valor) = 3 Then  '1. con archivos  2.-vacio
'    'Se ha producido un error al crear una de las carpetas
'    ProcesarLinea = 1
'    Exit Function
'    End If
'
'
''Construimos en directorio temporal el archivo de integracion -doc
'If ConstruyeWord(nombre) = 1 Then
'    'Error constuyendo el archivo word
'    ProcesarLinea = 1
'    Exit Function
'    End If
'
''Ahora el archivo esta en c:\windows\temp\aridoc12.tmp
'' luego lo asignamos a nombre, no sin antes eliminar el txt anterior
'' Puede que deberiamos de trabajar directamente sobre el nombre final
'Kill nombre
'nombre = App.Path & "\temp\aridoc12.tmp"
'
'If mImg.Agregar = 0 Then
'    FileCopy nombre, dirbase & "\" & mImg.NomPath & mImg.NomFich
'    Kill nombre
'    ProcesarLinea = 0
'    End If
'
'Set mImg = Nothing
'Exit Function
'ErrorHdle:
'    MsgBox "Err: " & Err.Number & vbCrLf & " Des:  " & Err.Description, vbExclamation
'End Function





'Private Sub AnchoColumna(opcion As Integer)
'' Almacenaremos o guardaremos el ancho de la columna
'' segun sea
''    1   leer -> Lo leeremos de un archivo
''    2   escribir -> Lo escribiremos al archivo
''    3   modificar -> Se produce cuando cambiamos de lista a iconos
'Dim a
'
''On Err GoTo ErrHandler
'
'If opcion = 3 Then
'    Col1 = ListView1.ColumnHeaders.Item(1).Width
'    Col2 = ListView1.ColumnHeaders.Item(2).Width
'    Col3 = ListView1.ColumnHeaders.Item(3).Width
'    Col4 = ListView1.ColumnHeaders.Item(4).Width
'    Col5 = ListView1.ColumnHeaders.Item(5).Width
'    Col6 = ListView1.ColumnHeaders.Item(6).Width
'    Col7 = ListView1.ColumnHeaders.Item(7).Width
'    Exit Sub
'    End If
'
'
'End Sub
'
'




'Private Sub InsertaDocWord()
'Dim rc
'        rc = DevuelveExtension("doc")
'        If rc = 0 Then Exit Sub
'        HaSidoCancelado = True ' Le ponemos el valor por defecto
'        Set Img = New CImag
'        Img.Siguiente
'        Img.NomPath = CarpetaW
'        Img.NomFich = Img.Id & ".doc"
'        Img.Extension = rc
'        FileCopy NombreDoc, Img.NomPath & Img.NomFich
'        Set frmImagen1 = New frmImagen
'        Set frmImagen1.mImg = Img
'        frmImagen1.modificar = False
'        frmImagen1.Show vbModal
'        Set frmImagen1 = Nothing
'        If Not HaSidoCancelado Then
'            Screen.MousePointer = vbHourglass
'            MostrarArchivos (inicial & NodoSeleccionado.FullPath & "\")
'            Screen.MousePointer = vbDefault
'            Else
'                ' ha sido cancelado. Habra que borrar el archivo
'                Kill Img.NomPath & Img.Id & ".*"
'        End If
'End Sub

Private Sub MoverArchivoErroneo(Origen As String, Destino As String)
On Error GoTo ErrMover
'    FileCopy origen, destino
'    Kill origen
    'Probamos con la opcion name
    Name Origen As Destino
    MsgBox "Con anterioridad no se borro el archivo " & vbCrLf & Origen & vbCrLf & _
        "Sa ha movido como: " & vbCrLf & "    " & Destino, vbCritical
Exit Sub
ErrMover:
    MostrarError Err.Number
End Sub
        
Private Sub BorrarTemporal()
Dim cad As String
On Error Resume Next
    If Dir(App.Path & "\temp\*.*") <> "" Then Kill App.Path & "\temp\*.*"
End Sub



''-------------------------------------------------------------------------------------
'' Modificacion del 25 de Abril
'
'
''LAs extensiones se cargaran en un vector de extensiones
'Private Function CargaExtensiones() As String
'Dim RS As Recordset
'
'On Error GoTo ECargaExtensiones
'CargaExtensiones = ""
'Set RS = Db.OpenRecordset("Select * from TExtension ORDER BY Cod", 2)
'While Not RS.EOF
'    If RS!Cod > 15 Then
'        CargaExtensiones = "El codigo de extension NO debe superar el nº15. Consulte en ARIADNA SOFTWARE."
'        Exit Function
'    End If
'    VectorExt(RS!Cod) = LCase(RS!Extension)
'    RS.MoveNext
'Wend
'RS.Close
'Set RS = Nothing
'
'
''Tambien cargamos esta variable
'LongitudIncial = Len(inicial & Carpeta & "\") + 1
'
'Exit Function
'ECargaExtensiones:
'    CargaExtensiones = Err.Description
'End Function
'
'
'Private Function DevuelveExtensionNuevo(Cad As String) As Integer
'Dim i As Integer
'
'For i = 0 To 15
'    If Cad = VectorExt(i) Then
'        DevuelveExtensionNuevo = i
'        Exit Function
'    End If
'Next i
'DevuelveExtensionNuevo = -1
'End Function




Private Sub AnyadeError(ByRef Nom As String, ByRef NoEncontrado As String)
'On Error GoTo EAnyadeError
'Dim Cad As String
'
'Db.Execute " INSERT INTO Temporal " _
'        & "(Id,Archivo,ParaBorrar) VALUES " _
'        & "(" & NErrores & ",'" & nom & "', " & NoEncontrado & ");"
'
'EAnyadeError:
'    Err.Clear
End Sub



Private Sub Integracion()
    
    Cortar11 = DevuelveDesdeBD("exeintegra", "equipos", "codequipo", vUsu.PC, "N")
    If Cortar11 <> "" Then
        If Dir(Cortar11, vbArchive) = "" Then
            MsgBox "Configuraciòn erronea para el equipo. No existe integrador: " & Cortar11, vbExclamation
        Else
            
            Cortar11 = """" & Cortar11 & """"
            pegar11 = Caption
            Caption = "INTEGRANDO FICHEROS"
            Me.Refresh
            
            
            
            'Esperms a que vuelva
            'LanzaArchivoModificar Cortar11
            Shell Cortar11, vbNormalFocus
            
            
            'CargaArbol
            'If TreeView1.Nodes.Count > 1 Then TreeView1.Nodes(2).EnsureVisible
            Caption = pegar11
            Cortar11 = ""
            pegar11 = ""
        End If
    End If
End Sub

'Private Sub Integracion()
'Dim rc As Byte
'
'SeHanCreadoCarpetas = False
''De pendiendo de la forma de integrar llamaremos al integrar de antes o al integrar nuevo
'If mConfig.TipoIntegracion = 0 Then
'    Integracion2
'    Else
'    'Nueva forma de integrar, pero comprobamos que hay archivos
'    'Ciertas comprobaciones
'
'    'EXISTE LA CARPETA DE INTEGRACION
'    If Dir(mConfig.carpetaInt, vbDirectory) = "" Then
'        MsgBox "La carpeta de integración : " & vbCrLf & _
'            "     " & CarpetaIntegracion & vbCrLf & _
'            "  No existe.   ", vbCritical
'        Exit Sub
'    End If
'
'    If Dir(mConfig.carpetaInt, vbArchive) = "" Then
'        'Vacia , no abrmos nada
'        Exit Sub
'    End If
'
'    'Reconoce la nueva extension
'    'Vemos si el word es soportado por la aplicacion
'    rc = DevuelveExtension("nfi")
'    If rc = 0 Then
'        MsgBox "Tienes que agregar la extension NFI en la configuración, con el programa ARIVISOR " & vbCrLf, vbInformation
'        Exit Sub
'    End If
'
'    If MsgBox("Hay archivo pendientes de integrar. Realizar la integración ahora?", vbQuestion + vbYesNoCancel) <> vbYes Then _
'        Exit Sub
'    'Llegados a este punto abriremos el formulario de integracion
'    frmIntegracion.NFI_Extension = rc
'    frmIntegracion.Show vbModal
'End If  'Del tipo de integracicon
'
'If SeHanCreadoCarpetas Then
'        Screen.MousePointer = vbHourglass
'        CargaArbolRec
'        Screen.MousePointer = vbDefault
'End If


'end Sub




'-------------------------------------------------------
'Los iconos iran desplazados UNO para poder poner el DEFAULT


Private Sub CargarListviews()
Dim i As Integer
Dim J As Integer
Dim Pos As Integer
Dim SQL As String
Dim Errores As String
    Errores = False
    If Dir(App.Path & "\Defaultico.dat", vbArchive) = "" Then
        MsgBox "Archivos necesarios para la aplicacion han sido borrados(Defaultico.dat)", vbCritical
        End
    End If

    

'Cargamos IMAGELIST con los iconos de las imagenes
    SQL = "SELECT * FROM extensionpc where codequipo=" & vUsu.PC & "  order by codext"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Errores = ""
    Pos = 1
    
    While Not miRSAux.EOF
        J = miRSAux!codext + 1
        If J > Pos Then
            For i = Pos To J - 1
                CargaIcono i, App.Path & "\Defaultico.dat"
            Next i
        End If
        Pos = J + 1
            
        'Cargamos el icono
        SQL = App.Path & "\imagenes\" & miRSAux!codext & ".ico"
        If Dir(SQL) <> "" Then
            
        Else
            Errores = Errores & SQL & vbCrLf
            SQL = App.Path & "\Defaultico.dat"
        End If
        CargaIcono J, SQL
    
        'Siguiente
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    If Pos = 1 Then CargaIcono 1, App.Path & "\Defaultico.dat"
    'Si han habido errores
    'Proponemos carga iconos
    If Errores <> "" Then
        vUsu.CargaIconosExtensiones = True
        Conn.Execute "UPDATE equipos SET cargaIconsExt= 1 WHERE codequipo=" & vUsu.PC
        SQL = "Se han producido errores cargando iconos. Reinicie la aplicacion y si continua el problema consulte con el soporte técnico:" & vbCrLf & Errores
        SQL = SQL & vbCrLf & vbCrLf & "¿Finalizar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then End
    End If
End Sub


Private Sub CargaIcono(Cod As Integer, vpath As String)
    ImageList2.ListImages.Add , "C" & Cod, LoadPicture(vpath)
    ImageList3.ListImages.Add , "C" & Cod, LoadPicture(vpath)
End Sub





Private Sub CargaList4()
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\th1.bmp")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\th2.bmp")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\th3.bmp")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\CONVT04B.ico")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\imgtool.bmp")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\NEW2.bmp")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\camera2.bmp")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\delete.bmp")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\save.bmp")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\exportar.bmp")
    ImageList4.ListImages.Add , , LoadPicture(App.Path & "\imagenes2\print.bmp")
End Sub

'Private Sub CargaList1()
'
'
'
'    ImageList1.ListImages.Add , "find", LoadPicture(App.Path & "\imagenes\find.bmp")
'    ImageList1.ListImages.Add , "abierto", LoadPicture(App.Path & "\imagenes\fileo.ico")
'    ImageList1.ListImages.Add , "cerrado", LoadPicture(App.Path & "\imagenes\filec.ico")
'
'    'ImageList1.ListImages.Add , "importar", LoadPicture(App.Path & "\imagenes\importar.bmp")
'    ImageList1.ListImages.Add , "importar", LoadPicture(App.Path & "\imagenes\modificar.ico")
'
'    ImageList1.ListImages.Add , "cortar", LoadPicture(App.Path & "\imagenes\cut.bmp")
'    ImageList1.ListImages.Add , "pegar", LoadPicture(App.Path & "\imagenes\paste.bmp")
'    ImageList1.ListImages.Add , "carpeta", LoadPicture(App.Path & "\imagenes\carpeta.bmp")
'    ImageList1.ListImages.Add , "delete", LoadPicture(App.Path & "\imagenes\delete.bmp")
'    ImageList1.ListImages.Add , "lista", LoadPicture(App.Path & "\imagenes\vw-list.bmp")
'    ImageList1.ListImages.Add , "iconos", LoadPicture(App.Path & "\imagenes\vw-lrgic.bmp")
'    ImageList1.ListImages.Add , "prop", LoadPicture(App.Path & "\imagenes\prop.bmp")
'    ImageList1.ListImages.Add , "nuevo", LoadPicture(App.Path & "\imagenes\new.bmp")
'    ImageList1.ListImages.Add , "eliminar", LoadPicture(App.Path & "\imagenes\discnet.bmp")
'    ImageList1.ListImages.Add , "imprimir", LoadPicture(App.Path & "\imagenes\print.bmp")
'
'    'Neuevo
'
'
'End Sub


Private Sub PonerPreferenciasPersonales()
        
       If vUsu.preferencias.Vista = lvwReport Then
            Me.Toolbar1.Buttons(10).Value = tbrPressed
            Me.Toolbar1.Buttons(11).Value = tbrUnpressed
       Else
            Me.Toolbar1.Buttons(11).Value = tbrPressed
            Me.Toolbar1.Buttons(10).Value = tbrUnpressed
        End If
            
End Sub

Private Function CompruebaCarpeta(ClaveCarpeta As String) As Boolean
Dim C As Ccarpetas


    CompruebaCarpeta = False
    Set C = New Ccarpetas
    If C.Leer(CInt(Mid(ClaveCarpeta, 2)), (ModoTrabajo = 1)) = 0 Then
        If vUsu.codusu = 0 Then
            CompruebaCarpeta = True
        Else
            If C.userprop = vUsu.codusu Or (C.escriturag And vUsu.Grupo) Then
                CompruebaCarpeta = True
            Else
                MsgBox "No tiene permiso sobre la carpeta destino", vbExclamation
            End If
        End If
    End If
    
    Set C = Nothing

End Function


Public Function ComprobarCarpetasOcultasDentroCarpeta(ByRef Carpe As Ccarpetas) As Byte
Dim Nod As Node
    If NodoOrigen Is Nothing Then Exit Function
    ComprobarCarpetasOcultasDentroCarpeta = 1
    Set miRSAux = New ADODB.Recordset
    If Not TieneCarpetasOcultas(NodoOrigen) Then ComprobarCarpetasOcultasDentroCarpeta = 0
    Set miRSAux = Nothing
End Function



Private Function TieneCarpetasOcultas(No As Node) As Boolean
Dim cad As String
Dim ContNodos As Long
Dim vHijo As Node
Dim Fin As Boolean
Dim B As Boolean
    TieneCarpetasOcultas = True
    If No.Children > 0 Then
        Set vHijo = No.Child
        Fin = False
        Do
            B = TieneCarpetasOcultas(vHijo)
            If B Then
                Exit Function
            Else
                If vHijo = No.Child.LastSibling Then
                    Fin = True
                Else
                    Set vHijo = vHijo.Next
                End If
            End If
        Loop Until Fin
    End If
    cad = "Select count(*) from carpetas where padre = " & Mid(No.Key, 2)
    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ContNodos = 0
    If Not miRSAux.EOF Then
        If Not IsNull(miRSAux.Fields(0)) Then ContNodos = miRSAux.Fields(0)
    End If
    miRSAux.Close
    If No.Children = ContNodos Then TieneCarpetasOcultas = False
    
End Function




Private Function CopiarCarpetas()
    Set miRSAux = New ADODB.Recordset
    'Sera un procesor recursivo. Para cada carpeta iremos metiendo los archivos
    CopiaArchivosCarpetaRecursiva NodoOrigen, NodoSeleccionado.Key
    Set miRSAux = Nothing
End Function

    
Private Sub CopiaArchivosCarpetaRecursiva(No As Node, ElPadre As String)
Dim Nod As Node
Dim J As Integer
Dim i As Integer
Dim CodigoPadre As Integer
Dim CodigoNuevaCarpeta As Integer

    'Primero copiamos la carpeta
    CodigoPadre = CInt(Mid(ElPadre, 2))
    If CopiaArchivosDeLaCarpeta(No, CodigoPadre, CodigoNuevaCarpeta) Then
        If No.Children > 0 Then
            J = No.Children
            Set Nod = No.Child
            For i = 1 To J
               CopiaArchivosCarpetaRecursiva Nod, "C" & CodigoNuevaCarpeta
               If i <> J Then Set Nod = Nod.Next
            Next i
        End If
    End If
End Sub


Private Function CopiaArchivosDeLaCarpeta(No As Node, vPadre As Integer, ByRef CodigoCarpetaNueva As Integer) As Boolean
Dim NuevaCarpeta As Ccarpetas
Dim vOrigen As Ccarpetas
Dim cad As String
Dim TieneArchivos As Boolean

    Set vOrigen = New Ccarpetas
    CopiaArchivosDeLaCarpeta = False
    If vOrigen.Leer(CInt(Mid(No.Key, 2)), (ModoTrabajo = 1)) = 0 Then
        
        Set NuevaCarpeta = New Ccarpetas
        'Le metemos las propiedades
        With NuevaCarpeta
            .padre = vPadre
            .Almacen = vOrigen.Almacen
            .escriturag = vOrigen.escriturag
            .lecturag = vOrigen.lecturag
            .Nombre = vOrigen.Nombre
            .pathreal = vOrigen.pathreal
            .pwd = vOrigen.pwd
            .SRV = vOrigen.SRV
            .user = vOrigen.user
            .version = vOrigen.version
            'Cambiamos el propietario
            .userprop = vUsu.codusu
            .groupprop = vUsu.GrupoPpal
        End With
        If NuevaCarpeta.Agregar = 0 Then
            'Una vez agregado tenemos k leerlo otra vez, para k meta los valores
            'Ya ha creado la carpeta
            CopiaArchivosDeLaCarpeta = True
            CodigoCarpetaNueva = NuevaCarpeta.codcarpeta
            BorrarTemporal
            'INSERT INTO tmpfich (codusu, codequipo, imagen) VALUES (1, 0, NULL)
            'Cad = "INSERT INTO tmpfich(codusu,codequipo,imagen) SELECT " & vUsu.codusu & "," & vUsu.PC & ",codigo from timagen where "
            'Cad = Cad & "codcarpeta = " & Mid(No.Key, 2)
            'Conn.Execute Cad
            ' Val(Mid(No.Key, 2))
            
            'Compruebo si tiene archivos
            TieneArchivos = False
            miRSAux.Open "Select codigo from timagen where codcarpeta =" & Mid(No.Key, 2), Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not miRSAux.EOF Then
                TieneArchivos = True
                While Not miRSAux.EOF
                    
                    InsertaTemporal miRSAux!codigo
                    miRSAux.MoveNext
                Wend
            End If
            miRSAux.Close
            
            If TieneArchivos Then
                With frmMovimientoArchivo
                    .Opcion = 4
                    .Destino = Mid(NodoSeleccionado.Key, 2)
              
        
                        Set .vDestino = NuevaCarpeta
                        Set .vOrigen = vOrigen
  
                    .Show vbModal
                End With
            End If
            
            
            
        End If
    End If
    Set vOrigen = Nothing
End Function



Private Sub LanzaArchivoModificar(CadenaShell As String)
Dim PID As Long


    
    
    PID = Shell(CadenaShell, vbNormalFocus)
    If PID <> 0 Then
        'Esperar a que finalice
        WaitForTerm PID
    End If

End Sub





Private Sub MueveSeparador(X As Single, Y As Single, ElList As Boolean)
    
    If ElList Then
        'Ha soltado en el list
       Y = Me.FrameSeparador.Left + X
    Else
        'En el treee
       Y = Me.TreeView1.Left + X
    End If

    X = Me.Width
    Y = Round((Y / X), 2) * 100
    If Y > 80 Then
        X = 80
    Else
        If Y < 20 Then
            X = 20
        Else
            X = Abs(Y)
        End If
    End If
    
    vUsu.preferencias.Ancho = CInt(X)
    Form_Resize
End Sub

'dopcion= 0.- Multiple normal
'         1.- Carpeta
Private Sub InsertarMultiple(dOpcion As Byte)
Dim i As Integer

  If NodoSeleccionado Is Nothing Then Exit Sub
  If ModoTrabajo <> vbNorm Then
        Mensajes1 (5)
        Exit Sub
  End If
  If NodoSeleccionado.Parent Is Nothing Then
        MsgBox "No se pueden insertar archivos en la carpeta Raiz", vbInformation
        Exit Sub
  End If
        

    
    'Comprobamos si el usuarios tiene permiso de
    Set Car = New Ccarpetas
    
    If Car.Leer(CInt(Mid(NodoSeleccionado.Key, 2)), (ModoTrabajo = 1)) = 0 Then
        
        'OK. VEMOS EL Permiso
        If Car.userprop = vUsu.codusu Or (Car.escriturag And vUsu.Grupo) Then
            
            If MsgBox("La carpeta donde va insertar es: " & vbCrLf & UCase(Text1.Text) & vbCrLf & _
                "¿Desea continuar ?", vbQuestion + vbYesNoCancel) = vbYes Then
            
            
                Set frmNuevoArchivo.Mc = Car
                If dOpcion = 0 Then
                    frmNuevoArchivo.Opcion = 1
                Else
                    frmNuevoArchivo.Opcion = 5
                End If
                frmNuevoArchivo.Carpeta = Text1.Text
                DatosMOdificados = False
                frmNuevoArchivo.Show vbModal
                If DatosMOdificados Then
                    Me.Refresh
                    Screen.MousePointer = vbHourglass
                    If dOpcion = 1 Then
                        'Tenemos que recargar el arbol
                        Cortar11 = NodoSeleccionado.Key
                        
                        'Volvemos a cargar el nodo
                        ListView1.ListItems.Clear
                        CargaArbol
                        
                        'Volvemos a situarlos en el nodo
                        Cortar11 = ""
                        For i = 1 To TreeView1.Nodes.Count
                            If TreeView1.Nodes(i).Key = NodoSeleccionado.Key Then
                                Set NodoSeleccionado = TreeView1.Nodes(i)
                                NodoSeleccionado.Expanded = True
                                Cortar11 = "OK"
                                Exit For
                            End If
                        Next i
                        If Cortar11 = "" Then Set NodoSeleccionado = TreeView1.Nodes(1)
                        
                        Set TreeView1.SelectedItem = NodoSeleccionado
                    End If
                    MostrarArchivos NodoSeleccionado.Key
                    Screen.MousePointer = vbDefault
                End If
            End If
        Else
            MsgBox "No tiene permiso ", vbExclamation
        End If
    Else
        MsgBox "Error leyendo carpeta : " & NodoSeleccionado.Text
    End If
  Screen.MousePointer = vbDefault

End Sub

'Enviaremos al form pregunta los datos
'   0.- Selecion de archivos
'   1.- Propiedades carpeta
'--> ParaImportes :  0- Tamaños
'                    1.- Seleccionados
'                    2.- Carpetas
'                    3.- Carpetas subcarpetas
'
Private Sub VerEspacio(miOpcion As Byte, ParaImportes As Byte)
Dim cad As String
Dim Tamaño As Currency
Dim i As Long
Dim J As Integer
Dim Seleccionados As Integer
Dim Haber As Currency

    If ModoTrabajo <> vbNorm Then
        Mensajes1 16
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
    Set miRSAux = New ADODB.Recordset
    DatosCopiados = Text1.Text & "|"
    
    '##############################################
    '##############################################
    'Tamaño archivos seleccionados
    '##############################################
    '##############################################
    '---------------------------------
    If ParaImportes <= 1 Then
        J = 0
        cad = "-1"
        Seleccionados = 0
        Tamaño = 0
        Haber = 0
        For i = 1 To ListView1.ListItems.Count
           If ListView1.ListItems(i).Selected Then
                cad = cad & "," & Mid(ListView1.ListItems(i).Key, 2)
                J = J + 1
                Seleccionados = Seleccionados + 1
            End If
            If J > 15 Then
                If ParaImportes = 0 Then
                    Tamaño = Tamaño + DameTamañoArchivos2(cad)
                Else
                    ImportesArchivos cad, Tamaño, Haber
                End If
                cad = "-1"
                J = 0
            End If
        Next i
    
        If J > 0 Then
            If ParaImportes = 0 Then
                Tamaño = Tamaño + DameTamañoArchivos2(cad)
            Else
                ImportesArchivos cad, Tamaño, Haber
            End If
        End If
            
    
        '-->>>>>
        If ParaImportes = 0 Then
            
            Tamaño = Round(Tamaño, 2)
            DatosCopiados = DatosCopiados & Seleccionados & "|" & Tamaño & "|"   'SELECCIONADOS
        
        Else
            DatosCopiados = DatosCopiados & Seleccionados & "|" & Tamaño & "|" & Haber & "|"  'SELECCIONADOS
        
        End If
    End If
    
    '##############################################
    '##############################################
    'Tamaño archivos en carpetas
    '##############################################
    '##############################################
    '----------------------------------------------------
    
    If ParaImportes = 0 Then
        
        miRSAux.Open "select sum(tamnyo),count(*) from timagen where codcarpeta=" & Mid(TreeView1.SelectedItem.Key, 2), Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Tamaño = 0
        i = 0
        If Not miRSAux.EOF Then
            Tamaño = DBLet(miRSAux.Fields(0), "N")
            i = DBLet(miRSAux.Fields(1), "N")
        End If
        miRSAux.Close
        
        '-->>>>>
        Tamaño = Round(Tamaño, 2)
        DatosCopiados = DatosCopiados & i & "|" & Tamaño & "|"   'En carpeta
        'Archivos ocultos
        DatosCopiados = DatosCopiados & i - ListView1.ListItems.Count & "|"
        
    
    Else
        If ParaImportes = 2 Then
            'Importes dentro de la carpeta
            
            cad = "SELECT sum(importe1),sum(importe2),count(*) from timagen where codcarpeta=" & Mid(TreeView1.SelectedItem.Key, 2)
            miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                
            If Not miRSAux.EOF Then
                DatosCopiados = DatosCopiados & DBLet(miRSAux.Fields(0), "N")
                DatosCopiados = DatosCopiados & "|" & DBLet(miRSAux.Fields(1), "N")
                DatosCopiados = DatosCopiados & "|" & DBLet(miRSAux.Fields(2), "N") & "|"
                
            Else
                DatosCopiados = DatosCopiados & "0|0|0|"
            End If
            miRSAux.Close

         End If
    End If
    
    '##############################################
    '##############################################
    'Total carpetas dentro de treevie
    '##############################################
    '##############################################
    '------------------------------------------------------------------
    If ParaImportes = 0 Then
        cad = "Select count(*) from carpetas where padre=" & Mid(TreeView1.SelectedItem.Key, 2)
        i = 0
        miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not miRSAux.EOF Then
            i = DBLet(miRSAux.Fields(0), "N")
        End If
        miRSAux.Close
        J = i - TreeView1.SelectedItem.Children
        
        '-->>>>>
        DatosCopiados = DatosCopiados & i & "|" & J & "|"    'Carpetas visibles y ocultas
        'DatosCopiados = DatosCopiados & TreeView1.SelectedItem.Children & "|" & J & "|"    'Carpetas visibles y ocultas
        
        
        frmPregunta.Opcion = 5 + miOpcion
        frmPregunta.Show vbModal
        
    Else
        If ParaImportes = 3 Then
            Cortar11 = CarpetasSubcarpetas(TreeView1.SelectedItem)
            'Contaremos cuantas subcarpetas son
            Seleccionados = 0
            J = 1
            cad = ""
            Tamaño = 0
            Haber = 0
            Do
                i = InStr(J, Cortar11, "|")
                If i > 0 Then
                    Seleccionados = Seleccionados + 1
                    If cad <> "" Then cad = cad & ","
                    cad = cad & Mid(Cortar11, J, i - J)
                    J = i + 1
                    
                    If Seleccionados > 15 Then
                        'Lanzamos el sumador de importes
                        SumaImportesArchivosCarpetas cad, Tamaño, Haber
                        
                        'Reestablecemos
                        Seleccionados = 0
                        cad = ""
                    End If
                End If
            Loop Until i = 0
            
            If Seleccionados > 0 Then
                SumaImportesArchivosCarpetas cad, Tamaño, Haber
            End If
            
            'Ahora unimos en la cadena
            DatosCopiados = DatosCopiados & Tamaño & "|" & Haber & "|"
        End If
    End If
    
    
    
    'Si para importes es mayor k cero= IMPORTES.
    'LLamaremos a pregunta asandole diversos valores
    If ParaImportes > 0 Then
        frmPregunta.Opcion = 7 + ParaImportes
        frmPregunta.Show vbModal
    End If
    
    Screen.MousePointer = vbDefault
    DatosCopiados = ""
    
End Sub


Private Function DameTamañoArchivos2(ByRef cadWHERE As String) As Currency
    DameTamañoArchivos2 = 0
    cadWHERE = "SELECT sum(tamnyo) from timagen where codigo in (" & cadWHERE & ")"
    miRSAux.Open cadWHERE, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then
        If Not IsNull(miRSAux.Fields(0)) Then DameTamañoArchivos2 = miRSAux.Fields(0)
    End If
    miRSAux.Close
End Function

'Dados dos importes les suma , si tene valor lo k coresponde
Private Function ImportesArchivos(ByRef cadWHERE As String, ByRef importe1 As Currency, ByRef importe2 As Currency)
    
    cadWHERE = "SELECT sum(importe1),sum(importe2) from timagen where codigo in (" & cadWHERE & ")"
    miRSAux.Open cadWHERE, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then
        If Not IsNull(miRSAux.Fields(0)) Then
            If Not IsNull(miRSAux.Fields(0)) Then importe1 = importe1 + miRSAux.Fields(0)
        End If
        If Not IsNull(miRSAux.Fields(1)) Then
            If Not IsNull(miRSAux.Fields(0)) Then importe2 = importe2 + miRSAux.Fields(1)
        End If
    End If
    miRSAux.Close
End Function



Private Function SumaImportesArchivosCarpetas(ByRef cadWHERE As String, ByRef importe1 As Currency, ByRef importe2 As Currency)
    
    cadWHERE = "SELECT sum(importe1),sum(importe2) from timagen where codcarpeta in (" & cadWHERE & ")"
    miRSAux.Open cadWHERE, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not miRSAux.EOF Then
        If Not IsNull(miRSAux.Fields(0)) Then
            If Not IsNull(miRSAux.Fields(0)) Then importe1 = importe1 + miRSAux.Fields(0)
        End If
        If Not IsNull(miRSAux.Fields(1)) Then
            If Not IsNull(miRSAux.Fields(0)) Then importe2 = importe2 + miRSAux.Fields(1)
        End If
    End If
    miRSAux.Close
End Function


Private Sub PonerCaption()
    Caption = "ARIDOC. Gestión documental.  Usuario: " & vUsu.Nombre & "  (" & vUsu.codusu & ")"
End Sub




Private Function EliminarCarpeta(ByRef Nod As Node) As Boolean
Dim cad As String
Dim i As Integer


    EliminarCarpeta = False
    
    If ModoTrabajo <> vbNorm Then
        Mensajes1 1
        Exit Function
    End If
    
    
    Set Car = New Ccarpetas
    If Car.Leer(Val(Mid(Nod.Key, 2)), (ModoTrabajo = 1)) = 0 Then
        i = 0
        If vUsu.codusu = 0 Then
                        
                        
            pegar11 = "Va a eliminar la carpeta y todo los datos que contiene:"
            pegar11 = pegar11 & vbCrLf & vbCrLf & Car.Nombre & vbCrLf & vbCrLf & "¿Desesa continuar?"
            
            If MsgBox(pegar11, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
                
            pegar11 = "Este proceso es irreversible. ¿Desea eliminar la carpeta " & Nod.FullPath & " ?"
            If MsgBox(pegar11, vbQuestion + vbYesNo) = vbNo Then Exit Function
            'Es usuario ROOT. Puede eliminar lo k le de la gana
            'Es ROOT.  Elimina lo k le da la gana
            BorrarCarpetaRoot Car
            EliminarCarpeta = True
            Set Car = Nothing
            Exit Function
            
            
        Else
            If Car.userprop = vUsu.codusu Or (Car.escriturag And vUsu.Grupo) Then i = 1
        End If
        
        If i = 1 Then
            
            'Tienen permiso
                
            'Comprobaremos k no existen carpetas no mostradas para este
            'usuario
            If ComprobarCarpetasOcultasDentroCarpeta(Car) = 0 Then
                'Vemos si tienen archivos
                Set miRSAux = New ADODB.Recordset
                cad = "Select count(*) from  timagen where codcarpeta=" & Car.codcarpeta
                miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                cad = "0"
                If Not miRSAux.EOF Then cad = CStr(DBLet(miRSAux.Fields(0), "N"))
                miRSAux.Close
                Set miRSAux = Nothing
            
                If Val(cad) = 0 Then
                    
                    
                    'Comprobamos k no tiene subcarpetas
                    Set miRSAux = New ADODB.Recordset
                    cad = "Select count(*) from  carpetas where padre=" & Car.codcarpeta
                    miRSAux.Open cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    cad = "0"
                    If Not miRSAux.EOF Then cad = CStr(DBLet(miRSAux.Fields(0), "N"))
                    miRSAux.Close
                    Set miRSAux = Nothing
                    
                    If Val(cad) = 0 Then
                        'Hacemos la pregunta
                        cad = "Seguro que desea eliminar la carpeta: " & Nod.Text & "?"
                        If MsgBox(cad, vbQuestion + vbYesNoCancel) = vbYes Then
                            If Car.Eliminar() = 0 Then
                                EliminarCarpeta = True
                                'La quitamos de la cadena de busqueda
                                cad = "|" & Car.codcarpeta & "|"
                                i = InStr(1, CadenaCarpetas, cad)
                                If i > 0 Then CadenaCarpetas = Mid(CadenaCarpetas, 1, i) & Mid(CadenaCarpetas, i + Len(cad))
                            End If
                        End If
                        
                    Else
                        MsgBox "Contiene subcarpetas", vbExclamation
                    End If
                Else
                    MsgBox "La carpeta contiene archivos. Err(" & cad & ")", vbExclamation
                End If
            
            Else
                MsgBox "La carpeta contiene archivos / carpetas ocultos.", vbExclamation
                  
            End If
        Else
            MsgBox "No tiene permiso", vbExclamation
        End If
    End If
    Set Car = Nothing
End Function



Private Function CarpetasSubcarpetas(Nodo As Node) As String
Dim cad As String
Dim Aux As String
Dim Nod As Node


        'Cad = Mid(Nodo.Key, 2)
        CarpetasSubcarpetas = Mid(Nodo.Key, 2) & "|"
        'Debug.Print Nodo.Text
        Aux = ""
        If Nodo.Children > 0 Then
            
            Set Nod = Nodo.Child
            Do
                'Debug.Print Nod.Text
                Aux = Aux & CarpetasSubcarpetas(Nod)
                'If Nod.Key <> Nod.LastSibling.Key Then
                Set Nod = Nod.Next
            Loop Until Nod Is Nothing
            
                
        End If
        CarpetasSubcarpetas = CarpetasSubcarpetas & Aux
        
End Function


Private Sub CargarImagenEncabezado(QueColumna As Integer)
Dim i As Integer

    For i = 1 To ListView1.ColumnHeaders.Count

        'If ListView1.ColumnHeaders(I).Alignment > 0 Then Stop
        If i = QueColumna Then
            ListView1.ColumnHeaders(i).Icon = 3 + Abs(OrderAscendente)
        Else
            ListView1.ColumnHeaders(i).Icon = 0
        End If
        
    Next i
End Sub








''listaimpresion
'Private Sub ImprimirListaArchivos()
'
'    'FALTA###
'End Sub





Private Sub PonerMenuNuevoDocumentos()
    
    
    
    For ListviewSHIFTPresionado = 1 To 6
        Me.mnNuevoN1(ListviewSHIFTPresionado).Visible = False
    Next ListviewSHIFTPresionado
    
    pegar11 = "Select * from extension where aparecemenu=1"
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open pegar11, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ListviewSHIFTPresionado = 0
    While Not miRSAux.EOF
        ListviewSHIFTPresionado = ListviewSHIFTPresionado + 1
        If ListviewSHIFTPresionado <= 6 Then
            With Me.mnNuevoN1(ListviewSHIFTPresionado)
                .Caption = miRSAux!Descripcion
                .Visible = True
                .Tag = miRSAux!codext
            End With
            
        End If
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
    pegar11 = ""
    ListviewSHIFTPresionado = 0
End Sub






Private Function BorrarCarpetaRoot(ByRef Carpe As Ccarpetas) As Boolean

        
    
    Screen.MousePointer = vbHourglass
    pegar11 = "INSERT INTO tmpfich(codusu,codequipo,imagen) VALUES (" & vUsu.codusu & "," & vUsu.PC & ","
 
    If ProcesoBorreCarpeta(Carpe) Then
       ' CargaArbol
        'If TreeView1.Nodes.Count > 1 Then TreeView1.Nodes(1).Expanded = True
    End If
    Set miRSAux = Nothing
    Screen.MousePointer = vbDefault
         

End Function



Private Function ProcesoBorreCarpeta(Carpeta As Ccarpetas) As Boolean
Dim Carpetas As String
Dim SubC As Ccarpetas
Dim i As Integer

    ProcesoBorreCarpeta = False

    'Compruebo si tiene archivos esta carpeta
    Carpetas = "DELETE FROM tmpfich WHERE codusu =" & vUsu.codusu & " AND codequipo = " & vUsu.PC
    Conn.Execute Carpetas
    Carpetas = "Select * from timagen where codcarpeta =" & Carpeta.codcarpeta
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open Carpetas, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRSAux.EOF
        Carpetas = ""
        Conn.Execute pegar11 & miRSAux!codigo & ")"
        
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    
    
    
    'Si hay archivos para borrar...
    If Carpetas = "" Then
            'Ahora lanzaremos movimientos de archivos a piñon
            Set frmMovimientoArchivo.vDestino = Carpeta
            frmMovimientoArchivo.Opcion = 8
            frmMovimientoArchivo.Show vbModal
    End If
    
    
    
    'A lo mejor deberiamos comprobar si se han quedado archivos
    
    
    
    'Comprobamos las subcarpetas
    Carpetas = "Select * from carpetas where padre =" & Carpeta.codcarpeta
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open Carpetas, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Carpetas = ""
    While Not miRSAux.EOF
        Carpetas = Carpetas & miRSAux!codcarpeta & "|"
        miRSAux.MoveNext
    Wend
    miRSAux.Close
    Set miRSAux = Nothing
    
    i = 0
    While Carpetas <> ""
        Set SubC = New Ccarpetas
        i = InStr(1, Carpetas, "|")
        Cortar11 = Mid(Carpetas, 1, i - 1)
        Carpetas = Mid(Carpetas, i + 1)
        i = 0
        If SubC.Leer(CInt(Cortar11), (ModoTrabajo = 1)) = 1 Then
            i = 1
        Else
            If Not ProcesoBorreCarpeta(SubC) Then
                i = 1
            Else
                i = 0
            End If
        End If
        If i = 1 Then
            'Paramos el proceso
            Carpetas = ""
        
        End If
        Set SubC = Nothing
    Wend
    
    
    'Borro la carpeta
    If i = 0 Then
        'Si k podemos borrar la carpeta y seguir
        
        If Carpeta.Eliminar() = 0 Then ProcesoBorreCarpeta = True
        
                
    End If
End Function



Private Sub CompruebaMail()


    Set miRSAux = New ADODB.Recordset
    Cortar11 = "Select * from mailc where destino=" & vUsu.codusu & " AND leido = 0"
    miRSAux.Open Cortar11, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cortar11 = "mail"
    If Not miRSAux.EOF Then
        If Not IsNull(miRSAux.Fields(0)) Then
            If miRSAux.Fields(0) > 0 Then Cortar11 = "tienemail"
        End If
    End If
    miRSAux.Close
    Set miRSAux = Nothing
        
    Toolbar1.Buttons(22).Image = Cortar11
    Cortar11 = ""
        
    
End Sub


Private Sub ComprobarRefrescar()
    Screen.MousePointer = vbHourglass
    If LeerBDRefresco Then
        'Se ha producido alguna modificacion en la carpeta. Hay que volver a cargar el arbol
        Screen.MousePointer = vbHourglass
        Me.Tag = Caption
        Caption = "Leyendo carpetas ....."
        Me.Refresh
        RefrescarCarpetas
        Caption = Me.Tag
        Me.Tag = ""
    End If
    Screen.MousePointer = vbDefault
End Sub
'
Private Sub RefrescarCarpetas()
 Dim cad As String
 Dim i As Integer
 
    If TreeView1.SelectedItem Is Nothing Then
        cad = ""
    Else
        cad = TreeView1.SelectedItem.Key
    End If
    
    CargaArbol
    ListView1.ListItems.Clear
    
    If cad <> "" Then
        For i = 1 To TreeView1.Nodes.Count
            If TreeView1.Nodes(i).Key = cad Then
                cad = ""
                TreeView1.Nodes(i).EnsureVisible
                Exit For
            End If
        Next i
        If cad = "" Then
            'NODO SELECCIONADO
            Set TreeView1.SelectedItem = TreeView1.Nodes(i)
            Set NodoSeleccionado = TreeView1.Nodes(i)
            'mostramos archivos
            MostrarArchivos TreeView1.SelectedItem.Key
        End If
    Else
        If TreeView1.Nodes.Count > 1 Then
            TreeView1.Nodes(2).EnsureVisible
        End If
    End If
End Sub


Private Function LeerBDRefresco() As Boolean
Dim F As Date

    LeerBDRefresco = False
    If ModoTrabajo <> vbNorm Then Exit Function
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select * from actualiza where codigo=1", Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
    F = DateAdd("yyyy", -1, Now)
    If Not miRSAux.EOF Then
        If Not IsNull(miRSAux.Fields(1)) Then F = miRSAux.Fields(1)
    End If
    miRSAux.Close
    Set miRSAux = Nothing
    If F > Now Then LeerBDRefresco = True

End Function

Private Sub CompruebaFechaMYSQL()
Dim F As Date
Dim L As Long
    Set miRSAux = New ADODB.Recordset
    miRSAux.Open "Select curdate(),curtime()", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If miRSAux.EOF Then
        MsgBox "Error leyendo fecha servidor MYSQL", vbExclamation
    Else
        F = CDate(Format(miRSAux.Fields(0), "dd/mm/yyyy") & " " & Format(miRSAux.Fields(1), "hh:mm:ss"))
        L = DateDiff("n", F, Now)
        If Abs(L) > 5 Then Caption = Caption & " COMPROBAR HORA"
    End If
    miRSAux.Close
End Sub


Private Sub GuardarCarpetasAbiertas()
Dim i As Integer
    CarpetasAbiertas = ""
    For i = 2 To TreeView1.Nodes.Count
        If TreeView1.Nodes(i).Expanded Then
            CarpetasAbiertas = CarpetasAbiertas & TreeView1.Nodes(i).Key & "|"
        End If
    Next i
    If CarpetasAbiertas <> "" Then CarpetasAbiertas = "|" & CarpetasAbiertas
End Sub
