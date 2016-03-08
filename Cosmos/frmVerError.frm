VERSION 5.00
Begin VB.Form frmVerError 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCan 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   7680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7095
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmVerError.frx":0000
      Top             =   480
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "frmVerError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public opcion As Byte
    '0.-
    '1.-
    
Dim PrimeraVez As Boolean



Private Sub cmdCan_Click()
 Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaDatos
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
End Sub


Private Sub CargaDatos()
Dim C As String
Dim Rs As ADODB.Recordset
Dim I As Integer
    On Error GoTo ECargaDatos
    Set Rs = New ADODB.Recordset
    If opcion = 0 Then
        Label2.Caption = "Archivos fisicos SIN referencia en BD"
        
    Else
        Label2.Caption = "Archivos en BD SIN referencia en un archivo fisico"
        
    End If
    Label2.Refresh
    C = "Select * from temporal "
    Rs.Open C, ConnAntiguoAridoc, adOpenForwardOnly, adLockOptimistic, adCmdText
    Text1.Text = ""
    I = 0
    While Not Rs.EOF
        I = I + 1
        C = Right("           " & Rs!id, 10)
        Text1.Text = Text1.Text & C & vbCrLf
        
        Rs.MoveNext
        If I > 50 Then
            I = 0
            Me.Refresh
        End If
    Wend
    Rs.Close
    
    Exit Sub
ECargaDatos:
    MsgBox Err.Description

End Sub
