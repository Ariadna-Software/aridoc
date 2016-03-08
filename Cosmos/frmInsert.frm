VERSION 5.00
Begin VB.Form frmInsert 
   Caption         =   "Genera Estrucutura basica"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   9765
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Index           =   1
      Left            =   8520
      TabIndex        =   5
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   7200
      TabIndex        =   4
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmInsert.frx":0000
      Top             =   3840
      Width           =   9495
   End
   Begin VB.TextBox Text1 
      Height          =   3015
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmInsert.frx":0493
      Top             =   360
      Width           =   9495
   End
   Begin VB.Label Label1 
      Caption         =   "INSERTS"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Estructura"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    
    If ProcesaText(0) Then
        ProcesaText 1
    End If

'    ConnNuevoAridoc.Execute Text1(0).Text
    
'    ConnNuevoAridoc.Execute Text1(1).Text
    
End Sub


Private Function ProcesaText(indice As Integer) As Boolean
Dim i As Integer
Dim Cad As String
Dim Fin As Boolean
Dim Inicio As Long
Dim J As Integer
    Inicio = 1
    ProcesaText = False
    Fin = False
    Do
        i = InStr(Inicio, Text1(indice).Text, vbCrLf)
        If i = 0 Then
            Cad = Text1(indice).Text
            Fin = True
        Else
            Cad = Mid(Text1(indice).Text, Inicio, i - Inicio + 1)
            Inicio = i + 2
        End If
        Debug.Print "i: " & i & Cad
        Cad = Trim(Cad)
        'Limpiamos de sltos
        Do
            J = InStr(1, Cad, vbCr)
            If J > 0 Then Cad = Mid(Cad, 1, J - 1) & Mid(Cad, J + 1)
        Loop Until J = 0
    
        Do
            J = InStr(1, Cad, vbLf)
            If J > 0 Then Cad = Mid(Cad, 1, J - 1) & Mid(Cad, J + 1)
        Loop Until J = 0
    
        
        
        
        
        If Cad = vbCr Or Cad = vbCrLf Then Cad = ""
        
        If Cad <> "" Then
            If Not EjecutaSQL(Cad) Then
                If MsgBox("¿Desea continuar?", vbQuestion + vbYesNo) = vbNo Then
                    Exit Function
                End If
            End If
        End If
     Loop Until i = 0
     ProcesaText = True
End Function

Private Function EjecutaSQL(SQL As String) As Boolean
    On Error GoTo Einicio
    EjecutaSQL = False

    
    ConnNuevoAridoc.Execute SQL
    
    EjecutaSQL = True
    Exit Function
Einicio:
    MsgBox "SQL: " & SQL & vbCrLf & vbCrLf & "Error: " & Err.Description, vbExclamation
    Err.Clear
End Function


Private Sub Form_Load()
    If Not LeerFicheros Then Command1(0).Enabled = False
End Sub


Private Function LeerFicheros() As Boolean
Dim N As String
Dim Cad As String
Dim NF As Integer
    On Error GoTo ELeerFicheros
    Text1(0).Text = ""
    Text1(1).Text = ""
    
    
    N = App.Path & "\aridoc.sql"
    If Dir(N, vbArchive) = "" Then
        MsgBox "No se enecuentra el fichero 1: " & N, vbExclamation
        Exit Function
    End If
    
    NF = FreeFile
    Open N For Input As #NF
    While Not EOF(NF)
        Line Input #NF, Cad
        Text1(0).Text = Text1(0).Text & Cad & vbCrLf
    Wend
    Close #NF
    
    
    N = App.Path & "\aridocDAT.sql"
    If Dir(N, vbArchive) = "" Then
        MsgBox "No se enecuentra el fichero 2: " & N, vbExclamation
        Exit Function
    End If
    
    NF = FreeFile
    Open N For Input As #NF
    While Not EOF(NF)
        Line Input #NF, Cad
        Text1(1).Text = Text1(1).Text & Cad & vbCrLf
    Wend
    Close #NF
    
    
    
    
    
    'aridocDAT.sql
    LeerFicheros = True
    Exit Function
ELeerFicheros:
    MsgBox Err.Description
End Function
