VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de MiApl"
   ClientHeight    =   4515
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7440
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3116.333
   ScaleMode       =   0  'User
   ScaleWidth      =   6986.546
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1215
      Left            =   3480
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   1155
      ScaleWidth      =   3555
      TabIndex        =   5
      Top             =   1920
      Width           =   3615
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   5520
      Picture         =   "frmAbout.frx":E3F4
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   5640
      TabIndex        =   0
      Top             =   3960
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   2
      X1              =   563.431
      X2              =   6648.487
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   3
      X1              =   450.745
      X2              =   5761.083
      Y1              =   1242.392
      Y2              =   1242.392
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ariadna Software S.L."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "C/ Franco Tormo N.3 Bajo Izq. 46007 Valencia"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel: 96 3580547"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   7
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "administracion@ariadnasoftware.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   6
      Top             =   3240
      Width           =   3735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   450.745
      X2              =   6648.487
      Y1              =   2567.61
      Y2              =   2567.61
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Gesti�n documental"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   480
      TabIndex        =   2
      Top             =   1080
      Width           =   2085
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "T�tulo de"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   960
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   4605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   338.059
      X2              =   5648.397
      Y1              =   2567.61
      Y2              =   2567.61
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Versi�n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   480
      TabIndex        =   4
      Top             =   1440
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'' Opciones de seguridad de clave del Registro...
'Const READ_CONTROL = &H20000
'Const KEY_QUERY_VALUE = &H1
'Const KEY_SET_VALUE = &H2
'Const KEY_CREATE_SUB_KEY = &H4
'Const KEY_ENUMERATE_SUB_KEYS = &H8
'Const KEY_NOTIFY = &H10
'Const KEY_CREATE_LINK = &H20
'Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
'                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
'                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
'
'' Tipos ROOT de clave del Registro...
'Const HKEY_LOCAL_MACHINE = &H80000002
'Const ERROR_SUCCESS = 0
'Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
'Const REG_DWORD = 4                      ' N�mero de 32 bits
'
'Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
'Const gREGVALSYSINFOLOC = "MSINFO"
'Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
'Const gREGVALSYSINFO = "PATH"
'
'Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
'Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
'Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Dim PulsadoCombinacion As Boolean


Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    If ShiftDown And CtrlDown Then PulsadoCombinacion = True
End Sub

Private Sub cmdOK_KeyUp(KeyCode As Integer, Shift As Integer)
PulsadoCombinacion = False
End Sub

'Private Sub cmdSysInfo_Click()
'  Call StartSysInfo
'End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub cmdSysInfo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    If ShiftDown And CtrlDown Then PulsadoCombinacion = True
End Sub

Private Sub cmdSysInfo_KeyUp(KeyCode As Integer, Shift As Integer)
PulsadoCombinacion = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim ShiftDown, AltDown, CtrlDown
    ShiftDown = (Shift And vbShiftMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    If ShiftDown And CtrlDown Then PulsadoCombinacion = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
PulsadoCombinacion = False
End Sub

Private Sub Form_Load()
    Me.Caption = "Acerca de " & App.Title
    lblVersion.Caption = "Versi�n " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    PulsadoCombinacion = False
    'Carga imagen fondo
    CargaImagen
End Sub

Private Sub CargaImagen()
    On Error Resume Next
    Me.Picture = LoadPicture(App.Path & "\entrada.dat")
    If Err.Number <> 0 Then
        lblTitle.Visible = True
        Me.lblVersion.ForeColor = vbBlack
        picIcon.Visible = True
        Err.Clear
    Else
        picIcon.Visible = False
        lblTitle.Visible = False
        Me.lblVersion.ForeColor = vbWhite
    End If
    
End Sub

'Public Sub StartSysInfo()
'    On Error GoTo SysInfoErr
'
'    Dim Rc As Long
'    Dim SysInfoPath As String
'
'    ' Intentar obtener ruta de acceso y nombre del programa de Info. del sistema a partir del Registro...
'    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
'    ' Intentar obtener s�lo ruta del programa de Info. del sistema a partir del Registro...
'    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
'        ' Validar la existencia de versi�n conocida de 32 bits del archivo
'        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
'            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
'
'        ' Error: no se puede encontrar el archivo...
'        Else
'            GoTo SysInfoErr
'        End If
'    ' Error: no se puede encontrar la entrada del Registro...
'    Else
'        GoTo SysInfoErr
'    End If
'
'    Call Shell(SysInfoPath, vbNormalFocus)
'
'    Exit Sub
'SysInfoErr:
'    MsgBox "La informaci�n del sistema no est� disponible en este momento", vbInformation + vbOKOnly
'End Sub
'
'Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'    Dim I As Long                                           ' Contador de bucle
'    Dim Rc As Long                                          ' C�digo de retorno
'    Dim hKey As Long                                        ' Controlador de una clave de Registro abierta
'    Dim hDepth As Long                                      '
'    Dim KeyValType As Long                                  ' Tipo de datos de una clave de Registro
'    Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave de Registro
'    Dim KeyValSize As Long                                  ' Tama�o de variable de clave de Registro
'    '------------------------------------------------------------
'    ' Abrir clave de registro bajo KeyRoot {HKEY_LOCAL_MACHINE...}
'    '------------------------------------------------------------
'    Rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir clave de Registro
'
'    If (Rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error de controlador...
'
'    tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
'    KeyValSize = 1024                                       ' Marcar tama�o de variable
'
'    '------------------------------------------------------------
'    ' Obtener valor de clave de Registro...
'    '------------------------------------------------------------
'    Rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
'                         KeyValType, tmpVal, KeyValSize)    ' Obtener o crear valor de clave
'
'    If (Rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
'
'    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agregar cadena terminada en valor nulo...
'        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Encontrado valor nulo, se va a quitar de la cadena
'    Else                                                    ' En WinNT las cadenas no terminan en valor nulo...
'        tmpVal = Left(tmpVal, KeyValSize)                   ' No se ha encontrado valor nulo, s�lo se va a extraer la cadena
'    End If
'    '------------------------------------------------------------
'    ' Determinar tipo de valor de clave para conversi�n...
'    '------------------------------------------------------------
'    Select Case KeyValType                                  ' Buscar tipos de datos...
'    Case REG_SZ                                             ' Tipo de datos String de clave de Registro
'        KeyVal = tmpVal                                     ' Copiar valor de cadena
'    Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
'        For I = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
'            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Generar valor car�cter a car�cter
'        Next
'        KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a cadena
'    End Select
'
'    GetKeyValue = True                                      ' Se ha devuelto correctamente
'    Rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
'    Exit Function                                           ' Salir
'
'GetKeyError:      ' Borrar despu�s de que se produzca un error...
'    KeyVal = ""                                             ' Establecer valor a cadena vac�a
'    GetKeyValue = False                                     ' Fallo de retorno
'    Rc = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
'End Function


Private Sub lblDescription_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim m

If Button = 2 Then

'Vamos a hacer un Huevo de pascua

If PulsadoCombinacion Then
    m = vbCrLf & vbCrLf & vbCrLf & vbCrLf
    m = m & "Hola querido Usuario:  " & vbCrLf
    m = m & "________________________" & vbCrLf & vbCrLf
    m = m & " Has encontrado  la combinaci�n de teclas y rat�n" & vbCrLf
    m = m & " para que te aparezca esta pantalla, tambi�n" & vbCrLf
    m = m & " llamada Huevo de Pascua.    " & vbCrLf
    m = m & "          Era facil �no?." & vbCrLf & vbCrLf
    m = m & "                       Aaaaadeu"
    m = m & "           � Ariadna Software.                   " & vbCrLf & vbCrLf
    m = m & "                           "
    m = m & "                                   DABIZ" & vbCrLf
    
    MsgBox m, vbExclamation

    PulsadoCombinacion = False
    End If
End If
End Sub

