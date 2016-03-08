Attribute VB_Name = "LibSincronizar"
Option Explicit

Private Const SYNCHRONIZE = &H100000
Private Const INFINITE = &HFFFFFFFF

Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, _
         ByVal dwProcessId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long




'----------------------------------------
'Estos son para el proceso DOS  Task
'---------------------------------------
'
Private Declare Function GetWindow Lib "user32" (ByVal Hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal Hwnd As Long, _
'ByVal lpString As String, ByVal cch As Long) As Long
'Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long
'Const GW_CHILD = 5
'Const GW_HWNDFIRST = 0
'Const GW_HWNDLAST = 1
'Const GW_HWNDNEXT = 2
'Const GW_HWNDPREV = 3
'Const GW_OWNER = 4


'--------------------------------------
'Tercera prueba
'--------------------------------------
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Const GW_HWNDNEXT = 2


Public Sub WaitForTerm(ByVal PID As Long)
    On Error GoTo Gestion_Error

    'Variables locales
    Dim phnd As Long

    phnd = OpenProcess(SYNCHRONIZE, 1, PID)
    If phnd <> 0 Then
        Call WaitForSingleObject(phnd, INFINITE)
        Call CloseHandle(phnd)
    End If
Exit Sub
Gestion_Error:
    Call MsgBox(Err.Number & ": " & Err.Description)
End Sub
 
 





'Public Function LoadTaskList(Hwnd As Long) As String
'    Dim CurrWnd As Long, Length As Long, ListItem As String
'
'    ' Recoge el primer handle del primer programa en proceso
'    ' solo los de primer nivel
'    CurrWnd = GetWindow(Hwnd, GW_HWNDFIRST)
'
'    ' Bucle mientras el handle devuelto por GetWindow sea valido
'    While CurrWnd <> 0
'        ' Devuelve la longitud del nombre de la tarea para crear el buffer
'        Length = GetWindowTextLength(CurrWnd)
'
'        'Recoge el nombre de la aplicación
'        ListItem = Space$(Length + 1)
'        Length = GetWindowText(CurrWnd, ListItem, Length + 1)
'
'        ' Si devuelve un valor valido, se incluye en la lista
'        If Length > 0 Then
'            LoadTaskList = LoadTaskList & CurrWnd & "·"
'        Else

'        End If
'
'        ' Recoge el siguiente indice de la lista de tareas
'        CurrWnd = GetWindow(CurrWnd, GW_HWNDNEXT)
'    Wend
'    LoadTaskList = "·" & LoadTaskList
'End Function



Public Function ExistePId() As String
    Dim test_hwnd As Long, test_pid As Long, test_thread_id As Long
    Dim Existe As Boolean

    Existe = False
    'Find the first window
    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
        'Check if the window isn't a child
        'If GetParent(test_hwnd) = 0 Then
            'Get the window's thread
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)
            ExistePId = ExistePId & test_pid & "·"
        'End If
        'retrieve the next window
        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
    Loop
    ExistePId = "·" & ExistePId
End Function
