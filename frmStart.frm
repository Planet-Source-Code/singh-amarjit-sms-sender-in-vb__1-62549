VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SMS SENDER"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10965
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picInfo1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   10695
      TabIndex        =   7
      Top             =   2220
      Width           =   10755
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   29
         Top             =   2820
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "S&ave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   28
         Top             =   2340
         Width           =   975
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5520
         TabIndex        =   27
         Top             =   4020
         Width           =   975
      End
      Begin VB.TextBox txtMsg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1605
         Left            =   6600
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   3660
         Width           =   4035
      End
      Begin VB.TextBox txtTotalNums 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6600
         TabIndex        =   23
         Top             =   3300
         Width           =   2895
      End
      Begin VB.ListBox lstActNos 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   120
         TabIndex        =   22
         Top             =   3420
         Width           =   2895
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   21
         Top             =   1860
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.ListBox lstNos 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   6600
         TabIndex        =   18
         Top             =   1440
         Width           =   2895
      End
      Begin VB.CheckBox chkDebug 
         Caption         =   "Debug"
         Height          =   285
         Left            =   60
         TabIndex        =   17
         Top             =   120
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.PictureBox picDebug 
         Height          =   2775
         Left            =   60
         ScaleHeight     =   2715
         ScaleWidth      =   4995
         TabIndex        =   14
         Top             =   480
         Width           =   5055
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2205
            Left            =   60
            TabIndex        =   16
            Top             =   60
            Width           =   4875
         End
         Begin VB.TextBox txtOut 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            TabIndex        =   15
            Top             =   2280
            Width           =   4875
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total No. of Mobile Numbers: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   4500
         TabIndex        =   26
         Top             =   3360
         Width           =   2115
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Message: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5700
         TabIndex        =   25
         Top             =   3660
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile Number:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   5280
         TabIndex        =   19
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturer: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   6900
         TabIndex        =   13
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label lblManufacturer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8280
         TabIndex        =   12
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label lblDevType 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8280
         TabIndex        =   11
         Top             =   60
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Device Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   6900
         TabIndex        =   10
         Top             =   120
         Width           =   960
      End
      Begin VB.Label lblProvider 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8280
         TabIndex        =   9
         Top             =   660
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Provider: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   6900
         TabIndex        =   8
         Top             =   720
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "&Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4680
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox picBott 
      Align           =   2  'Align Bottom
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   10905
      TabIndex        =   3
      Top             =   7755
      Width           =   10965
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   60
         Width           =   8655
      End
   End
   Begin VB.ComboBox cmbPorts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   8160
      Top             =   660
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10905
      TabIndex        =   0
      Top             =   0
      Width           =   10965
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Port: "
      Height          =   285
      Left            =   2940
      TabIndex        =   5
      Top             =   780
      Width           =   690
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mOK
Dim mErr
Dim mResult
Dim doit As Boolean

Private Sub fresh_Click1()
    List1.AddItem "Done..."
End Sub


Private Sub chkDebug_Click()
    Me.picDebug.Visible = Me.chkDebug.Value
End Sub

Private Sub cmdAdd_Click()
    frmAdd.Show 1
    Dim SNo
    If frmAdd.Tag = "O" Then
        'User Entered some Mobile number(s)
        'First check is it a list or a single number
        With frmAdd
            If .Check1.Value = 1 Then
                'YES
                st = .txtStart
                en = .txtEnd
                Me.lstNos.AddItem "" & st & " - " & en
                For i = st To en
                    Me.lstActNos.AddItem "" & i
                Next
            Else
                'NO
                SNo = .txtStart
                Me.lstNos.AddItem SNo
                Me.lstActNos.AddItem SNo
            End If
        End With
    End If
    Unload frmAdd
    Me.txtTotalNums = Me.lstActNos.ListCount
End Sub

Private Sub cmdConnect_Click()
    On Error GoTo p1
    If Me.cmdConnect.Caption = "&Connect" Then
        If Len(Me.cmbPorts.Text) = 0 Then MsgBox "Please select a valid port please...": Me.cmbPorts.SetFocus: Exit Sub
        Me.cmdConnect.Caption = "&Disconnect"
        setStatus "Connecting..."
        MSComm1.RThreshold = 1
        MSComm1.InputLen = 0
        MSComm1.Settings = "19200,N,8,1"
        MSComm1.DTREnable = True
        MSComm1.InBufferSize = 32
        MSComm1.OutBufferSize = 0
        MSComm1.CommPort = Me.cmbPorts.Text
        MSComm1.RTSEnable = True
        DoEvents
        MSComm1.PortOpen = True
        DoEvents
        setStatus "Connected to Port: " & Me.cmbPorts.Text
        DoEvents
        Me.picInfo1.Enabled = True
        
        setStatus "Getting Status...."
        
        getMobileInfo
        setStatus "Done...Connected to COM" & Me.cmbPorts.Text
    ElseIf Me.cmdConnect.Caption = "&Disconnect" Then
        Me.cmdConnect.Caption = "&Connect"
        Me.MSComm1.PortOpen = False
        Me.picInfo1.Enabled = False
    End If
    Exit Sub
p1:
    MsgBox "Sorry, Cannot connect, Please check Port and Connection...and try again" & vbCrLf & Err.Description
    End
End Sub


Function getProvider(ByVal s)
    s1 = ""
    If Len(s) > 0 Then
        p = InStr(s, Chr(34))
        s1 = Mid(s, p + 1)
        p1 = InStr(s1, Chr(34))
        If p1 > 0 Then
            s1 = Mid(s1, 1, p1 - 1)
        End If
    End If
    getProvider = s1
End Function


Function getManufacturer(ByVal s)
    s1 = ""
    If Len(s) > 0 Then
        s1 = Mid(s, 11)
        p = InStr(s1, Chr(13))
        If p = 0 Then p = InStr(s1, Chr(10))
        If p > 0 Then
            s1 = Mid(s1, 1, p - 1)
        End If
    End If
    getManufacturer = s1
End Function

Function getDevType(ByVal s)
    s1 = ""
    If Len(s) > 0 Then
        s1 = Mid(s, 7)
        p = InStr(s1, Chr(10))
        If p = 0 Then p = InStr(s1, Chr(13))
        If p > 0 Then
            s1 = Mid(s1, 1, p)
        End If
    End If
    getDevType = s1
End Function

Private Sub cmdLoad_Click()
    Open App.Path & "\List.txt" For Input As 1
    Me.lstActNos.Clear
    While Not EOF(1)
        Line Input #1, s
        Me.lstActNos.AddItem "" & s
        DoEvents
    Wend
    Close #1
    MsgBox "File List.txt Loaded..."
End Sub

Private Sub cmdRefresh_Click()
    setStatus "Getting Ports list..."
    ListComPorts
    setStatus ""
End Sub

Sub getMobileInfo()
    Dim st As Boolean
    Me.txtOut = ""
    st = sendIt("AT", "OK", "ERROR")
    If st = True Then
        'Everything OK
    Else
        'Not Connected
        MsgBox "GSM Modem Not found"
        End
    End If
    
    Me.txtOut = ""
    st = sendIt("ATI", "OK", "ERROR")
    If st = True Then
        Me.lblDevType.Caption = getDevType(Me.txtOut)
    Else
        Me.lblDevType.Caption = ""
    End If
    
    Me.txtOut = ""
    
    st = sendIt("AT+CGMI", "OK", "ERROR")
    If st = True Then
        Me.lblManufacturer.Caption = getManufacturer(Me.txtOut)
    Else
        Me.lblManufacturer.Caption = ""
    End If
    Me.txtOut = ""

    st = sendIt("AT+COPS?", "OK", "ERROR")
    If st = True Then
        Me.lblProvider.Caption = getProvider(Me.txtOut)
    Else
        Me.lblProvider.Caption = ""
    End If
    'Me.txtOut = ""

End Sub

Private Sub ListComPorts()
    Dim i As Integer
    
    Me.cmbPorts.Clear
    setStatus "Getting Available Com Ports..."
    For i = 1 To 16
        If COMAvailable(i) Then
            Me.cmbPorts.AddItem i
            setStatus "Com " & i & " found"
        End If
    Next
    Me.cmbPorts.ListIndex = 0
End Sub

Sub setStatus(ByVal s)
    Me.lblStatus.Caption = "" & s
End Sub

Function sendIt(ByVal s, ByVal ok, ByVal eror, Optional ByVal TOut = 2) As Boolean
    mOK = ok
    mErr = eror
    Me.List1.AddItem "sending..." & s
    MSComm1.Output = s & Chr(13)
    Dim p As Double, p1 As Double, p2 As Double
    p = 0.0001 * TOut
    p2 = 0#
    doit = False
    sdata = ""
    Dim dt1 As Date, dt2 As Date
    dt1 = Now
    s1 = ""
    While doit = False
        dt2 = Now
        p1 = (dt2 - dt1)
        'p2 = p1 * 10000#
        If p1 >= p Then
            doit = True
            sendIt = False
            Exit Function
        'ElseIf p2 Mod 10 = 0 Then
        '    s1 = s1 & "-"
        '    If Len(s1) > 20 Then s1 = "-"
        '    setStatus s1
        '    DoEvents
        End If
        DoEvents
    Wend
    sendIt = True
End Function

Private Sub cmdRemove_Click()
    p = Me.lstNos.ListIndex
    If Len(Me.lstNos.Text) > 0 Then Me.lstNos.RemoveItem Me.lstNos.ListIndex
    If p > 0 Then If Me.lstNos.ListCount > 0 Then Me.lstNos.ListIndex = p - 1
    Me.txtTotalNums = Me.lstActNos.ListCount
    'updateActNos
End Sub

Function SendSMS(ByVal MobNo, ByVal Msg)
    Dim what As Boolean
    what = sendIt("AT+CMGS=" & Chr(34) & MobNo & Chr(34), ">", "ERROR")
    If what = True Then
        what = sendIt(Msg & Chr(26), "OK", "ERROR")
    End If
    SendSMS = what
End Function

Private Sub cmdSave_Click()
    Open App.Path & "\List.txt" For Output As 1
    For i = 0 To Me.lstActNos.ListCount - 1
        Print #1, Me.lstActNos.List(i)
        DoEvents
    Next
    Close #1
    MsgBox "File saved as List.txt"
End Sub

Private Sub cmdSend_Click()
    Dim what As Boolean
    Dim s
    what = sendIt("AT+CMGF=1", "OK", "ERROR")
    If what = True Then
        y = Me.lstActNos.ListCount
        For i = 0 To y - 1
            s = Me.lstActNos.List(i)
            setStatus "Sending Message " & (i + 1) & " of " & y & " to " & s
            what = SendSMS(s, Me.txtMsg)
            If what = False Then GoTo p
            setStatus "Sent..."
            DoEvents
        Next
    End If
    setStatus "Done..."
    Exit Sub
p:
    MsgBox "Some Error Occured..." & Err.Description
End Sub


Private Sub MSComm1_OnComm()
    Dim sdata1
    Me.List1.AddItem "In OnComm"
    If MSComm1.CommEvent = comEvReceive Then
        sdata1 = MSComm1.Input
        sdata = sdata & sdata1
        If InStr(sdata, mOK) > 0 Then
            doit = True
        ElseIf InStr(sdata, mErr) > 0 Then
            doit = True
        ElseIf InStr(sdata, ">") > 0 Then
            doit = True
        End If
        mResult = sdata
        If Len(sdata) > 0 Then Me.List1.AddItem "--> " & sdata
        txtOut = txtOut & sdata
    End If
End Sub

