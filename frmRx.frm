VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmRx 
   Caption         =   "Rx From Basic Stamp"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8580
   Icon            =   "frmRx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMToSend 
      Height          =   735
      Left            =   0
      TabIndex        =   7
      Top             =   4320
      Width           =   5415
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   2640
      TabIndex        =   6
      Top             =   2040
      Width           =   4935
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox txtMsg 
      Height          =   2175
      Left            =   0
      TabIndex        =   4
      Text            =   "Demo Msg"
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtMNo 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Text            =   "9876091077"
      Top             =   1440
      Width           =   2535
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7920
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   4
      DTREnable       =   -1  'True
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current RCTime:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1410
   End
End
Attribute VB_Name = "frmRx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mOK
Dim mErr
Dim mResult
Dim doit As Boolean
Dim sdata As String

Private Sub cmdSend_Click()
    Dim n
    ' Setup PictureBox for Scale
    List1.Clear
    List1.AddItem "Starting..."
    ' Fire Rx Event Every Byte
    MSComm1.RThreshold = 1
    ' When Inputting Data, Input All Bytes
    MSComm1.InputLen = 0
    ' 19200 Baud, No Parity, 8 Data Bits, 1 Stop Bit
    MSComm1.Settings = "19200,N,8,1"
    ' Make sure DTR line is low to prevent Stamp reset
    MSComm1.DTREnable = True
    MSComm1.InBufferSize = 32
    MSComm1.OutBufferSize = 0
    ' Open COM1
    MSComm1.CommPort = 5
    MSComm1.RTSEnable = True
    'Me.MSComm1.Handshaking = 2 - comRTS
    MSComm1.PortOpen = True
    List1.AddItem "Port Opened"

    Dim what As Boolean
    
    what = sendIt("AT+CMGF=1", "OK", "ERROR")
    If what = True Then
        what = sendIt("AT+CMGS=" & Chr(34) & Me.txtMNo & Chr(34), ">", "ERROR")
        If what = True Then
            n = Now
            Me.txtMToSend = Me.txtMsg & n
            'MSComm1.Output = Me.txtMsg & n & Chr(26) & Chr(13)
            what = sendIt(Me.txtMToSend & Chr(26), "OK", "ERROR")
        End If
    End If
    Me.MSComm1.PortOpen = False
    List1.AddItem "Done..."
End Sub

Function sendIt(ByVal s, ByVal ok, ByVal eror, Optional ByVal TOut = 5) As Boolean
    mOK = ok
    mErr = eror
    List1.AddItem "Sending.." & s
    MSComm1.Output = s & Chr(13)
    Dim p
    p = 0.0001 * TOut
    doit = False
    sdata = ""
    Dim dt1 As Date, dt2 As Date
    dt1 = Now
    Dim p1
    While doit = False
        dt2 = Now
        p1 = (dt2 - dt1)
        If p1 >= p Then
            List1.AddItem "Timeout..."
            doit = True
            sendIt = False
            Exit Function
        End If
        DoEvents
    Wend
    sendIt = True
End Function


Sub wait()
    Dim p
    p = 0.0005
    doit = False
    Me.List1.AddItem "Waiting..."
    sdata = ""
    Dim dt1 As Date, dt2 As Date
    dt1 = Now
    Dim p1
    While doit = False
        dt2 = Now
        p1 = (dt2 - dt1)
        If p1 >= p Then
            List1.AddItem "Timeout..."
            doit = True
        End If
        DoEvents
    Wend
End Sub

Private Sub MSComm1_OnComm()
    List1.AddItem "In OnComm"
    Dim sdata1
    If MSComm1.CommEvent = comEvReceive Then
        sdata1 = MSComm1.Input
        sdata = sdata & sdata1
        If InStr(sdata, mOK) > 0 Then
            doit = True
            mResult = "OK"
            List1.AddItem "--> " & sdata
        ElseIf InStr(sdata, mErr) > 0 Then
            doit = True
            List1.AddItem "Err--->" & sdata
            mResult = "ERR"
        ElseIf InStr(sdata, ">") > 0 Then
            doit = True
            List1.AddItem ">>---> " & sdata
            mResult = sdata
        Else
            List1.AddItem "?---> " & sdata
            mResult = sdata
        End If
    End If
End Sub
