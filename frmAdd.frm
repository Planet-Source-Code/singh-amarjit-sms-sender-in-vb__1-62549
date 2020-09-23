VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Number(s)"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   3180
      TabIndex        =   8
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Mobile No.:"
      Height          =   315
      Left            =   180
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtEnd 
      Enabled         =   0   'False
      Height          =   390
      Left            =   1980
      TabIndex        =   4
      Top             =   900
      Width           =   2475
   End
   Begin VB.TextBox txtStart 
      Height          =   390
      Left            =   1980
      TabIndex        =   1
      Top             =   60
      Width           =   2475
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
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
      Left            =   1860
      TabIndex        =   7
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label lblNos 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1980
      TabIndex        =   6
      Top             =   1440
      Width           =   2475
   End
   Begin VB.Label Label3 
      Caption         =   "to"
      Height          =   315
      Left            =   2940
      TabIndex        =   2
      Top             =   540
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Nums:"
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   1500
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile No.: "
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    Me.txtEnd.Enabled = Me.Check1.Value
End Sub

Private Sub cmdCancel_Click()
    Me.Tag = "C"
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Len(Me.txtStart) < 10 Then MsgBox "Please check the number you have entered...": Exit Sub
    If Me.Check1.Value = True Then If Len(Me.txtEnd) < 10 Then MsgBox "Please check the number you have entered...": Exit Sub
    Me.Tag = "O"
    Me.Hide
End Sub

Private Sub Form_Activate()
    Me.lblNos.Caption = ""
    Me.Check1.Value = 0
    Me.txtEnd = ""
    Me.txtStart = ""
End Sub

Private Sub txtEnd_Change()
    doit
End Sub

Sub doit()
    If Len(Me.txtStart) < 10 Or Len(Me.txtEnd) < 10 Then
        Me.lblNos.Caption = ""
    Else
        Me.lblNos.Caption = Abs(Val(Me.txtEnd) - Val(Me.txtStart)) + 1
    End If
End Sub

Private Sub txtStart_Change()
    doit
End Sub
