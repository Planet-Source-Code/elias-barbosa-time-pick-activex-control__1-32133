VERSION 5.00
Object = "*\A..\..\..\..\..\VISUAL~1\SAMPLE~1\_PROGR~1\MYCLOC~1\TIMEPI~1\vbpTimePick.vbp"
Begin VB.Form frmMain 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Pick Example"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin vbpTimePick.ctlTimePick ctlTimePick1 
      Height          =   480
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   847
      HourValue       =   "12"
      MinuteValue     =   "00"
      AMPMValue       =   "AM"
      BackColor       =   -2147483647
      TimeColor       =   255
      TitleColor      =   -2147483624
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Border Style"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "American/British"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Enabled"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "You can, also, use this Control with the British (Military) time style, which divides a day into 24 hours."
      ForeColor       =   &H80000018&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   5775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0442
      ForeColor       =   &H80000018&
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   5775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMain.frx":0530
      ForeColor       =   &H80000018&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
    If (ctlTimePick1.ShowFrame = Border) Then
        ctlTimePick1.ShowFrame = Flat
        
    Else
        ctlTimePick1.ShowFrame = Border
        
    End If
    
End Sub

Private Sub Command4_Click()
    If (ctlTimePick1.Enabled) Then
        ctlTimePick1.Enabled = False
        
    Else
        ctlTimePick1.Enabled = True
        
    End If
    
End Sub

Private Sub Command1_Click()
    If (ctlTimePick1.TimeStyle = American) Then
        ctlTimePick1.TimeStyle = British
        
    Else
        ctlTimePick1.TimeStyle = American
        
    End If
    
End Sub

Private Sub ctlTimePick1_Change()
    Text1.Text = ctlTimePick1.Value
    
End Sub

Private Sub ctlTimePick1_Finished()
    Text1.SetFocus
    
End Sub

Private Sub Form_Load()
    Text1.Text = ctlTimePick1.Value
    
End Sub
