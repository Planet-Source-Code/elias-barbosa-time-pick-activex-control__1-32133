VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ctlTimePick 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2235
   PropertyPages   =   "ctlTimePick.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   2235
   ToolboxBitmap   =   "ctlTimePick.ctx":0026
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   550
      Left            =   0
      TabIndex        =   3
      Top             =   -30
      Width           =   2285
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1620
         TabIndex        =   2
         Text            =   "AM"
         Top             =   155
         Width           =   615
      End
      Begin VB.Frame Frame5 
         Caption         =   "Hours"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   505
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   740
         Begin VB.TextBox txtHours 
            Height          =   285
            Left            =   60
            MaxLength       =   2
            TabIndex        =   0
            Text            =   "12"
            Top             =   160
            Width           =   375
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   435
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   160
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   5
            OrigLeft        =   435
            OrigTop         =   240
            OrigRight       =   675
            OrigBottom      =   525
            Max             =   12
            Min             =   1
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Minutes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   505
         Left            =   850
         TabIndex        =   4
         Top             =   0
         Width           =   740
         Begin VB.TextBox txtMinutes 
            Height          =   285
            Left            =   60
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "00"
            Top             =   160
            Width           =   375
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   285
            Left            =   435
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   160
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            Value           =   30
            Max             =   59
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   765
         TabIndex        =   8
         Top             =   150
         Width           =   135
      End
   End
End
Attribute VB_Name = "ctlTimePick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Dim txtFinalTime As String
Dim intStopUpDown As Boolean

Dim ChangeCount As Integer
Dim StopTextChange As Integer

Dim ChangeCount2 As Integer
Dim StopTextChange2 As Integer

Dim ChangeCount3 As Boolean

Dim intKeyUp As Boolean
Dim intKeyDown As Boolean
Dim intTimeSelStart As Integer
Dim intTimeSelLength As Integer

'Default Property Values:
Const m_def_tatata = 0
Const m_def_TimeStyle = 0
Const m_def_BackColor = vbButtonFace
Const m_def_TimeColor = &H80000008
Const m_def_TitleColor = &H80000008
Const m_def_Value = "12:00 AM"
Const m_def_Enabled = True

'Property Variables:
Dim m_tatata As Variant
Dim m_TimeStyle As AmericBritish
Dim m_BackColor As OLE_COLOR
Dim m_TimeColor As OLE_COLOR
Dim m_TitleColor As OLE_COLOR
Dim m_Value As String
Dim m_Enabled As Boolean

'Event Declarations:
Event Change()
Attribute Change.VB_Description = "Occurs when the time on the TimePick Control have changed."
Event Finished()

Public Enum AmericBritish
    American = 0
    British = 1
End Enum

Public Enum BorderStyle
   Flat = 0
   Border = 1
End Enum

Public Enum AMPM_Type
   AM = 0
   PM = 1
End Enum

Private Sub UserControl_Initialize()
    Combo1.AddItem "AM", 0
    Combo1.AddItem "PM", 1
    
    
End Sub

Private Sub UserControl_Resize()
    With UserControl
        .Height = 480
        If (m_TimeStyle = American) Then
            .Width = 2230
        Else
            .Width = 1590
        End If
    End With
    
End Sub

'===========================================
'This Subs were created to costumize
'the time GUI for this Control.
'===========================================

'**********************************
'******** Seting the hour *********
'**********************************
Private Sub txtHours_KeyDown(KeyCode As Integer, Shift As Integer)
    'Prevent any Shift key combination
    'from going through...
    If (Shift = vbShiftMask) Then
        txtHours.Locked = True
        Exit Sub
    End If
    
    'Find out whether the user has
    'typed a number or not.
    Select Case KeyCode
        Case vbKeyLButton, vbKeyRButton, vbKeyBack, _
             vbKeyDelete, vbKeyTab, vbKeyControl, _
             vbKey0, vbKey1, vbKey2, vbKey3, _
             vbKey4, vbKey5, vbKey6, vbKey7, _
             vbKey8, vbKey9, _
             vbKeyNumlock, vbKeyNumpad0, vbKeyNumpad1, _
             vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, _
             vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, _
             vbKeyNumpad8, vbKeyNumpad9
            'Allow key to go through.
            
        Case vbKeyUp
            intKeyUp = True
            intTimeSelStart = txtHours.SelStart
            intTimeSelLength = txtHours.SelLength
            
        Case vbKeyDown
            intKeyDown = True
            intTimeSelStart = txtHours.SelStart
            intTimeSelLength = txtHours.SelLength
            
        Case vbKeyReturn, vbKeySeparator
            txtMinutes.SetFocus
            
        Case Else
            txtHours.Locked = True
            
    End Select
    
End Sub

Private Sub txtHours_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intValue As Integer
    Dim intTempTotalValue As String
    Dim intTotalValue As Integer
    
    txtHours.Locked = False
    
    '=================================
    '===== When the user presses =====
    '===== the Up Key Arrow.     =====
    '=================================
    If (intKeyUp) Then
        '==================================
        '===== If the first character =====
        '===== is selected.           =====
        '==================================
        If ((intTimeSelStart = 0) _
        Or (intTimeSelLength = 2)) _
        And (Len(txtHours.Text) = 2) Then
            intValue = CInt(Right(txtHours.Text, 1))
            
            If (m_TimeStyle = American) _
            And (intValue < 3) _
            And (intValue <> 0) Then
                 If (Left(txtHours.Text, 1) = "0") Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = 1 & Right(txtHours.Text, 1)
                 Else
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = 0 & Right(txtHours.Text, 1)
                 End If
                 
            ElseIf (m_TimeStyle = British) _
            And (intValue < 5) Then
                 If (Left(txtHours.Text, 1) = "0") Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = 1 & Right(txtHours.Text, 1)
                    
                ElseIf (Left(txtHours.Text, 1) = "1") Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = 2 & Right(txtHours.Text, 1)
                    
                 ElseIf (Left(txtHours.Text, 1) = "2") Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = 0 & Right(txtHours.Text, 1)
                    
                 End If
            End If
            
            txtHours.SelStart = 0
            txtHours.SelLength = 1
            
        '===================================
        '===== If the second character =====
        '===== is selected.            =====
        '===================================
        ElseIf ((intTimeSelStart = 1) _
        Or (intTimeSelStart = 2)) _
        And (Len(txtHours.Text) = 2) Then
            
            intValue = CInt(txtHours.Text) + 1
            
            If (m_TimeStyle = American) Then
                
                If (intValue < 13) Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = intValue
                    
                Else
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = "01"
                    
                End If
                
            ElseIf (m_TimeStyle = British) Then
                
                If (intValue < 24) Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = intValue
                    
                Else
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = "00"
                    
                End If
                
            End If
            txtHours.SelStart = 1
            txtHours.SelLength = 1
            
        End If
        
        intKeyUp = False
    End If
    
    
    '=================================
    '===== When the user presses =====
    '===== the Down Key Arrow.   =====
    '=================================
    If (intKeyDown) Then
        '==================================
        '===== If the first character =====
        '===== is selected.           =====
        '==================================
        If ((intTimeSelStart = 0) _
        Or (intTimeSelLength = 2)) _
        And (Len(txtHours.Text) = 2) Then
            intValue = CInt(Right(txtHours.Text, 1))
            
            If (m_TimeStyle = American) _
            And (intValue < 3) _
            And (intValue <> 0) Then
                
                If (Left(txtHours.Text, 1) = "0") Then
                   'Change immediately...
                   ChangeCount = 1
                   txtHours.Text = 1 & Right(txtHours.Text, 1)
                Else
                   'Change immediately...
                   ChangeCount = 1
                   txtHours.Text = 0 & Right(txtHours.Text, 1)
                End If
                 
            ElseIf (m_TimeStyle = British) _
            And (intValue < 5) Then
                 If (Left(txtHours.Text, 1) = "0") Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = 2 & Right(txtHours.Text, 1)
                    
                ElseIf (Left(txtHours.Text, 1) = "2") Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = 1 & Right(txtHours.Text, 1)
                    
                 ElseIf (Left(txtHours.Text, 1) = "1") Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = 0 & Right(txtHours.Text, 1)
                    
                 End If
                 
            End If
            txtHours.SelStart = 0
            txtHours.SelLength = 1
            
        '===================================
        '===== If the second character =====
        '===== is selected.            =====
        '===================================
        ElseIf ((intTimeSelStart = 1) _
        Or (intTimeSelStart = 2)) _
        And (Len(txtHours.Text) = 2) Then
            
            intValue = CInt(txtHours.Text) - 1
            
            If (m_TimeStyle = American) Then
                If (intValue > 0) Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = intValue
                    
                Else
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = "12"
                    
                End If
                
            ElseIf (m_TimeStyle = British) Then
                
                If (intValue >= 0) Then
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = intValue
                    
                Else
                    'Change immediately...
                    ChangeCount = 1
                    txtHours.Text = "23"
                    
                End If
                
            End If
            txtHours.SelStart = 1
            txtHours.SelLength = 1
            
        End If
        intKeyDown = False
    End If
    
    
End Sub

Private Sub txtHours_Change()
    If (StopTextChange = 0) Then
        If (ChangeCount = 1) Then
            'If Time Style is American, show time
            'separated in to sets of 12 hours and
            'show the AM/PM information.
            If (m_TimeStyle = American) Then
                If (Val(txtHours.Text) <> UpDown1.Value) _
                And (Val(txtHours.Text) <= 12) _
                And (Val(txtHours.Text) > 0) Then
                    UpDown1.Value = txtHours.Text
                    
                End If
                
                'Transform from the British (Military)
                'Time to the American time format.
                If (Val(txtHours.Text) > 12) _
                And (Val(txtHours.Text) <= 23) Then
                    txtHours.Text = Val(txtHours.Text) - 12
                    Combo1.Text = "PM"
                    UpDown1.Value = txtHours.Text
                    txtHours.SetFocus
                    Call UpdateFinalTime
                    
                ElseIf (Val(txtHours.Text) > 23) Then
                    txtHours.Text = "11"
                    Combo1.Text = "PM"
                    UpDown1.Value = Val(txtHours.Text)
                    txtHours.SetFocus
                    Call UpdateFinalTime
                    
                ElseIf (Val(txtHours.Text) = 0) Then
                    txtHours.Text = "12"
                    Combo1.Text = "AM"
                    UpDown1.Value = Val(txtHours.Text)
                    txtHours.SetFocus
                    
                End If
                
                If (Len(txtHours.Text) = 1) Then
                    txtHours.Text = "0" & txtHours.Text
                    
                End If
                
                ChangeCount = 0
                txtHours.SelStart = 0
                txtHours.SelLength = 2
                
                Call UpdateFinalTime
                
                RaiseEvent Change
                
            'If Time Style is British (Military),
            'show time in one set of 24 hours
            'and discart the AM/PM information.
            Else
                If (Val(txtHours.Text) <> UpDown1.Value) _
                And (Val(txtHours.Text) <= 23) _
                And (Val(txtHours.Text) >= 0) Then
                    UpDown1.Value = txtHours.Text
                    
                End If
                
                If (Val(txtHours.Text) > 23) Then
                    txtHours.Text = "23"
                    UpDown1.Value = Val(txtHours.Text)
                    txtHours.SetFocus
                    Call UpdateFinalTime
                    
                End If
                
                If (Len(txtHours.Text) = 1) Then
                    txtHours.Text = "0" & txtHours.Text
                    
                End If
                
                ChangeCount = 0
                txtHours.SelStart = 0
                txtHours.SelLength = 2
                
                Call UpdateFinalTime
                
                RaiseEvent Change
                
            End If
            
        Else
            ChangeCount = 1
            
        End If
    End If
    
End Sub

Private Sub txtHours_GotFocus()
    txtHours.SelStart = 0
    txtHours.SelLength = 2
    
End Sub

Private Sub txtHours_LostFocus()
    Call txtHours_Change
    txtHours.Enabled = True
    
End Sub

Private Sub UpDown1_Change()
    If Not (intStopUpDown) Then
        StopTextChange = 1
        If (Val(txtHours.Text) = 0) _
        And (m_TimeStyle = American) Then
            txtHours.Text = "12"
            Combo1.Text = "AM"
            UpDown1.Value = 12
            txtHours.SetFocus
        End If
        
        If (ChangeCount = 1) Then
            If (m_TimeStyle = American) Then
                If (Val(txtHours.Text) <> UpDown1.Value) _
                And (Val(txtHours.Text) <= 12) _
                And (Val(txtHours.Text) <> 0) Then
                    UpDown1.Value = txtHours.Text
                End If
                ChangeCount = 0
                
            Else
                If (Val(txtHours.Text) <> UpDown1.Value) _
                And (Val(txtHours.Text) <= 23) _
                And (Val(txtHours.Text) >= 0) Then
                    UpDown1.Value = txtHours.Text
                End If
                ChangeCount = 0
                
            End If
        Else
            txtHours.Text = UpDown1.Value
        End If
        
        If (Len(txtHours.Text) = 1) Then
            txtHours.Text = "0" & txtHours.Text
        End If
        
        txtHours.SelStart = 0
        txtHours.SelLength = 2
    
        StopTextChange = 0
        
        Call UpdateFinalTime
        
        RaiseEvent Change
    End If
    
End Sub

Private Sub UpDown1_UpClick()
    If (Val(txtHours.Text) = 12) Then
        If (Combo1.Text = "AM") Then
            Combo1.Text = "PM"
        Else
            Combo1.Text = "AM"
        End If
    End If
    txtHours.SetFocus
End Sub

Private Sub UpDown1_DownClick()
    If (Val(txtHours.Text) = 11) Then
        If (Combo1.Text = "AM") Then
            Combo1.Text = "PM"
        Else
            Combo1.Text = "AM"
        End If
    End If
    txtHours.SetFocus
    
End Sub

'**********************************
'******* Seting the Minutes *******
'**********************************
Private Sub txtMinutes_KeyDown(KeyCode As Integer, Shift As Integer)
    'Prevent any Shift key combination
    'from going through...
    If (Shift = vbShiftMask) Then
        txtMinutes.Locked = True
        Exit Sub
    End If
    
    'Find out whether the user has
    'typed a number or not.
    Select Case KeyCode
        Case vbKeyLButton, vbKeyRButton, vbKeyBack, _
             vbKeyDelete, vbKeyTab, vbKeyControl, _
             vbKey0, vbKey1, vbKey2, vbKey3, _
             vbKey4, vbKey5, vbKey6, vbKey7, _
             vbKey8, vbKey9, _
             vbKeyNumlock, vbKeyNumpad0, vbKeyNumpad1, _
             vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, _
             vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, _
             vbKeyNumpad8, vbKeyNumpad9
            'Allow key to go through.
            
        Case vbKeyUp
            intKeyUp = True
            intTimeSelStart = txtMinutes.SelStart
            intTimeSelLength = txtMinutes.SelLength
            
        Case vbKeyDown
            intKeyDown = True
            intTimeSelStart = txtMinutes.SelStart
            intTimeSelLength = txtMinutes.SelLength
            
        Case vbKeyReturn, vbKeySeparator
            Combo1.SetFocus
            
        Case Else
            txtMinutes.Locked = True
            
    End Select
End Sub

Private Sub txtMinutes_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intValue As Integer
    Dim intTempTotalValue As String
    Dim intTotalValue As Integer
    
    txtMinutes.Locked = False
    
    '=================================
    '===== When the user presses =====
    '===== the Up Key Arrow.     =====
    '=================================
    If (intKeyUp) Then
        '==================================
        '===== If the first character =====
        '===== is selected.           =====
        '==================================
        If ((intTimeSelStart = 0) _
        Or (intTimeSelLength = 2)) _
        And (Len(txtMinutes.Text) = 2) Then
            intValue = CInt(Left(txtMinutes.Text, 1))
            intTotalValue = Right(txtMinutes.Text, 1)
            
            If (intValue < 5) Then
                'Change immediately...
                ChangeCount2 = 1
                txtMinutes.Text = (intValue + 1) & intTotalValue
                
            Else
                'Change immediately...
                ChangeCount2 = 1
                txtMinutes.Text = 0 & intTotalValue
                
            End If
            
            txtMinutes.SelStart = 0
            txtMinutes.SelLength = 1
            
        '===================================
        '===== If the second character =====
        '===== is selected.            =====
        '===================================
        ElseIf ((intTimeSelStart = 1) _
        Or (intTimeSelStart = 2)) _
        And (Len(txtMinutes.Text) = 2) Then
            intValue = CInt(txtMinutes.Text) + 1
            
            If (intValue < 60) Then
                'Change immediately...
                ChangeCount2 = 1
                txtMinutes.Text = intValue
                    
            Else
                'Change immediately...
                ChangeCount2 = 1
                txtMinutes.Text = "00"
                
            End If
            
            txtMinutes.SelStart = 1
            txtMinutes.SelLength = 1
            
        End If
        
        intKeyUp = False
    End If
    
    '=================================
    '===== When the user presses =====
    '===== the Down Key Arrow.   =====
    '=================================
    If (intKeyDown) Then
        '==================================
        '===== If the first character =====
        '===== is selected.           =====
        '==================================
        If ((intTimeSelStart = 0) _
        Or (intTimeSelLength = 2)) _
        And (Len(txtMinutes.Text) = 2) Then
            
            intValue = CInt(Left(txtMinutes.Text, 1)) - 1
            intTotalValue = Right(txtMinutes.Text, 1)
            
            If (intValue >= 0) Then
                'Change immediately...
                ChangeCount2 = 1
                txtMinutes.Text = intValue & intTotalValue
                
            Else
                'Change immediately...
                ChangeCount2 = 1
                txtMinutes.Text = 5 & intTotalValue
                
            End If
            
            txtMinutes.SelStart = 0
            txtMinutes.SelLength = 1
            
        '===================================
        '===== If the second character =====
        '===== is selected.            =====
        '===================================
        ElseIf ((intTimeSelStart = 1) _
        Or (intTimeSelStart = 2)) _
        And (Len(txtMinutes.Text) = 2) Then
            
            intValue = CInt(txtMinutes.Text) - 1
            
            If (intValue >= 0) Then
                'Change immediately...
                ChangeCount2 = 1
                txtMinutes.Text = intValue
                    
            Else
                'Change immediately...
                ChangeCount2 = 1
                txtMinutes.Text = "59"
                
            End If
            
            txtMinutes.SelStart = 1
            txtMinutes.SelLength = 1
            
        End If
        intKeyDown = False
    End If
    
    txtMinutes.Locked = False
    
End Sub

Private Sub txtMinutes_Change()
    If (StopTextChange2 = 0) Then
        If (ChangeCount2 = 1) Then
            If (Val(txtMinutes.Text) <> UpDown2.Value) _
            And (Val(txtMinutes.Text) <= 59) Then
                UpDown2.Value = txtMinutes.Text
            ElseIf (Val(txtMinutes.Text) > 59) Then
                txtMinutes.Text = "59"
                UpDown2.Value = 59
            End If
            
            If (Len(txtMinutes.Text) = 1) Then
                txtMinutes.Text = "0" & txtMinutes.Text
            End If
            ChangeCount2 = 0
            txtMinutes.SelStart = 0
            txtMinutes.SelLength = 2
            
            Call UpdateFinalTime
            
            RaiseEvent Change
            
        Else
            ChangeCount2 = 1
        End If
    End If
    
End Sub

Private Sub txtMinutes_GotFocus()
    txtMinutes.SelStart = 0
    txtMinutes.SelLength = 2
    
End Sub

Private Sub txtMinutes_LostFocus()
    Call txtMinutes_Change
    txtMinutes.Enabled = True
    
End Sub

Private Sub UpDown2_Change()
    If Not (intStopUpDown) Then
        StopTextChange2 = 1
        If (ChangeCount2 = 1) Then
            If (Val(txtMinutes.Text) <> UpDown2.Value) _
            And (Val(txtMinutes.Text) <= 59) Then
                UpDown2.Value = txtMinutes.Text
            End If
            ChangeCount2 = 0
        Else
            txtMinutes.Text = UpDown2.Value
        End If
        
        If (Len(txtMinutes.Text) = 1) Then
            txtMinutes.Text = "0" & txtMinutes.Text
        End If
        
        txtMinutes.SelStart = 0
        txtMinutes.SelLength = 2
    
        StopTextChange2 = 0
        
        Call UpdateFinalTime
        
        RaiseEvent Change
    End If
    
End Sub

Private Sub UpDown2_UpClick()
    txtMinutes.SetFocus
    
End Sub

Private Sub UpDown2_DownClick()
    txtMinutes.SetFocus
    
End Sub

'**********************************
'********* Seting the AMPM ********
'**********************************
Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown
            'Allow key to go through.
            Exit Sub
        Case vbKeyA
            Combo1.Text = "AM"
            
        Case vbKeyP
            Combo1.Text = "PM"
            
        Case vbKeyReturn, vbKeySeparator
            RaiseEvent Finished

    End Select
    Combo1.SelStart = 0
    Combo1.SelLength = 2
    Combo1.Locked = True
    
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    Combo1.Locked = False
    
End Sub

Private Sub Combo1_Change()
    test = Combo1
    If Not (ChangeCount3) Then
        Call UpdateFinalTime
        RaiseEvent Change
    End If
    
End Sub

Private Sub Combo1_Click()
    Combo1_Change
    
End Sub

Private Sub Combo1_LostFocus()
    Combo1.SelLength = 0
    
End Sub

Private Function UpdateFinalTime()
    Dim intValue As Integer
    
    If (m_TimeStyle = American) Then
        txtFinalTime = txtHours.Text & ":" & txtMinutes.Text & " " & Combo1.Text
        
    Else
        txtFinalTime = txtHours.Text & ":" & txtMinutes.Text
        
    End If
    
    'The following line of code was added to prevent
    'the UpDown1_Change and the UpDown2_Change events
    'from firing when the hours and minutes are been set.
    intStopUpDown = True
    
    UpDown1.Value = txtHours.Text
    UpDown2.Value = txtMinutes.Text
    
    intStopUpDown = False
    
    m_Value = txtFinalTime
    
End Function

Private Sub Command1_Click()
    txtHours.Text = Left(txtInitialTime.Text, 2)
    txtMinutes.Text = Mid(txtInitialTime.Text, 4, 2)
    Combo1.Text = Right(txtInitialTime.Text, 2)
    
End Sub

'==================================================
'======= Following is a list of all the ===========
'======= Properties of this Control.    ===========
'==================================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "This property returns/sets a value that determines whether the TimePick Control can respond to user-generated events or not."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "General"
    Enabled = m_Enabled
    
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
    
    txtHours.Enabled = m_Enabled
    UpDown1.Enabled = m_Enabled
    txtMinutes.Enabled = m_Enabled
    UpDown2.Enabled = m_Enabled
    Combo1.Enabled = m_Enabled
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,1,1,12:00 AM
Public Property Get Value() As String
Attribute Value.VB_Description = "This property returns the time currently displayed on the TimePick Control."
Attribute Value.VB_ProcData.VB_Invoke_Property = "General"
    Value = m_Value
    
End Property

Public Property Let Value(ByVal New_Value As String)
    'This property is read only.

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtHours,txtHours,-1,Text
Public Property Get HourValue() As String
Attribute HourValue.VB_Description = "This property will Get/Set the value for the Hour Text Box."
Attribute HourValue.VB_ProcData.VB_Invoke_Property = "General"
    HourValue = txtHours.Text
    
End Property

Public Property Let HourValue(ByVal New_HourValue As String)
    txtHours.Text() = New_HourValue
    PropertyChanged "HourValue"
    
    Call txtHours_Change
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtMinutes,txtMinutes,-1,Text
Public Property Get MinuteValue() As String
Attribute MinuteValue.VB_Description = "This property will Get/Set the value for the Minute Text Box."
Attribute MinuteValue.VB_ProcData.VB_Invoke_Property = "General"
    MinuteValue = txtMinutes.Text
    
End Property

Public Property Let MinuteValue(ByVal New_MinuteValue As String)
    txtMinutes.Text() = New_MinuteValue
    PropertyChanged "MinuteValue"
    
    Call txtMinutes_Change
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Text
Public Property Get AMPMValue() As AMPM_Type
Attribute AMPMValue.VB_Description = "This property will Get/Set the value for the AMPM Combo Box."
Attribute AMPMValue.VB_ProcData.VB_Invoke_Property = ";Misc"
    If (Combo1.Text = "AM") Then
        AMPMValue = AM
    Else
        AMPMValue = PM
    End If
End Property

Public Property Let AMPMValue(ByVal New_AMPMValue As AMPM_Type)
    Combo1.Text() = New_AMPMValue
    
    If (New_AMPMValue = AM) Then
        Combo1.Text = "AM"
    Else
        Combo1.Text = "PM"
    End If
    
    PropertyChanged "AMPMValue"
    
    Call Combo1_Change
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,3,0,&H8000000F&
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "This property will Get/Set the Background collor for the TimePick Control."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If Ambient.UserMode Then Err.Raise 393
    BackColor = m_BackColor
    
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    If Ambient.UserMode Then Err.Raise 382
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    Frame1.BackColor = m_BackColor
    Frame5.BackColor = m_BackColor
    Frame6.BackColor = m_BackColor
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,3,0,&H80000008&
Public Property Get TimeColor() As OLE_COLOR
Attribute TimeColor.VB_Description = "This property will Get/Set the color of the text that display the time on the TimePick Control."
Attribute TimeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If Ambient.UserMode Then Err.Raise 393
    TimeColor = m_TimeColor
    
End Property

Public Property Let TimeColor(ByVal New_TimeColor As OLE_COLOR)
    If Ambient.UserMode Then Err.Raise 382
    m_TimeColor = New_TimeColor
    PropertyChanged "TimeColor"
    
    txtHours.ForeColor = m_TimeColor
    txtMinutes.ForeColor = m_TimeColor
    Combo1.ForeColor = m_TimeColor
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,3,0,&H80000008&
Public Property Get TitleColor() As OLE_COLOR
Attribute TitleColor.VB_Description = "This property will Get/Set the color of the Title on the TimePick Control."
Attribute TitleColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    If Ambient.UserMode Then Err.Raise 393
    TitleColor = m_TitleColor
    
End Property

Public Property Let TitleColor(ByVal New_TitleColor As OLE_COLOR)
    If Ambient.UserMode Then Err.Raise 382
    m_TitleColor = New_TitleColor
    PropertyChanged "TitleColor"
    
    Frame5.ForeColor = m_TitleColor
    Frame6.ForeColor = m_TitleColor
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Frame5,Frame5,-1,BorderStyle
Public Property Get ShowFrame() As BorderStyle
Attribute ShowFrame.VB_Description = "Enable/Disable borders and titles."
Attribute ShowFrame.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ShowFrame = Frame5.BorderStyle
End Property

Public Property Let ShowFrame(ByVal New_ShowFrame As BorderStyle)
    Frame5.BorderStyle = New_ShowFrame
    PropertyChanged "ShowFrame"
    
    Frame6.BorderStyle = Frame5.BorderStyle
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=28,0,0,0
Public Property Get TimeStyle() As AmericBritish
Attribute TimeStyle.VB_Description = "Set/Get the time style. It can be the American style that divides a day into two sets of 12 hrs and requires the AM/PM Combo Box or the British style that displays the time in one set of 24 hrs."
Attribute TimeStyle.VB_ProcData.VB_Invoke_Property = ";Behavior"
    TimeStyle = m_TimeStyle
    
End Property

Public Property Let TimeStyle(ByVal New_TimeStyle As AmericBritish)
    Dim intValue As Integer
    
    m_TimeStyle = New_TimeStyle
    PropertyChanged "TimeStyle"
    
    intValue = CInt(txtHours.Text)
    
    If (m_TimeStyle = British) Then
        Combo1.Visible = False
        UpDown1.Max = 23
        UpDown1.Min = 0
        
        If (Combo1.Text = "AM") Then
            
            If (intValue = 12) Then
                
                'The two "ChangeCount = 0" statements are
                'here to prevent the event Change from
                'been raised.
                ChangeCount = 0
                txtHours.Text = "00"
                ChangeCount = 0
            End If
            
        Else
            intValue = intValue + 12
            
            'The two "ChangeCount = 0" statements are
            'here to prevent the event Change from
            'been raised.
            ChangeCount = 0
            
            If (intValue > 23) Then
                txtHours.Text = "00"
                
            Else
                txtHours.Text = Format(intValue, "00")
                
            End If
            
            ChangeCount = 0
                            
        End If
        
    Else
        
        Combo1.Visible = True
        UpDown1.Max = 12
        UpDown1.Min = 1
        
        'The two "ChangeCount = 0" statements are
        'here to prevent the event Change from
        'been raised.
        ChangeCount = 0
        ChangeCount3 = True
        
        If (intValue = 0) Then
            txtHours.Text = "12"
            Combo1.Text = "AM"
        
        ElseIf (intValue < 13) Then
            Combo1.Text = "AM"
            
        Else
            txtHours.Text = Format((intValue - 12), "00")
            Combo1.Text = "PM"
            
        End If
                
        ChangeCount3 = False
        ChangeCount = 0
        
    End If
    
    UpDown1.Value = txtHours.Text
    
    Call UserControl_Resize
    Call UpdateFinalTime
    RaiseEvent Change
    
End Property

'==================================================
'======= Following are the Subs that will    ======
'======= initialize and save the properties. ======
'==================================================

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_Value = m_def_Value
    m_BackColor = m_def_BackColor
    m_TimeColor = m_def_TimeColor
    m_TitleColor = m_def_TitleColor
    m_TimeStyle = m_def_TimeStyle
    
    m_tatata = m_def_tatata
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'This property should be first on the list.
    m_TimeStyle = PropBag.ReadProperty("TimeStyle", m_def_TimeStyle)
    
    If (m_TimeStyle = British) Then
        Combo1.Visible = False
        UpDown1.Max = 23
        UpDown1.Min = 0
        
    Else
        Combo1.Visible = True
        UpDown1.Max = 12
        UpDown1.Min = 1
        
    End If
    
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    
    'The following line of code was added to prevent the
    'txtHours_Change from firing when the hours
    'are been set for the first time.
    StopTextChange = 1
    txtHours.Text = PropBag.ReadProperty("HourValue", "05")
    
    'The following line of code was added to prevent the
    'txtMinutes_Change from firing when the minutes
    'are been set for the first time.
    StopTextChange2 = 1
    
    txtMinutes.Text = PropBag.ReadProperty("MinuteValue", "30")
    Combo1.Text = PropBag.ReadProperty("AMPMValue", "PM")
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_TimeColor = PropBag.ReadProperty("TimeColor", m_def_TimeColor)
    m_TitleColor = PropBag.ReadProperty("TitleColor", m_def_TitleColor)
    
    Frame5.BorderStyle = PropBag.ReadProperty("ShowFrame", 1)
    
    Call UserControl_Resize
    
    txtHours.Enabled = m_Enabled
    UpDown1.Enabled = m_Enabled
    txtMinutes.Enabled = m_Enabled
    UpDown2.Enabled = m_Enabled
    Combo1.Enabled = m_Enabled
    
    Frame1.BackColor = m_BackColor
    Frame5.BackColor = m_BackColor
    Frame6.BackColor = m_BackColor
    
    txtHours.ForeColor = m_TimeColor
    txtMinutes.ForeColor = m_TimeColor
    Combo1.ForeColor = m_TimeColor
    
    Frame5.ForeColor = m_TitleColor
    Frame6.ForeColor = m_TitleColor
    
    Frame6.BorderStyle = Frame5.BorderStyle
    
    txtFinalTime = m_Value
    
    StopTextChange = 0
    StopTextChange2 = 0
    
    Call UpdateFinalTime
    
    m_tatata = PropBag.ReadProperty("tatata", m_def_tatata)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("HourValue", txtHours.Text, "05")
    Call PropBag.WriteProperty("MinuteValue", txtMinutes.Text, "30")
    Call PropBag.WriteProperty("AMPMValue", Combo1.Text, "PM")
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("TimeColor", m_TimeColor, m_def_TimeColor)
    Call PropBag.WriteProperty("TitleColor", m_TitleColor, m_def_TitleColor)
    Call PropBag.WriteProperty("ShowFrame", Frame5.BorderStyle, 1)
    Call PropBag.WriteProperty("TimeStyle", m_TimeStyle, m_def_TimeStyle)
    
    Call PropBag.WriteProperty("tatata", m_tatata, m_def_tatata)
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get tatata() As Variant
    tatata = m_tatata
End Property

Public Property Let tatata(ByVal New_tatata As Variant)
    m_tatata = New_tatata
    PropertyChanged "tatata"
End Property

