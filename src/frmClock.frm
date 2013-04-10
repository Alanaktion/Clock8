VERSION 5.00
Begin VB.Form frmClock 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Clock8"
   ClientHeight    =   2100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Segoe UI Semilight"
      Size            =   8.25
      Charset         =   0
      Weight          =   350
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmClock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   140
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrClose 
      Interval        =   1200
      Left            =   5280
      Top             =   120
   End
   Begin VB.Timer tmrClock 
      Interval        =   20
      Left            =   120
      Top             =   120
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5625
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblMonthDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "March 26"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   3600
      TabIndex        =   2
      Top             =   1020
      Width           =   1725
   End
   Begin VB.Label lblWeekday 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tuesday"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   21.75
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   3600
      TabIndex        =   1
      Top             =   450
      Width           =   1500
   End
   Begin VB.Label lblTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "11:41"
      BeginProperty Font 
         Name            =   "Segoe UI Light"
         Size            =   60
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1590
      Left            =   600
      TabIndex        =   0
      Top             =   210
      Width           =   2730
   End
End
Attribute VB_Name = "frmClock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public IsTopMost As Boolean
 

Private Sub Form_Activate()

    ' Set Opacity
    Dim bytOpacity As Byte
    bytOpacity = 230
    Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, GetWindowLong(Me.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(Me.hwnd, 0, bytOpacity, LWA_ALPHA)
    
    ' Set Always on Top
    lR = SetTopMostWindow(Me.hwnd, True)
    IsTopMost = True
    
    ' Position Form
    Me.Left = Screen.TwipsPerPixelX * 50
    Me.Top = (Screen.Height - Me.Height - 50 * Screen.TwipsPerPixelY)
    
    ' Initialize Clock
    tmrClock.Enabled = True
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Show close button and reset button hide timer
    lblClose.Visible = True
    lblClose.Left = Me.ScaleWidth - lblClose.Width
    tmrClose.Enabled = False
    tmrClose.Enabled = True
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ' Toggle Always On Top with right click
    If Button = 2 Then
        IsTopMost = Not IsTopMost
        lR = SetTopMostWindow(Me.hwnd, IsTopMost)
    End If
    
End Sub

Private Sub lblClose_Click()

    ' Exit Clock
    Unload Me
    
End Sub

Private Sub lblMonthDate_Click()

End Sub

Private Sub tmrClock_Timer()

    ' Update label captions
    If Not lblTime.Caption = Format(Now, "h:mm   a/p") Then _
           lblTime.Caption = Format(Now, "h:mm   a/p")          ' Extra space to prevent a/p from showing
    If Not lblWeekday.Caption = Format(Now, "dddd") Then _
           lblWeekday.Caption = Format(Now, "dddd")
    If Not lblMonthDate.Caption = Format(Now, "mmmm d") Then _
           lblMonthDate.Caption = Format(Now, "mmmm d")
    
    ' Position date around time
    If lblWeekday.Left <> lblTime.Left + lblTime.Width + 18 Then
        lblWeekday.Left = lblTime.Left + lblTime.Width + 18
        lblMonthDate.Left = lblWeekday.Left
    End If
    
    ' Size window around date
    If lblWeekday.Width > lblMonthDate.Width Then
        If Me.Width <> lblWeekday.Width + lblWeekday.Left + 40 Then _
            Me.Width = lblWeekday.Width + lblWeekday.Left + 40
    Else
        If Me.Width <> lblMonthDate.Width + lblMonthDate.Left + 40 Then _
            Me.Width = lblMonthDate.Width + lblMonthDate.Left + 40
    End If
    
End Sub

Private Sub tmrClose_Timer()

    lblClose.Visible = False
    
End Sub
