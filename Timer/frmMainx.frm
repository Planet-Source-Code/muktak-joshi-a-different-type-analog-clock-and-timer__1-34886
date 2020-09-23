VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "BigTimer"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   193
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   199
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   0
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   159
      TabIndex        =   0
      ToolTipText     =   "Click Here"
      Top             =   0
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   1440
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Click Here"
      Top             =   1920
      Width           =   2415
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSec As Single
Dim LastMin As Single
Dim LastHour As Single


Private Function ConvertMinutes(Mins As Integer) As Single

If (Mins < 15) And (Mins > 0) Then
    ConvertMinutes = (15 - Mins) * 0.1075
    Exit Function
End If
If Mins = 15 Then ConvertMinutes = 0
If Mins = 0 Then ConvertMinutes = 1.68
If Mins > 15 Then
    Mins = Mins - 15
    ConvertMinutes = 6.3685 - (0.1075 * Mins)
    Exit Function
End If

End Function
Private Function ConvertHours(Hrs As Integer) As Single
If Hrs > 12 Then Hrs = Hrs - 12

If (Hrs < 3) And (Hrs > 0) Then
    ConvertHours = (3 - Hrs) * 0.53
    Exit Function
End If
If Hrs > 3 Then
    Hrs = Hrs - 3
    ConvertHours = 6.3685 - (0.53 * Hrs)
    Exit Function
End If
If Hrs = 3 Then ConvertHours = 0
If Hrs = 0 Then ConvertHours = 1.68
End Function

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
'LastSec = 1.68
LastSec = ConvertMinutes(DatePart("s", Time))
LastMin = ConvertMinutes(Minute(Time))
LastHour = ConvertHours(Hour(Time))

End Sub

Private Sub Form_Resize()
Picture1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - lblTime.Height
Me.ScaleHeight = Me.ScaleHeight + lblTime.Height
lblTime.Move 0, Picture1.Top + Picture1.Height, Picture1.Width
Call Form_Load
End Sub


Private Sub Picture1_Click()
frmTimer.Show
End Sub



Private Sub Timer1_Timer()
Dim Seconds As Integer

On Error Resume Next

If LastSec < 0.279 Then
LastSec = 6.35
End If
LastSec = LastSec - 0.1
Seconds = (LastSec / 0.1) - 1
Seconds = 60 - Seconds
Picture1.Cls

For X = Picture1.ScaleHeight - (Picture1.ScaleHeight / 1.6) + 2 To Picture1.ScaleHeight - (Picture1.ScaleHeight / 1.7)
Picture1.Circle (Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2), X, vbBlack, LastSec, LastSec - 0.1
Next

LastMin = ConvertMinutes(Minute(Time))
LastHour = ConvertHours(Hour(Time))

For X = 1 To Picture1.ScaleHeight - (Picture1.ScaleHeight / 1.6)
Picture1.Circle (Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2), X, vbRed, LastMin, LastMin - 0.1
Next

For X = Picture1.ScaleHeight - (Picture1.ScaleHeight / 1.2) To Picture1.ScaleHeight - (Picture1.ScaleHeight / 1.4)
Picture1.Circle (Picture1.ScaleWidth / 2, Picture1.ScaleHeight / 2), X, &H800000, LastHour, LastHour - 0.1
Next
lblTime.Caption = Time
End Sub

