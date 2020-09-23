VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BigTimer 1.00"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4455
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Setting Section"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton cmdstart 
         Caption         =   "Start Timer"
         Height          =   375
         Left            =   1680
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtfile 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CommandButton cmdfile 
         Caption         =   "&Browse"
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox cbounit 
         Height          =   315
         ItemData        =   "frmmain.frx":0442
         Left            =   1320
         List            =   "frmmain.frx":0452
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtint 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Run File :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Unit :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Set Interval :"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4080
      Top             =   3480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   15
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   4455
   End
End
Attribute VB_Name = "frmTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim timeis As Long
Dim part As Integer

Private Sub cmdfile_Click()
CommonDialog1.ShowOpen
txtfile.Text = CommonDialog1.FileName
End Sub

Private Sub cmdset_Click()

End Sub

Private Sub cmdstart_Click()
If Not txtint.Text = "" Then
If txtint.Text > 0 Then
Select Case cbounit.Text


Case "MilliSeconds"
timeis = Int(txtint.Text / 1000)
Case "Seconds"
timeis = Int(txtint.Text)
Case "Minutes"
timeis = Int(txtint.Text * 60)
Case "Hours"
timeis = Int(txtint.Text * 3600)
End Select
If timeis = 0 Then timeis = Int(txtint.Text)
part = 4455 / timeis

Timer1.Enabled = True
Screen.MousePointer = 11
End If
End If
End Sub

Private Sub Form_Load()
Timer1.Enabled = False

End Sub

Private Sub Timer1_Timer()

If timeis > 0 Then
timeis = timeis - 1
Label5.Width = Label5.Width + part
Else
ShellExecute frmmain.hwnd, "Open", txtfile.Text, 0, App.Path, 0
Timer1.Enabled = False
Label5.Width = 0
Screen.MousePointer = 0
End If
End Sub

