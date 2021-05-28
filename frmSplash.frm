VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   3000
   ClientTop       =   3000
   ClientWidth     =   5745
   ForeColor       =   &H00FF0000&
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":08CA
   ScaleHeight     =   3270
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2760
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4920
      Top             =   1200
   End
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   4920
      Top             =   2520
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1260
      Index           =   1
      Left            =   2040
      Picture         =   "frmSplash.frx":3D98C
      Top             =   960
      Width           =   1770
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Magic Mail v1.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   540
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Private Sub Timer1_Timer()
    Progress.Value = Progress.Value + 2
    If Progress.Value = 100 Then
        tmrEffect.Enabled = True
        Timer1.Enabled = False
    End If
End Sub

Private Sub tmrEffect_Timer()
    Me.Height = Me.Height - (Me.Height / 2)
    Me.Top = Me.Top - (Me.Top / 2)
    Me.Left = Me.Left - (Me.Left / 2)
    Me.Width = Me.Width - (Me.Width / 2)
    If Me.Height = 0 Then
            Unload Me
            frmMail.Show
        End If
End Sub
