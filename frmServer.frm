VERSION 5.00
Begin VB.Form frmServer 
   Caption         =   "Serer"
   ClientHeight    =   1560
   ClientLeft      =   3555
   ClientTop       =   3840
   ClientWidth     =   3720
   ControlBox      =   0   'False
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3720
   Begin VB.Frame Frame1 
      Caption         =   "Add Server"
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Text            =   "25"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   " "
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Port"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   3
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Server "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public inder As Boolean

Private Sub cmdAdd_Click()
    If Len(Trim(txtServer.Text)) = 0 Then
        MsgBox "Please enter a server"
        txtServer.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtPort.Text)) = 0 Or Not IsNumeric(txtPort.Text) Then
        MsgBox "Please enter a valid port"
        txtPort.SetFocus
        Exit Sub
    End If
    inder = True
    Name1 = Trim(txtServer.Text)
    Port1 = Trim(txtPort.Text)
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    inder = False
    Me.Hide
End Sub

