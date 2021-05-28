VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMail 
   Caption         =   "Magic Mail v1.0"
   ClientHeight    =   7005
   ClientLeft      =   2565
   ClientTop       =   1350
   ClientWidth     =   5520
   Icon            =   "frmMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   5520
   Begin MSFlexGridLib.MSFlexGrid MSFlex 
      Height          =   3135
      Left            =   2880
      TabIndex        =   7
      Top             =   2760
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.TextBox txtLog 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   6360
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   5295
      Begin VB.TextBox txtHost 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtIp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Date :"
         Height          =   255
         Left            =   3720
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblDate 
         Height          =   255
         Left            =   4200
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMail.frx":0896
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   390
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   688
      BandCount       =   2
      BandBorders     =   0   'False
      VariantHeight   =   0   'False
      _CBWidth        =   5295
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinHeight1      =   330
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      MinHeight2      =   330
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   10
         Top             =   30
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   582
         ButtonWidth     =   1693
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Send"
               Object.ToolTipText     =   "Send"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Clear"
               Object.ToolTipText     =   "Clear"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtSender 
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Tag             =   "0"
      Text            =   "Sender@anydomain.com"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtSenderName 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Tag             =   "0"
      Text            =   "Sender Name"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtReceiver 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Tag             =   "0"
      Text            =   "Reciever@anydomain.com"
      Top             =   1560
      Width           =   2535
   End
   Begin MSWinsockLib.Winsock Smtp 
      Left            =   600
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   3480
   End
   Begin VB.CheckBox chkHtml 
      Caption         =   "Send as HTML"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
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
      Left            =   4320
      TabIndex        =   9
      Top             =   5880
      Width           =   975
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
      Left            =   3000
      TabIndex        =   8
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtMessage 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "0"
      Text            =   "frmMail.frx":09AA
      Top             =   3360
      Width           =   2655
   End
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Tag             =   "0"
      Text            =   "Type your subject here . . . ."
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtRecName 
      Height          =   285
      Left            =   120
      MaxLength       =   255
      TabIndex        =   0
      Tag             =   "0"
      Text            =   "Reciever Name"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4320
      TabIndex        =   11
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Choose Server"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   18
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Subject"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Reciever's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Reciever's E-Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Sender's E-Mail "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Sender's Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private timer As Long
Private data As Boolean
Private inder As Boolean
Private Change As Boolean
Dim inData As String
Const SmtpName = 0
Private Const TIME_OUT = 30
Const PortNo = 1

Private Sub chkHtml_Click()
    If chkHtml.Value = 1 Then
        chkHtml.Tag = "Html;"
    Else
        chkHtml.Tag = "Plain;"
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim ServerAdded As Long
    Dim Port As Integer
    frmServer.Show vbModal
    If frmServer.inder Then
        MSFlex.Row = 1
        If MSFlex.Text = "" Then
            MSFlex.Row = 1
            MSFlex.Col = 0
            MSFlex.Text = Name1
            MSFlex.Col = 1
            MSFlex.Text = Port1
        Else
            MSFlex.AddItem Name1
            ServerAdded = MSFlex.Rows - 1
            MSFlex.Row = ServerAdded
            MSFlex.Col = 1
            MSFlex.Text = Port1
        End If
        SaveConfig
    End If
    Unload frmServer
End Sub
Private Sub SaveConfig()
Dim TotalRows As Integer
Dim CurrentRow As Integer
Dim SName As String
Dim SPort As String

On Error GoTo Err
    TotalRows = MSFlex.Rows - 1
    Open "Server.ini" For Output As #1
    For CurrentRow = 1 To TotalRows
        MSFlex.Row = CurrentRow
        MSFlex.Col = SmtpName
        SName = MSFlex.Text
        MSFlex.Col = PortNo
        SPort = MSFlex.Text
        Write #1, SName; SPort
    Next
    Close #1
    Exit Sub

Err:
    MsgBox "Server.ini Save Error", vbInformation, "Save Error"
End Sub

Private Sub cmdRemove_Click()
    If Change = True Then
        If MSFlex.Rows = 2 Then
            MSFlex.Col = SmtpName
            MSFlex.Text = ""
            MSFlex.Col = PortNo
            MSFlex.Text = ""
        Else
            MSFlex.RemoveItem (MSFlex.Row)
        End If
    Else
        MsgBox "You must select a server name to remove from the server list!", vbInformation, "Remove Error"
    End If
    Change = False
    MSFlex.Row = 0
    SaveConfig
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim str As String
    inder = False
    timer = 0
    
    MSFlex.ColAlignment(SmtpName) = flexAlignLeftCenter
    MSFlex.ColAlignment(PortNo) = flexAlignCenterCenter
    MSFlex.ColWidth(0) = 1790
    MSFlex.ColWidth(1) = 400
    MSFlex.Row = 0
    MSFlex.Col = 0
    MSFlex.Text = "Server Name"
    MSFlex.Col = 1
    MSFlex.Text = "Port"
    MSFlex.SelectionMode = flexSelectionByRow
    Call LoadIni
    
    lblDate.Caption = Date
    txtIp = Smtp.LocalIP
    txtPort = Smtp.LocalPort
    txtHost = Smtp.LocalHostName
    
End Sub
Private Sub LoadIni()
    Dim SmtpServer As String
    Dim SmtpPort As String
    Dim SmtpAdd As Integer
    On Error GoTo Err
    Open "Server.ini" For Input As #1
    While Not EOF(1)
        Input #1, SmtpServer, SmtpPort
        MSFlex.Row = 1
        If MSFlex.Text = "" Then
            MSFlex.Col = SmtpName
            MSFlex.Text = SmtpServer
            MSFlex.Col = 1
            MSFlex.Text = SmtpPort
        Else
            MSFlex.AddItem SmtpServer
            SmtpAdd = MSFlex.Rows - 1
            MSFlex.Row = SmtpAdd
            MSFlex.Col = PortNo
            MSFlex.Text = SmtpPort
        End If
    Wend
    Close #1
    Exit Sub
Err:
    MsgBox "Error in opening Server.ini", vbInformation, "File Error"
    End
End Sub

Private Sub Label10_Click()
    MsgBox "Programmed By Inderpal Singh" + vbCrLf + "E-Mail : inderpal0@hotmail.com" + vbCrLf + "http://connect.to/lanserver", vbInformation, "About Inderpal Singh"
End Sub

Private Sub Label8_Click()
    Unload Me
End Sub

Private Sub MSFlex_Click()
    Change = True
End Sub

Private Sub Timer1_Timer()
    timer = timer + 1
    If timer = TIME_OUT Then
        CConnection True  'Disconnected if time out
        MsgBox "Could not connected to the Selected Host" + vbCrLf + "operation timed out", vbInformation, "Time Out"
        Timer1.Enabled = False
    End If
End Sub
Private Sub CConnection(Err As Boolean)
    Smtp.Close  'Close connection & Enable Controls
    txtLog = ""
    txtRecName.Enabled = True
    txtReceiver.Enabled = True
    txtSenderName.Enabled = True
    txtSender.Enabled = True
    txtSubject.Enabled = True
    txtMessage.Enabled = True
    MSFlex.Enabled = True
    cmdAdd.Enabled = True
    cmdRemove.Enabled = True
    Toolbar1.Enabled = True
    chkHtml.Enabled = True
    If Err Then If MsgBox("Do you want to remove this service ", vbYesNo) = vbYes Then cmdRemove_Click
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    
        Case 1:
                Call SendButton
        Case 3:
                 Call Clear
    End Select
End Sub
Private Sub Clear()
    txtRecName.Text = ""
    txtReceiver = ""
    txtSenderName = ""
    txtSender = ""
    txtSubject = ""
    txtMessage = ""
    
    txtRecName.Tag = 0
    txtReceiver.Tag = 0
    txtSenderName.Tag = 0
    txtSender.Tag = 0
    txtSubject.Tag = 0
    txtMessage.Tag = 0
    txtRecName.SetFocus
End Sub
Private Sub txtSender_GotFocus()
    If txtSender.Tag = 0 Then
        txtSender.Tag = 1
        txtSender.Text = ""
    End If
End Sub
Private Sub txtSender_Validate(KeepFocus As Boolean)
    If txtSender.Text = "" Then
        txtSender.Text = "you@anydomain.com"
        KeepFocus = False
        txtSender.Tag = 0
    End If
End Sub
Private Sub txtSenderName_GotFocus()
    If txtSenderName.Tag = 0 Then
        txtSenderName.Tag = 1
        txtSenderName.Text = ""
    End If
End Sub
Private Sub txtSenderName_Validate(KeepFocus As Boolean)
    If txtSenderName.Text = "" Then
        txtSenderName.Text = "Sender Name"
        KeepFocus = False
        txtSenderName.Tag = 0
    End If
End Sub
Private Sub txtReceiver_GotFocus()
    If txtReceiver.Tag = 0 Then
        txtReceiver.Tag = 1
        txtReceiver.Text = ""
    End If
End Sub
Private Sub txtReceiver_Validate(KeepFocus As Boolean)
    If txtReceiver.Text = "" Then
        txtReceiver.Text = "Receiver@anydomain.com"
        KeepFocus = False
        txtReceiver.Tag = 0
    End If
End Sub

Private Sub txtRecName_GotFocus()
    If txtRecName.Tag = 0 Then
        txtRecName.Tag = 1
        txtRecName.Text = ""
    End If
End Sub
Private Sub txtRecName_Validate(KeepFocus As Boolean)
    If txtRecName.Text = "" Then
        txtRecName.Text = "Reciever Name"
        KeepFocus = False
        txtRecName.Tag = 0
    End If
End Sub

Private Sub txtSubject_GotFocus()
    If txtSubject.Tag = 0 Then
        txtSubject.Tag = 1
        txtSubject.Text = ""
    End If
End Sub

Private Sub txtSubject_Validate(KeepFocus As Boolean)
    If txtSubject = "" Then
        txtSubject = "Type your subject here . . . ."
        KeepFocus = False
        txtSubject.Tag = 0
    End If
End Sub
Private Sub txtMessage_GotFocus()
    If txtMessage.Tag = 0 Then
        txtMessage.Tag = 1
        txtMessage.Text = ""
    End If
End Sub
Private Sub txtMessage_Validate(KeepFocus As Boolean)
    If txtMessage.Text = "" Then
        txtMessage.Text = "Type your message here . . .  ."
        KeepFocus = False
        txtMessage.Tag = 0
    End If
End Sub

Private Sub Smtp_Connect()
    txtLog = "Connected"
    timer = 0
    Timer1.Enabled = True
    While Not inder         'Wait for reply
        If Smtp.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    
    Dim reply As String
    Dim temp() As String
    reply = inData
    inData = ""
    inder = False
    temp = Split(reply, " ")
    If Not Val(temp(0)) = 220 Then           'Error occured
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CConnection False
        Exit Sub
    End If
    txtLog = "Receiving Message"
    'Start the process
    Smtp.SendData "HELLO " + Smtp.LocalHostName + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not inder         'Wait for reply
        If Smtp.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    inder = False
    temp = Split(reply, " ")
    If Not Val(temp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CConnection False
        Exit Sub
    End If
    'Send MAIL FROM
    Smtp.SendData "MAIL FROM:<" + txtSender.Text + ">" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not inder         'Wait for reply
        If Smtp.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    inder = False
    temp = Split(reply, " ")
    If Not Val(temp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CConnection True
        Exit Sub
    End If
    'Send RCPT TO
    Smtp.SendData "RCPT TO:<" + txtReceiver.Text + ">" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not inder         'Wait for reply
        If Smtp.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    inder = False
    temp = Split(reply, " ")
    If Not Val(temp(0)) = 250 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CConnection True
        Exit Sub
    End If
    'Send DATA
    DoEvents
    Smtp.SendData "DATA" + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not inder         'Wait for reply
        If Smtp.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    inder = False
    temp = Split(reply, " ")
    If Not Val(temp(0)) = 354 Then
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CConnection False
        Exit Sub
    End If
    txtLog = "Sending Mail . . . . "
    'Send the E-Mail
    Smtp.SendData "From: <" + txtSender.Text + ">" + vbCrLf + _
                      "To: " + txtReceiver.Text + vbCrLf + _
                      "Subject: " + txtSubject.Text + vbCrLf + _
                      "Mailer: Magic Mail v1.0" + vbCrLf + _
                      "Mime-Version: 1.0" + vbCrLf + _
                      "Content-Type: text/" + chkHtml.Tag + vbTab + "charset=us-ascii" + vbCrLf + vbCrLf + _
                      txtMessage.Text
    Smtp.SendData vbCrLf + "." + vbCrLf
    DoEvents
    timer = 0
    Timer1.Enabled = True
    While Not inder             'Wait for reply
        If Smtp.State = sckClosed Then Exit Sub
        DoEvents
    Wend
    Timer1.Enabled = False
    reply = inData
    inData = ""
    inder = False
    temp = Split(reply, " ")
    If Not Val(temp(0)) = 250 Then               'Error occured
        MsgBox "Server returned the following error:" + vbCrLf + reply
        CConnection False
        Exit Sub
    End If
    Smtp.SendData "QUIT"
    MsgBox "Message Successfully Sent" + vbCrLf + "Thanks For Using Magic Mail", vbInformation, "Successful"
    CConnection False
End Sub
Private Sub Smtp_DataArrival _
(ByVal bytesTotal As Long)
    Dim data As String
    Smtp.GetData data, vbString
    inData = inData + data
    If strcmp(Right$(inData, 2), vbCrLf) = 0 Then inder = True
End Sub

Private Sub Smtp_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    If Not Number = sckSuccess Then
        MsgBox Description          'Display error
        Timer1.Enabled = False
        CConnection True
    End If
End Sub

Private Sub SendButton()
    Dim tmp As String
    Dim tmp1 As String
    Dim Row1 As String
    If Change = False Then
        MsgBox "Please Choose a Server First", vbInformation, "Server"
        Exit Sub
    Else
        Row1 = MSFlex.Row
        tmp = MSFlex.Text
        MSFlex.Col = 1
        tmp1 = MSFlex.Text
    End If
    If txtRecName.Tag = 0 Then
        MsgBox "Please enter Receiver's Name"
        txtRecName.SetFocus
        Exit Sub
    End If
    If txtReceiver.Tag = 0 Then
        MsgBox "Please enter Receiver's E-mail"
        txtReceiver.SetFocus
        Exit Sub
    End If
    If txtSenderName.Tag = 0 Then
        MsgBox "Please enter Sender's Name"
        txtSenderName.SetFocus
        Exit Sub
    End If
    If txtSender.Tag = 0 Then
        MsgBox "Please enter Sender's E-mail"
        txtSender.SetFocus
        Exit Sub
    End If
    
    If txtSubject.Tag = 0 Then
        MsgBox "Please enter a subject"
        txtSubject.SetFocus
        Exit Sub
    End If
    
    If InStr(1, txtSender, "@") = 0 Then
        MsgBox "The Senders email address must contain an @ character"
        txtSender.SetFocus
        Exit Sub
    End If
    
    If InStr(1, txtReceiver, "@") = 0 Then
        MsgBox "The Receiver email address must contain an @ character"
        txtReceiver.SetFocus
        Exit Sub
    End If
    
    txtLog = "Connecting . . . ."
    Smtp.Connect tmp, Val(tmp1)    'Connect to server
    txtSenderName.Enabled = False
    txtSender.Enabled = False
    txtRecName.Enabled = False
    txtReceiver.Enabled = False
    txtSubject.Enabled = False
    txtMessage.Enabled = False
    MSFlex.Enabled = False
    Toolbar1.Enabled = False
    cmdAdd.Enabled = False
    cmdRemove.Enabled = False
    chkHtml.Enabled = False
End Sub

