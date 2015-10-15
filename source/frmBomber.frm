VERSION 5.00
Begin VB.Form frmSkypeBomber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SkypeBomber V1.0"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   Icon            =   "frmBomber.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Height          =   1095
      Left            =   3960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   5760
      Width           =   4215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel (Quit)"
      Height          =   495
      Left            =   6120
      TabIndex        =   10
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtMsg 
      Height          =   1935
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      ToolTipText     =   "Type message here"
      Top             =   3120
      Width           =   4215
   End
   Begin VB.HScrollBar barTimes 
      Height          =   255
      LargeChange     =   10
      Left            =   5520
      Max             =   999
      Min             =   1
      TabIndex        =   6
      Top             =   2400
      Value           =   1
      Width           =   1335
   End
   Begin VB.ListBox lstContacts 
      Height          =   4350
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Label lblVictim 
      Alignment       =   2  'Center
      Caption         =   "Select a contact"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   3960
      TabIndex        =   16
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label9 
      Caption         =   "Victim:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "memin.cotorrito@gmail.com"
      Height          =   255
      Left            =   1680
      TabIndex        =   14
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Please report any bug to"
      Height          =   255
      Left            =   1680
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "by Underdog1987"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Type your message here:"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "times."
      Height          =   255
      Left            =   6960
      TabIndex        =   5
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label lblTimes 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Send message"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Select a contact to send message(s)"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   120
      Picture         =   "frmBomber.frx":030A
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Skype Bomber"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmSkypeBomber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub barTimes_Change()
lblTimes.Caption = barTimes.Value
End Sub

Private Sub barTimes_Scroll()
lblTimes.Caption = barTimes.Value
End Sub

Private Sub Command1_Click()
If lstContacts.ListIndex = -1 Then
    MsgBox ("Select a contact to send message"), vbCritical, "SkypeBomber"
    Exit Sub
ElseIf RTrim(LTrim(txtMsg.Text)) = "" Then
    MsgBox ("Cannot send empty message"), vbCritical, "SkypeBomber"
    Exit Sub
Else
    If LCase$(lstContacts.List(lstContacts.ListIndex)) = "all-contacts" Then
        For c = 1 To lstContacts.ListCount
            lstContacts.ListIndex = lstContacts.ListIndex + 1
            'MsgBox lstContacts.List(lstContacts.ListIndex)
            sendMsg lstContacts.List(lstContacts.ListIndex), barTimes.Value, txtMsg.Text
        Next
    Else
        sendMsg lstContacts.List(lstContacts.ListIndex), barTimes.Value, txtMsg.Text
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo notInstalled
barTimes.Value = 1
Dim oSkype As Object
Set oSkype = CreateObject("Skype4COM.Skype")
txtStatus.Text = "Starting..."
If Not oSkype.Client.IsRunning Then
    oSkype.Client.Start
End If
oSkype.Attach
DoEvents
txtStatus.Text = txtStatus.Text & " Done." & vbCrLf & "Getting Skype Friends..."
lstContacts.AddItem "All-contacts"
For Each user In oSkype.Friends
    lstContacts.AddItem user.Handle
Next
txtStatus.Text = txtStatus.Text & " Done" & vbCrLf & "Ready."
    'Set oCall = oSkype.PlaceCall("skaziunderdog")
notInstalled:
If Err <> 0 Then
    Dim doInstall As Integer
    doInstall = MsgBox("Ops! I think that Skype is not installed. Install now?", vbYesNo + vbQuestion, "SkypeBomber")
    If doInstall = vbYes Then
        Dim wsh As Object
        Set wsh = CreateObject("Wscript.shell")
        wsh.run "iexplore.exe http://www.skype.com"
    End If
    End
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim rExit As Integer
rExit = MsgBox("Exit?", vbQuestion + vbYesNo, "SkypeBomber")
If rExit = vbNo Then
    Cancel = True
Else
    MsgBox ("SkypeBomber coded by Underdog1987 http://underdog1987.wordpress.com")
    Cancel = False
End If

End Sub

Private Sub sendMsg(Contact As String, iTimes As Integer, Message As String)
    Dim oSkype As Object
    Set oSkype = CreateObject("Skype4COM.Skype")
    If Not oSkype.Client.IsRunning Then
        oSkype.Client.Start
    End If
    oSkype.Attach
    DoEvents
    txtStatus.Text = txtStatus.Text & "Sending messages to " & Contact & "......"
    For x = 1 To iTimes
        oSkype.SendMessage Contact, Message
        For y = 1 To 5000
            DoEvents
        Next
    Next
    txtStatus.Text = txtStatus.Text & "Done!" & vbCrLf & "Ready"
    Set oSkype = Nothing
End Sub

Private Sub lstContacts_Click()
lblVictim.Caption = lstContacts.List(lstContacts.ListIndex)
End Sub
