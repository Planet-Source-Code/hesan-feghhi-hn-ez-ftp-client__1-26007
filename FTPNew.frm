VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form FTP 
   Caption         =   "HN FTP"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   Icon            =   "FTPNew.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Web Browser"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   10095
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   4695
      Left            =   120
      TabIndex        =   27
      Top             =   1320
      Visible         =   0   'False
      Width           =   10095
      ExtentX         =   17806
      ExtentY         =   8281
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Local"
      Height          =   255
      Left            =   3600
      TabIndex        =   25
      Top             =   1440
      Width           =   6495
   End
   Begin VB.Frame Frame3 
      Height          =   4335
      Left            =   3480
      TabIndex        =   19
      Top             =   1440
      Width           =   6735
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   615
         Width           =   3135
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   120
         TabIndex        =   23
         Top             =   975
         Width           =   3135
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H80000001&
         DragIcon        =   "FTPNew.frx":0442
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3450
         Left            =   3240
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         System          =   -1  'True
         TabIndex        =   22
         Top             =   660
         Width           =   3375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "File List"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3240
         TabIndex        =   21
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Directory/Drive List"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Remote"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      Height          =   4335
      Left            =   0
      TabIndex        =   13
      Top             =   1440
      Width           =   3375
      Begin VB.CommandButton Command8 
         BackColor       =   &H80000010&
         Caption         =   "Create"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   26
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000010&
         Caption         =   "Directory List"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2415
      End
      Begin VB.ListBox List1 
         BackColor       =   &H80000001&
         Height          =   1620
         Left            =   120
         TabIndex        =   16
         Top             =   525
         Width           =   3135
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000010&
         Caption         =   "File List"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   3135
      End
      Begin VB.ListBox List2 
         BackColor       =   &H80000001&
         Height          =   1815
         Left            =   120
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   14
         Top             =   2325
         Width           =   3135
      End
   End
   Begin VB.TextBox txtDir 
      Height          =   225
      Left            =   9840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   360
      TabIndex        =   4
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton Command1 
         Caption         =   "Quit"
         CausesValidation=   0   'False
         Height          =   375
         Index           =   3
         Left            =   8040
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Refresh"
         Height          =   375
         Index           =   2
         Left            =   6720
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Logoff"
         Height          =   375
         Index           =   1
         Left            =   8040
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Logon"
         Height          =   375
         Index           =   0
         Left            =   6720
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Text            =   "ftp://ftp.geocities.com"
         Top             =   240
         Width           =   5295
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   4320
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Server URL:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "User Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
   End
   Begin InetCtlsObjects.Inet ITC1 
      Left            =   9240
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblStatus 
      Caption         =   "Request Complete"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5880
      Width           =   10215
   End
End
Attribute VB_Name = "FTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
  Select Case Index
      Case 0  'Logon
          txtDir.Text = ""
          Call Logon
      Case 1  'Logoff
          txtDir.Text = ""
          ITC1.Cancel
          Call SetCmdButtonState(False)
      Case 2  'Refresh
          Call RefreshDirList
      Case 3  'Quit
          ITC1.Cancel
          End
  End Select
End Sub

Private Sub Command3_Click()
 If Frame2.Height = 4335 Then Frame2.Height = 350 Else Frame2.Height = 4335
End Sub

Private Sub Command4_Click()
 If Frame3.Height = 4335 Then Frame3.Height = 350 Else Frame3.Height = 4335
End Sub

Private Sub Command8_Click()
 Dim A$
 A = InputBox("Enetr the name of the directory you want to create", "Create Directory", "newdir")
 If StrPtr(A) <> 0 Then
  SendCommand "MKDIR " + A
 End If
End Sub

Private Sub Command9_Click()
 WB.Visible = False
 Command9.Value = False
End Sub

Private Sub Dir1_Change()
 File1 = Dir1
End Sub

Private Sub Drive1_Change()
 Dir1 = Drive1
End Sub

Private Sub File1_DblClick()
    WB.Navigate File1.Path + "\" + File1.FileName
    WB.Visible = True
    Command9.Visible = True
End Sub

Private Sub Form_Load()
    Call SetCmdButtonState(False)
End Sub

Private Sub File1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 ChDir File1.Path
 GetRemoteFile
 File1.Refresh
End Sub

Private Sub ITC1_StateChanged(ByVal State As Integer)
 On Error Resume Next
 Dim Data1 As String, Data2 As String, R As String
 Select Case State
     Case icResolvingHost
         lblStatus.Caption = "Looking up host Computer IP"
         lblStatus.MousePointer = 13
     Case icHostResolved
         lblStatus.Caption = "IP Address found"
         lblStatus.MousePointer = 13
     Case icConnecting
         lblStatus.Caption = "Connecting to host Computer"
         lblStatus.MousePointer = 13
     Case icConnected
         lblStatus.Caption = "Connected to host Computer"
         lblStatus.MousePointer = 13
     Case icRequesting
         lblStatus.Caption = "Sending a request to host Computer"
         lblStatus.MousePointer = 13
     Case icRequestSent
         lblStatus.Caption = "Request Sent"
         lblStatus.MousePointer = 13
     Case icReceivingResponse
         lblStatus.Caption = "Receiving Response from host Computer"
         lblStatus.MousePointer = 13
     Case icResponseReceived
         lblStatus.Caption = "Response Received"
         lblStatus.MousePointer = 13
     Case icDisconnecting
         lblStatus.Caption = "Disconnecting from host Computer"
         lblStatus.MousePointer = 13
     Case icDisconnected
         lblStatus.Caption = "Disconnected from host Computer"
         lblStatus.MousePointer = 13
     Case icError
         lblStatus.Caption = "Error " + Str$(ITC1.ResponseCode) + " " + ITC1.ResponseInfo
         lblStatus.MousePointer = 13
     Case icResponseCompleted
         Dim LastData1 As String
         lblStatus.Caption = "Request Completed Successfully"
         Do While True
             Data1 = ITC1.GetChunk(512, icString)
             If Len(Data1) = 0 Then Exit Do
             DoEvents
             Data2 = Data2 + Data1
         Loop
         txtDir.Text = Data2
         List1.Clear
         List2.Clear
         Open "C:\RemoteTEMP.tmp" For Output As #1
             Print #1, txtDir.Text
             Close #1
         Open "C:\RemoteTEMP.tmp" For Input As #1
         List1.AddItem ".."
         While Not EOF(1)
             Line Input #1, R$
             If Right$(R$, 1) = "/" Then
                 List1.AddItem Left$(R$, Len(R$) - 1)
             Else
                 List2.AddItem R$
             End If
         Wend
         Close #1
         lblStatus.MousePointer = vbDefault
 End Select
End Sub

Public Sub Logon()
    On Error GoTo ErrorHandler
    If txtURL = "" Or txtPassword = "" Then
        MsgBox "You must specify a URL and Password"
        Exit Sub
    End If
    Command1(0).Enabled = False
    ITC1.Protocol = icFTP
    ITC1.URL = txtURL.Text
    If txtUserName.Text = "" Then
        ITC1.UserName = "anonymous"
    Else
        ITC1.UserName = txtUserName.Text
    End If
    ITC1.Password = txtPassword.Text
    ITC1.Execute , "DIR"
    Call SetCmdButtonState(True)
    Exit Sub
ErrorHandler:
    If Err = 35754 Then
        MsgBox "Cannot Connect to remote host"
    Else
        MsgBox Err.Description
    End If
    Call SetCmdButtonState(False)
    ITC1.Cancel
End Sub

Public Sub PutLocalFile()
    Dim RemoteFileName As String, Cmd As String
    If File1.ListIndex > -1 Then RemoteFileName = File1.List(File1.ListIndex)
    Cmd = "PUT " + RemoteFileName + " " + File1.List(File1.ListIndex)
    Do
        DoEvents
    Loop Until ITCReady(False)
    Call SendCommand(Cmd)
    Exit Sub
End Sub

Public Sub GetRemoteFile()
    Dim LocalFileName As String, Cmd As String
    On Error GoTo ErrorHandler
    LocalFileName = List2.List(List2.ListIndex)
    Cmd = "GET " + List2.List(List2.ListIndex) + " " + LocalFileName
    If ITCReady(True) Then
        Call SendCommand(Cmd)
    End If
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
End Sub

Public Sub ChangeDir()
    Dim Cmd As String
    
    If List1.ListIndex < 0 Then
     If List1.ListCount > 0 Then List1.ListIndex = 0 Else Exit Sub
    End If
    Cmd = "CD " + List1.List(List1.ListIndex)
    If ITCReady(True) Then
       Call SendCommand(Cmd)
       Do
           DoEvents
       Loop Until ITCReady(False)
       txtDir.Text = ""
       Call SendCommand("DIR")
   End If
End Sub

Public Sub SetCmdButtonState(Loggedon As Boolean)
    Dim X As Integer
    If Loggedon Then
        Command1(0).Enabled = False  'Logon
        Command1(1).Enabled = True   'Logoff
        Command1(2).Enabled = True   'Refresh
    Else
        Command1(0).Enabled = True   'Logon
        Command1(1).Enabled = False  'Logoff
        Command1(2).Enabled = False  'Refresh
   End If
End Sub

Public Sub RefreshDirList()
    If ITCReady(True) Then
        txtDir.Text = ""
        Call SendCommand("DIR")
    End If
End Sub

Public Sub SendCommand(Cmd As String)
    On Error GoTo ErrorHandler
    ITC1.Execute , Cmd
    Exit Sub
ErrorHandler:
    MsgBox Err.Description
    Resume Next
End Sub

Public Function ITCReady(Message As Boolean) As Boolean
    Dim msg As String
    If ITC1.StillExecuting Then
        If Message Then
            msg = "The Program has Not Finished" + Chr$(13)
            msg = msg = "executing your Last request." + Chr$(13)
            msg = msg + "Plase Wait then try again later."
            MsgBox msg
        End If
        ITCReady = False
    Else
        ITCReady = True
    End If
End Function

Private Sub List1_DblClick()
 ChangeDir
End Sub

Private Sub List1_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim A
    If KeyCode = 46 Then
        A = MsgBox("Do you really want to remove this directory from the remote computer?", vbYesNo, "HN FTP")
        If A = vbYes Then
            If ITCReady(True) Then
                If List1.ListIndex > -1 Then ITC1.Execute , "RMDIR " + List1.List(List1.ListIndex)
            End If
        End If
    End If
End Sub

Private Sub List2_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim A
    If KeyCode = 46 Then
        A = MsgBox("Do you really want to remove this file from the remote computer?", vbYesNo, "HN FTP")
        If A = vbYes Then
            If ITCReady(True) Then
                If List2.ListIndex > -1 Then ITC1.Execute , "DELETE " + List2.List(List2.ListIndex)
            End If
        End If
    End If
End Sub

Private Sub List2_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
 PutLocalFile
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)

End Sub

