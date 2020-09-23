VERSION 5.00
Begin VB.Form BPMCounter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Jim's BPM Counter"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2790
   Icon            =   "BPMCounter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2205
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&BEAT"
      Default         =   -1  'True
      Height          =   735
      Left            =   960
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Reset"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "BPM counter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00101010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "BPMCounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private BPMArray(14) As Single
Private LastBPM
Private MaxBPMs
Private BPM As Integer
Private OldBPM As Integer

Private OutBuffer As String

Private Sub Command1_Click()
    For I = 0 To 14
        BPMArray(I) = 0
    Next
    MaxBPMs = 0
Label1.Caption = Format(OldBPM / 100, "##0.00")
Command2.SetFocus

End Sub

Private Sub Command2_Click()
On Error Resume Next
Static LastClick
If LastClick <> 0 And LastClick < Timer Then
    LastBPM = (LastBPM + 1) Mod 15
    If MaxBPMs < 15 Then MaxBPMs = MaxBPMs + 1
    BPMArray(LastBPM) = 60 / (Timer - LastClick)
    For I = 0 To 14
        cBPM = cBPM + BPMArray(I)
    Next
    Label1.Caption = Format(cBPM / MaxBPMs, "##0.00")
    BPM = (cBPM / MaxBPMs) * 100
End If

LastClick = Timer
End Sub

Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = o Then Command2.BackColor = &HFF
End Sub

Private Sub Command2_KeyUp(KeyCode As Integer, Shift As Integer)
    Command2.BackColor = &H8000000F
    If Shift = 0 And KeyCode <> 32 Then Command2_Click
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command2.BackColor = &HFF
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command2.BackColor = &H8000000F
End Sub

Sub ProcessStream(Stream As String)
OldBPM = Asc(Left$(Stream, 1)) + Asc(Right$(Stream, 1)) * 256
End Sub
Function GetLatest() As String
GetLatest = Chr$(Val(Label2.Caption)) + OutBuffer + Chr$(0)
OutBuffer = ""
End Function
Function InitializeDevice(id As Byte) As String
Dim TempID As Long
TempID = Device_Channel + Channel_ChannelID + Channel_Commands + Channel_BPM
'Debug.Print TempID
InitializeDevice = Chr$(TempID And &HFF&) + Chr$((TempID And &HFF00&) / &H100&) + Chr$((TempID And &HFF0000) / &H10000) + Chr$(0)

End Function

Private Sub Command3_Click()
Unload Me
End Sub


