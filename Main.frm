VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jim's Mixer"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6045
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Show BPM Counter »"
      Height          =   375
      Left            =   4185
      TabIndex        =   18
      Top             =   3465
      Width           =   1770
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2790
      Picture         =   "Main.frx":030A
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   3735
      Width           =   480
   End
   Begin VB.CommandButton Show_File_Finder 
      Caption         =   "Show File Finder"
      Height          =   375
      Left            =   90
      TabIndex        =   11
      Top             =   3480
      Width           =   1770
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1590
      Top             =   4290
   End
   Begin VB.CheckBox Deck2_Mute 
      Caption         =   "Mute Deck 2"
      Height          =   195
      Left            =   4050
      TabIndex        =   10
      Top             =   2700
      Width           =   1410
   End
   Begin VB.CheckBox Deck1_Mute 
      Caption         =   "Mute Deck 1"
      Height          =   195
      Left            =   675
      TabIndex        =   9
      Top             =   2700
      Width           =   1410
   End
   Begin VB.HScrollBar Cross_Fader 
      Height          =   375
      Left            =   45
      Max             =   9640
      Min             =   -9640
      TabIndex        =   6
      Top             =   2970
      Width           =   5955
   End
   Begin VB.VScrollBar Deck2_Volume 
      Height          =   2310
      Left            =   3105
      Max             =   -9640
      TabIndex        =   5
      Top             =   90
      Width           =   240
   End
   Begin VB.VScrollBar Deck1_Volume 
      Height          =   2310
      Left            =   2700
      Max             =   -9640
      TabIndex        =   4
      Top             =   90
      Width           =   240
   End
   Begin VB.CommandButton Deck2_Open 
      Caption         =   "Open"
      Height          =   375
      Left            =   3510
      TabIndex        =   3
      Top             =   45
      Width           =   2445
   End
   Begin VB.CommandButton Deck1_Open 
      Caption         =   "Open"
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   2445
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   1080
      Top             =   4215
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MediaPlayerCtl.MediaPlayer Deck1 
      Height          =   645
      Left            =   45
      TabIndex        =   0
      Top             =   1710
      Width           =   2535
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -60
      WindowlessVideo =   0   'False
   End
   Begin MediaPlayerCtl.MediaPlayer Deck2 
      Height          =   645
      Left            =   3420
      TabIndex        =   1
      Top             =   1710
      Width           =   2535
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -60
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Deck1_File 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<NO FILE>"
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   180
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   585
      Width           =   2310
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "By James Compton"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1980
      TabIndex        =   17
      Top             =   3510
      Width           =   2070
   End
   Begin VB.Label Deck2_Remain 
      Alignment       =   1  'Right Justify
      Caption         =   "(00:00:00 remaining)"
      Height          =   240
      Left            =   4410
      TabIndex        =   15
      Top             =   2385
      Width           =   1500
   End
   Begin VB.Label Deck1_Remain 
      Alignment       =   1  'Right Justify
      Caption         =   "(00:00:00 remaining)"
      Height          =   240
      Left            =   1080
      TabIndex        =   14
      Top             =   2385
      Width           =   1500
   End
   Begin VB.Label Deck2_Time 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3555
      TabIndex        =   13
      Top             =   2385
      Width           =   1005
   End
   Begin VB.Label Deck1_Time 
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   12
      Top             =   2385
      Width           =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   3015
      X2              =   3015
      Y1              =   3420
      Y2              =   2880
   End
   Begin VB.Label Deck2_File 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<NO FILE>"
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   3555
      OLEDropMode     =   1  'Manual
      TabIndex        =   8
      Top             =   585
      Width           =   2310
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   1185
      Left            =   90
      Top             =   495
      Width           =   2490
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   1185
      Left            =   3465
      Top             =   495
      Width           =   2490
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------
'|     Jim's Mixer       |
'-------------------------
'This version has the following additions:
'
' - BPM Counter
' - Drag and drop ability (i.e. Windows Explorer)
' - File find form to help changing tracks
' - Time elapsed and time remaining display
'
'Anyway, hope you like it, and don't forget to drop me a line!


Private Sub Command1_Click()
On Error Resume Next
BPMCounter.Show
BPMCounter.Top = Form1.Top + Form1.Height / 2 - BPMCounter.Height / 2
BPMCounter.Left = Form1.Left + Form1.Width
End Sub

Private Sub Cross_Fader_Change()
' Right, this is where the crossfading is done, 2 lines of code! Simple!
If Cross_Fader.Value > 0 Then Deck1_Volume.Value = (9640 - Cross_Fader.Value) - 9640
If Cross_Fader.Value < 0 Then Deck2_Volume.Value = Cross_Fader.Value
End Sub

Private Sub Cross_Fader_Scroll()
Cross_Fader_Change
End Sub



Private Sub Deck1_File_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim File
On Error GoTo Error
For Each File In Data.Files
Deck1.filename = File
Deck1_File.Caption = Mid(File, InStrRevVB5(File, "\") + 1, Len(File))
Next File
Exit Sub

Error:
MsgBox "Not a valid file!", vbCritical, "Error"
End Sub

Private Sub Deck1_Mute_Click()
If Deck1_Mute.Value = 1 Then Deck1.Mute = True
If Deck1_Mute.Value = 0 Then Deck1.Mute = False
End Sub



Private Sub Deck1_Open_Click()
On Error GoTo Error
Dialog.CancelError = True 'This is to stop the track resetting when playing
                          'if cancel is pressed
Dialog.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
Dialog.ShowOpen

Deck1.filename = Dialog.filename
'Visual basic 6 users may want to get rid of the module...since it is a feature
'that is already on VB6 (InStrRev)
Deck1_File.Caption = Mid(Dialog.filename, InStrRevVB5(Dialog.filename, "\") + 1, Len(Dialog.filename))
If Dialog.filename = "" Then Deck1_File.Caption = "<NO FILE>"
Exit Sub

Error:
If Err.Number <> 32755 Then
MsgBox "Error loading file - " & Err.Number & " : " & Err.Description
Else
End If
End Sub

Private Sub Deck1_Volume_Change()
Deck1.Volume = Deck1_Volume.Value
End Sub

Private Sub Deck1_Volume_Scroll()
Deck1.Volume = Deck1_Volume.Value
End Sub

Private Sub Deck2_File_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim File
On Error GoTo Error
For Each File In Data.Files
Deck2.filename = File
Deck2_File.Caption = Mid(File, InStrRevVB5(File, "\") + 1, Len(File))
Next File
Exit Sub

Error:
MsgBox "Not a valid file!", vbCritical, "Error"
End Sub

Private Sub Deck2_Mute_Click()
If Deck2_Mute.Value = 1 Then Deck2.Mute = True
If Deck2_Mute.Value = 0 Then Deck2.Mute = False
End Sub

Private Sub Deck2_Open_Click()
On Error GoTo Error
Dialog.CancelError = True
Dialog.Filter = "All supported files |*.wav;*.wma;*.mp3;*.mid|MP3 Files *.mp3|*.mp3|Wave Files *.wav|*.wav|Midi Files *.mid|*.mid"
Dialog.ShowOpen

Deck2.filename = Dialog.filename
'Visual basic 6 users may want to get rid of the module...since it is a feature
'that is already on VB6 (InStrRev)
Deck2_File.Caption = Mid(Dialog.filename, InStrRevVB5(Dialog.filename, "\") + 1, Len(Dialog.filename))
If Dialog.filename = "" Then Deck2_File.Caption = "<NO FILE>"
Exit Sub

Error:
If Err.Number <> 32755 Then ' Cancel was pressed?
MsgBox "Error loading file - " & Err.Number & " : " & Err.Description
Else
End If
End Sub

Private Sub Deck2_Volume_Change()
Deck2.Volume = Deck2_Volume.Value
End Sub

Private Sub Deck2_Volume_Scroll()
Deck2.Volume = Deck2_Volume.Value
End Sub

Private Sub Form_Load()
Cross_Fader.Value = 0
Me.Caption = "Jim's Mixer v" & App.Major & "." & App.Minor & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form2
Unload BPMCounter
End Sub

Private Sub Show_File_Finder_Click()
On Error Resume Next
Form2.Show
Form2.Top = Form1.Top + Form1.Height
Form2.Left = Form1.Left
End Sub



Private Sub Timer1_Timer()
' Show time

' > DECK 1
On Error Resume Next
If Deck1.CurrentPosition > 0 Then
Deck1_Time.Caption = TimeSerial(0, 0, Int(Deck1.CurrentPosition))
End If
'Remaining time
Deck1_Remain.Caption = "(" & TimeSerial(0, 0, Int(Deck1.Duration) - Int(Deck1.CurrentPosition)) & " remaining)"

' > DECK 2
On Error Resume Next
If Deck2.CurrentPosition > 0 Then
Deck2_Time.Caption = TimeSerial(0, 0, Int(Deck2.CurrentPosition))
End If
'Remaining time
Deck2_Remain.Caption = "(" & TimeSerial(0, 0, Int(Deck2.Duration) - Int(Deck2.CurrentPosition)) & " remaining)"
' Turn mp3 name to red if 20 seconds or less left in track

'DECK 1

If Deck1.CurrentPosition >= (Deck1.Duration - 20) Then
Deck1_File.ForeColor = vbRed
Else
Deck1_File.ForeColor = vbWhite
End If

'DECK 2

If Deck2.CurrentPosition >= (Deck2.Duration - 20) Then
Deck2_File.ForeColor = vbRed
Else
Deck2_File.ForeColor = vbWhite
End If
End Sub
