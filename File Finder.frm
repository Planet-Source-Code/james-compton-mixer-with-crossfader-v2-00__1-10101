VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Finder"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6045
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      DragIcon        =   "File Finder.frx":0000
      Height          =   4770
      Left            =   2430
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      Pattern         =   "*.wav;*.wma;*.mp3;*.mid"
      TabIndex        =   2
      Top             =   315
      Width           =   3570
   End
   Begin VB.DirListBox Dir1 
      Height          =   4815
      Left            =   0
      TabIndex        =   1
      Top             =   315
      Width           =   2400
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6000
   End
   Begin VB.Label Label1 
      Caption         =   "Simply click once on the file to select it, then drag it over to one of the two decks, where it shows the filename."
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   5220
      Width           =   5910
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
On Error GoTo Error
File1.Path = Dir1.Path
Exit Sub

Error:
MsgBox "Directory unavailable", vbCritical, "Error"
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo Error
Dir1.Path = Drive1.Drive
Exit Sub

Error:
MsgBox "Drive unavailable", vbCritical, "Error"
Drive1.Drive = Dir1.Path
End Sub

Private Sub File1_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Dim File
Dim Temp
For Each File In Data.Files
Temp = File
Next File
End Sub
