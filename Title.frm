VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3510
      Top             =   1125
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   2025
      Picture         =   "Title.frx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   1215
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   765
      TabIndex        =   3
      Top             =   810
      Width           =   3030
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "By James Compton - 2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   810
      TabIndex        =   2
      Top             =   1980
      Width           =   3030
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Jim's Mixer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   810
      TabIndex        =   0
      Top             =   270
      Width           =   3030
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      BorderWidth     =   4
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2430
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      BorderWidth     =   4
      X1              =   0
      X2              =   4590
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderWidth     =   4
      X1              =   4635
      X2              =   4635
      Y1              =   45
      Y2              =   3105
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   4
      X1              =   0
      X2              =   4650
      Y1              =   2475
      Y2              =   2475
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Label3.Caption = "Version " & App.Major & "." & App.Minor & App.Revision
End Sub

Private Sub Timer1_Timer()
If Timer1.Interval = 2000 Then
Form1.Show
Unload Me
End If
End Sub
