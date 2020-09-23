VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   """Trans"" - made by Jotaf98"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   3255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2640
      Picture         =   "Trans.frx":0000
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   6
      Top             =   120
      Width           =   330
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2370
      Top             =   3120
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "0"
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Text            =   "0"
      Top             =   1080
      Width           =   495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Movement"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   720
      Picture         =   "Trans.frx":0906
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "This picture is just to restore the image before repeating the effect. It will stay hidden at run time."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1800
      Left            =   120
      Picture         =   "Trans.frx":B3E8
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   121
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   ";)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   2450
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Y:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Sec. Image:"
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   180
      Width           =   495
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label8 
      Caption         =   "Made by Jotaf98 (João F. S. Henriques)"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   3015
   End
   Begin VB.Label Label9 
      Caption         =   "E-mail me at: jotaf98@hotmail.com"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label5 
      Caption         =   "Hit ""Start"" to see this beautiful effect in action"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label Label6 
      Caption         =   "Then, try changing the X and Y values, and checking and unchecking the ""Movement"" while running!"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'+--------------------------------------+
'+ (Check the .bas file)                |
'+--------------------------------------+

Private Sub Command1_Click()
    'Toggle timer
    Timer1.Enabled = Not Timer1.Enabled
    
    'Change the button's caption
    If Timer1.Enabled Then
        Command1.Caption = "Stop"
    Else
        Command1.Caption = "Start"
    End If
End Sub

Private Sub Timer1_Timer()
    'If "Movement" is checked, increase X and Y
    If Check1.Value Then
        Text1.Text = Int(Text1.Text) + 2
        Text2.Text = Int(Text2.Text) + 1
    End If
    
    'Repaint the original picture
    'For some reason, you must repaint the background
    'before using this function. You may have to
    'experiment a bit on this.
    Picture1.PaintPicture Picture3.Picture, 0, 0
    
    'Draw Picture2 on Picture1
    PaintTrans Picture1.hdc, Picture2.hdc, Text1.Text, Text2.Text, Picture2.ScaleWidth, Picture2.ScaleHeight
End Sub
