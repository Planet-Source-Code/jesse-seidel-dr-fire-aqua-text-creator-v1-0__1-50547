VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "How to use this"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "< Back to First Step"
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3720
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next Step >"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   3720
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next Step >"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.PictureBox step4 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   7875
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   7935
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "YOUR DONE! "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   13
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "You can now do whatever editing you want with it (moving stuff around, change colors, ect.)"
         Height          =   855
         Left            =   3720
         TabIndex        =   12
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Image Image4 
         Height          =   2820
         Left            =   120
         Picture         =   "frmHelp.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3465
      End
      Begin VB.Label Label7 
         Caption         =   "Step 4:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox step3 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   7875
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   7935
      Begin VB.Label Label6 
         Caption         =   $"frmHelp.frx":D1E02
         Height          =   855
         Left            =   3720
         TabIndex        =   9
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Image Image3 
         Height          =   2820
         Left            =   120
         Picture         =   "frmHelp.frx":D1EA3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3465
      End
      Begin VB.Label Label5 
         Caption         =   "Step 3:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.PictureBox step2 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   6315
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   6375
      Begin VB.Image Image1 
         Height          =   660
         Left            =   120
         Picture         =   "frmHelp.frx":1A3CA5
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Label2 
         Caption         =   "Open up MSPaint"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "Step 2:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox step1 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin VB.Image Image2 
         Height          =   2505
         Left            =   840
         Picture         =   "frmHelp.frx":1A6FC7
         Top             =   720
         Width           =   3195
      End
      Begin VB.Label Label4 
         Caption         =   "Press the Print Screen button on your keyboard (See image)"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Step 1: "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next Step >"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   7935
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
step1.Visible = False
step2.Visible = True
step2.ZOrder 0
Command2.Visible = True
Command1.Visible = False
End Sub

Private Sub Picture2_Click()

End Sub

Private Sub Command2_Click()
step2.Visible = False
step3.Visible = True
step3.ZOrder 0
Command3.Visible = True
Command2.Visible = False
End Sub

Private Sub Command3_Click()
step3.Visible = False
step4.Visible = True
step4.ZOrder 0
Command4.Visible = True
Command3.Visible = False
End Sub

Private Sub Command4_Click()
step4.Visible = False
step1.Visible = True
step1.ZOrder 0
Command1.Visible = True
Command4.Visible = False
End Sub
