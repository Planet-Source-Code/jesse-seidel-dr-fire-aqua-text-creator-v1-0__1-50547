VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aqua Letter Maker - by Dr. Fire"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3240
      Top             =   3480
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Auto refresh"
      Height          =   255
      Left            =   1680
      TabIndex        =   28
      Top             =   4200
      Width           =   4455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   5880
      TabIndex        =   27
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6240
      TabIndex        =   25
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   495
      Left            =   840
      TabIndex        =   24
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Help"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   4080
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   4920
      TabIndex        =   22
      Top             =   3360
      Width           =   735
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   21
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3960
      MaxLength       =   1
      TabIndex        =   19
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   3480
      MaxLength       =   1
      TabIndex        =   17
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3000
      MaxLength       =   1
      TabIndex        =   15
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   13
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   11
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   9
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   7
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   600
      MaxLength       =   1
      TabIndex        =   5
      Top             =   3600
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MaxLength       =   1
      TabIndex        =   3
      Top             =   3600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   2715
      ScaleWidth      =   6675
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      Begin VB.Image Image1 
         Height          =   1200
         Left            =   -1080
         Picture         =   "Form1.frx":15F942
         Top             =   -1080
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Image L10 
         Height          =   1200
         Left            =   5400
         Picture         =   "Form1.frx":164844
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Image L9 
         Height          =   1200
         Left            =   4080
         Picture         =   "Form1.frx":169746
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Image L8 
         Height          =   1200
         Left            =   2760
         Picture         =   "Form1.frx":16E648
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Image L7 
         Height          =   1200
         Left            =   1440
         Picture         =   "Form1.frx":17354A
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Image L6 
         Height          =   1200
         Left            =   120
         Picture         =   "Form1.frx":17844C
         Top             =   1440
         Width           =   1260
      End
      Begin VB.Image L5 
         Height          =   1200
         Left            =   5400
         Picture         =   "Form1.frx":17D34E
         Top             =   120
         Width           =   1260
      End
      Begin VB.Image L4 
         Height          =   1200
         Left            =   4080
         Picture         =   "Form1.frx":182250
         Top             =   120
         Width           =   1260
      End
      Begin VB.Image L3 
         Height          =   1200
         Left            =   2760
         Picture         =   "Form1.frx":187152
         Top             =   120
         Width           =   1260
      End
      Begin VB.Image L2 
         Height          =   1200
         Left            =   1440
         Picture         =   "Form1.frx":18C054
         Top             =   120
         Width           =   1260
      End
      Begin VB.Image L1 
         Height          =   1200
         Left            =   120
         Picture         =   "Form1.frx":190F56
         Top             =   120
         Width           =   1260
      End
   End
   Begin VB.Line Line2 
      X1              =   5760
      X2              =   5760
      Y1              =   3840
      Y2              =   3360
   End
   Begin VB.Label Label12 
      Caption         =   "Note: Dont mind the m's, I just put them there for max letter size"
      Height          =   255
      Left            =   2400
      TabIndex        =   26
      Top             =   120
      Width           =   4455
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6840
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label11 
      Caption         =   "10"
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "9"
      Height          =   255
      Left            =   3960
      TabIndex        =   18
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label9 
      Caption         =   "8"
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "7"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "6"
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label6 
      Caption         =   "5"
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "4"
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "3"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "2"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "Letters:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case Text1.Text
    Case "a"
    L1.Picture = ResLetters.La.Picture
    
    Case "b"
    L1.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L1.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L1.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L1.Picture = ResLetters.Le.Picture
    
    Case "f"
    L1.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L1.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L1.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L1.Picture = ResLetters.Li.Picture
    
    Case "j"
    L1.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L1.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L1.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L1.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L1.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L1.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L1.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L1.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L1.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L1.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L1.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L1.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L1.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L1.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L1.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L1.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L1.Picture = ResLetters.Lz.Picture
    
    Case ""
    L1.Picture = Image1.Picture
    
    Case " "
    L1.Picture = Image1.Picture
End Select
    
Select Case Text2.Text
    Case "a"
    L2.Picture = ResLetters.La.Picture
    
    Case "b"
    L2.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L2.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L2.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L2.Picture = ResLetters.Le.Picture
    
    Case "f"
    L2.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L2.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L2.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L2.Picture = ResLetters.Li.Picture
    
    Case "j"
    L2.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L2.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L2.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L2.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L2.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L2.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L2.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L2.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L2.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L2.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L2.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L2.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L2.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L2.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L2.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L2.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L2.Picture = ResLetters.Lz.Picture
    
    Case ""
    L2.Picture = Image1.Picture
    
    Case " "
    L2.Picture = Image1.Picture
End Select

Select Case Text3.Text
    Case "a"
    L3.Picture = ResLetters.La.Picture
    
    Case "b"
    L3.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L3.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L3.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L3.Picture = ResLetters.Le.Picture
    
    Case "f"
    L3.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L3.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L3.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L3.Picture = ResLetters.Li.Picture
    
    Case "j"
    L3.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L3.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L3.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L3.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L3.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L3.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L3.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L3.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L3.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L3.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L3.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L3.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L3.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L3.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L3.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L3.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L3.Picture = ResLetters.Lz.Picture
    
    Case ""
    L3.Picture = Image1.Picture
    
    Case " "
    L3.Picture = Image1.Picture
End Select

Select Case Text4.Text
    Case "a"
    L4.Picture = ResLetters.La.Picture
    
    Case "b"
    L4.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L4.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L4.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L4.Picture = ResLetters.Le.Picture
    
    Case "f"
    L4.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L4.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L4.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L4.Picture = ResLetters.Li.Picture
    
    Case "j"
    L4.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L4.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L4.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L4.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L4.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L4.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L4.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L4.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L4.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L4.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L4.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L4.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L4.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L4.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L4.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L4.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L4.Picture = ResLetters.Lz.Picture
    
    Case ""
    L4.Picture = Image1.Picture
    
    Case " "
    L4.Picture = Image1.Picture
End Select

Select Case Text5.Text
    Case "a"
    L5.Picture = ResLetters.La.Picture
    
    Case "b"
    L5.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L5.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L5.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L5.Picture = ResLetters.Le.Picture
    
    Case "f"
    L5.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L5.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L5.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L5.Picture = ResLetters.Li.Picture
    
    Case "j"
    L5.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L5.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L5.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L5.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L5.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L5.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L5.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L5.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L5.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L5.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L5.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L5.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L5.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L5.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L5.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L5.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L5.Picture = ResLetters.Lz.Picture
    
    Case ""
    L5.Picture = Image1.Picture
    
    Case " "
    L5.Picture = Image1.Picture
End Select

Select Case Text6.Text
    Case "a"
    L6.Picture = ResLetters.La.Picture
    
    Case "b"
    L6.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L6.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L6.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L6.Picture = ResLetters.Le.Picture
    
    Case "f"
    L6.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L6.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L6.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L6.Picture = ResLetters.Li.Picture
    
    Case "j"
    L6.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L6.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L6.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L6.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L6.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L6.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L6.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L6.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L6.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L6.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L6.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L6.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L6.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L6.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L6.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L6.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L6.Picture = ResLetters.Lz.Picture
    
    Case ""
    L6.Picture = Image1.Picture
    
    Case " "
    L6.Picture = Image1.Picture
End Select

Select Case Text7.Text
    Case "a"
    L7.Picture = ResLetters.La.Picture
    
    Case "b"
    L7.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L7.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L7.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L7.Picture = ResLetters.Le.Picture
    
    Case "f"
    L7.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L7.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L7.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L7.Picture = ResLetters.Li.Picture
    
    Case "j"
    L7.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L7.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L7.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L7.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L7.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L7.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L7.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L7.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L7.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L7.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L7.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L7.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L7.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L7.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L7.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L7.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L7.Picture = ResLetters.Lz.Picture
    
    Case ""
    L7.Picture = Image1.Picture
    
    Case " "
    L7.Picture = Image1.Picture
End Select

Select Case Text8.Text
    Case "a"
    L8.Picture = ResLetters.La.Picture
    
    Case "b"
    L8.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L8.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L8.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L8.Picture = ResLetters.Le.Picture
    
    Case "f"
    L8.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L8.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L8.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L8.Picture = ResLetters.Li.Picture
    
    Case "j"
    L8.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L8.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L8.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L8.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L8.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L8.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L8.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L8.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L8.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L8.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L8.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L8.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L8.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L8.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L8.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L8.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L8.Picture = ResLetters.Lz.Picture
    
    Case ""
    L8.Picture = Image1.Picture
    
    Case " "
    L8.Picture = Image1.Picture
End Select

Select Case Text9.Text
    Case "a"
    L9.Picture = ResLetters.La.Picture
    
    Case "b"
    L9.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L9.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L9.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L9.Picture = ResLetters.Le.Picture
    
    Case "f"
    L9.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L9.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L9.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L9.Picture = ResLetters.Li.Picture
    
    Case "j"
    L9.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L9.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L9.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L9.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L9.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L9.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L9.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L9.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L9.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L9.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L9.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L9.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L9.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L9.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L9.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L9.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L9.Picture = ResLetters.Lz.Picture
    
    Case ""
    L9.Picture = Image1.Picture
    
    Case " "
    L9.Picture = Image1.Picture
End Select

Select Case Text10.Text
    Case "a"
    L10.Picture = ResLetters.La.Picture
    
    Case "b"
    L10.Picture = ResLetters.Lb.Picture
    
    Case "c"
    L10.Picture = ResLetters.Lc.Picture
    
    Case "d"
    L10.Picture = ResLetters.Ld.Picture
    
    Case "e"
    L10.Picture = ResLetters.Le.Picture
    
    Case "f"
    L10.Picture = ResLetters.Lf.Picture
    
    Case "g"
    L10.Picture = ResLetters.Lg.Picture
    
    Case "h"
    L10.Picture = ResLetters.Lh.Picture
    
    Case "i"
    L10.Picture = ResLetters.Li.Picture
    
    Case "j"
    L10.Picture = ResLetters.Lj.Picture
    
    Case "k"
    L10.Picture = ResLetters.Lk.Picture
    
    Case "l"
    L10.Picture = ResLetters.Ll.Picture
    
    Case "m"
    L10.Picture = ResLetters.Lm.Picture
    
    Case "n"
    L10.Picture = ResLetters.Ln.Picture
    
    Case "o"
    L10.Picture = ResLetters.Lo.Picture
    
    Case "p"
    L10.Picture = ResLetters.Lp.Picture
    
    Case "q"
    L10.Picture = ResLetters.Lq.Picture
    
    Case "r"
    L10.Picture = ResLetters.Lr.Picture
    
    Case "s"
    L10.Picture = ResLetters.Ls.Picture
    
    Case "t"
    L10.Picture = ResLetters.Lt.Picture
    
    Case "u"
    L10.Picture = ResLetters.Lu.Picture
    
    Case "v"
    L10.Picture = ResLetters.Lv.Picture
    
    Case "w"
    L10.Picture = ResLetters.Lw.Picture
    
    Case "x"
    L10.Picture = ResLetters.Lx.Picture
    
    Case "y"
    L10.Picture = ResLetters.Ly.Picture
    
    Case "z"
    L10.Picture = ResLetters.Lz.Picture
    
    Case ""
    L10.Picture = Image1.Picture
    
    Case " "
    L10.Picture = Image1.Picture
End Select
End Sub

Private Sub Command2_Click()
frmHelp.Visible = True
End Sub

Private Sub Command3_Click()
MsgBox "Made by: Jesse Seidel, A.K.A Dr. Fire"

End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
L1.Picture = Image1.Picture
L2.Picture = Image1.Picture
L3.Picture = Image1.Picture
L4.Picture = Image1.Picture
L5.Picture = Image1.Picture
L6.Picture = Image1.Picture
L7.Picture = Image1.Picture
L8.Picture = Image1.Picture
L9.Picture = Image1.Picture
L10.Picture = Image1.Picture
End Sub

Private Sub Text1_Change()
Text2.Text = ""
Text2.SetFocus
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Text10_Change()
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text10_Click()
Text10.Text = ""
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Text9.SetFocus
Text9.Text = ""
End If
End Sub

Private Sub Text2_Change()
Text3.Text = ""
Text3.SetFocus
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text2_Click()
Text2.Text = ""
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Text1.SetFocus
Text1.Text = ""
End If
End Sub

Private Sub Text3_Change()
Text4.Text = ""
Text4.SetFocus
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text3_Click()
Text3.Text = ""
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Text2.SetFocus
Text2.Text = ""
End If
End Sub

Private Sub Text4_Change()
Text5.Text = ""
Text5.SetFocus
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text4_Click()
Text4.Text = ""
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Text3.SetFocus
Text3.Text = ""
End If
End Sub

Private Sub Text5_Change()
Text6.Text = ""
Text6.SetFocus
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text5_Click()
Text5.Text = ""
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Text4.SetFocus
Text4.Text = ""
End If
End Sub

Private Sub Text6_Change()
Text7.Text = ""
Text7.SetFocus
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text6_Click()
Text6.Text = ""
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Text5.SetFocus
Text5.Text = ""
End If
End Sub

Private Sub Text7_Change()
Text8.Text = ""
Text8.SetFocus
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text7_Click()
Text7.Text = ""
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Text6.SetFocus
Text6.Text = ""
End If
End Sub

Private Sub Text8_Change()
Text9.Text = ""
Text9.SetFocus
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text8_Click()
Text8.Text = ""
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Text7.SetFocus
Text7.Text = ""
End If
End Sub

Private Sub Text9_Change()
Text10.Text = ""
Text10.SetFocus
If Check1.Value = 1 Then
Command1_Click
Else

End If
End Sub

Private Sub Text9_Click()
Text9.Text = ""
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
Text8.SetFocus
Text8.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
Text1.SetFocus
Timer1.Enabled = False
End Sub
