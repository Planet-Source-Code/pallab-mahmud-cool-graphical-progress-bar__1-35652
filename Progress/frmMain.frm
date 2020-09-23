VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Progress Bar !"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmLoop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3855
      Top             =   690
   End
   Begin VB.CommandButton cmLoop 
      Caption         =   "Start"
      Height          =   420
      Left            =   1920
      TabIndex        =   9
      Top             =   945
      Width           =   1185
   End
   Begin VB.CommandButton cmStep 
      Caption         =   "i = i+1"
      Height          =   420
      Left            =   105
      TabIndex        =   8
      Top             =   945
      Width           =   1185
   End
   Begin VB.PictureBox picBase 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   240
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   105
      ScaleWidth      =   2700
      TabIndex        =   5
      Top             =   780
      Width           =   2730
      Begin VB.PictureBox picImg 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         Picture         =   "frmMain.frx":1122
         ScaleHeight     =   13
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   6
         Top             =   420
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label lblInf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   5.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   120
         Left            =   1335
         TabIndex        =   7
         Top             =   -15
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdStep 
      Caption         =   "i = i+1"
      Height          =   420
      Left            =   90
      TabIndex        =   4
      Top             =   330
      Width           =   1185
   End
   Begin VB.Timer tmrLoop 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3885
      Top             =   105
   End
   Begin VB.CommandButton cmdLoop 
      Caption         =   "Start"
      Height          =   420
      Left            =   1920
      TabIndex        =   2
      Top             =   330
      Width           =   1185
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   75
      ScaleHeight     =   195
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   75
      Width           =   3030
      Begin VB.PictureBox picBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         ScaleHeight     =   13
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   0
         Width           =   345
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------'
'|Progress bar [cool!!]                           |'
'|------------------------------------------------|'
'|Written by Pallab Mahmud                        |'
'|© Copyright 2001 by Pallab Mahmud               |'
'|email: pallmahmud@yahoo.com                     |'
'|                                                |'
'|This sample code is a FREEWARE. Use it in your  |'
'|own prointPercentect as it fits You but do not re-sale   |'
'|this code or destroy the original authors name. |'
'|                                                |'
'|Warning: No warranty is provided with this set  |'
'|of code so use it in your own risk. The author  |'
'|is not responsible for the Damage caused by     |'
'|this code.                                      |'
'--------------------------------------------------'
'--------------------------------------------------'
'Comments:This is a cool progress bar.You can change
'it base and bar picture whatever you want.It uses
'only one api call.I think it is great.What do you think?
'Hey,listen i am new in programing and i am 14 years old
'So,don't mind and Please please........vote for me
'--------------------------------------------------'
Option Explicit
Dim intComplete, intComp
Private Sub cmdLoop_Click()
    tmrLoop.Enabled = True
End Sub
Private Sub cmdStep_Click()
    intComplete = intComplete + 1
    curPercent intComplete, picMain, picBar
    lblInfo = intPercent
End Sub
Private Sub cmLoop_Click()
    tmLoop.Enabled = True
End Sub
Private Sub cmStep_Click()
    intComp = intComp + 1
    curPercent intComp, picBase, picImg
    lblInf = intPercent
End Sub
Private Sub tmLoop_Timer()
    intComp = intComp + 1
    curPercent intComp, picBase, picImg
    lblInf = intPercent
    If intPercent = 100 & "%" Then tmLoop.Enabled = False
End Sub
Private Sub tmrLoop_Timer()
    intComplete = intComplete + 1
    curPercent intComplete, picMain, picBar
    lblInfo = intPercent
    If intPercent = 100 & "%" Then tmrLoop.Enabled = False
End Sub
