VERSION 5.00
Begin VB.Form FrmInstruction 
   Caption         =   "Game Instruction - Tic Tac Game By Biny"
   ClientHeight    =   5490
   ClientLeft      =   4395
   ClientTop       =   1680
   ClientWidth     =   5805
   Icon            =   "Instuction.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Instuction.frx":2A8B2
   ScaleHeight     =   5490
   ScaleWidth      =   5805
   Begin VB.CommandButton CmdPlay 
      BackColor       =   &H80000010&
      Caption         =   "&Start Game"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Game Instruction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
      Begin VB.Label Label3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Enjoy the Game!"
         BeginProperty Font 
            Name            =   "Pristina"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   4335
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"Instuction.frx":4280F
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "            Tic Tac Game           Progrmmed By: Biniam Asnake"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "FrmInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdPlay_Click()

FrmInstruction.Visible = False
FrmTicTac.Visible = True ~dulla^@204~ 