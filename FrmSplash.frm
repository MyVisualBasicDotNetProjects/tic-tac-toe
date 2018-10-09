VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H80000018&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   3915
   ClientTop       =   1470
   ClientWidth     =   7440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000003&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7200
      Begin VB.Image Image1 
         Height          =   1155
         Left            =   4080
         Picture         =   "Splash.frx":000C
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000003&
         Caption         =   "Can only be given by Biny for his best buddies!"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   3360
         Width           =   4455
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000003&
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   2535
      End
      Begin VB.Image imgLogo 
         Height          =   1185
         Left            =   360
         Picture         =   "Splash.frx":3123
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H80000003&
         Caption         =   "Date: 05 Jan 08 - 06 Jan 08"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   4
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H80000003&
         Caption         =   "All Rights Protected!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5400
         TabIndex        =   3
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H80000003&
         Caption         =   "Not for Commercial Purpose. For Home use only. Illegal Copying can get U into Serious Trouble!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "Version 1.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5385
         TabIndex        =   5
         Top             =   2640
         Width           =   1470
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "Programmed By: Biniam Asnake"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   6
         Top             =   2040
         Width           =   4905
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "Tic Tac Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2280
         TabIndex        =   8
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000003&
         Caption         =   "LicenseTo Biniam Asnake"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "Thug Life Softwares"
         BeginProperty Font 
            Name            =   "Copperplate Gothic Bold"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1920
         TabIndex        =   7
         Top             =   600
         Width           =   5040
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3120
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalName ~dulla^@204~ ~dulla^@204~ _Creatable = False
Attribute VB_Pred~dulla^@204~ rue
Attribute VB_Exposed = False

~dulla^@204~ it

Private Sub Form_KeyPress(KeyAs~dulla^@204~ r)
    
    Unload Me
    FrmTicTa~dulla^@204~ rue
    
End Sub

Private Sub For~dulla^@204~  lblVersion.Caption = "Version " & Ap~dulla^@204~  & App.Minor & "." & App.Revision
  ~dulla^@204~ ame.Caption = App.Title
End Sub

P~dulla^@204~ ame1_Click()

    Unload Me
    FrmTicTac.Visible = True

End Sub
