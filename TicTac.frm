VERSION 5.00
Begin VB.Form FrmTicTac 
   BackColor       =   &H80000003&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Game     Programmed By Biny A."
   ClientHeight    =   5085
   ClientLeft      =   4755
   ClientTop       =   2085
   ClientWidth     =   4845
   Icon            =   "TicTac.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TicTac.frx":2A8B2
   ScaleHeight     =   5085
   ScaleWidth      =   4845
   Begin VB.CommandButton CmdAbout 
      Caption         =   "&About ..."
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   1
      Top             =   4785
      Width           =   1335
   End
   Begin VB.CommandButton CmdAgain 
      BackColor       =   &H80000012&
      Caption         =   "&Play Again"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Copperplate Gothic Bold"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   7
      Visible         =   0   'False
      X1              =   720
      X2              =   3840
      Y1              =   3480
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   6
      Visible         =   0   'False
      X1              =   720
      X2              =   3840
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   5
      Visible         =   0   'False
      X1              =   720
      X2              =   3840
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   4
      Visible         =   0   'False
      X1              =   720
      X2              =   3840
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   3
      Visible         =   0   'False
      X1              =   3480
      X2              =   3480
      Y1              =   480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   2
      Visible         =   0   'False
      X1              =   2280
      X2              =   2280
      Y1              =   480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   1
      Visible         =   0   'False
      X1              =   1200
      X2              =   1200
      Y1              =   480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      Index           =   0
      Visible         =   0   'False
      X1              =   720
      X2              =   3840
      Y1              =   480
      Y2              =   3480
   End
   Begin VB.Shape Cir 
      Height          =   615
      Index           =   8
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Rec 
      Height          =   615
      Index           =   8
      Left            =   3120
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Cir 
      Height          =   615
      Index           =   7
      Left            =   1920
      Shape           =   2  'Oval
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Rec 
      Height          =   615
      Index           =   7
      Left            =   1920
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Cir 
      Height          =   615
      Index           =   6
      Left            =   840
      Shape           =   2  'Oval
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Rec 
      Height          =   615
      Index           =   6
      Left            =   840
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Cir 
      Height          =   615
      Index           =   5
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Rec 
      Height          =   615
      Index           =   5
      Left            =   3120
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Cir 
      Height          =   615
      Index           =   4
      Left            =   1920
      Shape           =   2  'Oval
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Rec 
      Height          =   615
      Index           =   4
      Left            =   1920
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Cir 
      Height          =   615
      Index           =   3
      Left            =   840
      Shape           =   2  'Oval
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Rec 
      Height          =   615
      Index           =   3
      Left            =   840
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Cir 
      Height          =   615
      Index           =   2
      Left            =   3120
      Shape           =   2  'Oval
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Rec 
      Height          =   615
      Index           =   2
      Left            =   3120
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Cir 
      Height          =   615
      Index           =   1
      Left            =   1920
      Shape           =   2  'Oval
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Rec 
      Height          =   615
      Index           =   1
      Left            =   1920
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Rec 
      Height          =   615
      Index           =   0
      Left            =   840
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Shape Cir 
      Height          =   615
      Index           =   0
      Left            =   840
      Shape           =   2  'Oval
      Top             =   600
      Visible        ~dulla^@204~ ~dulla^@204~ th           =   615
   End
   Begi~dulla^@204~ age1 
      Height          =   855~dulla^@204~           =   8
      Left          ~dulla^@204~      Top             =   2640
      ~dulla^@204~    =   855
   End
   Begin VB.Image~dulla^@204~    Height          =   855
      Ind~dulla^@204~ =   7
      Left            =   1800~dulla^@204~            =   2640
      Width     ~dulla^@204~ 
   End
   Begin VB.Image Image1 
~dulla^@204~          =   855
      Index        ~dulla^@204~    Left            =   720
      Top~dulla^@204~ =   2640
      Width           =   8~dulla^@204~   Begin VB.Image Image1 
      Heigh~dulla^@204~   855
      Index           =   5
 ~dulla^@204~         =   3000
      Top          ~dulla^@204~       Width           =   855
   End~dulla^@204~ .Image Image1 
      Height         ~dulla^@204~    Index           =   4
      Left ~dulla^@204~   1800
      Top             =   156~dulla^@204~ h           =   855
   End
   Begin~dulla^@204~ ge1 
      Height          =   855
~dulla^@204~          =   3
      Left           ~dulla^@204~    Top             =   1560
      Wi~dulla^@204~  =   855
   End
   Begin VB.Image I~dulla^@204~  Height          =   855
      Index~dulla^@204~   2
      Left            =   3000
~dulla^@204~          =   480
      Width        ~dulla^@204~   End
   Begin VB.Image Image1 
   ~dulla^@204~       =   855
      Index           ~dulla^@204~ Left            =   1800
      Top  ~dulla^@204~   480
      Width           =   855~dulla^@204~ egin VB.Image Image1 
      Height  ~dulla^@204~ 55
      Index           =   0
    ~dulla^@204~      =   720
      Top             =~dulla^@204~  Width           =   855
   End
   ~dulla^@204~  Line5 
      X1              =   72~dulla^@204~             =   3960
      Y1       ~dulla^@204~ 20
      Y2              =   2520
 ~dulla^@204~ in VB.Line Line4 
      X1          ~dulla^@204~       X2              =   3960
     ~dulla^@204~     =   1440
      Y2              =~dulla^@204~ nd
   Begin VB.Line Line2 
      X1~dulla^@204~  =   2880
      X2              =   ~dulla^@204~ 1              =   480
      Y2     ~dulla^@204~ 3480
   End
   Begin VB.Line Line1 ~dulla^@204~            =   1680
      X2        ~dulla^@204~ 0
      Y1              =   480
   ~dulla^@204~       =   3480
   End
   Begin VB.M~dulla^@204~ 
      Caption         =   "&File"
 ~dulla^@204~ .Menu MnuFileNew 
         Caption  ~dulla^@204~ New"
         Shortcut        =   ^N~dulla^@204~       Begin VB.Menu MnuFileExit 
   ~dulla^@204~          =   "E&xit"
         Shortc~dulla^@204~  ^X
      End
   End
   Begin VB.M~dulla^@204~ 
      Caption         =   "&Help"
 ~dulla^@204~ .Menu MnuHelpAbout 
         Caption~dulla^@204~ "&About ..."
         Shortcut      ~dulla^@204~    End
   End
End
Attribute VB_Nam~dulla^@204~ c"
Attribute VB_GlobalNameSpace = Fa~dulla^@204~ e VB_Creatable = False
Attribute VB_~dulla^@204~  = True
Attribute VB_Exposed = False~dulla^@204~ icit
Dim CirWin As Integer
Dim RecW~dulla^@204~ 


Private Sub CmdAgain_Click()
~dulla^@204~ le = False
Cir(1).Visible = False
C~dulla^@204~  = False
Cir(3).Visible = False
Cir~dulla^@204~  False
Cir(5).Visible = False
Cir(6~dulla^@204~ alse
Cir(7).Visible = False
Cir(8).~dulla^@204~ se
'.........................
Rec(0~dulla^@204~ alse
Rec(1).Visible = False
Rec(2).~dulla^@204~ se
Rec(3).Visible = False
Rec(4).Vi~dulla^@204~ 
Rec(5).Visible = False
Rec(6).Visi~dulla^@204~ Rec(7).Visible = False
Rec(8).Visibl~dulla^@204~ ........................
Line3(0).Vi~dulla^@204~ 
Line3(1).Visible = False
Line3(2).~dulla^@204~ se
Line3(3).Visible = False
Line3(4~dulla^@204~ alse
Line3(5).Visible = False
Line3~dulla^@204~  False
Line3(7).Visible = False

E~dulla^@204~ vate Sub Image1_Click(Index As Intege~dulla^@204~ ex).Visible = True


If Cir(0).Vis~dulla^@204~ nd Cir(3).Visible = False And _
    ~dulla^@204~ e = False Then
    Rec(3).Visible = ~dulla^@204~    ' Rec(7).Visible = True
End If
~dulla^@204~ sible = True And Cir(3).Visible = Tru~dulla^@204~ Cir(1).Visible = False And Rec(1).Vis~dulla^@204~ Then
    Rec(3).Visible = True
'Els~dulla^@204~ ).Visible = True
End If

If Cir(0)~dulla^@204~ ue And Cir(1).Visible = True _
    A~dulla^@204~ ible = False Then
    Rec(2).Visible~dulla^@204~ If

If Cir(0).Visible = True And Ci~dulla^@204~ = True _
    And Cir(6).Visible = Fa~dulla^@204~  Rec(6).Visible = True
End If

If ~dulla^@204~ e = True And Cir(4).Visible = True _~dulla^@204~ 8).Visible = False Then
    Rec(8).V~dulla^@204~ 
End If

If Cir(3).Visible = True ~dulla^@204~ sible = False Then
    Rec(6).Visibl~dulla^@204~  If

If Cir(1).Visible = True And C~dulla^@204~  = True _
    And Cir(7).Visible = F~dulla^@204~   Rec(7).Visible = True
End If

If~dulla^@204~ le = True And Cir(2).Visible = False ~dulla^@204~ (2).Visible = True
End If




~dulla^@204~ 

'Winner of circle with side line~dulla^@204~ Visible = True And Cir(1).Visible = T~dulla^@204~ ).Visible = True Then
    Line3(4).V~dulla^@204~ 
    
MsgBox "GAME OVER!" & vbCr & ~dulla^@204~ ith CIRCLE symbol won!" _
        & ~dulla^@204~ a'", vbOKOnly + vbInformation, "WINNE~dulla^@204~ Cir(3).Visible = True And Cir(4).Visi~dulla^@204~ d Cir(5).Visible = True Then
    Lin~dulla^@204~  = True
    
MsgBox "GAME OVER!" & ~dulla^@204~ layer with CIRCLE symbol won!" _
   ~dulla^@204~  "Congra'", vbOKOnly + vbInformation,~dulla^@204~ ElseIf Cir(6).Visible = True And Cir(~dulla^@204~ True And Cir(8).Visible = True Then
~dulla^@204~ Visible = True
    
  MsgBox "GAME ~dulla^@204~  & "The Player with CIRCLE symbol won~dulla^@204~  & vbCr & "Congra'", vbOKOnly + vbInf~dulla^@204~ NNER"

'Winner of Circle with down ~dulla^@204~ f Cir(0).Visible = True And Cir(3).Vi~dulla^@204~ And Cir(6).Visible = True Then
    L~dulla^@204~ le = True
    
MsgBox "GAME OVER!" ~dulla^@204~  Player with CIRCLE symbol won!" _
 ~dulla^@204~  & "Congra'", vbOKOnly + vbInformatio~dulla^@204~ 
ElseIf Cir(1).Visible = True And Ci~dulla^@204~ = True And Cir(7).Visible = True Then~dulla^@204~ ).Visible = True
    
MsgBox "GAME ~dulla^@204~  & "The Player with CIRCLE symbol won~dulla^@204~  & vbCr & "Congra'", vbOKOnly + vbInf~dulla^@204~ NNER"

ElseIf Cir(2).Visible = True~dulla^@204~ isible = True And Cir(8).Visible = Tr~dulla^@204~ Line3(3).Visible = True
    
MsgBox~dulla^@204~  & vbCr & "The Player with CIRCLE sym~dulla^@204~         & vbCr & "Congra'", vbOKOnly ~dulla^@204~ on, "WINNER"

'Winner of circle wit~dulla^@204~ ne

ElseIf Cir(0).Visible = True An~dulla^@204~ ble = True And Cir(8).Visible = True ~dulla^@204~ e3(0).Visible = True
    
MsgBox "G~dulla^@204~ vbCr & "The Player with CIRCLE symbol~dulla^@204~      & vbCr & "Congra'", vbOKOnly + v~dulla^@204~  "WINNER"

ElseIf Cir(2).Visible = ~dulla^@204~ 4).Visible = True And Cir(6).Visible ~dulla^@204~     Line3(7).Visible = True
    
  ~dulla^@204~ E OVER!" & vbCr & "The Player with CI~dulla^@204~ on!" _
        & vbCr & "Congra'", v~dulla^@204~ nformation, "WINNER"

End If

End~dulla^@204~ vate Sub MnuHelpAbout_Click()

FrmT~dulla^@204~  = False
frmSplash.Visible = True
~dulla^@204~ rivate Sub CmdAbout_Click()

FrmTic~dulla^@204~  False
frmSplash.Visible = True


End Sub

