VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Beginner"
   ClientHeight    =   3615
   ClientLeft      =   5925
   ClientTop       =   6840
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   3615
   ScaleWidth      =   4680
   Begin VB.CommandButton Command3 
      Caption         =   "Next Lesson"
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2760
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Info2"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Local2"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Info"
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Local"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Text3.Text = Winsock1.LocalIP  ' When The Command2 Button Is Clicked Your LocalIP Is Displayed In Text3.Text '
Text4.Text = Winsock1.LocalHostName  ' When The Command2 Button Is Clicked Your LocalHostName Is Displayed In Text4.Text '
End Sub

Private Sub Command3_Click()
Form1.Hide 'When Command3 Is Clicked Form1 Is Hidden'
Form2.Show 'When Command3 Is Clicked Form2 Is Shown '
End Sub

Private Sub Form_Load()
Text1.Text = Winsock1.LocalIP  ' When The Form Loads Your LocalIP Is Displayed In Text1.Text '
Text2.Text = Winsock1.LocalHostName   ' When The Form Loads Your LocalHostName Is Displayed In Text1.Text '
End Sub


Private Sub Command1_Click()
MsgBox "This Is Your LocalIP And LocalHost", vbInformation, "Info"  ' When The Command1 Button Is Clicked A Message Box PopsUp Telling You Whats Between The " " '
End Sub
'##################################################################################
' Remember To Add Your Winsock Control,To Add This GoTo Project Then Components   #
'                                                                                 #
' Scroll Down Until You See Microsoft Winsock Control 6.0,Tick The Box,Now Goto   #
' #Your Menu Bar On The Left,Click The Little Icon Looks Like 2 PC'S,On MouseOver #
' #It Will Say Winsock,Just Double Click It !!!!'                                 #
' #################################################################################

