VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   Caption         =   "IP Getter"
   ClientHeight    =   1005
   ClientLeft      =   4935
   ClientTop       =   7215
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   1005
   ScaleWidth      =   5025
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get IP"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   240
      TabIndex        =   1
      Text            =   "www.somesitehere.com"
      Top             =   360
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Get Site IP"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next 'Does This Really Need Explaining??'
Winsock1.Connect Text1.Text, "80" 'Tells Winsock To Connect To The Site Address In Text1.Text On The Given Port This One Being 80 As Its Most Commonly Your Internet Port '
DoEvents:
End Sub

Private Sub Command2_Click()
Form2.Hide 'When Command2 Is Clicked Form2 Is Hidden'
Form3.Show 'When Command2 Is Clicked Form3 Is Shown '
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next
Text1.Text = Winsock1.RemoteHostIP 'When Winsock Connects This Shows The Site IP Of The Site You Searced For'
End Sub


