VERSION 5.00
Begin VB.Form FrmNotepad 
   Caption         =   "Untitled - Notepad"
   ClientHeight    =   5040
   ClientLeft      =   3615
   ClientTop       =   3315
   ClientWidth     =   8415
   Icon            =   "FrmNotepad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8415
   Begin VB.TextBox txtWindow 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
   End
   Begin VB.Menu Mnuserach 
      Caption         =   "&Edit"
   End
   Begin VB.Menu Mnusearch 
      Caption         =   "&Search"
   End
   Begin VB.Menu Mnuhelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "FrmNotepad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Me.txtWindow.Height = Me.ScaleHeight
Me.txtWindow.Width = Me.ScaleWidth
End Sub

Private Sub txtWindow_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And txtWindow.Text = "chat" Then
FrmNotepad.Visible = False
FrmChat.Visible = True
End If
End Sub

