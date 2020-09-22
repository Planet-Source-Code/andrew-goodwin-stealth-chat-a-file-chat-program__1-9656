VERSION 5.00
Begin VB.Form FrmChat 
   Caption         =   "Stealth-Chat File Edition Beta V7.0 EzY-Software "
   ClientHeight    =   3705
   ClientLeft      =   4935
   ClientTop       =   3900
   ClientWidth     =   5535
   Icon            =   "FrmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "FrmChat.frx":0442
   ScaleHeight     =   3705
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicStaticRoomList 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      Height          =   3045
      Left            =   0
      Picture         =   "FrmChat.frx":052C
      ScaleHeight     =   2985
      ScaleWidth      =   5475
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
      Begin VB.PictureBox PicUsers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2955
         Left            =   0
         Picture         =   "FrmChat.frx":0836
         ScaleHeight     =   2955
         ScaleWidth      =   5535
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   5535
         Begin VB.PictureBox picStart 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2955
            Left            =   0
            Picture         =   "FrmChat.frx":0B40
            ScaleHeight     =   2955
            ScaleWidth      =   5535
            TabIndex        =   31
            Top             =   0
            Width           =   5535
            Begin VB.PictureBox PicAbout 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   2955
               Left            =   0
               Picture         =   "FrmChat.frx":0E4A
               ScaleHeight     =   2955
               ScaleWidth      =   5535
               TabIndex        =   39
               Top             =   0
               Visible         =   0   'False
               Width           =   5535
               Begin VB.PictureBox Picerror 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   2955
                  Left            =   0
                  Picture         =   "FrmChat.frx":128C
                  ScaleHeight     =   2955
                  ScaleWidth      =   5535
                  TabIndex        =   53
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   5535
                  Begin VB.PictureBox picquick 
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   2955
                     Left            =   0
                     Picture         =   "FrmChat.frx":16CE
                     ScaleHeight     =   2955
                     ScaleWidth      =   5535
                     TabIndex        =   58
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   5535
                     Begin VB.CommandButton Command5 
                        Caption         =   "Back"
                        Height          =   375
                        Left            =   480
                        TabIndex        =   59
                        Top             =   2400
                        Width           =   975
                     End
                     Begin VB.Label Label32 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "type this in and then press enter. Then type your new name in and press enter again."
                        Height          =   495
                        Left            =   1320
                        TabIndex        =   67
                        Top             =   960
                        Width           =   3615
                     End
                     Begin VB.Label Label24 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "<name>"
                        Height          =   255
                        Left            =   600
                        TabIndex        =   66
                        Top             =   960
                        Width           =   615
                     End
                     Begin VB.Label Label31 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Once in notepad type chat and then enter, then it will go back to normal chat again."
                        Height          =   495
                        Left            =   720
                        TabIndex        =   65
                        Top             =   1800
                        Width           =   4215
                     End
                     Begin VB.Label Label30 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Turns the program into notepad"
                        Height          =   255
                        Left            =   960
                        TabIndex        =   64
                        Top             =   1440
                        Width           =   3255
                     End
                     Begin VB.Label Label29 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "  n"
                        Height          =   255
                        Left            =   600
                        TabIndex        =   63
                        Top             =   1440
                        Width           =   375
                     End
                     Begin VB.Label Label27 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Minimizes the program ."
                        Height          =   255
                        Left            =   960
                        TabIndex        =   62
                        Top             =   600
                        Width           =   1695
                     End
                     Begin VB.Label Label28 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "Quick Keys:"
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   255
                        Left            =   600
                        TabIndex        =   61
                        Top             =   120
                        Width           =   1095
                     End
                     Begin VB.Label Label25 
                        BackColor       =   &H00FFFFFF&
                        Caption         =   "  m"
                        Height          =   255
                        Left            =   600
                        TabIndex        =   60
                        Top             =   600
                        Width           =   375
                     End
                  End
                  Begin VB.CommandButton Command6 
                     Caption         =   "Exit Chat"
                     Height          =   375
                     Left            =   480
                     TabIndex        =   57
                     Top             =   2400
                     Width           =   975
                  End
                  Begin VB.Label lblerror 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Error Description:"
                     Height          =   1215
                     Left            =   480
                     TabIndex        =   56
                     Top             =   960
                     Width           =   4815
                  End
                  Begin VB.Label Label26 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "Description:"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   480
                     TabIndex        =   55
                     Top             =   480
                     Width           =   1095
                  End
                  Begin VB.Label lblheading 
                     BackColor       =   &H00FFFFFF&
                     Caption         =   "An error has occured."
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   255
                     Left            =   480
                     TabIndex        =   54
                     Top             =   120
                     Width           =   4815
                  End
               End
               Begin VB.CommandButton Command4 
                  Caption         =   "Go Back"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   40
                  Top             =   2640
                  Width           =   1095
               End
               Begin VB.Label Label23 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "- A Hot-Key to replica Notepad"
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   52
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.Label Label22 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "- A small and compact interface"
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   51
                  Top             =   600
                  Width           =   2295
               End
               Begin VB.Label Label21 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "- No popup forms or errors"
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   50
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.Label Label20 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Includes"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   3000
                  TabIndex        =   49
                  Top             =   120
                  Width           =   855
               End
               Begin VB.Line Line1 
                  X1              =   2880
                  X2              =   2880
                  Y1              =   120
                  Y2              =   1320
               End
               Begin VB.Label Label19 
                  Caption         =   "E-Mail Cybergod82@Hotmail.com"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   48
                  Top             =   2640
                  Width           =   3495
               End
               Begin VB.Label Label18 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   $"FrmChat.frx":1B10
                  Height          =   975
                  Left            =   600
                  TabIndex        =   47
                  Top             =   1560
                  Width           =   4935
               End
               Begin VB.Label Label17 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Why Stealth-Chat was created"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   600
                  TabIndex        =   46
                  Top             =   1320
                  Width           =   2775
               End
               Begin VB.Label Label16 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "- Finished 8/7/00"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   45
                  Top             =   1080
                  Width           =   1335
               End
               Begin VB.Label Label14 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "- Started 1/5/00"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   44
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.Label Label13 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "- Version 7.0"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   43
                  Top             =   600
                  Width           =   2295
               End
               Begin VB.Label Label12 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "- Created by Andrew Goodwin"
                  Height          =   255
                  Left            =   600
                  TabIndex        =   42
                  Top             =   360
                  Width           =   2295
               End
               Begin VB.Label Label15 
                  BackStyle       =   0  'Transparent
                  Caption         =   "Stealth-Chat"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   600
                  TabIndex        =   41
                  Top             =   120
                  Width           =   1095
               End
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Start Chat"
               Height          =   375
               Left            =   600
               TabIndex        =   36
               Top             =   2400
               Width           =   1095
            End
            Begin VB.TextBox txtusername 
               Height          =   285
               Left            =   600
               MaxLength       =   14
               TabIndex        =   35
               Top             =   2040
               Width           =   4215
            End
            Begin VB.TextBox txtpath 
               Height          =   285
               Left            =   600
               TabIndex        =   38
               Top             =   1200
               Width           =   4215
            End
            Begin VB.Label lblwarning 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H000000FF&
               Height          =   495
               Left            =   1800
               TabIndex        =   37
               Top             =   2400
               Width           =   3135
            End
            Begin VB.Label Label11 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Enter your nick name that will be used in the chat rooms. you can change this later if you need to. MAX 15 chars"
               ForeColor       =   &H00000000&
               Height          =   495
               Left            =   600
               TabIndex        =   34
               Top             =   1560
               Width           =   4095
            End
            Begin VB.Label Label10 
               BackColor       =   &H00FFFFFF&
               Caption         =   $"FrmChat.frx":1C46
               Height          =   615
               Left            =   600
               TabIndex        =   33
               Top             =   480
               Width           =   4095
            End
            Begin VB.Label lbltitle 
               BackStyle       =   0  'Transparent
               Caption         =   "Welcome to Stealth-Chat"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   600
               TabIndex        =   32
               Top             =   240
               Width           =   2295
            End
         End
         Begin VB.Timer TmeOpenUsers 
            Enabled         =   0   'False
            Interval        =   1
            Left            =   4920
            Top             =   1080
         End
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            Caption         =   "O.K"
            Height          =   375
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   2400
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.TextBox txtUsers 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   600
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   28
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label lblCurrentRoom2 
            Caption         =   "Current Room:"
            Height          =   255
            Left            =   2160
            TabIndex        =   29
            Top             =   2520
            Width           =   2535
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "This shows the current users that are in the room you are currently in:"
            Height          =   495
            Left            =   600
            TabIndex        =   27
            Top             =   600
            Width           =   4095
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Users:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Cancel"
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.Timer TmrUpdateDisplay 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   4560
         Top             =   2400
      End
      Begin VB.Label lblerrorDesciption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Static Rooms: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   540
         TabIndex        =   24
         Top             =   180
         Width           =   3615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Choose a room below by clicking on the room name:    Or click on the cancel button to go back to chat room"
         Height          =   495
         Left            =   480
         TabIndex        =   23
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "1) Work"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   480
         TabIndex        =   22
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "2) General"
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   1320
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "3) Lobby"
         Height          =   255
         Left            =   480
         TabIndex        =   20
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "4) Hacking"
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   1800
         Width           =   3735
      End
      Begin VB.Label Label6 
         Caption         =   "5) Life"
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label Lblcurrentroom 
         Caption         =   "Current Room:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lblWork 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label LblLobby 
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label LblLife 
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label lblGeneral 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   13
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label LblHacking 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4440
         TabIndex        =   12
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Timer TmrUsers 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   360
   End
   Begin VB.Timer TmrRefresh 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   360
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      MaxLength       =   170
      TabIndex        =   6
      Top             =   3120
      Width           =   5535
   End
   Begin VB.PictureBox Menu12Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      Picture         =   "FrmChat.frx":1CDF
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Menu7Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1680
      Picture         =   "FrmChat.frx":1F29
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Menu0Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      Picture         =   "FrmChat.frx":2173
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Menu9Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      Picture         =   "FrmChat.frx":225D
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox Menu8Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      Picture         =   "FrmChat.frx":24A7
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtChatWindow 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label Label9 
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblroom 
      Height          =   255
      Left            =   3480
      TabIndex        =   8
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label lblInformation 
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3480
      Width           =   3375
   End
   Begin VB.Menu MnuChatOptions 
      Caption         =   "&Chat Options"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu MnuChangeRoom 
         Caption         =   "Change &Room"
         Index           =   3
         Shortcut        =   ^R
      End
      Begin VB.Menu Mnuline2 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu MnuViewUsers 
         Caption         =   "View Users"
         Index           =   8
      End
      Begin VB.Menu MnuSmartKeys 
         Caption         =   "View Smart Keys"
         Index           =   9
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuauto 
         Caption         =   "Auto Scroll"
         Checked         =   -1  'True
         Shortcut        =   {F2}
      End
      Begin VB.Menu Mnuline3 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About Stealth-Chat"
         Index           =   11
      End
      Begin VB.Menu MnuExit 
         Caption         =   "&Exit"
         Index           =   12
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "FrmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Close_Cancel

End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0C0C0
Label3.BackColor = &HFFFFFF
Label4.BackColor = &HC0C0C0
Label5.BackColor = &HFFFFFF
Label6.BackColor = &HC0C0C0
lblWork.BackColor = &HC0C0C0
lblGeneral.BackColor = &HFFFFFF
LblLife.BackColor = &HC0C0C0
LblHacking.BackColor = &HFFFFFF
LblLobby.BackColor = &HC0C0C0
End Sub

Private Sub Command2_Click()

Close_Cancel

End Sub

Private Sub Command3_Click()
If txtpath.Text = "" And txtusername.Text = "" And Stop_Handle = 0 Then
lblwarning.Caption = "Please enter the path to the location where the room files are. and your nick name"
Exit Sub
End If
If txtusername = "" And Stop_Handle = 0 Then
lblwarning.Caption = "Please enter a nick name"
Exit Sub
End If
If txtpath.Text = "" Then
lblwarning.Caption = "Please enter the path to the location where the room files are."
Exit Sub
End If
NameOfPerson = txtusername.Text
Temp_Filepath = txtpath.Text
Close 1#
Open (App.Path & "\location.ini") For Output As 1#
Print #1, Temp_Filepath
Close #1
Open (App.Path & "\Name.txt") For Output As 1#
Print #1, NameOfPerson
Close 1#
PicStaticRoomList.Visible = False
PicUsers.Visible = False
picStart.Visible = False
Create_Rooms
End Sub

Private Sub Command4_Click()
PicStaticRoomList.Visible = False
PicUsers.Visible = False
picStart.Visible = False
PicAbout.Visible = False
lblInformation.Visible = True
lblroom.Visible = True
txtMessage.Enabled = True
End Sub

Private Sub Command5_Click()
PicStaticRoomList.Visible = False
PicUsers.Visible = False
picStart.Visible = False
Picerror.Visible = False
PicAbout.Visible = False
picquick.Visible = False

lblInformation.Visible = True
lblroom.Visible = True
txtMessage.Enabled = True
End Sub

Private Sub Command6_Click()
Left_Chat
End
End Sub

Private Sub Form_Load()
Change_Name = 0
Auto_Scroll = 1
Stop_Handle = 0
Names_All = "Lobby"
Check_File
End Sub

Public Sub Check_File()
On Error GoTo Check_error

Open (App.Path & "\location.ini") For Append As 1#
Close 1#
Open (App.Path & "\name.txt") For Append As 1#
Close 1#
Read_File
Exit Sub
Check_error:
Name_Error
End Sub

Public Sub Read_File()
Open (App.Path & "\location.ini") For Input As 1#
Start_FileOpen = Input(LOF(1#), 1#)
Start_FilePath = Start_FileOpen
Close #1
Open (App.Path & "\name.txt") For Input As 1#
NameOfPerson = Input(LOF(1#), 1#)
Close 1#
If NameOfPerson = "" And Start_FilePath > "" Then
Name_Error
Else
If Start_FilePath = "" And NameOfPerson > "" Then
Room_Create_Error
Else
If Start_FilePath = "" And NameOfPerson = "" Then
lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False
MnuChatOptions.Enabled = False
PicStaticRoomList.Visible = True
PicUsers.Visible = True
picStart.Visible = True
Else
Create_Rooms
End If
End If
End If
End Sub

Public Sub Create_Rooms()
Shorten_File
Room_General = Op_ChatPath & "\General.txt"
Room_Hacking = Op_ChatPath & "\Hacking.txt"
Room_Life = Op_ChatPath & "\Life.txt"
Room_Work = Op_ChatPath & "\Work.txt"
Room_Lobby = Op_ChatPath & "\Lobby.txt"

Users_General = Op_ChatPath & "\General(Ro).txt"
Users_Hacking = Op_ChatPath & "\Hacking(Ro).txt"
Users_Life = Op_ChatPath & "\Life(Ro).txt"
Users_Work = Op_ChatPath & "\Work(Ro).txt"
Users_Lobby = Op_ChatPath & "\Lobby(Ro).txt"
    
Open_Static_Rooms
    
End Sub

Private Sub Form_Terminate()
On Error GoTo Check_error
Close 1#
Open (Room_Lobby) For Append As 1#
Close 1#
Open (Room_Hacking) For Append As 1#
Close 1#
Open (Room_General) For Append As 1#
Close 1#
Open (Room_Life) For Append As 1#
Close 1#
Open (Room_Work) For Append As 1#
Close 1#

Open (Users_Lobby) For Append As 1#
Close 1#
Open (Users_Hacking) For Append As 1#
Close 1#
Open (Users_General) For Append As 1#
Close 1#
Open (Users_Life) For Append As 1#
Close 1#
Open (Users_Work) For Append As 1#
Close 1#
Close 1#
Open (Users_Lobby) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_Lobby)
Close 1#
Kill (Users_Lobby)
Close 1#
End If
Close 1#
Open (Users_Work) For Append As 1#
Close 1#
Open (Users_Work) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_Work)
Close 1#
Kill (Users_Work)
Close 1#
End If
Close 1#
Open (Users_General) For Append As 1#
Close 1#
Open (Users_General) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_General)
Close 1#
Kill (Users_General)
Close 1#
End If
Close 1#
Open (Users_Life) For Append As 1#
Close 1#
Open (Users_Life) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_Life)
Close 1#
Kill (Users_Life)
Close 1#
End If
Close 1#
Open (Users_Hacking) For Append As 1#
Close 1#
Open (Users_Hacking) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_Hacking)
Close 1#
Kill (Users_Hacking)
Close 1#
End If

Left_Chat
Exit Sub
Check_error:
Exit Sub
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0C0C0
Label3.BackColor = &HFFFFFF
Label4.BackColor = &HC0C0C0
Label5.BackColor = &HFFFFFF
Label6.BackColor = &HC0C0C0
lblWork.BackColor = &HC0C0C0
lblGeneral.BackColor = &HFFFFFF
LblLife.BackColor = &HC0C0C0
LblHacking.BackColor = &HFFFFFF
LblLobby.BackColor = &HC0C0C0
End Sub

Private Sub Label2_Click()
Kill (Set_Users)
left_Room
set_room = Room_Work
Set_Users = Users_Work
TmeOpenUsers.Enabled = True
lblroom.Caption = "Room | Work"
Current_Room = "Current Room: Work"
Close_Room_List
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HFFC0C0
Label3.BackColor = &HFFFFFF
Label4.BackColor = &HC0C0C0
Label5.BackColor = &HFFFFFF
Label6.BackColor = &HC0C0C0
lblWork.BackColor = &HFFC0C0
lblGeneral.BackColor = &HFFFFFF
LblLife.BackColor = &HC0C0C0
LblHacking.BackColor = &HFFFFFF
LblLobby.BackColor = &HC0C0C0
End Sub

Private Sub Label3_Click()
Kill (Set_Users)
left_Room
TmeOpenUsers.Enabled = True
set_room = Room_General
Set_Users = Users_General
lblroom.Caption = "Room | General"
Names_All = "General"
Current_Room = "Current Room: General"
Close_Room_List
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0C0C0
Label3.BackColor = &HFFC0C0
Label4.BackColor = &HC0C0C0
Label5.BackColor = &HFFFFFF
Label6.BackColor = &HC0C0C0
lblWork.BackColor = &HC0C0C0
lblGeneral.BackColor = &HFFC0C0
LblLife.BackColor = &HC0C0C0
LblHacking.BackColor = &HFFFFFF
LblLobby.BackColor = &HC0C0C0
End Sub

Private Sub Label4_Click()
Kill (Set_Users)
left_Room
set_room = Room_Lobby
Set_Users = Users_Lobby
lblroom.Caption = "Room | Lobby"
Names_All = "Lobby"
Current_Room = "Current Room: Lobby"
Close_Room_List
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0C0C0
Label3.BackColor = &HFFFFFF
Label4.BackColor = &HFFC0C0
Label5.BackColor = &HFFFFFF
Label6.BackColor = &HC0C0C0
lblWork.BackColor = &HC0C0C0
lblGeneral.BackColor = &HFFFFFF
LblLife.BackColor = &HC0C0C0
LblHacking.BackColor = &HFFFFFF
LblLobby.BackColor = &HFFC0C0

End Sub

Private Sub Label5_Click()
Kill (Set_Users)
left_Room
set_room = Room_Hacking
Set_Users = Users_Hacking
lblroom.Caption = "Room | Hacking"
Names_All = "Hacking"
Current_Room = "Current Room: Hacking"
Close_Room_List
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0C0C0
Label3.BackColor = &HFFFFFF
Label4.BackColor = &HC0C0C0
Label5.BackColor = &HFFC0C0
Label6.BackColor = &HC0C0C0
lblWork.BackColor = &HC0C0C0
lblGeneral.BackColor = &HFFFFFF
LblLife.BackColor = &HC0C0C0
LblHacking.BackColor = &HFFC0C0
LblLobby.BackColor = &HC0C0C0
End Sub

Private Sub Label6_Click()
Kill (Set_Users)
left_Room
set_room = Room_Life
Set_Users = Users_Life
lblroom.Caption = "Room | Life"
Names_All = "Life"
Current_Room = "Current Room: Life"
Close_Room_List
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0C0C0
Label3.BackColor = &HFFFFFF
Label4.BackColor = &HC0C0C0
Label5.BackColor = &HFFFFFF
Label6.BackColor = &HFFC0C0
lblWork.BackColor = &HC0C0C0
lblGeneral.BackColor = &HFFFFFF
LblLife.BackColor = &HFFC0C0
LblHacking.BackColor = &HFFFFFF
LblLobby.BackColor = &HC0C0C0
End Sub

Private Sub MnuAbout_Click(Index As Integer)
PicStaticRoomList.Visible = True
PicUsers.Visible = True
picStart.Visible = True
PicAbout.Visible = True
lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False
End Sub

Private Sub mnuauto_Click()
If mnuauto.Checked = True Then
mnuauto.Checked = False
Auto_Scroll = 0
Else
If mnuauto.Checked = False Then
mnuauto.Checked = True
Auto_Scroll = 1
End If
End If
End Sub

Private Sub MnuChangeRoom_Click(Index As Integer)
On Error GoTo Check_error
Lblcurrentroom.Caption = Current_Room
PicStaticRoomList.Visible = True
PicUsers.Visible = False
Close 1#
Open (Room_Lobby) For Append As 1#
Close 1#
Open (Room_Hacking) For Append As 1#
Close 1#
Open (Room_General) For Append As 1#
Close 1#
Open (Room_Life) For Append As 1#
Close 1#
Open (Room_Work) For Append As 1#
Close 1#
Close 1#
Open (Users_Lobby) For Append As 1#
Close 1#
Open (Users_Hacking) For Append As 1#
Close 1#
Open (Users_General) For Append As 1#
Close 1#
Open (Users_Life) For Append As 1#
Close 1#
Open (Users_Work) For Append As 1#
Close 1#
lblroom.Visible = False
txtMessage.Enabled = False
lblInformation.Visible = False
Exit Sub
Check_error:
lblerror.Caption = "An error ocurred in the MnuChangeRoom_Click() function when trying to create the room files. This maybe because come on else has cleared them when leaving Stealth-Chat. Restart Stealth-Chat to correct the problem."
PicStaticRoomList.Visible = True
PicUsers.Visible = True
MnuChatOptions.Enabled = False
picStart.Visible = True
PicAbout.Visible = True
Picerror.Visible = True
lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False
End Sub

Private Sub MnuExit_Click(Index As Integer)
On Error GoTo Check_error
Close 1#
Close 1#
Open (Room_Lobby) For Append As 1#
Close 1#
Open (Room_Hacking) For Append As 1#
Close 1#
Open (Room_General) For Append As 1#
Close 1#
Open (Room_Life) For Append As 1#
Close 1#
Open (Room_Work) For Append As 1#
Close 1#

Open (Users_Lobby) For Append As 1#
Close 1#
Open (Users_Hacking) For Append As 1#
Close 1#
Open (Users_General) For Append As 1#
Close 1#
Open (Users_Life) For Append As 1#
Close 1#
Open (Users_Work) For Append As 1#
Close 1#
Open (Users_Lobby) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_Lobby)
Close 1#
Kill (Users_Lobby)
Close 1#
End If
Close 1#
Open (Users_Work) For Append As 1#
Close 1#
Open (Users_Work) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_Work)
Close 1#
Kill (Users_Work)
Close 1#
End If
Close 1#
Open (Users_General) For Append As 1#
Close 1#
Open (Users_General) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_General)
Close 1#
Kill (Users_General)
Close 1#
End If
Close 1#
Open (Users_Life) For Append As 1#
Close 1#
Open (Users_Life) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_Life)
Close 1#
Kill (Users_Life)
Close 1#
End If
Close 1#
Open (Users_Hacking) For Append As 1#
Close 1#
Open (Users_Hacking) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
Kill (Room_Hacking)
Close 1#
Kill (Users_Hacking)
Close 1#
End If
Left_Chat
End
Exit Sub
Check_error:
Exit Sub
End Sub

Private Sub MnuSmartKeys_Click(Index As Integer)
PicStaticRoomList.Visible = True
PicUsers.Visible = True
picStart.Visible = True
Picerror.Visible = True
PicAbout.Visible = True
picquick.Visible = True

lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False
End Sub

Private Sub MnuViewUsers_Click(Index As Integer)
PicStaticRoomList.Visible = True
PicUsers.Visible = True
picStart.Visible = False
lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False
lblCurrentRoom2 = Current_Room
End Sub

Private Sub PicStaticRoomList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.BackColor = &HC0C0C0
Label3.BackColor = &HFFFFFF
Label4.BackColor = &HC0C0C0
Label5.BackColor = &HFFFFFF
Label6.BackColor = &HC0C0C0
lblWork.BackColor = &HC0C0C0
lblGeneral.BackColor = &HFFFFFF
LblLife.BackColor = &HC0C0C0
LblHacking.BackColor = &HFFFFFF
LblLobby.BackColor = &HC0C0C0
End Sub

Private Sub TmeOpenUsers_Timer()
On Error GoTo Check_error
Close 1#
Open (Set_Users) For Append As 1#
Close 1#

Close 1#
Open (Set_Users) For Input As 1#
Users_Names = Input(LOF(1#), 1#)
Close 1#

If txtUsers = "" Then
Close #1
Open (Set_Users) For Append As 1#
Print #1, NameOfPerson
Close #1

End If
Exit Sub
Check_error:
lblerror.Caption = "An error occred in the TmeOpenUsers_Timer() function when trying to add your name to the users text file. Re-start Stealth-Chat to try and correct the problem."
PicStaticRoomList.Visible = True
PicUsers.Visible = True
MnuChatOptions.Enabled = False
picStart.Visible = True
PicAbout.Visible = True
Picerror.Visible = True
lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False
End Sub
Private Sub TmrRefresh_Timer()
On Error GoTo Check_error
If Auto_Scroll = 1 Then
file_size = Len(txtChatWindow.Text)
txtChatWindow.SelStart = file_size
End If
Close 1#
Open (set_room) For Append As 1#

Close 1#
Open (set_room) For Input As 1#
txtChatWindow.Text = Input(LOF(1#), 1#)
Close 1#
Exit Sub

Check_error:
lblerror.Caption = "An error occred in the TmrRefresh_Timer() function when trying to open the room file.This may be because the room file may of become corrupt or has been deleted. Re-start Stealth-Chat to try and correct the problem."
PicStaticRoomList.Visible = True
PicUsers.Visible = True
MnuChatOptions.Enabled = False
picStart.Visible = True
PicAbout.Visible = True
Picerror.Visible = True
lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False
End Sub

Public Sub Open_Static_Rooms()

On Error GoTo Error_Create_Rooms

Close 1#
Open (Room_Lobby) For Append As 1#
Close 1#
Open (Room_Hacking) For Append As 1#
Close 1#
Open (Room_General) For Append As 1#
Close 1#
Open (Room_Life) For Append As 1#
Close 1#
Open (Room_Work) For Append As 1#
Close 1#

Open (Users_Lobby) For Append As 1#
Close 1#
Open (Users_Hacking) For Append As 1#
Close 1#
Open (Users_General) For Append As 1#
Close 1#
Open (Users_Life) For Append As 1#
Close 1#
Open (Users_Work) For Append As 1#
Close 1#

set_room = Room_Lobby
Set_Users = Users_Lobby
Joined_Room
lblInformation.Caption = "Enter Message Then Press Enter."
lblroom.Caption = "Room | Lobby"
Current_Room = "Current Room: Lobby"
TmrUsers.Enabled = True
TmrRefresh.Enabled = True
TmrUpdateDisplay.Enabled = True
Exit Sub
Error_Create_Rooms:
 Room_Create_Error
End Sub


Private Sub TmrUpdateDisplay_Timer()
On Error GoTo Check_error
Close 1#
Open (Users_Lobby) For Append As 1#
Close 1#
Open (Users_Lobby) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
LblLobby.Caption = "Empty"
Else
LblLobby.Caption = "Chatting"
Close 1#
End If

Close 1#
Open (Users_Work) For Append As 1#
Close 1#
Open (Users_Work) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
lblWork.Caption = "Empty"
Else
lblWork.Caption = "Chatting"
Close 1#
End If

Close 1#
Open (Users_General) For Append As 1#
Close 1#
Open (Users_General) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
lblGeneral.Caption = "Empty"
Else
lblGeneral.Caption = "Chatting"
Close 1#
End If

Close 1#
Open (Users_Life) For Append As 1#
Close 1#
Open (Users_Life) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
LblLife.Caption = "Empty"
Else
LblLife.Caption = "Chatting"
Close 1#
End If

Close 1#
Open (Users_Hacking) For Append As 1#
Close 1#
Open (Users_Hacking) For Input As 1#
File_Temp = Input(LOF(1#), 1#)
If File_Temp = "" Then
LblHacking.Caption = "Empty"
Else
LblHacking.Caption = "Chatting"
Close 1#
End If
Exit Sub
Check_error:
lblerror.Caption = "An error occred in the TmrUpdateDisplay_Timer() function when trying to check whether the rooms are empty or not. Re-start Stealth-Chat to try and correct the problem."
PicStaticRoomList.Visible = True
PicUsers.Visible = True
MnuChatOptions.Enabled = False
picStart.Visible = True
PicAbout.Visible = True
Picerror.Visible = True
lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False

End Sub

Public Sub Close_Room_List()

PicStaticRoomList.Visible = False
txtMessage.Enabled = False
lblInformation.Visible = True
txtMessage.Enabled = True
lblroom.Visible = True

Joined_Room
End Sub

Public Sub Joined_Room()


Close 1#
Open (set_room) For Append As 1#
Print #1, "(" + NameOfPerson + " has enterd the room " + ")"
Close 1#

Add_To_User_List

End Sub

Public Sub Close_Cancel()

PicStaticRoomList.Visible = False
txtMessage.Enabled = False
lblInformation.Visible = True
txtMessage.Enabled = True
lblroom.Visible = True


End Sub


Public Sub Add_To_User_List()

Close #1
Open (Set_Users) For Append As 1#
Print #1, NameOfPerson
Close 1#
TmeOpenUsers.Enabled = True
End Sub

Private Sub TmrUsers_Timer()

Close 1#
Open (Set_Users) For Append As 1#
Close 1#

Close 1#
Open (Set_Users) For Input As 1#
Users_Names = Input(LOF(1#), 1#)
Close 1#

txtUsers.Text = Users_Names

End Sub

Public Sub Left_Chat()
On Error GoTo exit_sub
Close 1#
Open (set_room) For Append As 1#
Print #1, "(" + NameOfPerson + " has left Stealth-Chat " + ")"
Close 1#
Kill (Set_Users)
Exit Sub
exit_sub:
Exit Sub
End Sub

Public Sub left_Room()
Close 1#
Open (set_room) For Append As 1#
Print #1, "(" + NameOfPerson + " has left the room " + ")"
Close 1#
End Sub

Public Sub Add_UserText()
If txtMessage.Text = "n" Then
FrmNotepad.Visible = True
FrmChat.Visible = False
txtMessage.Text = ""
FrmNotepad.txtWindow.Text = ""
Exit Sub
End If
If txtMessage.Text = "m" Then
FrmChat.WindowState = 1
txtMessage.Text = ""
Exit Sub
End If
User_Message = txtMessage.Text
Close 1#
Open (set_room) For Append As 1#
Print #1, NameOfPerson + ":" + User_Message
Close 1#
txtMessage.Text = ""
End Sub

Private Sub txtMessage_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And txtMessage.Text = "<name>" Then
Change_Name = 1
lblInformation.Caption = "Enter name then press enter."
txtMessage.Text = ""
Exit Sub
End If
If KeyCode = 13 And Change_Name = 1 Then
first_name = NameOfPerson
NameOfPerson = txtMessage.Text
Close 1#
Open (App.Path & "\name.txt") For Output As 1#
Print #1, NameOfPerson
Close 1#
Open (set_room) For Append As 1#
Print #1, "( " + first_name + " has changed there name to > " + NameOfPerson + " )"
Close 1#
lblInformation.Caption = "Enter message then press enter"
Change_Name = 0
txtMessage.Text = ""
Exit Sub
End If
If KeyCode = 13 Then
Add_UserText
End If
End Sub

Public Sub Shorten_File()
'***********************************************
'This part gets rid of the enter char at the
'of the text line in the file.
'***********************************************
Open (App.Path & "\location.ini") For Input As 1#
Start_FileOpen = Input(LOF(1#), 1#)
Start_FilePath = Start_FileOpen
Start_FileSize = Len(Start_FileOpen)
Start_FilePath = Left(Start_FilePath, Start_FileSize - 2)
Close #1
Open (App.Path & "\name.txt") For Input As 1#
NameOfPerson = Input(LOF(1#), 1#)
Start_FileSize = Len(NameOfPerson)
NameOfPerson = Left(NameOfPerson, Start_FileSize - 2)
If NameOfPerson = "" Then
Name_Error
Exit Sub
Else
If Start_FilePath = "" Then
Room_Create_Error
End If
End If
lblInformation.Visible = True
lblroom.Visible = True
txtMessage.Enabled = True
MnuChatOptions.Enabled = True
Op_ChatPath = Start_FilePath
End Sub

Public Sub Room_Create_Error()
Close 1#
Open (App.Path & "\name.txt") For Input As 1#
NameOfPerson = Input(LOF(1#), 1#)
Start_FileSize = Len(NameOfPerson)
Temp_Filepath = Left(NameOfPerson, Start_FileSize - 2)
Close 1#
Close 1#
Open (App.Path & "\location.ini") For Input As 1#
Start_FileOpen = Input(LOF(1#), 1#)
Start_FilePath = Start_FileOpen
Start_FileSize = Len(Start_FileOpen)
Temp_F = Left(Start_FilePath, Start_FileSize - 2)
Close #1
txtpath.Text = Temp_F
txtusername.Text = Temp_Filepath
Stop_Handle = 1
TmrUsers.Enabled = False
TmrRefresh.Enabled = False
TmrUpdateDisplay.Enabled = False
lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False
MnuChatOptions.Enabled = False
PicStaticRoomList.Visible = True
PicUsers.Visible = True
picStart.Visible = True
lbltitle.Caption = "Error starting Stealth-Chat"
txtusername.Enabled = False
Label11.Enabled = False
Label10.ForeColor = &HFF&
Label10.Caption = "There was an error starting Stealth-Chat please re-type the correct path in the box below where the room files are located"
End Sub

Public Sub Name_Error()
Close 1#
Open (App.Path & "\location.ini") For Input As 1#
Start_FileOpen = Input(LOF(1#), 1#)
Start_FilePath = Start_FileOpen
Start_FileSize = Len(Start_FileOpen)
Temp_F = Left(Start_FilePath, Start_FileSize - 2)
Close #1
txtpath.Text = Temp_F
TmrUsers.Enabled = False
TmrRefresh.Enabled = False
TmrUpdateDisplay.Enabled = False
lblInformation.Visible = False
lblroom.Visible = False
txtMessage.Enabled = False
MnuChatOptions.Enabled = False
PicStaticRoomList.Visible = True
PicUsers.Visible = True
picStart.Visible = True
Label10.Enabled = False
lbltitle.Caption = "Error starting Stealth-Chat"
txtpath.Enabled = False
Label11.ForeColor = &HFF&
Label11.Caption = "There was an problem reading your name from the name file please re-type your name again in the box below"
End Sub
