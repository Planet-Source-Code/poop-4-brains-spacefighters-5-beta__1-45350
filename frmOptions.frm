VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Spacefighters 5 - Options"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   346
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssOpts 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "About"
      TabPicture(0)   =   "frmOptions.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAbout"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblEAbout"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblWAbout"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblTAbout"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtEmail"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtWeb"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Keys"
      TabPicture(1)   =   "frmOptions.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblNot"
      Tab(1).Control(1)=   "lblAKeys"
      Tab(1).Control(2)=   "lblTKeys"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Speed"
      TabPicture(2)   =   "frmOptions.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "hsSpeed"
      Tab(2).Control(1)=   "lblScSpeed"
      Tab(2).Control(2)=   "lblFcSpeed"
      Tab(2).Control(3)=   "lblASpeed"
      Tab(2).Control(4)=   "lblFsSpeed"
      Tab(2).Control(5)=   "lblSlSpeed"
      Tab(2).Control(6)=   "lblTSpeed"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Help"
      TabPicture(3)   =   "frmOptions.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "wbHelp"
      Tab(3).Control(1)=   "lblAHelp"
      Tab(3).Control(2)=   "lblTHelp"
      Tab(3).ControlCount=   3
      Begin SHDocVwCtl.WebBrowser wbHelp 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   19
         Top             =   1080
         Width           =   6135
         ExtentX         =   10821
         ExtentY         =   5741
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.HScrollBar hsSpeed 
         Height          =   375
         LargeChange     =   1000
         Left            =   -74160
         Max             =   -10000
         SmallChange     =   20
         TabIndex        =   9
         Top             =   2880
         Width           =   4695
      End
      Begin VB.TextBox txtWeb 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "http://kevgames.hostkingdom.net/"
         Top             =   3840
         Width           =   3495
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "kevinmf89@hotmail.com"
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label lblNot 
         BackStyle       =   0  'Transparent
         Caption         =   "NOT YET IMPLEMENTED!!!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   -74640
         TabIndex        =   21
         Top             =   1920
         Width           =   5655
      End
      Begin VB.Label lblAHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "HTML Help Library"
         Height          =   255
         Left            =   -72480
         TabIndex        =   20
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label lblTHelp 
         BackStyle       =   0  'Transparent
         Caption         =   "Help Files"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   -74640
         TabIndex        =   18
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label lblScSpeed 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Slower Computer"
         Height          =   255
         Left            =   -71520
         TabIndex        =   17
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label lblFcSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Faster Computer"
         Height          =   255
         Left            =   -74160
         TabIndex        =   16
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblAKeys 
         Caption         =   "Just select a key from the combo-box next to the player's number to adjust the key controls."
         Height          =   855
         Left            =   -74520
         TabIndex        =   15
         Top             =   1080
         Width           =   5415
      End
      Begin VB.Label lblTKeys 
         Caption         =   "Key Controls"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   -74640
         TabIndex        =   14
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label lblASpeed 
         Caption         =   "Move the scrollbar below to adjust the speed of the game. The speed of the game also involves the speed of your computer."
         Height          =   855
         Left            =   -74520
         TabIndex        =   13
         Top             =   1080
         Width           =   5295
      End
      Begin VB.Label lblFsSpeed 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Faster"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70440
         TabIndex        =   12
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblSlSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Slower"
         Height          =   255
         Left            =   -74160
         TabIndex        =   11
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblTSpeed 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed Control"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   -74640
         TabIndex        =   10
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label lblTAbout 
         BackStyle       =   0  'Transparent
         Caption         =   "About Spacefighters 5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   5655
      End
      Begin VB.Label lblWAbout 
         Caption         =   "Website:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label lblEAbout 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   3480
         Width           =   975
      End
      Begin VB.Label lblAbout 
         BackStyle       =   0  'Transparent
         Height          =   2295
         Left            =   480
         TabIndex        =   3
         Top             =   1080
         Width           =   5295
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   4680
      Width           =   2055
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOk_Click()
Speed = hsSpeed.Value + 10000

Unload Me
End Sub

Private Sub Form_Load()
lblAbout.Caption = "This game was though up and developed by Kevin Fleet. I also made the graphics"

hsSpeed.Value = Speed - 10000

wbHelp.Navigate App.Path & "\help\index.htm"
End Sub
