VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "comct232.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ControlsXP"
   ClientHeight    =   4968
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6288
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4968
   ScaleWidth      =   6288
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   816
      ItemData        =   "Form1.frx":0000
      Left            =   4080
      List            =   "Form1.frx":0013
      TabIndex        =   15
      Top             =   3720
      Width           =   1812
   End
   Begin VB.ListBox List1 
      Height          =   696
      ItemData        =   "Form1.frx":0058
      Left            =   4080
      List            =   "Form1.frx":006B
      Style           =   1  'Checkbox
      TabIndex        =   14
      Top             =   2880
      Width           =   1812
   End
   Begin VB.ComboBox Combo1 
      Height          =   288
      ItemData        =   "Form1.frx":00CE
      Left            =   4080
      List            =   "Form1.frx":00DB
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   2280
      Width           =   1812
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1812
      LargeChange     =   25
      Left            =   3480
      Max             =   100
      TabIndex        =   12
      Top             =   2760
      Value           =   50
      Width           =   252
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   252
      LargeChange     =   33
      Left            =   360
      Max             =   99
      TabIndex        =   11
      Top             =   4320
      Value           =   49
      Width           =   2772
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   612
      Left            =   2760
      TabIndex        =   10
      Top             =   360
      Width           =   1212
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Option3"
      Height          =   252
      Left            =   2880
      TabIndex        =   9
      Top             =   1920
      Width           =   852
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Option2"
      Height          =   252
      Left            =   2880
      TabIndex        =   8
      Top             =   1560
      Width           =   1104
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   252
      Left            =   2880
      TabIndex        =   7
      Top             =   1200
      Value           =   -1  'True
      Width           =   852
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1572
      Left            =   4320
      TabIndex        =   3
      Top             =   360
      Width           =   1572
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   288
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   960
         Width           =   288
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   252
         Left            =   420
         TabIndex        =   4
         Top             =   480
         Width           =   852
      End
      Begin ComCtl2.UpDown UpDown1 
         Height          =   288
         Left            =   768
         TabIndex        =   6
         Top             =   960
         Width           =   252
         _ExtentX        =   445
         _ExtentY        =   508
         _Version        =   327681
         Value           =   5
         BuddyControl    =   "Text1"
         BuddyDispid     =   196618
         OrigLeft        =   1800
         OrigTop         =   840
         OrigRight       =   2052
         OrigBottom      =   1092
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
   End
   Begin ComctlLib.Slider Slider1 
      Height          =   492
      Left            =   360
      TabIndex        =   2
      Top             =   2760
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   868
      _Version        =   327682
      LargeChange     =   10
      Max             =   100
      TickFrequency   =   10
   End
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   360
      TabIndex        =   1
      Top             =   3720
      Width           =   2772
      _ExtentX        =   4890
      _ExtentY        =   445
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   2052
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   3620
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tab 1"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tab 2"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Tab 3"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()
  InitControlsXP
End Sub

Private Sub Timer1_Timer()
  ProgressBar1.Value = ProgressBar1.Value + 0.5
  If ProgressBar1.Value >= 100 Then ProgressBar1.Value = 0
End Sub
