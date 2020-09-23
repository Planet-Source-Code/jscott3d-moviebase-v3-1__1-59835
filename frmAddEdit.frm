VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{8BF89B33-148A-4373-8964-8BA63FDEA636}#1.0#0"; "FBButton.ocx"
Begin VB.Form frmAddEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmAddEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   1  'CenterOwner
   Begin FBButton.FlatButton cmdGetBrowser 
      Height          =   255
      Left            =   7080
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      Caption         =   "Get Info From Browser"
      HasFocusRect    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolbarImageList 
      Left            =   3120
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":1976
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":1F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":24AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":2A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":2FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":3578
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgs 
      Left            =   2400
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":3B12
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddEdit.frx":3E64
            Key             =   "IMG2"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboType 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6360
      TabIndex        =   5
      Top             =   690
      Width           =   2055
   End
   Begin VB.CommandButton cmdEditSupport 
      Caption         =   "..."
      Height          =   255
      Index           =   9
      Left            =   8520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   735
      Width           =   255
   End
   Begin VB.CommandButton cmdTitleSwitch 
      Caption         =   "< >"
      Height          =   285
      Left            =   5160
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox txtTitle 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   85
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   953
      ButtonWidth     =   1111
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "imgToolbarImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sv/Ex"
            Key             =   "SaveExit"
            Object.ToolTipText     =   "SaveExit"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Search"
            Key             =   "Search"
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Stop"
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Default"
            Key             =   "Default"
            Object.ToolTipText     =   "Default"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Edit"
                  Object.Tag             =   "Edit"
                  Text            =   "Edit"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reset"
            Key             =   "Reset"
            Object.ToolTipText     =   "Reset all Fields to Blank"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Help"
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmAddEdit.frx":41B6
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Technical"
      TabPicture(1)   =   "frmAddEdit.frx":41D2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Pictures"
      TabPicture(2)   =   "frmAddEdit.frx":41EE
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Search IMDB"
      TabPicture(3)   =   "frmAddEdit.frx":420A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame7 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   105
         Top             =   340
         Width           =   9135
         Begin VB.TextBox txtTemp 
            Alignment       =   2  'Center
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1515
            Left            =   1320
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   118
            Text            =   "frmAddEdit.frx":4226
            Top             =   1680
            Width           =   4815
         End
         Begin VB.CommandButton cmdGetChecked 
            Caption         =   "Get Checked"
            Default         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6720
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   5010
            Width           =   2295
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Fill Empty Fields Only."
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   4200
            Width           =   2295
         End
         Begin VB.TextBox txtMaxSearch 
            Enabled         =   0   'False
            Height          =   285
            Left            =   8160
            MaxLength       =   3
            TabIndex        =   107
            TabStop         =   0   'False
            Text            =   "20"
            ToolTipText     =   "Maximum Movies to Display"
            Top             =   3600
            Width           =   855
         End
         Begin VB.CommandButton cmdIMDBSearch 
            Caption         =   "Search"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6720
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   4680
            Width           =   2295
         End
         Begin MSComctlLib.ListView lstIMDBResults 
            Height          =   5055
            Left            =   120
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   240
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   8916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            Enabled         =   0   'False
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Match"
               Object.Width           =   7056
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "IMDB ID"
               Object.Width           =   3528
            EndProperty
         End
         Begin VB.Label Label35 
            Caption         =   "Note: You must be online to use this feature."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6720
            TabIndex        =   117
            Top             =   2880
            Width           =   2295
         End
         Begin VB.Label Label34 
            Caption         =   "Note: Please Check only one."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   116
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label Label32 
            Caption         =   " Double click on the title on the list to download the detail information of the movie."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   6720
            TabIndex        =   113
            Top             =   1800
            Width           =   2295
         End
         Begin VB.Label Label31 
            Caption         =   "If multiple matches found, they will be listed in the list on the left. "
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6720
            TabIndex        =   112
            Top             =   1320
            Width           =   2295
         End
         Begin VB.Label Label30 
            Caption         =   "Then press the Search Button on the Toolbar or below."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6720
            TabIndex        =   111
            Top             =   720
            Width           =   2295
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label29 
            Caption         =   "To preform a search, enter the title of the movie.  "
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6720
            TabIndex        =   110
            Top             =   240
            Width           =   2295
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label20 
            Caption         =   "Max # of Results:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   109
            Top             =   3615
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   96
         Top             =   360
         Width           =   9135
         Begin VB.Frame fras 
            Caption         =   "Front Cover"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Left            =   120
            TabIndex        =   104
            Top             =   600
            Width           =   3495
            Begin VB.Image imgFront 
               Height          =   4305
               Left            =   120
               Top             =   240
               Width           =   3255
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Back Cover"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Left            =   4680
            TabIndex        =   103
            Top             =   600
            Width           =   3495
            Begin VB.Image imgBack 
               Height          =   4305
               Left            =   120
               Top             =   240
               Width           =   3255
            End
         End
         Begin VB.CommandButton cmdBrowseFront 
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3720
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   960
            Width           =   735
         End
         Begin VB.CommandButton cmdBrowseBack 
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8280
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   960
            Width           =   735
         End
         Begin VB.CommandButton cmdClearFront 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3720
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   1440
            Width           =   735
         End
         Begin VB.CommandButton cmdClearBack 
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8280
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox txtFrontCoverLocation 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   240
            Width           =   3495
         End
         Begin VB.TextBox txtBackCoverLocation 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   240
            Width           =   3495
         End
      End
      Begin VB.Frame Frame5 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   84
         Top             =   340
         Width           =   9135
         Begin VB.CommandButton cmdAddRemoveFeatures 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   5880
            TabIndex        =   66
            Top             =   1440
            Width           =   255
         End
         Begin VB.CommandButton cmdAddRemoveFeatures 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   5880
            TabIndex        =   67
            Top             =   1920
            Width           =   255
         End
         Begin VB.Frame Frame11 
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   120
            TabIndex        =   95
            Top             =   2961
            Width           =   3075
            Begin VB.OptionButton optColor 
               Caption         =   "Color"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   62
               Top             =   240
               Width           =   915
            End
            Begin VB.OptionButton optColor 
               Caption         =   "Black/White"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1620
               TabIndex        =   63
               Top             =   240
               Width           =   1395
            End
         End
         Begin VB.ListBox lstSpecialFeatures 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            ItemData        =   "frmAddEdit.frx":424B
            Left            =   3360
            List            =   "frmAddEdit.frx":424D
            TabIndex        =   64
            Top             =   960
            Width           =   2415
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   11
            Left            =   5880
            TabIndex        =   65
            Top             =   960
            Width           =   255
         End
         Begin VB.ListBox lstTrailers 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            ItemData        =   "frmAddEdit.frx":424F
            Left            =   3360
            List            =   "frmAddEdit.frx":4251
            TabIndex        =   72
            Top             =   3240
            Width           =   2415
         End
         Begin VB.CommandButton cmdAddRemoveFeatures 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   5880
            TabIndex        =   75
            Top             =   4320
            Width           =   255
         End
         Begin VB.CommandButton cmdAddRemoveFeatures 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   5880
            TabIndex        =   74
            Top             =   3840
            Width           =   255
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   12
            Left            =   5880
            TabIndex        =   73
            Top             =   3360
            Width           =   255
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   10
            Left            =   4200
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   270
            Width           =   255
         End
         Begin VB.ComboBox cboScreenRatio 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1200
            Sorted          =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   2895
         End
         Begin VB.ListBox lstAudioTracks 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            ItemData        =   "frmAddEdit.frx":4253
            Left            =   6360
            List            =   "frmAddEdit.frx":4255
            TabIndex        =   68
            Top             =   960
            Width           =   2295
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   13
            Left            =   8760
            TabIndex        =   69
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton cmdAddRemoveFeatures 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   8760
            TabIndex        =   70
            Top             =   1440
            Width           =   255
         End
         Begin VB.CommandButton cmdAddRemoveFeatures 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   8760
            TabIndex        =   71
            Top             =   1920
            Width           =   255
         End
         Begin VB.ListBox lstSubtitles 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1860
            ItemData        =   "frmAddEdit.frx":4257
            Left            =   6360
            List            =   "frmAddEdit.frx":4259
            TabIndex        =   76
            Top             =   3240
            Width           =   2295
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   14
            Left            =   8760
            TabIndex        =   77
            Top             =   3240
            Width           =   255
         End
         Begin VB.CommandButton cmdAddRemoveFeatures 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   7
            Left            =   8760
            TabIndex        =   78
            Top             =   3720
            Width           =   255
         End
         Begin VB.CommandButton cmdAddRemoveFeatures 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   8760
            TabIndex        =   79
            Top             =   4200
            Width           =   255
         End
         Begin VB.Frame Frame10 
            Caption         =   "NTSC / PAL"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   675
            Left            =   120
            TabIndex        =   90
            Top             =   2128
            Width           =   3075
            Begin VB.OptionButton optNTSCPAL 
               Caption         =   "PAL"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   1620
               TabIndex        =   61
               Top             =   240
               Width           =   915
            End
            Begin VB.OptionButton optNTSCPAL 
               Alignment       =   1  'Right Justify
               Caption         =   "NTSC"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   60
               Top             =   240
               Width           =   915
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Tape Info"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   120
            TabIndex        =   87
            Top             =   3795
            Width           =   3075
            Begin VB.ComboBox cboTapeMode 
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               ItemData        =   "frmAddEdit.frx":425B
               Left            =   1200
               List            =   "frmAddEdit.frx":4268
               TabIndex        =   83
               Text            =   "cboTapeMode"
               Top             =   945
               Width           =   1695
            End
            Begin VB.TextBox txtNumberMovies 
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1200
               TabIndex        =   81
               Top             =   240
               Width           =   1695
            End
            Begin VB.TextBox txtFreeTime 
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1200
               TabIndex        =   82
               Top             =   600
               Width           =   1695
            End
            Begin VB.Label Label24 
               Caption         =   "Mode:"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   89
               Top             =   975
               Width           =   615
            End
            Begin VB.Label Label25 
               Caption         =   "# of Movies:"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   80
               Top             =   255
               Width           =   1095
            End
            Begin VB.Label Label26 
               Caption         =   "Free Time:"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   88
               Top             =   615
               Width           =   1095
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Disc Format"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   86
            Top             =   875
            Width           =   3075
            Begin VB.OptionButton optFormat 
               Caption         =   "Single Layer"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   1320
               TabIndex        =   57
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton optFormat 
               Caption         =   "Flipper"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   1320
               TabIndex        =   59
               Top             =   720
               Width           =   975
            End
            Begin VB.OptionButton optFormat 
               Caption         =   "Dual Layer"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   56
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton optFormat 
               Caption         =   "Dual-Sided"
               BeginProperty Font 
                  Name            =   "Comic Sans MS"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   120
               TabIndex        =   58
               Top             =   720
               Width           =   1215
            End
         End
         Begin VB.Label Label27 
            Caption         =   "Special Features:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   94
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label28 
            Caption         =   "Trailers:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3360
            TabIndex        =   93
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "Screen Ratio:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   300
            Width           =   1575
         End
         Begin VB.Label Label21 
            Caption         =   "Audio Tracks:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6360
            TabIndex        =   92
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label23 
            Caption         =   "Subtitles:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6360
            TabIndex        =   91
            Top             =   3000
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5415
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9135
         Begin VB.CommandButton cmdDate 
            Caption         =   "..."
            Height          =   255
            Index           =   0
            Left            =   5760
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   4410
            Width           =   255
         End
         Begin VB.CommandButton cmdDate 
            Caption         =   "..."
            Height          =   255
            Index           =   1
            Left            =   8760
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   4410
            Width           =   255
         End
         Begin VB.CommandButton cmdDate 
            Caption         =   "..."
            Height          =   255
            Index           =   2
            Left            =   8760
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   5025
            Width           =   255
         End
         Begin VB.ComboBox cboGenre 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   960
            Sorted          =   -1  'True
            TabIndex        =   10
            Top             =   225
            Width           =   3135
         End
         Begin VB.ComboBox cboEdition 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   960
            Sorted          =   -1  'True
            TabIndex        =   15
            Top             =   930
            Width           =   3135
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   270
            Width           =   255
         End
         Begin VB.ComboBox cboLocation 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5520
            Sorted          =   -1  'True
            TabIndex        =   30
            Top             =   2310
            Width           =   3135
         End
         Begin VB.TextBox txtDatePurched 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4680
            TabIndex        =   44
            Top             =   4395
            Width           =   975
         End
         Begin VB.TextBox txtCost 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   48
            Top             =   5010
            Width           =   1095
         End
         Begin VB.ComboBox cboNumberDisc 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmAddEdit.frx":4279
            Left            =   4680
            List            =   "frmAddEdit.frx":429B
            TabIndex        =   50
            Text            =   "cboNumberDisc"
            Top             =   4995
            Width           =   975
         End
         Begin VB.TextBox txtLength 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   960
            TabIndex        =   41
            Top             =   4395
            Width           =   1095
         End
         Begin VB.TextBox txtDVDDate 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7560
            TabIndex        =   52
            Top             =   5010
            Width           =   1095
         End
         Begin VB.TextBox txtMovieDate 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7560
            TabIndex        =   46
            Top             =   4395
            Width           =   1095
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   5
            Left            =   8760
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   3030
            Width           =   255
         End
         Begin VB.ComboBox cboRating 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5520
            TabIndex        =   36
            Top             =   3000
            Width           =   3135
         End
         Begin VB.ComboBox cboRegion 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   960
            Sorted          =   -1  'True
            TabIndex        =   33
            Top             =   3000
            Width           =   3135
         End
         Begin VB.ComboBox cboStudio 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   960
            Sorted          =   -1  'True
            TabIndex        =   21
            Top             =   1620
            Width           =   3135
         End
         Begin VB.ComboBox cboDirector 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5520
            Sorted          =   -1  'True
            TabIndex        =   18
            Top             =   930
            Width           =   3135
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   6
            Left            =   8760
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   960
            Width           =   255
         End
         Begin VB.ComboBox cboSubGenre 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5520
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   225
            Width           =   3135
         End
         Begin VB.ComboBox cboSeries 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   5520
            Sorted          =   -1  'True
            TabIndex        =   24
            Top             =   1620
            Width           =   3135
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   7
            Left            =   8760
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1650
            Width           =   255
         End
         Begin VB.ComboBox cboPackaging 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   960
            Sorted          =   -1  'True
            TabIndex        =   27
            Top             =   2310
            Width           =   3135
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   8
            Left            =   8760
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   2340
            Width           =   255
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   4
            Left            =   4200
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   3030
            Width           =   255
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   2
            Left            =   4200
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1650
            Width           =   255
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   1
            Left            =   4200
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton cmdEditSupport 
            Caption         =   "..."
            Height          =   255
            Index           =   3
            Left            =   4200
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   2340
            Width           =   255
         End
         Begin VB.ComboBox cboUserReview 
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "frmAddEdit.frx":42BE
            Left            =   1320
            List            =   "frmAddEdit.frx":42E4
            TabIndex        =   39
            Top             =   3690
            Width           =   3135
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   9
            Left            =   8400
            Picture         =   "frmAddEdit.frx":4398
            Top             =   3600
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   8
            Left            =   8080
            Picture         =   "frmAddEdit.frx":4922
            Top             =   3720
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   7
            Left            =   7755
            Picture         =   "frmAddEdit.frx":4EAC
            Top             =   3600
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   6
            Left            =   7440
            Picture         =   "frmAddEdit.frx":5436
            Top             =   3720
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   5
            Left            =   7125
            Picture         =   "frmAddEdit.frx":59C0
            Top             =   3600
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label2 
            Caption         =   "Genre:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   300
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Location:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4680
            TabIndex        =   29
            Top             =   2370
            Width           =   975
         End
         Begin VB.Label Label18 
            Caption         =   "Date Purchased:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   43
            Top             =   4440
            Width           =   1335
         End
         Begin VB.Label Label17 
            Caption         =   "Cost:        $"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   5055
            Width           =   975
         End
         Begin VB.Label Label16 
            Caption         =   "# of Discs (Tape):"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   49
            Top             =   5055
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Min."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2160
            TabIndex        =   42
            Top             =   4440
            Width           =   375
         End
         Begin VB.Label Label14 
            Caption         =   "Length:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   40
            Top             =   4440
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "DVD Date:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6600
            TabIndex        =   51
            Top             =   5055
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Movie Date:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6600
            TabIndex        =   45
            Top             =   4440
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Rating:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4680
            TabIndex        =   35
            Top             =   3060
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Region:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Top             =   3060
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Studio:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Director:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4680
            TabIndex        =   17
            Top             =   990
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Sub-Genre:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4560
            TabIndex        =   12
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label7 
            Caption         =   "Edition:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   990
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Series:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   4680
            TabIndex        =   23
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Packaging:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   26
            Top             =   2370
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "User Review:"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   4
            Left            =   6800
            Picture         =   "frmAddEdit.frx":5F4A
            Top             =   3720
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   3
            Left            =   6480
            Picture         =   "frmAddEdit.frx":64D4
            Top             =   3615
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   2
            Left            =   6160
            Picture         =   "frmAddEdit.frx":6A5E
            Top             =   3735
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   1
            Left            =   5835
            Picture         =   "frmAddEdit.frx":6FE8
            Top             =   3615
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image picStar 
            Height          =   240
            Index           =   0
            Left            =   5520
            Picture         =   "frmAddEdit.frx":7572
            Top             =   3734
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
   Begin VB.Label Label33 
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   735
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   735
      Width           =   495
   End
End
Attribute VB_Name = "frmAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim response As String

'------------------------------------------------------------------
' Publics
'------------------------------------------------------------------
'----- Adds to each table for types
Public Sub AddToTbls()
    On Error Resume Next
    With rsType
        .FindFirst "[Type] Like '" & cboType.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Type = cboType.Text
            .Update
        End If
    End With
    With rsGenre
        .FindFirst "[Genre] Like '" & cboGenre.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Genre = cboGenre.Text
            .Update
        End If
    End With
    With rsEdition
        .FindFirst "[Edition] Like '" & cboEdition.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Edition = cboEdition.Text
            .Update
        End If
    End With
    With rsDirector
        .FindFirst "[Director] Like '" & cboDirector.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Director = cboDirector.Text
            .Update
        End If
    End With
    With rsStudio
        .FindFirst "[Studio] Like '" & cboStudio.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Studio = cboStudio.Text
            .Update
        End If
    End With
    With rsSeries
        .FindFirst "[Series] Like '" & cboSeries.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Series = cboSeries.Text
            .Update
        End If
    End With
    With rsPackaging
        .FindFirst "[Packaging] Like '" & cboPackaging.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Packaging = cboPackaging.Text
            .Update
        End If
    End With
    With rsLocation
        .FindFirst "[Location] Like '" & cboLocation.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Location = cboLocation.Text
            .Update
        End If
    End With
    With rsRegion
        .FindFirst "[Region] Like '" & cboRegion.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Region = cboRegion.Text
            .Update
        End If
    End With
    With rsRating
        .FindFirst "[Ratings] Like '" & cboRating.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !Ratings = cboRating.Text
            .Update
        End If
    End With
    With rsScreenRatio
        .FindFirst "[ScreenRatio] Like '" & cboScreenRatio.Text & "'"
        If .NoMatch = True Then
            .AddNew
            !ScreenRatio = cboScreenRatio.Text
            .Update
        End If
    End With
End Sub

'----- Reset all
Public Sub Reset()
    On Error GoTo err
    txtTitle.Text = ""
    '    cboType.Text = ""
    '    cboGenre.Text = ""
    '    cboSubGenre.Text = ""
    '    cboEdition.Text = ""
    '    cboDirector.Text = ""
    '    cboStudio.Text = ""
    '    cboSeries.Text = ""
    '    cboPackaging.Text = ""
    '    cboLocation.Text = ""
    '    cboRegion.Text = ""
    '    cboRating.Text = ""
    
    cboUserReview.Text = ""
    
    txtLength.Text = ""
    txtCost.Text = ""
    txtDatePurched.Text = ""
    cboNumberDisc.Text = ""
    txtMovieDate.Text = ""
    txtDVDDate.Text = ""
    '// ------ Tech
    '    cboScreenRatio.Text = ""

    optFormat.Item(0).Value = False
    optFormat.Item(1).Value = False
    optFormat.Item(2).Value = False
    optFormat.Item(3).Value = False

    optNTSCPAL.Item(0).Value = False
    optNTSCPAL.Item(1).Value = False

    optColor.Item(0).Value = False
    optColor.Item(1).Value = False

    txtNumberMovies.Text = ""
    cboTapeMode.Text = ""
    txtFreeTime.Text = ""

    lstSpecialFeatures.Clear
    lstTrailers.Clear
    lstAudioTracks.Clear
    lstSubtitles.Clear

    txtFrontCoverLocation.Text = ""
    txtBackCoverLocation.Text = ""

    imgFront.Picture = LoadPicture
    imgBack.Picture = LoadPicture

    ClearCBO

    If DefaultStartup = 1 Then LoadDefault
    Exit Sub
err:
    MsgBox err.Description & "      " & err.Number
End Sub
'----- Save
Public Sub SaveTitle()
    On Error Resume Next
    With rsMovies
    
        If Edit = True Then
            If txtTitle.Text = .Fields("Title") Then
                .Edit
            ElseIf txtTitle.Text <> .Fields("Title") Then
                .AddNew
            End If
        Else
            .AddNew
        End If
        '// ----- General

        If Left$(LCase(txtTitle.Text), 4) = "the " Then txtTitle.Text = Mid$(txtTitle.Text, 5) & ", " & Left$(txtTitle.Text, 3)

        !Title = txtTitle.Text

        !Type = cboType.Text
        !Genre = cboGenre.Text
        !SubGenre = cboSubGenre.Text
        !Edition = cboEdition.Text
        !Director = cboDirector.Text
        !Studio = cboStudio.Text
        !Series = cboSeries.Text
        !Packaging = cboPackaging.Text
        !Location = cboLocation.Text
        !Region = cboRegion.Text
        !Rating = cboRating.Text
        !UserReview = cboUserReview.Text
        !Length = txtLength.Text
        !Cost = txtCost.Text
        !DatePurched = txtDatePurched.Text
        !NumberDisc = cboNumberDisc.Text
        !MovieDate = txtMovieDate.Text
        !DVDDate = txtDVDDate.Text
        '// ------ Tech
        !ScreenRatio = cboScreenRatio.Text

        If optFormat.Item(0).Value = True Then !DiscFormat = 0
        If optFormat.Item(1).Value = True Then !DiscFormat = 1
        If optFormat.Item(2).Value = True Then !DiscFormat = 2
        If optFormat.Item(3).Value = True Then !DiscFormat = 3

        If optNTSCPAL.Item(0).Value = True Then !NTSCPAL = 0
        If optNTSCPAL.Item(1).Value = True Then !NTSCPAL = 1

        If optColor.Item(0).Value = True Then !Color = 0
        If optColor.Item(1).Value = True Then !Color = 1

        !NumberMovies = txtNumberMovies.Text
        !TapeMode = cboTapeMode.Text
        !FreeTime = txtFreeTime.Text
        lstBreakdown lstSpecialFeatures
        !SpecialFeatures = lstBroken
        lstBreakdown lstTrailers
        !Trailers = lstBroken
        lstBreakdown lstAudioTracks
        !AudioTracks = lstBroken
        lstBreakdown lstSubtitles
        !Subtitles = lstBroken
        !FrontCover = txtFrontCoverLocation.Text
        !BackCover = txtBackCoverLocation.Text
        .Update
    End With
End Sub

'----- Gets Data from dbase for editings
Public Function GetViewData(SearchText As String)
    On Error Resume Next '// iff error then exit the current sub
    Dim strSearch As String '// creats a string

    strSearch = "[Title] Like '" & SearchText & "'" '// compares frmMain.lstnames.selected to a record
    With rsMovies '// opens the recordset
        .FindFirst strSearch '// searches for the record
        txtTitle.Text = .Fields("Title")
        cboType.Text = .Fields("Type")
        '// ----- General

        cboGenre.Text = .Fields("Genre")
        cboSubGenre.Text = .Fields("SubGenre")
        cboEdition.Text = .Fields("Edition")
        cboDirector.Text = .Fields("Director")
        cboStudio.Text = .Fields("Studio")
        cboSeries.Text = .Fields("Series")
        cboPackaging.Text = .Fields("Packaging")
        cboLocation.Text = .Fields("Location")
        cboRegion.Text = .Fields("Region")
        cboRating.Text = .Fields("Rating")
        cboUserReview.Text = .Fields("UserReview")
        txtLength.Text = .Fields("Length")
        txtCost.Text = .Fields("Cost")
        txtDatePurched.Text = .Fields("DatePurched")
        cboNumberDisc.Text = .Fields("NumberDisc")
        txtMovieDate.Text = .Fields("MovieDate")
        txtDVDDate.Text = .Fields("DVDDate")
        '// ------ Technical
        cboScreenRatio.Text = .Fields("ScreenRatio")
        optFormat(.Fields("DiscFormat")).Value = True
        optNTSCPAL(.Fields("NTSCPAL")).Value = True
        optColor(.Fields("Color")).Value = True
        txtNumberMovies.Text = .Fields("NumberMovies")
        cboTapeMode.Text = .Fields("TapeMode")
        txtFreeTime.Text = .Fields("FreeTime")
        lstReBuild lstSpecialFeatures, .Fields("SpecialFeatures")
        lstReBuild lstTrailers, .Fields("Trailers")
        lstReBuild lstAudioTracks, .Fields("AudioTracks")
        lstReBuild lstSubtitles, .Fields("Subtitles")
        '// ------ Pictures
        txtFrontCoverLocation.Text = .Fields("FrontCover")
        imgFront.Picture = LoadPicture(.Fields("FrontCover"))
        txtBackCoverLocation.Text = .Fields("BackCover")
        imgBack.Picture = LoadPicture(.Fields("BackCover"))

    End With
End Function

'----- Loads Default Settings
Public Sub LoadDefault()
    On Error Resume Next
    With rsDefault
        .MoveFirst
        cboType.Text = .Fields("Type")
        cboPackaging.Text = .Fields("Packaging")
        cboLocation.Text = .Fields("Location")
        cboRegion.Text = .Fields("Region")
        cboScreenRatio.Text = .Fields("ScreenRatio")
        cboNumberDisc.Text = .Fields("NumofDisc")
        optFormat(.Fields("DiscFormat")).Value = True
        optNTSCPAL(.Fields("NTSCPAL")).Value = True
        optColor(.Fields("Color")).Value = True
    End With

End Sub
'----- Clears all cbos
Public Sub ClearCBO()
    On Error Resume Next
    cboGenre.Clear
    cboSubGenre.Clear
    cboType.Clear
    cboStudio.Clear
    cboEdition.Clear
    cboDirector.Clear
    cboPackaging.Clear
    cboSeries.Clear
    cboRegion.Clear
    cboRating.Clear
    cboLocation.Clear
    cboScreenRatio.Clear
        picStar(0).Visible = False
        picStar(1).Visible = False
        picStar(2).Visible = False
        picStar(3).Visible = False
        picStar(4).Visible = False
        picStar(5).Visible = False
        picStar(6).Visible = False
        picStar(7).Visible = False
        picStar(8).Visible = False
        picStar(9).Visible = False
    LoadCBO

End Sub
'---------- Load All cbos
Public Sub LoadCBO()
    On Error Resume Next
    rsType.MoveFirst
    Do While Not rsType.EOF
        cboType.AddItem rsType.Fields("Type")
        rsType.MoveNext
    Loop
    rsGenre.MoveFirst
    Do While Not rsGenre.EOF
        cboGenre.AddItem rsGenre.Fields("Genre")
        cboSubGenre.AddItem rsGenre.Fields("Genre")
        rsGenre.MoveNext
    Loop
    rsStudio.MoveFirst
    Do While Not rsStudio.EOF
        cboStudio.AddItem rsStudio.Fields("Studio")
        rsStudio.MoveNext
    Loop
    rsEdition.MoveFirst
    Do While Not rsEdition.EOF
        cboEdition.AddItem rsEdition.Fields("Edition")
        rsEdition.MoveNext
    Loop
    rsDirector.MoveFirst
    Do While Not rsDirector.EOF
        cboDirector.AddItem rsDirector.Fields("Director")
        rsDirector.MoveNext
    Loop
    rsPackaging.MoveFirst
    Do While Not rsPackaging.EOF
        cboPackaging.AddItem rsPackaging.Fields("Packaging")
        rsPackaging.MoveNext
    Loop
    rsSeries.MoveFirst
    Do While Not rsSeries.EOF
        cboSeries.AddItem rsSeries.Fields("Series")
        rsSeries.MoveNext
    Loop
    rsRegion.MoveFirst
    Do While Not rsRegion.EOF
        cboRegion.AddItem rsRegion.Fields("Region")
        rsRegion.MoveNext
    Loop
    rsRating.MoveFirst
    Do While Not rsRating.EOF
        cboRating.AddItem rsRating.Fields("Ratings")
        rsRating.MoveNext
    Loop
    rsLocation.MoveFirst
    Do While Not rsLocation.EOF
        cboLocation.AddItem rsLocation.Fields("Location")
        rsLocation.MoveNext
    Loop
    rsScreenRatio.MoveFirst
    Do While Not rsScreenRatio.EOF
        cboScreenRatio.AddItem rsScreenRatio.Fields("ScreenRatio")
        rsScreenRatio.MoveNext
    Loop

End Sub

Private Sub cmdDate_Click(Index As Integer)
On Error Resume Next
    Select Case Index
        Case 0
            txtDatePurched.Text = Date
        Case 1
            txtMovieDate.Text = Year(Date)
        Case 2
            txtDVDDate.Text = Year(Date)
    End Select

End Sub

Private Sub cmdGetBrowser_Click()
On Error Resume Next
    response = MsgBox("Would you like to add data from the browser?  " & vbCrLf & "If so you have to have the browser to the page you want." & vbCrLf & "Note:  It has to be from DVDEmpire.com", 36, "Message Box Title")
    If response = vbNo Then Exit Sub
    
End Sub

'------------------------------------------------------------------
' Form Controls
'------------------------------------------------------------------
'--------------------------- Load
Private Sub Form_Load()
    On Error Resume Next
    Me.Width = 9720 '640 pixels
    Me.Height = 7600 '480 pixels

    If Edit Then
        Me.Caption = "Edit Movie:       " & EditID 'txtDescription.Text
        
        StatusBarMsg "Oh Great!      Now you want me to edit (""" & EditID & """)  You must be Mad.", 1

    Else
        Me.Caption = "Add Movie"
        Toolbar1.Buttons.Item(1).Enabled = False
        Toolbar1.Buttons.Item(2).Enabled = False
        
        StatusBarMsg "Why in the Hell would you want to add a movie.", 1

    End If

    'LoadCBO
    Reset
End Sub
'---------- Unload
Private Sub Form_Unload(Cancel As Integer)
    StatusBarMsg "Welcome to the worst movie database program ever.", 1
End Sub


'------------------------------------------------------------------
' Toolbars
'------------------------------------------------------------------
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Save"
            SaveTitle
            AddToTbls
            fillDVDTreeView frmMain.treMovieList, "Title"
            Reset
        Case "SaveExit"
            SaveTitle
            AddToTbls
            fillDVDTreeView frmMain.treMovieList, "Title"
            Unload Me
        Case "Search"
        Case "Default"
            LoadDefault
        Case "Reset"
            Reset
        Case "Exit"
            Unload Me
    End Select
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    On Error Resume Next
    Select Case ButtonMenu.Key
        Case "Edit"
            frmDefault.Show , frmAddEdit
    End Select
End Sub
'------------------------------------------------------------------
' Controls
'------------------------------------------------------------------
'---------- Command Buttons
'----- Text Switcher ---- Changes Title  ex. " The Bloomers " to " Bloomers, The "
Private Sub cmdTitleSwitch_Click()
    On Error Resume Next
    If Left$(LCase(txtTitle.Text), 4) = "the " Then
        txtTitle.Text = Mid$(txtTitle.Text, 5) & ", " & Left$(txtTitle.Text, 3)
        Exit Sub
    End If
    If Right$(LCase(txtTitle.Text), 5) = ", the" Then
        txtTitle.Text = Right$(txtTitle.Text, 3) & " " & Left$(txtTitle.Text, Len(txtTitle.Text) - 5)
        Exit Sub
    End If

End Sub
'----- Edit Tables
Private Sub cmdEditSupport_Click(Index As Integer)
    On Error Resume Next

    intSelect = Index

    Select Case Index
        Case 11 'Special Features.
            frmSelect.Show , frmAddEdit
            frmSelect.Caption = "Special Features"
            frmSelect.GetChecked lstSpecialFeatures
        Case 12 'Trailers
            frmSelect.Show , frmAddEdit
            frmSelect.Caption = "Trailers"
            frmSelect.GetChecked lstTrailers
        Case 13 'Audio Tracks
            frmSelect.Show , frmAddEdit
            frmSelect.Caption = "Audio Tracks"
            frmSelect.GetChecked lstAudioTracks
        Case 14 'Subtitles
            frmSelect.Show , frmAddEdit
            frmSelect.Caption = "Subtitles"
            frmSelect.GetChecked lstSubtitles
        Case Else
            frmEditSupport.Show , frmAddEdit
    End Select

End Sub
'----- AddEdit Special features
Private Sub cmdAddRemoveFeatures_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            RemoveListBoxItem lstSpecialFeatures.Text, frmAddEdit.lstSpecialFeatures
        Case 1
            RemoveListBoxItem lstTrailers.Text, frmAddEdit.lstTrailers
        Case 2
            RemoveListBoxItem lstAudioTracks.Text, frmAddEdit.lstAudioTracks
        Case 3
            RemoveListBoxItem lstSubtitles.Text, frmAddEdit.lstSubtitles
        Case 4
            response = InputBox("Extra Feature", "Features")
            If response = "" Then Exit Sub
            With rsSpecialFeatures
                .FindFirst "[SpecialFeature] Like '" & response & "'"
                If .NoMatch = True Then
                    .AddNew
                    !SpecialFeature = response
                    .Update
                    lstSpecialFeatures.AddItem response
                Else
                    lstSpecialFeatures.AddItem .Fields("SpecialFeature")
                End If
            End With
        Case 5
            response = InputBox("Extra Feature", "Features")
            If response = "" Then Exit Sub
            With rsTrailers
                .FindFirst "[Trailers] Like '" & response & "'"
                If .NoMatch = True Then
                    .AddNew
                    !Trailers = response
                    .Update
                    lstTrailers.AddItem response
                Else
                    lstTrailers.AddItem .Fields("Trailers")
                End If
            End With
        Case 6
            response = InputBox("Extra Feature", "Features")
            If response = "" Then Exit Sub
            With rsAudio
                .FindFirst "[Audio] Like '" & response & "'"
                If .NoMatch = True Then
                    .AddNew
                    !Audio = response
                    .Update
                    lstAudioTracks.AddItem response
                Else
                    lstAudioTracks.AddItem .Fields("Audio")
                End If
            End With
        Case 7
            response = InputBox("Extra Feature", "Features")
            If response = "" Then Exit Sub
            With rsSubtitles
                .FindFirst "[Subtitles] Like '" & response & "'"
                If .NoMatch = True Then
                    .AddNew
                    !Subtitles = response
                    .Update
                    lstSubtitles.AddItem response
                Else
                    lstSubtitles.AddItem .Fields("Subtitles")
                End If
            End With
    End Select

End Sub

'---------- Picture Commads
'----- Clears
Private Sub cmdClearBack_Click()
    On Error Resume Next
    imgBack.Picture = LoadPicture()
    txtBackCoverLocation.Text = ""
End Sub
Private Sub cmdClearFront_Click()
    On Error Resume Next
    imgFront.Picture = LoadPicture()
    txtFrontCoverLocation.Text = ""
End Sub

Private Sub cmdBrowseFront_Click()
    CommonDialog1.CancelError = True

    On Error GoTo ErrHandler

    'CommonDialog1.Flags = cd10FNHideReadOnly

    CommonDialog1.Filter = "Image Files (*.jpg)|*.jpg"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen

    txtFrontCoverLocation.Text = CommonDialog1.FileName
    imgFront.Picture = LoadPicture(CommonDialog1.FileName)
ErrHandler:
    Exit Sub

End Sub
'----- Browser Buttons
Private Sub cmdBrowseBack_Click()
    CommonDialog1.CancelError = True

    On Error GoTo ErrHandler

    'CommonDialog1.Flags = cd10FNHideReadOnly

    CommonDialog1.Filter = "Image Files (*.jpg)|*.jpg"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.ShowOpen

    txtBackCoverLocation.Text = CommonDialog1.FileName
    imgBack.Picture = LoadPicture(CommonDialog1.FileName)
ErrHandler:
    Exit Sub

End Sub

'----- User Review Stars
Private Sub cboUserReview_Click()
    On Error Resume Next
    If cboUserReview.ListIndex = 0 Then
        picStar(0).Visible = False
        picStar(1).Visible = False
        picStar(2).Visible = False
        picStar(3).Visible = False
        picStar(4).Visible = False
        picStar(5).Visible = False
        picStar(6).Visible = False
        picStar(7).Visible = False
        picStar(8).Visible = False
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 1 Then
        picStar(0).Visible = True
        picStar(1).Visible = False
        picStar(2).Visible = False
        picStar(3).Visible = False
        picStar(4).Visible = False
        picStar(5).Visible = False
        picStar(6).Visible = False
        picStar(7).Visible = False
        picStar(8).Visible = False
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 2 Then
        picStar(0).Visible = True
        picStar(1).Visible = True
        picStar(2).Visible = False
        picStar(3).Visible = False
        picStar(4).Visible = False
        picStar(5).Visible = False
        picStar(6).Visible = False
        picStar(7).Visible = False
        picStar(8).Visible = False
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 3 Then
        picStar(0).Visible = True
        picStar(1).Visible = True
        picStar(2).Visible = True
        picStar(3).Visible = False
        picStar(4).Visible = False
        picStar(5).Visible = False
        picStar(6).Visible = False
        picStar(7).Visible = False
        picStar(8).Visible = False
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 4 Then
        picStar(0).Visible = True
        picStar(1).Visible = True
        picStar(2).Visible = True
        picStar(3).Visible = True
        picStar(4).Visible = False
        picStar(5).Visible = False
        picStar(6).Visible = False
        picStar(7).Visible = False
        picStar(8).Visible = False
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 5 Then
        picStar(0).Visible = True
        picStar(1).Visible = True
        picStar(2).Visible = True
        picStar(3).Visible = True
        picStar(4).Visible = True
        picStar(5).Visible = False
        picStar(6).Visible = False
        picStar(7).Visible = False
        picStar(8).Visible = False
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 6 Then
        picStar(0).Visible = True
        picStar(1).Visible = True
        picStar(2).Visible = True
        picStar(3).Visible = True
        picStar(4).Visible = True
        picStar(5).Visible = True
        picStar(6).Visible = False
        picStar(7).Visible = False
        picStar(8).Visible = False
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 7 Then
        picStar(0).Visible = True
        picStar(1).Visible = True
        picStar(2).Visible = True
        picStar(3).Visible = True
        picStar(4).Visible = True
        picStar(5).Visible = True
        picStar(6).Visible = True
        picStar(7).Visible = False
        picStar(8).Visible = False
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 8 Then
        picStar(0).Visible = True
        picStar(1).Visible = True
        picStar(2).Visible = True
        picStar(3).Visible = True
        picStar(4).Visible = True
        picStar(5).Visible = True
        picStar(6).Visible = True
        picStar(7).Visible = True
        picStar(8).Visible = False
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 9 Then
        picStar(0).Visible = True
        picStar(1).Visible = True
        picStar(2).Visible = True
        picStar(3).Visible = True
        picStar(4).Visible = True
        picStar(5).Visible = True
        picStar(6).Visible = True
        picStar(7).Visible = True
        picStar(8).Visible = True
        picStar(9).Visible = False
    ElseIf cboUserReview.ListIndex = 10 Then
        picStar(0).Visible = True
        picStar(1).Visible = True
        picStar(2).Visible = True
        picStar(3).Visible = True
        picStar(4).Visible = True
        picStar(5).Visible = True
        picStar(6).Visible = True
        picStar(7).Visible = True
        picStar(8).Visible = True
        picStar(9).Visible = True

    End If
End Sub

Private Sub txtTitle_Change()
    On Error Resume Next
    If txtTitle.Text = "" Then
        Toolbar1.Buttons.Item(1).Enabled = False
        Toolbar1.Buttons.Item(2).Enabled = False
    Else
        Toolbar1.Buttons.Item(1).Enabled = True
        Toolbar1.Buttons.Item(2).Enabled = True
    End If
End Sub

'------------  Combo boxs Autocomplete
Private Sub cboType_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboType, KeyAscii, iLeftOff
End Sub

Private Sub cboGenre_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboGenre, KeyAscii, iLeftOff
End Sub
Private Sub cboSubGenre_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboSubGenre, KeyAscii, iLeftOff
End Sub
Private Sub cboEdition_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboEdition, KeyAscii, iLeftOff
End Sub
Private Sub cboDirector_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboDirector, KeyAscii, iLeftOff
End Sub
Private Sub cboStudio_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboStudio, KeyAscii, iLeftOff
End Sub
Private Sub cboSeries_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboSeries, KeyAscii, iLeftOff
End Sub
Private Sub cboPackaging_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboPackaging, KeyAscii, iLeftOff
End Sub
Private Sub cboLocation_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboLocation, KeyAscii, iLeftOff
End Sub
Private Sub cboRegion_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboRegion, KeyAscii, iLeftOff
End Sub
Private Sub cboRating_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboRating, KeyAscii, iLeftOff
End Sub

Private Sub cboUserReview_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboUserReview, KeyAscii, iLeftOff
End Sub

Private Sub cboScreenRatio_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboScreenRatio, KeyAscii, iLeftOff
End Sub

Private Sub cboTapeMode_KeyPress(KeyAscii As Integer)
    Static iLeftOff As Long
    ComboAutoComplete cboTapeMode, KeyAscii, iLeftOff
End Sub
