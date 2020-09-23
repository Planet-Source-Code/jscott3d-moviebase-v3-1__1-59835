VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{8BF89B33-148A-4373-8964-8BA63FDEA636}#1.0#0"; "FBButton.ocx"
Object = "{EF3356C9-9F5E-4EBC-84EC-7A665520E422}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   9330
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   622
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.TreeView treMovieList 
      Height          =   6600
      Left            =   120
      TabIndex        =   54
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   11642
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   35
      LineStyle       =   1
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "imgTreeViewImages"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lstMovies 
      Height          =   2775
      Left            =   3960
      TabIndex        =   55
      Top             =   960
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "imgTreeViewImages"
      SmallIcons      =   "imgTreeViewImages"
      ColHdrIcons     =   "imgTreeViewImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Title"
         Text            =   "Title"
         Object.Width           =   6615
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Genre"
         Text            =   "Genre"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Director"
         Text            =   "Director"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Rating"
         Text            =   "Rating"
         Object.Width           =   5689
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "Year"
         Text            =   "Year"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "Format"
         Text            =   "Format"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "Price"
         Text            =   "Price"
         Object.Width           =   1852
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   2760
      Left            =   4680
      TabIndex        =   56
      Top             =   4800
      Width           =   6495
      ExtentX         =   11456
      ExtentY         =   4868
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Height          =   435
      Left            =   3480
      TabIndex        =   40
      Top             =   8280
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   767
      BandCount       =   2
      VariantHeight   =   0   'False
      _CBWidth        =   8520
      _CBHeight       =   435
      _Version        =   "6.7.8988"
      Child1          =   "Picture2"
      MinHeight1      =   25
      Width1          =   562
      NewRow1         =   0   'False
      BandStyle1      =   1
      Child2          =   "ProgressBar"
      MinHeight2      =   25
      NewRow2         =   0   'False
      BandStyle2      =   1
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   30
         ScaleHeight     =   375
         ScaleWidth      =   8430
         TabIndex        =   42
         Top             =   30
         Width           =   8430
         Begin FBButton.FlatButton cmdBack 
            Height          =   330
            Left            =   4080
            TabIndex        =   50
            ToolTipText     =   "Back"
            Top             =   25
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            PicturehDC      =   1
            HasPicture      =   -1  'True
            HasCaption      =   0   'False
            Caption         =   "GO"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdForward 
            Height          =   330
            Left            =   4440
            TabIndex        =   51
            ToolTipText     =   "Forward"
            Top             =   25
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            PicturehDC      =   1
            HasPicture      =   -1  'True
            HasCaption      =   0   'False
            Caption         =   "GO"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdCancel 
            Height          =   330
            Left            =   4800
            TabIndex        =   52
            ToolTipText     =   "Cancel"
            Top             =   25
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            PicturehDC      =   1
            HasPicture      =   -1  'True
            HasCaption      =   0   'False
            Caption         =   "GO"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdRefresh 
            Height          =   330
            Left            =   5160
            TabIndex        =   53
            ToolTipText     =   "Refresh"
            Top             =   25
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            PicturehDC      =   1
            HasPicture      =   -1  'True
            HasCaption      =   0   'False
            Caption         =   "GO"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdGo 
            Height          =   330
            Left            =   3720
            TabIndex        =   44
            ToolTipText     =   "Go"
            Top             =   25
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   582
            PicturehDC      =   1
            HasPicture      =   -1  'True
            HasCaption      =   0   'False
            Caption         =   "GO"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtAddress 
            Height          =   330
            Left            =   0
            TabIndex        =   43
            Top             =   25
            Width           =   3735
         End
      End
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   375
         Left            =   8490
         TabIndex        =   41
         Top             =   30
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.PictureBox picResizeTB 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   4440
      MousePointer    =   7  'Size N S
      ScaleHeight     =   135
      ScaleWidth      =   3255
      TabIndex        =   37
      Top             =   4200
      Width           =   3255
   End
   Begin VB.PictureBox picResizeLR 
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   3360
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3255
      ScaleWidth      =   135
      TabIndex        =   38
      Top             =   840
      Width           =   135
   End
   Begin MSComctlLib.ImageList imgToolbarImageList 
      Left            =   10680
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1976
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   9960
      Top             =   8520
      _ExtentX        =   847
      _ExtentY        =   847
      BmpCount        =   14
      Bmp:1           =   "frmMain.frx":2FDE
      Key:1           =   "#mnuViewOptions"
      Bmp:2           =   "frmMain.frx":3406
      Key:2           =   "#mnuToolsEditSupport"
      Bmp:3           =   "frmMain.frx":382E
      Key:3           =   "#mnuFileNew"
      Bmp:4           =   "frmMain.frx":3C56
      Key:4           =   "#mnuPrint"
      Bmp:5           =   "frmMain.frx":407E
      Key:5           =   "#mnuPrintPrintPreview"
      Bmp:6           =   "frmMain.frx":44A6
      Key:6           =   "#mnuPrintPrint"
      Bmp:7           =   "frmMain.frx":48CE
      Key:7           =   "#mnuFileExit"
      Bmp:8           =   "frmMain.frx":4CF6
      Key:8           =   "#mnuHelpContents"
      Bmp:9           =   "frmMain.frx":511E
      Key:9           =   "#mnuHelpAbout"
      Bmp:10          =   "frmMain.frx":5546
      Key:10          =   "#mnuTitleAddMovie"
      Bmp:11          =   "frmMain.frx":596E
      Key:11          =   "#mnuTitleEditThis"
      Bmp:12          =   "frmMain.frx":5D96
      Key:12          =   "#mnuTitleDeleteThis"
      Bmp:13          =   "frmMain.frx":61BE
      Key:13          =   "#mnuToolsDefaults"
      Mask:14         =   16711935
      Key:14          =   "#mnuDatabaseBackup"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MenuDrawStyle   =   2
   End
   Begin MSComctlLib.ImageList imgTreeViewImages 
      Left            =   11400
      Top             =   8520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65E6
            Key             =   "Closed"
            Object.Tag             =   "Closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B80
            Key             =   "File"
            Object.Tag             =   "File"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":711A
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   9045
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15531
            Text            =   "Status Bar Message"
            TextSave        =   "Status Bar Message"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   1482
      BandCount       =   5
      FixedOrder      =   -1  'True
      _CBWidth        =   12000
      _CBHeight       =   840
      _Version        =   "6.7.8988"
      Child1          =   "Toolbar1"
      MinHeight1      =   22
      Width1          =   280
      NewRow1         =   0   'False
      Child2          =   "Picture3"
      MinHeight2      =   25
      Width2          =   327
      NewRow2         =   0   'False
      Caption3        =   "Total Movies: 888"
      MinHeight3      =   25
      Width3          =   93
      UseCoolbarPicture3=   0   'False
      NewRow3         =   0   'False
      BandStyle3      =   1
      Child4          =   "Picture4"
      MinHeight4      =   25
      Width4          =   280
      BandPicture4    =   "frmMain.frx":76B4
      NewRow4         =   -1  'True
      Child5          =   "Picture1"
      MinHeight5      =   25
      Width5          =   500
      NewRow5         =   0   'False
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   165
         ScaleHeight     =   375
         ScaleWidth      =   4005
         TabIndex        =   35
         Top             =   435
         Width           =   4005
         Begin VB.ComboBox cboTree 
            Height          =   345
            Left            =   0
            TabIndex        =   36
            Text            =   "Combo1"
            Top             =   25
            Width           =   3200
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   30
         TabIndex        =   34
         Top             =   45
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "imgToolbarImageList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Add"
               Object.ToolTipText     =   "Add a Movie"
               Object.Tag             =   "Add"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit"
               Object.ToolTipText     =   "Edit Selected"
               Object.Tag             =   "Edit"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Remove"
               Object.ToolTipText     =   "Remove Selected"
               Object.Tag             =   "Remove"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditTables"
               Object.ToolTipText     =   "Edit Support Tables"
               Object.Tag             =   "EditTables"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditOption"
               Object.ToolTipText     =   "Edit Options"
               Object.Tag             =   "EditOptions"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditDefaults"
               Object.ToolTipText     =   "Edit Defaults For Movies"
               Object.Tag             =   "EditDefaults"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4395
         ScaleHeight     =   375
         ScaleWidth      =   6090
         TabIndex        =   30
         Top             =   30
         Width           =   6090
         Begin VB.ComboBox cboSearchType 
            Height          =   345
            ItemData        =   "frmMain.frx":B993
            Left            =   0
            List            =   "frmMain.frx":B995
            TabIndex        =   33
            Text            =   "Title"
            Top             =   25
            Width           =   1365
         End
         Begin VB.CommandButton cmdSearchGo 
            Caption         =   "Go"
            Height          =   315
            Left            =   5640
            TabIndex        =   32
            Top             =   25
            Width           =   420
         End
         Begin VB.TextBox txtSearchText 
            Height          =   315
            Left            =   1395
            TabIndex        =   31
            Top             =   25
            Width           =   4140
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4395
         ScaleHeight     =   375
         ScaleWidth      =   7515
         TabIndex        =   2
         Top             =   435
         Width           =   7515
         Begin FBButton.FlatButton cmdAll 
            Height          =   315
            Left            =   0
            TabIndex        =   3
            ToolTipText     =   "Show All"
            Top             =   25
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "Þ"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   0
            Left            =   273
            TabIndex        =   4
            ToolTipText     =   "Show All Movies Starting With An ""A"""
            Top             =   25
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "A"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   1
            Left            =   546
            TabIndex        =   5
            ToolTipText     =   "Show All Movies Starting With An ""B"""
            Top             =   25
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            PictureWidth    =   20
            PictureHeight   =   20
            Caption         =   "B"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   2
            Left            =   819
            TabIndex        =   6
            ToolTipText     =   "Show All Movies Starting With An ""C"""
            Top             =   25
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "C"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   3
            Left            =   1092
            TabIndex        =   7
            ToolTipText     =   "Show All Movies Starting With An ""D"""
            Top             =   25
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "D"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   4
            Left            =   1365
            TabIndex        =   8
            ToolTipText     =   "Show All Movies Starting With An ""E"""
            Top             =   25
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "E"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   5
            Left            =   1638
            TabIndex        =   9
            ToolTipText     =   "Show All Movies Starting With An ""F"""
            Top             =   25
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "F"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   6
            Left            =   1911
            TabIndex        =   10
            ToolTipText     =   "Show All Movies Starting With An ""G"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "G"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   7
            Left            =   2184
            TabIndex        =   11
            ToolTipText     =   "Show All Movies Starting With An ""H"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "H"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   8
            Left            =   2457
            TabIndex        =   12
            ToolTipText     =   "Show All Movies Starting With An ""I"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "I"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   9
            Left            =   2730
            TabIndex        =   13
            ToolTipText     =   "Show All Movies Starting With An ""J"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "J"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   10
            Left            =   3003
            TabIndex        =   14
            ToolTipText     =   "Show All Movies Starting With An ""K"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "K"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   11
            Left            =   3276
            TabIndex        =   15
            ToolTipText     =   "Show All Movies Starting With An ""L"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "L"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   12
            Left            =   3549
            TabIndex        =   16
            ToolTipText     =   "Show All Movies Starting With An ""M"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "M"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   13
            Left            =   3822
            TabIndex        =   17
            ToolTipText     =   "Show All Movies Starting With An ""N"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "N"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   14
            Left            =   4095
            TabIndex        =   18
            ToolTipText     =   "Show All Movies Starting With An ""O"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "O"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   15
            Left            =   4368
            TabIndex        =   19
            ToolTipText     =   "Show All Movies Starting With An ""P"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "P"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   16
            Left            =   4641
            TabIndex        =   20
            ToolTipText     =   "Show All Movies Starting With An ""Q"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "Q"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   17
            Left            =   4914
            TabIndex        =   21
            ToolTipText     =   "Show All Movies Starting With An ""R"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "R"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   18
            Left            =   5187
            TabIndex        =   22
            ToolTipText     =   "Show All Movies Starting With An ""S"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "S"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   19
            Left            =   5460
            TabIndex        =   23
            ToolTipText     =   "Show All Movies Starting With An ""T"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "T"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   20
            Left            =   5733
            TabIndex        =   24
            ToolTipText     =   "Show All Movies Starting With An ""U"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "U"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   21
            Left            =   6006
            TabIndex        =   25
            ToolTipText     =   "Show All Movies Starting With An ""V"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "V"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   22
            Left            =   6285
            TabIndex        =   26
            ToolTipText     =   "Show All Movies Starting With An ""W"""
            Top             =   30
            Width           =   310
            _ExtentX        =   556
            _ExtentY        =   556
            Alignment       =   0
            Caption         =   "W"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   23
            Left            =   6552
            TabIndex        =   27
            ToolTipText     =   "Show All Movies Starting With An ""X"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "X"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   24
            Left            =   6825
            TabIndex        =   28
            ToolTipText     =   "Show All Movies Starting With An ""Y"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "Y"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin FBButton.FlatButton cmdAlphabet 
            Height          =   315
            Index           =   25
            Left            =   7080
            TabIndex        =   29
            ToolTipText     =   "Show All Movies Starting With An ""Z"""
            Top             =   30
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   556
            Caption         =   "Z"
            HasFocusRect    =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin SHDocVwCtl.WebBrowser wb2 
      Height          =   480
      Left            =   9360
      TabIndex        =   39
      Top             =   8400
      Width           =   480
      ExtentX         =   847
      ExtentY         =   847
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8760
      Picture         =   "frmMain.frx":B997
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   49
      Top             =   8760
      Width           =   240
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8400
      Picture         =   "frmMain.frx":BF21
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   48
      Top             =   8760
      Width           =   240
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   8040
      Picture         =   "frmMain.frx":C4AB
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   47
      Top             =   8760
      Width           =   240
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7680
      Picture         =   "frmMain.frx":CA35
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   46
      Top             =   8760
      Width           =   240
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7320
      Picture         =   "frmMain.frx":CFBF
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   45
      Top             =   8760
      Width           =   240
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuPrintPageSetup 
            Caption         =   "P&age Setup"
         End
         Begin VB.Menu mnuPrintPrintPreview 
            Caption         =   "P&rint Preview"
         End
         Begin VB.Menu mnuPrintPrint 
            Caption         =   "Prin&t"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuPrintList 
            Caption         =   "Print Full &List"
         End
      End
      Begin VB.Menu mnuSepB 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Option"
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuDatabaseBackup 
         Caption         =   "&Backup"
      End
      Begin VB.Menu mnuDatabaseRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuDatabaseCompact 
         Caption         =   "Compact / Repair"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsEditSupport 
         Caption         =   "E&dit Support Tables"
      End
      Begin VB.Menu mnuToolsDefaults 
         Caption         =   "Edit Defaults"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchIMDB 
         Caption         =   "IMDB.com"
      End
      Begin VB.Menu mnuSearchDVDEmpire 
         Caption         =   "DVD Empire"
      End
      Begin VB.Menu mnuSearchFye 
         Caption         =   "FYE.com"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnua 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuRandom 
      Caption         =   "&Random"
   End
   Begin VB.Menu mnuSepz 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuComingSoon 
      Caption         =   "DVDs Coming Soon"
   End
   Begin VB.Menu mnuTitleMenu 
      Caption         =   "TitleMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuTitleAddMovie 
         Caption         =   "Add Movie"
      End
      Begin VB.Menu mnuP 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTitleEditThis 
         Caption         =   "Edit This Movie"
      End
      Begin VB.Menu mnuTitleDeleteThis 
         Caption         =   "Delete This Movie"
      End
      Begin VB.Menu mnuSepQ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTitleSearchIMDB 
         Caption         =   "Search IMDB"
      End
      Begin VB.Menu mnuTitleSearchDVDEmpire 
         Caption         =   "DVD Empire"
      End
      Begin VB.Menu mnuTitleSearchFYE 
         Caption         =   "FYE Search"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================================
'  Movie Base V3.5.2
'  An easy way to store all your move details and information.
'  Please bare in mind that I am still noob so most if this code
'  is more then likely its a mess.
'  But I hope that you can use and maybe learn from it.
'=========================================================================================
'  Created By: Jason Scott
'  Published Date: 4-3-05
'  WebSite: N/A
'  Legal Copyright: Jason Scott © 2005
'=========================================================================================
' Last Edit  4/3/2005 12:17:11 PM

Option Explicit

'=========================================================================================
'==================== Variables
'=========================================================================================
Dim strSearch, SearchText As String
Dim search As String
Dim response As String
Dim a As ListItem
Dim strInput As String
Public strSearch3 As String
Private BTN As Integer
Public SelNde As Node
Dim TitleName As String

'=========================================================================================
'==================== Public Functions
'=========================================================================================
'------------------------------------------------------------------
' Purpose   : Open Add New Movie
'------------------------------------------------------------------
Private Sub AddNewMovie()
    Edit = False
    frmAddEdit.Show , frmMain
End Sub 'AddNewMovie()
'------------------------------------------------------------------
' Purpose   : Open Edit Movie Form
'------------------------------------------------------------------
Private Sub EditMovie()
        Edit = True
        If treMovieList.SelectedItem.Text = "MovieBase.mdb" Then ' = False Then
            MsgBox "You must select a movie"
            Exit Sub
        Else
            EditID = treMovieList.SelectedItem.Text
            frmAddEdit.Show , frmMain
            frmAddEdit.GetViewData frmMain.treMovieList.SelectedItem.Text
        End If
End Sub 'EditMovie()
'------------------------------------------------------------------
' Purpose   : Delets Selected Movie
'------------------------------------------------------------------
Public Sub DeleteMovie()
    On Error GoTo err:
    SearchText = treMovieList.SelectedItem.Text '// set the search text as the selected in the listbox
    strSearch = "[Title] Like '" & SearchText & "'" '// compares to selected'

    response = MsgBox("Are You Sure That You Want To Delete " & SearchText & "?", vbYesNo)  '// makes sure that you want to delete the selected
    If response = vbYes Then ' User chose Yes.
        With rsMovies '// opens the recordset
            .FindFirst strSearch '// looks for selected in recordset
            .Delete '// deletes the selected
        End With '// closed the recordset
        fillDVDTreeView treMovieList, cboTree.Text '// refreshes the list
        lstMovies.ListItems.Clear
        Exit Sub
    Else
        fillDVDTreeView treMovieList, cboTree.Text   '// refreshes the list
        Exit Sub
    End If
err:
    MsgBox "Sorry But you need to select a movie to delete.", vbCritical, " Movies v2.0 Error"
End Sub 'DeleteMovie()
'------------------------------------------------------------------
' Purpose   : Loads Movie info from treeview to listview
'------------------------------------------------------------------
Public Sub Again(strSearch2 As String)
    On Error Resume Next
    With rsMovies
        .FindNext strSearch2
        If .NoMatch Then
            Exit Sub
        Else
            Set a = lstMovies.ListItems.Add(, , .Fields("Title"), , 2)
            a.SubItems(1) = .Fields("Genre")
            a.SubItems(2) = .Fields("Director")
            a.SubItems(3) = .Fields("Rating")
            a.SubItems(4) = .Fields("MovieDate")
            a.SubItems(5) = .Fields("Type")
            a.SubItems(6) = .Fields("Cost")
            Again strSearch2
        End If
    End With
End Sub 'Again(strSearch2 As String)
'------------------------------------------------------------------
' Purpose   : Loads Movie info from treeview to listview
'------------------------------------------------------------------
Public Function GetViewData(SearchText As String)
    On Error Resume Next
    lstMovies.ListItems.Clear 'clear table
    strSearch = "[Title] Like '" & SearchText & "'" ' compair Title to records
    With rsMovies ' open movie table
        .FindFirst strSearch 'find match to title
        Set a = lstMovies.ListItems.Add(, , .Fields("Title"), , 2) 'Loads Movie Info to listview
        a.SubItems(1) = .Fields("Genre")
        a.SubItems(2) = .Fields("Director")
        a.SubItems(3) = .Fields("Rating")
        a.SubItems(4) = .Fields("MovieDate")
        a.SubItems(5) = .Fields("Type")
        a.SubItems(6) = .Fields("Cost")
    End With
End Function 'GetViewData(SearchText As String)
'------------------------------------------------------------------
' Purpose   : Loads all cbos and lists
'------------------------------------------------------------------
Public Sub LoadList()
    On Error Resume Next
    cboTree.Clear ' clears cbo
    cboSearchType.Clear ' clears cbo
    
    With rsTreeType
        .MoveFirst ' move to first record
        Do While Not .EOF ' loop until last record
            cboTree.AddItem .Fields("TreeType") 'adds records
            .MoveNext 'move to next record
            DoEvents
        Loop
    End With

    With rsSearchType
        .MoveFirst ' move to first record
        Do While Not .EOF ' loop until last record
            cboSearchType.AddItem .Fields("SearchType") 'adds records
            .MoveNext 'move to next record
            DoEvents
        Loop
    End With

    cboTree.Text = cboTree.List(0) 'Sets list to first
    cboSearchType.Text = cboSearchType.List(0) 'Sets list to first
End Sub 'LoadList()

'=========================================================================================
'==================== Main Form Procedures
'=========================================================================================
Private Sub Form_Load()

    'Form Width \ Height
    Me.Width = 12120
    Me.Height = 9690
    
    ' Sets Button For Nav Bar
    cmdGo.PicturehDC = Picture5.hDC
    cmdBack.PicturehDC = Picture7.hDC
    cmdForward.PicturehDC = Picture8.hDC
    cmdCancel.PicturehDC = Picture9.hDC
    cmdRefresh.PicturehDC = Picture10.hDC
    
    'set the constants for TwipsPerPixel
    With Screen
        TPPX = .TwipsPerPixelX
        TPPY = .TwipsPerPixelY
    End With
    
    picResizeLR.Left = BarLR
    picResizeTB.Top = BarTB
    iSizerPosLR = picResizeLR.Left
    iSizerPosTB = picResizeTB.Top

    Me.Caption = "MovieBase v" & App.Major & "." & App.Minor & "." & App.Revision
    wb.Navigate (App.Path & "\HTML\Load.html") 'Loads Default Page
    fillDVDTreeView treMovieList, "Title" 'loads treview sorted by title
    LoadList 'loads cbos
    
    'lstMovies.SortKey = lstMovies.ColumnHeaders.Item(1)
    lstMovies.Sorted = True
    StatusBarMsg "Welcome to the worst movie database program ever.", 1

    'change the background pic color
    picResizeLR.BackColor = Me.BackColor
    picResizeTB.BackColor = Me.BackColor

End Sub 'Form_Load()

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub 'if minimized skip so no error

    '----- Start Splitter Resizing
        'resize the left side control
        treMovieList.Move CONTROL_PADDING, CoolBar1.Top + CoolBar1.Height + CONTROL_PADDING, _
            iSizerPosLR - CONTROL_PADDING, Me.ScaleHeight - (CONTROL_PADDING * 2) - CoolBar1.Height - StatusBar.Height
    
        'resize the Top control
        lstMovies.Move iSizerPosLR + CONTROL_PADDING_BAR, CoolBar1.Top + CoolBar1.Height + CONTROL_PADDING, _
            Me.ScaleWidth - (iSizerPosLR + CONTROL_PADDING_BAR), iSizerPosTB - (CoolBar1.Top + CoolBar1.Height)
    
        'resize the Bottom control
        wb.Move iSizerPosLR + CONTROL_PADDING_BAR, iSizerPosTB + CONTROL_PADDING_BAR, _
            Me.ScaleWidth - (iSizerPosLR + CONTROL_PADDING_BAR), Me.ScaleHeight - iSizerPosTB - StatusBar.Height - (CONTROL_PADDING_BAR) - CoolBar2.Height
        'resize the sizer
        picResizeLR.Move iSizerPosLR, CoolBar1.Top + CoolBar1.Height, CONTROL_PADDING_BAR, Me.ScaleHeight - (CONTROL_PADDING * 2) - CoolBar1.Height - StatusBar.Height
        picResizeTB.Move iSizerPosLR + CONTROL_PADDING_BAR, iSizerPosTB, lstMovies.Width + CONTROL_PADDING, CONTROL_PADDING_BAR
    '----- End Splitter Resizing
    '----- Start GUI Resizing
        CoolBar2.Move wb.Left, wb.Top + wb.Height, wb.Width, 28
    '----- End GUI Resizing
End Sub 'Form_Resize()

Private Sub Form_Unload(Cancel As Integer)
    Dim intFrmNum As Integer
    On Error Resume Next
    
    'saves form information
    SaveDataVaris
    
    intFrmNum = Forms.Count
    Do Until intFrmNum = 0
        Unload Forms(intFrmNum - 1)
        intFrmNum = intFrmNum - 1
    Loop
End Sub 'Form_Unload()

'=========================================================================================
'==================== Form Splitter Functions
'=========================================================================================
Private Sub picResizeLR_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'change the color of the sizer to dark-grey
    If Button = vbLeftButton Then picResizeLR.BackColor = RGB(128, 128, 128)
    If Button = vbRightButton Then
        
    End If

End Sub 'picResizeLR_MouseDown()

Private Sub picResizeLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        'if we we're actually moving the control....
        If CLng(X) <> iSizerPosLR Then
            'if we're not out of bounds...
            If (picResizeLR.Left + (X / TPPX)) > 100 And (picResizeLR.Left + (X / TPPX)) < (Me.ScaleWidth - 100) Then
                'move the slider
                picResizeLR.Move picResizeLR.Left + (X / TPPX), _
                               picResizeLR.Top, _
                               picResizeLR.Width, _
                               picResizeLR.Height
                
                'set the variable
                iSizerPosLR = picResizeLR.Left
            End If
        End If
        'bring the sizer to the front
        picResizeLR.ZOrder 0
    End If
End Sub 'picResizeLR_MouseMove()

Private Sub picResizeLR_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'reset the sizer color
    picResizeLR.BackColor = vbButtonFace
    'resize the controls
    If Button = 2 Then iSizerPosLR = 280
    Form_Resize
End Sub 'picResizeLR_MouseUp()

Private Sub picResizeTB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'change the color of the sizer to dark-grey
    picResizeTB.BackColor = RGB(128, 128, 128)
End Sub 'picResizeTB_MouseDown()

Private Sub picResizeTB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        'if we we're actually moving the control....
        If CLng(Y) <> iSizerPosTB Then
            'if we're not out of bounds...
            If (picResizeTB.Top + (Y / TPPY)) > 100 And (picResizeTB.Top + (Y / TPPY)) < (Me.ScaleHeight - 100) Then
                'move the slider
                picResizeTB.Move picResizeTB.Left, _
                               picResizeTB.Top + (Y / TPPY), _
                               picResizeTB.Width, _
                               picResizeTB.Height
                
                'set the variable
                iSizerPosTB = picResizeTB.Top
            End If
        End If
        'bring the sizer to the front
        picResizeTB.ZOrder 0
    End If
End Sub 'picResizeTB_MouseMove()

Private Sub picResizeTB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'reset the sizer color
    picResizeTB.BackColor = vbButtonFace
    
    'resize the controls
    If Button = 2 Then iSizerPosTB = 400
    Form_Resize
End Sub 'picResizeTB_MouseUp()

'=========================================================================================
'==================== Coolbar Commands
'=========================================================================================
Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
    On Error Resume Next
    Form_Resize ' calls the form resize to adjust for controls
End Sub 'CoolBar1_HeightChanged()

Private Sub CoolBar1_Resize()
On Error Resume Next
    txtSearchText.Move cboSearchType.Width + 30, 25, Picture3.ScaleWidth - cboSearchType.Width - cmdSearchGo.Width - 60
    cmdSearchGo.Move txtSearchText.Left + txtSearchText.Width + 50
    
    cboTree.Move 0, 25, Picture4.ScaleWidth - 30
End Sub 'CoolBar1_Resize()

Private Sub CoolBar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    CoolBar1_Resize
End Sub 'CoolBar1_MouseMove()

Private Sub CoolBar1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        CoolBar1.Bands.Item(4).Width = 280
        CoolBar1.Bands.Item(1).Width = 280
        CoolBar1_Resize
    End If
End Sub 'CoolBar1_MouseUp()

Private Sub CoolBar2_Resize()
On Error Resume Next
    CoolBar2.Bands.Item(2).MinWidth = 75
    txtAddress.Move 0, 25, _
    Picture2.ScaleWidth - (cmdGo.Width * 5.5), 330
    
    cmdGo.Move txtAddress.Left + txtAddress.Width + 30, 25, 330, 330
    cmdBack.Move cmdGo.Left + cmdGo.Width, 25, 330, 330
    cmdForward.Move cmdBack.Left + cmdBack.Width, 25, 330, 330
    cmdCancel.Move cmdForward.Left + cmdForward.Width, 25, 330, 330
    cmdRefresh.Move cmdCancel.Left + cmdCancel.Width, 25, 330, 330
End Sub 'CoolBar2_Resize()

'=========================================================================================
'==================== All ToolBar Commands
'=========================================================================================
'------------------------------------------------------------------
' Main Toolbar
'------------------------------------------------------------------
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Add"
            AddNewMovie
        Case "Edit"
            EditMovie
        Case "Remove"
            DeleteMovie
        Case "EditTables"
            frmEditSupport.Show , frmMain
        Case "EditOption"
            frmOptions.Show , frmMain
        Case "EditDefaults"
            frmDefault.Show , frmMain
    End Select
End Sub 'Toolbar1_ButtonClick()
'------------------------------------------------------------------
' Alphabet bar
'------------------------------------------------------------------
Private Sub cmdAll_Click() '----- Show All
    On Error Resume Next
    lstMovies.Visible = False
    lstMovies.ListItems.Clear
    With rsMovies
        .MoveFirst
        Do While Not .EOF
            Set a = lstMovies.ListItems.Add(, , .Fields("Title"), , 2)
            a.SubItems(1) = .Fields("Genre")
            a.SubItems(2) = .Fields("Director")
            a.SubItems(3) = .Fields("Rating")
            a.SubItems(4) = .Fields("MovieDate")
            a.SubItems(5) = .Fields("Type")
            a.SubItems(6) = .Fields("Cost")
            .MoveNext
            DoEvents
        Loop
    End With
    lstMovies.Visible = True
End Sub 'cmdAll_Click()

Private Sub cmdAlphabet_Click(Index As Integer) '----- Sorted by Letters
    On Error Resume Next
    lstMovies.Visible = False
    lstMovies.ListItems.Clear
    Select Case Index
        Case 0
            strSearch = "Left$([Title], 1) Like '" & "a" & "'"
        Case 1
            strSearch = "Left$([Title], 1) Like '" & "b" & "'"
        Case 2
            strSearch = "Left$([Title], 1) Like '" & "c" & "'"
        Case 3
            strSearch = "Left$([Title], 1) Like '" & "d" & "'"
        Case 4
            strSearch = "Left$([Title], 1) Like '" & "e" & "'"
        Case 5
            strSearch = "Left$([Title], 1) Like '" & "f" & "'"
        Case 6
            strSearch = "Left$([Title], 1) Like '" & "g" & "'"
        Case 7
            strSearch = "Left$([Title], 1) Like '" & "h" & "'"
        Case 8
            strSearch = "Left$([Title], 1) Like '" & "i" & "'"
        Case 9
            strSearch = "Left$([Title], 1) Like '" & "j" & "'"
        Case 10
            strSearch = "Left$([Title], 1) Like '" & "k" & "'"
        Case 11
            strSearch = "Left$([Title], 1) Like '" & "l" & "'"
        Case 12
            strSearch = "Left$([Title], 1) Like '" & "m" & "'"
        Case 13
            strSearch = "Left$([Title], 1) Like '" & "n" & "'"
        Case 14
            strSearch = "Left$([Title], 1) Like '" & "o" & "'"
        Case 15
            strSearch = "Left$([Title], 1) Like '" & "p" & "'"
        Case 16
            strSearch = "Left$([Title], 1) Like '" & "q" & "'"
        Case 17
            strSearch = "Left$([Title], 1) Like '" & "r" & "'"
        Case 18
            strSearch = "Left$([Title], 1) Like '" & "s" & "'"
        Case 19
            strSearch = "Left$([Title], 1) Like '" & "t" & "'"
        Case 20
            strSearch = "Left$([Title], 1) Like '" & "u" & "'"
        Case 21
            strSearch = "Left$([Title], 1) Like '" & "v" & "'"
        Case 22
            strSearch = "Left$([Title], 1) Like '" & "w" & "'"
        Case 23
            strSearch = "Left$([Title], 1) Like '" & "x" & "'"
        Case 24
            strSearch = "Left$([Title], 1) Like '" & "y" & "'"
        Case 25
            strSearch = "Left$([Title], 1) Like '" & "z" & "'"
    End Select
    search = strSearch
    With rsMovies
        .FindFirst search
        If .NoMatch Then
            'MsgBox "Your search is complete.", vbInformation + vbOKOnly, "Search Complete"
        Else
            Set a = lstMovies.ListItems.Add(, , .Fields("Title"), , 2)
            a.SubItems(1) = .Fields("Genre")
            a.SubItems(2) = .Fields("Director")
            a.SubItems(3) = .Fields("Rating")
            a.SubItems(4) = .Fields("MovieDate")
            a.SubItems(5) = .Fields("Type")
            a.SubItems(6) = .Fields("Cost")
            Again search
        End If
    End With
    lstMovies.Visible = True
End Sub 'cmdAlphabet_Click(Index As Integer)
'------------------------------------------------------------------
' Tree Sorting cbo
'------------------------------------------------------------------
Private Sub cboTree_Click()
    On Error Resume Next
    If cboTree.Text = "Title" Then fillDVDTreeView treMovieList, "Title" 'changes tree display
    If cboTree.Text = "Genre" Then fillDVDTreeView treMovieList, "Genre"
    If cboTree.Text = "Rating" Then fillDVDTreeView treMovieList, "Rating"
    If cboTree.Text = "Region" Then fillDVDTreeView treMovieList, "Region"
    If cboTree.Text = "Format" Then fillDVDTreeView treMovieList, "Format"
End Sub 'cboTree_Click()
'------------------------------------------------------------------
' Quick Search Bar
'------------------------------------------------------------------
Private Sub cmdSearchGo_Click()
    On Error Resume Next
    lstMovies.ListItems.Clear
    strSearch3 = "Mid$([" & cboSearchType.Text & "],1) Like '*" & txtSearchText.Text & "*'"
    With rsMovies
        .FindFirst strSearch3
        If .NoMatch Then
            MsgBox "Your search is complete. There are no movies with  """ & txtSearchText.Text & """  in the  """ & cboSearchType & """  section.", 16, "Search Complete"
        Else
            Set a = lstMovies.ListItems.Add(, , .Fields("Title"), , 2)
            a.SubItems(1) = .Fields("Genre")
            a.SubItems(2) = .Fields("Director")
            a.SubItems(3) = .Fields("Rating")
            a.SubItems(4) = .Fields("MovieDate")
            a.SubItems(5) = .Fields("Type")
            a.SubItems(6) = .Fields("Cost")
            Again strSearch3
        End If
    End With
End Sub 'cmdSearchGo_Click()
'=========================================================================================
'==================== Controls
'=========================================================================================
'------------------------------------------------------------------
' Treeview
'------------------------------------------------------------------
Private Sub treMovieList_Collapse(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    treMovieList.Nodes.Item(1).Image = "Closed" 'Change image to closed folder
End Sub 'treMovieList_Collapse()

Private Sub treMovieList_Expand(ByVal Node As MSComctlLib.Node)
    On Error Resume Next
    treMovieList.Nodes.Item(1).Image = "Open" 'Change image to open folder
End Sub 'treMovieList_Expand()

Private Sub treMovieList_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim d() As String

    d = Split(Node.Tag, "|")

    If UBound(d) <= 0 Then Exit Sub
    Set SelNde = Node

    If BTN = 2 Then
        Select Case UCase(d(0))
            Case "T" 'TItle
                TitleName = treMovieList.SelectedItem.Text
                
                mnuTitleSearchIMDB.Caption = "Search for '" & TitleName & "' at IMDB.com"
                mnuTitleSearchDVDEmpire.Caption = "Search for '" & TitleName & "' at DVDEmpire.com"
                mnuTitleSearchFYE.Caption = "Search for '" & TitleName & "' at FYE.com"
                PopupMenu mnuTitleMenu
                
            Case "G" 'group
                'MsgBox "Group" 'PopupMenu frmMain.mnuSMgroups
        End Select
        Exit Sub
    Else ' Click on a Title
        GetViewData treMovieList.SelectedItem.Text 'loads the listview
        CreatHTML treMovieList.SelectedItem.Text
        
        wb.Document.Script.Document.Clear
        wb.Document.Script.Document.Write txtHTML 'text1.Text
        wb.Document.Script.Document.Close
        
        StatusBarMsg "Good GOD!  Something must Really be wrong with you if you want to read about " & Chr(34) & treMovieList.SelectedItem.Text & Chr(34), 1
    End If
End Sub 'treMovieList_NodeClick()

Private Sub treMovieList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BTN = Button
End Sub 'treMovieList_MouseDown()

'------------------------------------------------------------------
' Listview
'------------------------------------------------------------------
Private Sub lstMovies_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lstMovies.SortOrder = 0 Then
        lstMovies.SortOrder = 1
    Else
        lstMovies.SortOrder = 0
    End If
    
    lstMovies.SortKey = ColumnHeader.Index - 1
    lstMovies.Sorted = True
End Sub 'lstMovies_ColumnClick()

Private Sub lstMovies_DblClick()

    CreatHTML lstMovies.SelectedItem.Text

    wb.Document.Script.Document.Clear
    wb.Document.Script.Document.Write txtHTML 'text1.Text
    wb.Document.Script.Document.Close
    StatusBarMsg "Good GOD!  Something must Really be wrong with you if you want to read about " & Chr(34) & lstMovies.SelectedItem.Text & Chr(34), 1
End Sub 'lstMovies_DblClick()()

Private Sub lstMovies_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
Dim SelectTemp As String
SelectTemp = treMovieList.SelectedItem.Text

    If BTN = 2 Then
        TitleName = Item.Text
        
        mnuTitleSearchIMDB.Caption = "Search for '" & TitleName & "' at IMDB.com"
        mnuTitleSearchDVDEmpire.Caption = "Search for '" & TitleName & "' at DVDEmpire.com"
        mnuTitleSearchFYE.Caption = "Search for '" & TitleName & "' at FYE.com"
        treMovieList.SelectedItem.Text = Item
        PopupMenu mnuTitleMenu
        treMovieList.SelectedItem.Text = SelectTemp
        Exit Sub
    End If
End Sub 'lstMovies_ItemClick()

Private Sub lstMovies_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BTN = Button
End Sub 'lstMovies_MouseDown()

'=========================================================================================
'==================== Menu Commands
'=========================================================================================
'------------------------------------------------------------------
' File Menu
'------------------------------------------------------------------
Private Sub mnuFileNew_Click() 'File/New
    AddNewMovie
End Sub 'mnuFileNew_Click()

Private Sub mnuPrintPageSetup_Click() 'File/Print/PageSetup
    wb.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_DODEFAULT: Exit Sub
End Sub 'mnuPrintPageSetup_Click()

Private Sub mnuPrintPrintPreview_Click() 'File/Print/PrintPreview
    wb.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT: Exit Sub
End Sub 'mnuPrintPrintPreview_Click()

Private Sub mnuPrintPrint_Click() 'File/Print/Print
    wb.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT: Exit Sub
End Sub 'mnuPrintPrint_Click()

Private Sub mnuPrintList_Click() 'File/Print/PrintFullList
    PrintList
    wb2.Document.Script.Document.Clear
    wb2.Document.Script.Document.Write listHTML 'text1.Text
    wb2.Document.Script.Document.Close
    wb2.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_DODEFAULT: Exit Sub
End Sub 'mnuPrintList_Click()

Private Sub mnuFileExit_Click() 'File/Exit
    On Error Resume Next
    Unload Me 'unloads form
End Sub 'mnuFileExit_Click()

'------------------------------------------------------------------
' View Menu
'------------------------------------------------------------------
Private Sub mnuViewOptions_Click() 'View/Options
    On Error Resume Next
    frmOptions.Show , frmMain
End Sub 'mnuViewOptions_Click()
'------------------------------------------------------------------
' Database Menu
'------------------------------------------------------------------
Private Sub mnuDatabaseBackup_Click() 'Database/Backup
    On Error Resume Next
    frmBackup.Show , frmMain
End Sub 'mnuDatabaseBackup_Click()

Private Sub mnuDatabaseRestore_Click() 'Database/Restore
    On Error Resume Next
    frmRestore.Show , frmMain
End Sub 'mnuDatabaseRestore_Click()

Private Sub mnuDatabaseCompact_Click() 'Database/CompactRepair
    On Error Resume Next
    frmCompact.Show , frmMain
End Sub 'mnuDatabaseCompact_Click()
'------------------------------------------------------------------
' Tools Menu
'------------------------------------------------------------------
Private Sub mnuToolsDefaults_Click() 'Tools/EditDefaults
    On Error Resume Next
    frmDefault.Show , frmMain
End Sub 'mnuToolsDefaults_Click()

Private Sub mnuToolsEditSupport_Click() 'Tools/EditSupportTables
    On Error Resume Next
    frmEditSupport.Show , frmMain
End Sub 'mnuToolsEditSupport_Click()
'------------------------------------------------------------------
' Search Menu
'------------------------------------------------------------------
Private Sub mnuSearchIMDB_Click() 'Search/IMDB
    strInput = InputBox("Please enter the title what you whish to search for.", "IMDB Title Search")
    wb.Navigate2 ("http://imdb.com/find?q=" & strInput & ";tt=on;nm=on;mx=20")
End Sub 'mnuSearchIMDB_Click()

Private Sub mnuSearchDVDEmpire_Click() 'Search/DVDEmpire
    strInput = InputBox("Please enter the title what you whish to search for.", "DVD Empire Title Search")
    wb.Navigate2 ("http://www.dvdempire.com/Exec/v5_search_item.asp?userid=00000865958310&string=" & strInput & "&media_id=&site_id=4&sort=")
End Sub 'mnuSearchDVDEmpire_Click()

Private Sub mnuSearchFye_Click() 'Search/FYE
    strInput = InputBox("Please enter the title what you whish to search for.", "FYE Title Search")
    wb.Navigate2 ("http://shop.fye.com/searchresults.aspx?qu=" & strInput & "&queryType=57&loc=50244&search_store=57")
End Sub 'mnuSearchFye_Click()
'------------------------------------------------------------------
' Help Menu
'------------------------------------------------------------------
Private Sub mnuHelpAbout_Click() 'Help/About
    On Error Resume Next
    frmAbout.Show , frmMain
End Sub 'mnuHelpAbout_Click()

'------------------------------------------------------------------
' Single Toolbar Buttons
'------------------------------------------------------------------
Private Sub mnuRandom_Click() 'Random
Randomize
    Dim sel As Integer
        With rsMovies
        .MoveFirst
        sel = Fix(Rnd() * .RecordCount)
        .Move sel
        MsgBox .Fields("Title"), 64, "A Random Movie To Watch!"
    End With
End Sub 'mnuRandom_Click()

Private Sub mnuComingSoon_Click() 'Coming Soon
    wb.Navigate2 "http://reel.com/reel.asp?node=dvd/comingsoon"
End Sub 'mnuComingSoon_Click()

'------------------------------------------------------------------
' TitleMenu
'------------------------------------------------------------------
Private Sub mnuTitleAddMovie_Click() 'TitleMenu/Add
    AddNewMovie
End Sub 'mnuTitleAddMovie_Click()

Private Sub mnuTitleDeleteThis_Click() 'TitleMenu/Delete
    DeleteMovie
End Sub 'mnuTitleDeleteThis_Click()

Private Sub mnuTitleEditThis_Click() 'TitleMenu/edit
    EditMovie
End Sub 'mnuTitleEditThis_Click()

Private Sub mnuTitleSearchIMDB_Click() 'TitleMenu/IMDB
    wb.Navigate2 ("http://imdb.com/find?q=" & TitleName & ";tt=on;nm=on;mx=20")
End Sub 'mnuTitleSearchIMDB_Click()

Private Sub mnuTitleSearchDVDEmpire_Click() 'TitleMenu/DVDEmpire
    wb.Navigate2 ("http://www.dvdempire.com/Exec/v5_search_item.asp?userid=00000865958310&string=" & TitleName & "&media_id=&site_id=4&sort=")
End Sub 'mnuTitleSearchDVDEmpire_Click()

Private Sub mnuTitleSearchFYE_Click() 'TitleMenu/FYE
    wb.Navigate2 ("http://shop.fye.com/searchresults.aspx?qu=" & TitleName & "&queryType=57&loc=50244&search_store=57")
End Sub 'mnuTitleSearchFYE_Click()

'=========================================================================================
'==================== Web Browser Commands
'=========================================================================================

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdGo_Click ' Return is clicked
End Sub

'------------------------------------------------------------------
' Runs Progress Bar
'------------------------------------------------------------------
Private Sub wb_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
    ProgressBar.Max = ProgressMax
    ProgressBar.Value = Progress
End Sub

'------------------------------------------------------------------
' Search Button
'------------------------------------------------------------------
Private Sub cmdGo_Click()
    On Error Resume Next
    If txtAddress.Text = "" Then Exit Sub
    wb.Navigate2 txtAddress.Text
End Sub 'cmdGo_Click()

Private Sub cmdBack_Click()
    On Error Resume Next
    wb.GoBack
End Sub

Private Sub cmdForward_Click()
    On Error Resume Next
    wb.GoForward
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    wb.Stop
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    wb.Refresh
End Sub


