VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "DVDBase - The Easy DVD Database"
   ClientHeight    =   5310
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGeneral 
      Caption         =   "General"
      Height          =   4515
      Left            =   2520
      TabIndex        =   2
      Top             =   420
      Width           =   6375
      Begin VB.CommandButton btnAddDVD 
         Caption         =   "Add DVD"
         Height          =   315
         Left            =   4470
         TabIndex        =   84
         ToolTipText     =   "Click to add this DVD"
         Top             =   180
         Width           =   945
      End
      Begin VB.ComboBox cboUserReview 
         Height          =   315
         ItemData        =   "frmMain.frx":0442
         Left            =   2550
         List            =   "frmMain.frx":0458
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   870
         Width           =   315
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   4
         Left            =   2220
         Picture         =   "frmMain.frx":046E
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   82
         Top             =   900
         Width           =   240
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   3
         Left            =   2010
         Picture         =   "frmMain.frx":0570
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   81
         Top             =   900
         Width           =   240
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   2
         Left            =   1800
         Picture         =   "frmMain.frx":0672
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   80
         Top             =   900
         Width           =   240
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   1
         Left            =   1590
         Picture         =   "frmMain.frx":0774
         ScaleHeight     =   195
         ScaleWidth      =   240
         TabIndex        =   79
         Top             =   900
         Width           =   240
      End
      Begin VB.PictureBox picStar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   200
         Index           =   0
         Left            =   1380
         Picture         =   "frmMain.frx":0876
         ScaleHeight     =   195
         ScaleWidth      =   225
         TabIndex        =   78
         Top             =   900
         Width           =   230
      End
      Begin VB.TextBox txtDirector 
         Height          =   315
         Left            =   4500
         TabIndex        =   17
         Text            =   "Text2"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox txtDatePurchased 
         Height          =   315
         Left            =   1620
         TabIndex        =   14
         Text            =   "66/66/66"
         Top             =   3720
         Width           =   795
      End
      Begin VB.TextBox txtLocationPurchased 
         Height          =   315
         Left            =   1620
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox txtCost 
         Height          =   315
         Left            =   1620
         TabIndex        =   15
         Text            =   "$666.66"
         Top             =   4080
         Width           =   705
      End
      Begin VB.TextBox txtStudio 
         Height          =   315
         Left            =   4500
         TabIndex        =   16
         Text            =   "Text2"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtRunningTime 
         Height          =   315
         Left            =   4500
         TabIndex        =   18
         Text            =   "66:66:66"
         ToolTipText     =   "Should be of the form hh:mm:ss, anything else will be rejected"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cboCurrentLocation 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2940
         Width           =   1815
      End
      Begin VB.ComboBox cboCaseType 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2580
         Width           =   1995
      End
      Begin VB.ComboBox cboRating 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2220
         Width           =   4935
      End
      Begin VB.ComboBox cboRegion 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1860
         Width           =   4935
      End
      Begin VB.TextBox txtDVDRelease 
         Height          =   315
         Left            =   1380
         TabIndex        =   8
         Text            =   "6666"
         Top             =   1500
         Width           =   495
      End
      Begin VB.TextBox txtMovieYear 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Text            =   "6666"
         Top             =   1140
         Width           =   495
      End
      Begin VB.ComboBox cboGenre 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   1380
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   180
         Width           =   3015
      End
      Begin VB.Label Label16 
         Caption         =   "Director"
         Height          =   255
         Left            =   3420
         TabIndex        =   76
         Top             =   990
         Width           =   675
      End
      Begin VB.Label Label13 
         Caption         =   "Running Time"
         Height          =   255
         Left            =   3420
         TabIndex        =   67
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Studio"
         Height          =   255
         Left            =   3420
         TabIndex        =   66
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label11 
         Caption         =   "Current Location"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   2970
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Case Type"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   2610
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   4110
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Location Purchased"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   3390
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Date Purchased"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   3750
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Rating"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   2250
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Region"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "DVD Release"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   1530
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Movie Year"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "User Review"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Genre"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Caption         =   "Title"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete DVD"
      Height          =   495
      Left            =   1290
      TabIndex        =   53
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "New DVD"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvwDiscs 
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   52917
      EndProperty
   End
   Begin VB.Frame fraAudioVideo 
      Caption         =   "Audio / Video"
      Height          =   4455
      Left            =   2520
      TabIndex        =   3
      Top             =   450
      Width           =   6375
      Begin VB.Frame Frame6 
         Caption         =   "NTSC/PAL"
         Height          =   555
         Left            =   180
         TabIndex        =   75
         Top             =   1800
         Width           =   2955
         Begin VB.OptionButton optPAL 
            Caption         =   "PAL"
            Height          =   255
            Left            =   1620
            TabIndex        =   25
            Top             =   240
            Width           =   915
         End
         Begin VB.OptionButton optNTSC 
            Caption         =   "NTSC"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Video Formats"
         Height          =   1515
         Left            =   180
         TabIndex        =   68
         Top             =   240
         Width           =   2955
         Begin VB.TextBox txtRatio 
            Height          =   315
            Left            =   2100
            TabIndex        =   23
            Text            =   "Text6"
            Top             =   270
            Width           =   555
         End
         Begin VB.CheckBox chk169 
            Caption         =   "16 x 9 Enhanced"
            Height          =   315
            Left            =   240
            TabIndex        =   22
            Top             =   1140
            Width           =   1635
         End
         Begin VB.CheckBox chkPanScan 
            Caption         =   "Pan & Scan"
            Height          =   315
            Left            =   240
            TabIndex        =   21
            Top             =   840
            Width           =   1095
         End
         Begin VB.CheckBox chkFullFrame 
            Caption         =   "Full Frame"
            Height          =   315
            Left            =   240
            TabIndex        =   20
            Top             =   540
            Width           =   1335
         End
         Begin VB.CheckBox chkWidescreen 
            Caption         =   "Widescreen"
            Height          =   315
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   " : 1"
            Height          =   255
            Left            =   2640
            TabIndex        =   72
            Top             =   300
            Width           =   255
         End
         Begin VB.Label Label14 
            Caption         =   "Ratio"
            Height          =   255
            Left            =   1620
            TabIndex        =   71
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Audio Formats"
         Height          =   3735
         Left            =   3240
         TabIndex        =   70
         Top             =   240
         Width           =   2955
         Begin VB.CheckBox chkAudioOther 
            Caption         =   "Other"
            Height          =   375
            Left            =   300
            TabIndex        =   41
            Top             =   2340
            Width           =   2235
         End
         Begin VB.CheckBox chkDolbySurround 
            Caption         =   "Dolby Surround"
            Height          =   375
            Left            =   300
            TabIndex        =   35
            Top             =   540
            Width           =   2235
         End
         Begin VB.CheckBox chkDolbyProLogic 
            Caption         =   "Dolby Pro-Logic"
            Height          =   375
            Left            =   300
            TabIndex        =   36
            Top             =   840
            Width           =   2235
         End
         Begin VB.CheckBox chkdd51 
            Caption         =   "Dolby Digital 5.1 (AC-3)"
            Height          =   375
            Left            =   300
            TabIndex        =   37
            Top             =   1140
            Width           =   2235
         End
         Begin VB.CheckBox chkDDEx 
            Caption         =   "Dolby Digital Surround EX"
            Height          =   375
            Left            =   300
            TabIndex        =   38
            Top             =   1440
            Width           =   2235
         End
         Begin VB.CheckBox chkDTS 
            Caption         =   "DTS"
            Height          =   375
            Left            =   300
            TabIndex        =   39
            Top             =   1740
            Width           =   2235
         End
         Begin VB.CheckBox chkSDDS 
            Caption         =   "Sony SDDS"
            Height          =   375
            Left            =   300
            TabIndex        =   40
            Top             =   2040
            Width           =   2235
         End
         Begin VB.CheckBox chkStereo 
            Caption         =   "Stereo"
            Height          =   375
            Left            =   300
            TabIndex        =   34
            Top             =   240
            Width           =   2235
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Subtitles"
         Height          =   1575
         Left            =   180
         TabIndex        =   69
         Top             =   2400
         Width           =   2955
         Begin VB.CheckBox chkFrench 
            Caption         =   "French"
            Height          =   315
            Left            =   240
            TabIndex        =   27
            Top             =   600
            Width           =   915
         End
         Begin VB.CheckBox chkGerman 
            Caption         =   "German"
            Height          =   315
            Left            =   240
            TabIndex        =   28
            Top             =   900
            Width           =   915
         End
         Begin VB.CheckBox chkSpanish 
            Caption         =   "Spanish"
            Height          =   315
            Left            =   240
            TabIndex        =   29
            Top             =   1200
            Width           =   915
         End
         Begin VB.CheckBox chkPortugese 
            Caption         =   "Portugese"
            Height          =   315
            Left            =   1620
            TabIndex        =   30
            Top             =   300
            Width           =   1035
         End
         Begin VB.CheckBox chkJapanese 
            Caption         =   "Japanese"
            Height          =   315
            Left            =   1620
            TabIndex        =   31
            Top             =   600
            Width           =   1035
         End
         Begin VB.CheckBox chkChinese 
            Caption         =   "Chinese"
            Height          =   315
            Left            =   1620
            TabIndex        =   32
            Top             =   900
            Width           =   915
         End
         Begin VB.CheckBox chkSubTitleOther 
            Caption         =   "Other"
            Height          =   315
            Left            =   1620
            TabIndex        =   33
            Top             =   1200
            Width           =   915
         End
         Begin VB.CheckBox chkEnglish 
            Caption         =   "English"
            Height          =   315
            Left            =   240
            TabIndex        =   26
            Top             =   300
            Width           =   915
         End
      End
   End
   Begin VB.Frame fraFeatures 
      Caption         =   "Features"
      Height          =   4455
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   6375
      Begin VB.Frame Frame5 
         Caption         =   "Disc Format"
         Height          =   3135
         Left            =   3060
         TabIndex        =   74
         Top             =   240
         Width           =   2115
         Begin VB.OptionButton optDualSided 
            Caption         =   "Dual-Sided"
            Height          =   195
            Left            =   240
            TabIndex        =   51
            Top             =   660
            Width           =   1335
         End
         Begin VB.OptionButton optFlipper 
            Caption         =   "Flipper"
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optDualLayer 
            Caption         =   "Dual Layer"
            Height          =   195
            Left            =   240
            TabIndex        =   50
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Disc Extras"
         Height          =   3135
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   2835
         Begin VB.CheckBox chkAnimatedMenus 
            Caption         =   "Animated Menus"
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   600
            Width           =   1695
         End
         Begin VB.CheckBox chkMakingOf 
            Caption         =   """Making Of"" Documentary"
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   900
            Width           =   2475
         End
         Begin VB.CheckBox chkBios 
            Caption         =   "Star Bios"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   2100
            Width           =   1035
         End
         Begin VB.CheckBox chkDeletedScenes 
            Caption         =   "Deleted Scenes"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   1500
            Width           =   1755
         End
         Begin VB.CheckBox chkCommentary 
            Caption         =   "Commentary"
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   1200
            Width           =   2235
         End
         Begin VB.CheckBox chkDVDROM 
            Caption         =   "DVD-ROM Content"
            Height          =   195
            Left            =   240
            TabIndex        =   49
            Top             =   2400
            Width           =   2295
         End
         Begin VB.CheckBox chkTheatricalTrailer 
            Caption         =   "Theatrical Trailer"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   1800
            Width           =   1875
         End
         Begin VB.CheckBox chkSceneAccess 
            Caption         =   "Scene Access"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   300
            Width           =   1635
         End
      End
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   5205
      Left            =   2460
      TabIndex        =   77
      Top             =   60
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   9181
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Object.Tag             =   "General"
            Object.ToolTipText     =   "Click to see the general properties of this DVD"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Audio / Video"
            Object.Tag             =   "AudioVideo"
            Object.ToolTipText     =   "Click to see the Audio and Video properties of this DVD"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Features"
            Object.Tag             =   "Features"
            Object.ToolTipText     =   "Click to see the features of this DVD and see the disc type"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAddDVD_Click()
  If intFormAction = ADD_NEW Then
    If Len(Me.txtTitle) > 0 Then
      If vbYes = MsgBox("Are you sure you want to add this DVD?", vbYesNo, "Confirm Add DVD") Then
        dvdCurrent.addDVD Me.txtTitle
        globalCode.fillDVDListView Me.lvwDiscs, dvdCurrent.getLatestDVD
        Me.Caption = strCaption & " : " & dvdCurrent.strTitle
      End If
    End If
  End If
End Sub

Private Sub btnDelete_Click()
  If checkSelected(Me.lvwDiscs) = -1 Then Exit Sub
  If vbYes = MsgBox("Are you sure you want to delete this DVD?", vbYesNo, "Deletion Warning") Then
    dvdCurrent.deleteDVD
    Me.Caption = strCaption
    globalCode.fillDVDListView Me.lvwDiscs
    discCode.resetAllFields
    discCode.disableDiscDisplay
  End If

End Sub

Private Sub btnNew_Click()
  intFormAction = ADD_NEW
  discCode.enableDiscDisplay
  discCode.resetAllFields
  discCode.enableAddDiscDisplay
  Me.btnAddDVD.Visible = True
End Sub

Private Sub cboCaseType_Click()
  If Me.cboCaseType.ListIndex <> -1 Then
    dvdCurrent.lngCaseType = lngCaseTypeArray(Me.cboCaseType.ListIndex + 1)
  Else
    dvdCurrent.lngCaseType = 0
  End If

End Sub

Private Sub cboCurrentLocation_Click()
  If Me.cboCurrentLocation.ListIndex <> -1 Then
    dvdCurrent.lngCurrentLocation = lngCurrentLocationArray(Me.cboCurrentLocation.ListIndex + 1)
  Else
    dvdCurrent.lngCurrentLocation = 0
  End If

End Sub

Private Sub cboGenre_Click()
  If Me.cboGenre.ListIndex <> -1 Then
    dvdCurrent.lngGenre = lngGenreArray(Me.cboGenre.ListIndex + 1)
  Else
    dvdCurrent.lngGenre = 0
  End If
End Sub

Private Sub cboRating_Click()
  If Me.cboRating.ListIndex <> -1 Then
    dvdCurrent.lngRating = lngRatingArray(Me.cboRating.ListIndex + 1)
  Else
    dvdCurrent.lngRating = 0
  End If
End Sub

Private Sub cboRegion_Click()
  If Me.cboRegion.ListIndex <> -1 Then
    dvdCurrent.lngRegion = lngRegionArray(Me.cboRegion.ListIndex + 1)
  Else
    dvdCurrent.lngRegion = 0
  End If
End Sub

Private Sub cboUserReview_Click()
  Dim intStars As Integer, intLoop As Integer
  intStars = Me.cboUserReview
  For intLoop = 0 To 4
    Me.picStar(intLoop).Visible = False
  Next intLoop
  For intLoop = 0 To intStars - 1
    Me.picStar(intLoop).Visible = True
  Next intLoop
  dvdCurrent.bytUserReview = CByte(intStars)
End Sub

Private Sub chk169_Validate(Cancel As Boolean)
  dvdCurrent.bln169 = Me.chk169
End Sub

Private Sub chkAnimatedMenus_Validate(Cancel As Boolean)
  dvdCurrent.blnAnimatedMenus = Me.chkAnimatedMenus
End Sub

Private Sub chkAudioOther_Validate(Cancel As Boolean)
  dvdCurrent.blnAudioOther = Me.chkAudioOther
End Sub

Private Sub chkBios_Validate(Cancel As Boolean)
  dvdCurrent.blnStarBios = Me.chkBios
End Sub

Private Sub chkChinese_Validate(Cancel As Boolean)
  dvdCurrent.blnChinese = Me.chkChinese
End Sub

Private Sub chkCommentary_Validate(Cancel As Boolean)
  dvdCurrent.blnCommentary = Me.chkCommentary
End Sub

Private Sub chkdd51_Validate(Cancel As Boolean)
  dvdCurrent.blnDD51 = Me.chkdd51
End Sub

Private Sub chkDDEx_Validate(Cancel As Boolean)
  dvdCurrent.blnDDEx = Me.chkDDEx
End Sub

Private Sub chkDeletedScenes_Validate(Cancel As Boolean)
  dvdCurrent.blnDeletedScenes = Me.chkDeletedScenes
End Sub

Private Sub chkDolbyProLogic_Validate(Cancel As Boolean)
  dvdCurrent.blnDolbyProLogic = Me.chkDolbyProLogic
End Sub

Private Sub chkDolbySurround_Validate(Cancel As Boolean)
  dvdCurrent.blnDolbySurround = Me.chkDolbySurround
End Sub

Private Sub chkDTS_Validate(Cancel As Boolean)
  dvdCurrent.blnDTS = Me.chkDTS
End Sub

Private Sub chkDVDROM_Validate(Cancel As Boolean)
  dvdCurrent.blnDVDRom = Me.chkDVDROM
End Sub

Private Sub chkEnglish_Validate(Cancel As Boolean)
  dvdCurrent.blnEnglish = Me.chkEnglish
End Sub

Private Sub chkFrench_Validate(Cancel As Boolean)
  dvdCurrent.blnFrench = Me.chkFrench
End Sub

Private Sub chkFullFrame_Validate(Cancel As Boolean)
  dvdCurrent.blnFullFrame = Me.chkFullFrame
End Sub

Private Sub chkGerman_Validate(Cancel As Boolean)
  dvdCurrent.blnGerman = Me.chkGerman
End Sub

Private Sub chkJapanese_Validate(Cancel As Boolean)
  dvdCurrent.blnJapanese = Me.chkJapanese
End Sub

Private Sub chkMakingOf_Validate(Cancel As Boolean)
  dvdCurrent.blnMakingOf = Me.chkMakingOf
End Sub

Private Sub chkPanScan_Validate(Cancel As Boolean)
  dvdCurrent.blnPanScan = Me.chkPanScan
End Sub

Private Sub chkPortugese_Validate(Cancel As Boolean)
  dvdCurrent.blnPortugese = Me.chkPortugese
End Sub

Private Sub chkSceneAccess_Validate(Cancel As Boolean)
  dvdCurrent.blnSceneAccess = Me.chkSceneAccess
End Sub

Private Sub chkSDDS_Validate(Cancel As Boolean)
  dvdCurrent.blnSDDS = Me.chkSDDS
End Sub

Private Sub chkSpanish_Validate(Cancel As Boolean)
  dvdCurrent.blnSpanish = Me.chkSpanish
End Sub

Private Sub chkStereo_Validate(Cancel As Boolean)
  dvdCurrent.blnStereo = Me.chkStereo
End Sub

Private Sub chkSubTitleOther_Validate(Cancel As Boolean)
  dvdCurrent.blnSubtitleOther = Me.chkSubTitleOther
End Sub

Private Sub chkTheatricalTrailer_Validate(Cancel As Boolean)
  dvdCurrent.blnTheatricalTrailer = Me.chkTheatricalTrailer
End Sub

Private Sub chkWidescreen_Validate(Cancel As Boolean)
  dvdCurrent.blnWidescreen = Me.chkWidescreen
End Sub

Private Sub Form_Load()
  globalCode.initialise

  globalCode.fillDVDListView Me.lvwDiscs

  Me.fraGeneral.Visible = True
  Me.fraAudioVideo.Visible = False
  Me.fraFeatures.Visible = False
  '
  'Fill combos
  '
  globalCode.fillSelectCombo Me.cboGenre.Name
  globalCode.fillSelectCombo Me.cboRegion.Name
  globalCode.fillSelectCombo Me.cboRating.Name
  globalCode.fillSelectCombo Me.cboCaseType.Name
  globalCode.fillSelectCombo Me.cboCurrentLocation.Name
  
  discCode.resetAllFields
  discCode.disableDiscDisplay
  
  Me.btnAddDVD.Visible = False
End Sub


Private Sub lvwDiscs_Click()
  If checkSelected(Me.lvwDiscs) <> -1 Then
    If Me.txtDirector.Enabled = False Then discCode.enableDiscDisplay
    intFormAction = EDIT
    Me.btnAddDVD.Visible = False
    dvdCurrent.fillDVD lngDVDIDArray(checkSelected(Me.lvwDiscs))
    dvdCurrent.displayDVD
    Me.Caption = strCaption & " : " & dvdCurrent.strTitle
  End If
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show
End Sub

Private Sub optDualLayer_Click()
  If Me.optDualLayer = True Then dvdCurrent.bytDiscFormat = 1
  If Me.optDualSided = True Then dvdCurrent.bytDiscFormat = 2
  If Me.optFlipper = True Then dvdCurrent.bytDiscFormat = 3
End Sub

Private Sub optDualSided_Click()
  If Me.optDualLayer = True Then dvdCurrent.bytDiscFormat = 1
  If Me.optDualSided = True Then dvdCurrent.bytDiscFormat = 2
  If Me.optFlipper = True Then dvdCurrent.bytDiscFormat = 3
End Sub

Private Sub optFlipper_Click()
  If Me.optDualLayer = True Then dvdCurrent.bytDiscFormat = 1
  If Me.optDualSided = True Then dvdCurrent.bytDiscFormat = 2
  If Me.optFlipper = True Then dvdCurrent.bytDiscFormat = 3
End Sub

Private Sub optNTSC_Click()
  dvdCurrent.blnNTSCPAL = Me.optNTSC
End Sub

Private Sub optPAL_Click()
  dvdCurrent.blnNTSCPAL = Me.optNTSC
End Sub

Private Sub tabMain_Click()
  Select Case Me.tabMain.SelectedItem.Tag
    Case "General"
      Me.fraGeneral.Visible = True
      Me.fraAudioVideo.Visible = False
      Me.fraFeatures.Visible = False
      dvdCurrent.displayDVD
    Case "AudioVideo"
      Me.fraGeneral.Visible = False
      Me.fraAudioVideo.Visible = True
      Me.fraFeatures.Visible = False
      dvdCurrent.displayDVD
    Case "Features"
      Me.fraGeneral.Visible = False
      Me.fraAudioVideo.Visible = False
      Me.fraFeatures.Visible = True
      dvdCurrent.displayDVD
  End Select
End Sub

Private Sub txtCost_Validate(Cancel As Boolean)
  If Len(Me.txtCost) = 0 Then Exit Sub
  If IsNumeric(Me.txtCost) = True Then
    dvdCurrent.curCost = Me.txtCost
  Else
    MsgBox "Cost must be a number"
    Cancel = True
  End If
End Sub

Private Sub txtDatePurchased_Validate(Cancel As Boolean)
  If Len(Me.txtDatePurchased) = 0 Then Exit Sub
  If IsDate(Me.txtDatePurchased) = True Then
    dvdCurrent.datDatePurchased = Me.txtDatePurchased
  Else
    MsgBox "Date purchased must be a date"
    Cancel = True
  End If
End Sub

Private Sub txtDirector_Validate(Cancel As Boolean)
  dvdCurrent.strDirector = Me.txtDirector
End Sub

Private Sub txtDVDRelease_Validate(Cancel As Boolean)
  If Len(Me.txtDVDRelease) = 0 Then Exit Sub
  If IsNumeric(Me.txtDVDRelease) = True And Len(Me.txtDVDRelease) = 4 Then
    dvdCurrent.datDVDRelease = "1/1/" & Me.txtDVDRelease
  Else
    MsgBox "DVD release must be a year"
    Cancel = True
  End If
End Sub

Private Sub txtLocationPurchased_Validate(Cancel As Boolean)
  dvdCurrent.strLocationPurchased = Me.txtLocationPurchased
End Sub


Private Sub txtMovieYear_Validate(Cancel As Boolean)
  If Len(Me.txtMovieYear) = 0 Then Exit Sub
  If IsNumeric(Me.txtMovieYear) = True And Len(Me.txtMovieYear) = 4 Then
    dvdCurrent.datMovieYear = "1/1/" & Me.txtMovieYear
  Else
    MsgBox "Movie year must be a year"
    Cancel = True
  End If
End Sub

Private Sub txtRatio_Validate(Cancel As Boolean)
  If Len(Me.txtRatio) = 0 Then Exit Sub
  If IsNumeric(Me.txtRatio) = True Then
    dvdCurrent.dblRatio = Me.txtRatio
  Else
    MsgBox "Ratio must be a number"
    Cancel = True
  End If

End Sub

Private Sub txtRunningTime_Validate(Cancel As Boolean)
  If Len(Me.txtRunningTime) = 0 Then Exit Sub
  Dim intRunningTime As Integer
  intRunningTime = discCode.parseTime(Me.txtRunningTime)
  If intRunningTime <> -1 Then
    dvdCurrent.intRunningTime = intRunningTime
  Else
    MsgBox "Running time must be of form hh, hh:mm, or hh:mm:ss, and total seconds must be less than 65000 seconds " & Chr(10) & _
    "which is 18 hours.  Show me a DVD with a running time more than this and I'll make this field a long."
    Cancel = True
  End If
End Sub

Private Sub txtStudio_Validate(Cancel As Boolean)
  dvdCurrent.strStudio = Me.txtStudio
End Sub

Private Sub txtTitle_Change()
  If intFormAction = ADD_NEW And Len(Me.txtTitle) > 0 And Len(Me.txtTitle) < 150 Then Me.btnAddDVD.Enabled = True
End Sub

Private Sub txtTitle_Validate(Cancel As Boolean)
  If intFormAction = ADD_NEW Then
    If Len(Me.txtTitle) = 0 Then
      MsgBox "DVD Title must have 1 or more characters"
      Cancel = True
    End If
  End If
  If intFormAction = EDIT Then
    If Len(Me.txtTitle) = 0 And Len(Me.txtTitle) < 150 Then
      MsgBox "Cannot update, DVD Title must have more than 1 character and less than 150 characters"
      Cancel = True
    Else
      dvdCurrent.strTitle = Me.txtTitle
    End If
  End If
End Sub

