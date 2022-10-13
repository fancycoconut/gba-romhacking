VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DayAndNight"
   ClientHeight    =   8355
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   557
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraROMInformation 
      Caption         =   "ROM Information"
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Tag             =   "48"
      Top             =   1560
      Width           =   8775
      Begin VB.Label lblLanguage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   11
         Top             =   720
         Width           =   270
      End
      Begin VB.Label lblROMLanguage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Tag             =   "51"
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   270
      End
      Begin VB.Label lblHeaderCode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Tag             =   "50"
         Top             =   480
         Width           =   435
      End
      Begin VB.Label lblROM 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "???"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   270
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Tag             =   "49"
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame fraImplement 
      Caption         =   "Implement Day Night"
      Height          =   2175
      Left            =   6720
      TabIndex        =   79
      Tag             =   "63"
      Top             =   2760
      Width           =   2175
      Begin VB.PictureBox pic3 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   1935
         TabIndex        =   80
         Top             =   240
         Width           =   1935
         Begin VB.TextBox txtDesign 
            Enabled         =   0   'False
            Height          =   284
            Left            =   480
            TabIndex        =   83
            Text            =   "0x"
            Top             =   840
            Width           =   255
         End
         Begin VB.TextBox txtOffset 
            Height          =   284
            Left            =   720
            MaxLength       =   6
            TabIndex        =   82
            Text            =   "AF0000"
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdImplement 
            Caption         =   "Implement!"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   81
            Tag             =   "66"
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Needs 0x270 bytes"
            Height          =   195
            Left            =   240
            TabIndex        =   85
            Tag             =   "64"
            Top             =   120
            Width           =   1395
         End
         Begin VB.Label lblOffset 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Offset"
            Height          =   195
            Left            =   720
            TabIndex        =   84
            Tag             =   "65"
            Top             =   600
            Width           =   465
         End
      End
   End
   Begin VB.Frame fraMainPrefrences 
      Caption         =   "Main Preferences"
      Height          =   2175
      Left            =   120
      TabIndex        =   59
      Tag             =   "53"
      Top             =   2760
      Width           =   6495
      Begin VB.PictureBox pic2 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   6255
         TabIndex        =   60
         Top             =   240
         Width           =   6255
         Begin VB.Frame fraTimePeriods 
            Caption         =   "Time Periods"
            Height          =   1695
            Left            =   120
            TabIndex        =   64
            Tag             =   "54"
            Top             =   0
            Width           =   3975
            Begin VB.PictureBox pic4 
               BorderStyle     =   0  'None
               ClipControls    =   0   'False
               Height          =   1335
               Left            =   120
               ScaleHeight     =   1335
               ScaleWidth      =   3735
               TabIndex        =   65
               Top             =   240
               Width           =   3735
               Begin VB.OptionButton optMorning 
                  Caption         =   "Morning"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   72
                  Tag             =   "55"
                  Top             =   80
                  Value           =   -1  'True
                  Width           =   1575
               End
               Begin VB.OptionButton optDay 
                  Caption         =   "Day"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   71
                  Tag             =   "56"
                  Top             =   320
                  Width           =   1575
               End
               Begin VB.OptionButton optAfternoon 
                  Caption         =   "Afternoon"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   70
                  Tag             =   "57"
                  Top             =   560
                  Width           =   1575
               End
               Begin VB.OptionButton optEvening 
                  Caption         =   "Evening"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   69
                  Tag             =   "58"
                  Top             =   780
                  Width           =   1575
               End
               Begin VB.OptionButton optNight 
                  Caption         =   "Night"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   68
                  Tag             =   "59"
                  Top             =   1000
                  Width           =   1575
               End
               Begin VB.ComboBox cmbMinutes 
                  Height          =   315
                  ItemData        =   "frmMain.frx":000C
                  Left            =   3000
                  List            =   "frmMain.frx":00C4
                  Style           =   2  'Dropdown List
                  TabIndex        =   67
                  Top             =   360
                  Width           =   540
               End
               Begin VB.ComboBox cmbHours 
                  Height          =   315
                  ItemData        =   "frmMain.frx":01B8
                  Left            =   2280
                  List            =   "frmMain.frx":0204
                  Style           =   2  'Dropdown List
                  TabIndex        =   66
                  Top             =   360
                  Width           =   540
               End
               Begin VB.Label lblMinutes 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "??"
                  Height          =   195
                  Left            =   3000
                  TabIndex        =   78
                  Top             =   840
                  Width           =   150
               End
               Begin VB.Label lblSeparator2 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  Height          =   195
                  Left            =   2760
                  TabIndex        =   77
                  Top             =   840
                  Width           =   60
               End
               Begin VB.Label lblHours 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "??"
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   76
                  Top             =   840
                  Width           =   150
               End
               Begin VB.Label lblSeparator 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ":"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Left            =   2880
                  TabIndex        =   75
                  Top             =   360
                  Width           =   75
               End
               Begin VB.Label lblFrom 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "From"
                  Height          =   195
                  Left            =   1800
                  TabIndex        =   74
                  Tag             =   "60"
                  Top             =   360
                  Width           =   360
               End
               Begin VB.Label lblTo 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "To"
                  Height          =   195
                  Left            =   1920
                  TabIndex        =   73
                  Tag             =   "61"
                  Top             =   840
                  Width           =   180
               End
            End
         End
         Begin VB.Frame fraExceptions 
            Caption         =   "Exceptions"
            Height          =   1695
            Left            =   4200
            TabIndex        =   61
            Tag             =   "62"
            Top             =   0
            Width           =   1935
            Begin VB.PictureBox pic1 
               BorderStyle     =   0  'None
               ClipControls    =   0   'False
               Height          =   1335
               Left            =   120
               ScaleHeight     =   1335
               ScaleWidth      =   1695
               TabIndex        =   62
               Top             =   240
               Width           =   1695
               Begin VB.ComboBox cmbExceptions3 
                  Height          =   315
                  ItemData        =   "frmMain.frx":0268
                  Left            =   120
                  List            =   "frmMain.frx":029C
                  Style           =   2  'Dropdown List
                  TabIndex        =   91
                  Top             =   840
                  Width           =   1455
               End
               Begin VB.ComboBox cmbExceptions2 
                  Height          =   315
                  ItemData        =   "frmMain.frx":033D
                  Left            =   120
                  List            =   "frmMain.frx":0371
                  Style           =   2  'Dropdown List
                  TabIndex        =   90
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.ComboBox cmbExceptions1 
                  Height          =   315
                  ItemData        =   "frmMain.frx":0412
                  Left            =   120
                  List            =   "frmMain.frx":0446
                  Style           =   2  'Dropdown List
                  TabIndex        =   63
                  Top             =   120
                  Width           =   1455
               End
            End
         End
      End
   End
   Begin VB.Frame fraPalette 
      Caption         =   "Palette"
      Height          =   2895
      Left            =   120
      TabIndex        =   12
      Tag             =   "67"
      Top             =   5040
      Width           =   8775
      Begin VB.PictureBox pic6 
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         Height          =   2535
         Left            =   4440
         ScaleHeight     =   2535
         ScaleWidth      =   4215
         TabIndex        =   13
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtOBJFlag 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   54
            Text            =   "0000"
            Top             =   1680
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txtBGFlag 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   53
            Text            =   "0000"
            Top             =   1680
            Width           =   495
         End
         Begin VB.CommandButton cmdI 
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   36
            Top             =   1680
            Width           =   255
         End
         Begin VB.CommandButton cmdU 
            Caption         =   "U"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   35
            Top             =   1680
            Width           =   255
         End
         Begin VB.CommandButton cmdS 
            Caption         =   "S"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2040
            TabIndex        =   34
            Top             =   1680
            Width           =   255
         End
         Begin VB.CheckBox chkContrast 
            Caption         =   "Lighter?"
            Height          =   255
            Left            =   3120
            TabIndex        =   17
            Top             =   1680
            Width           =   1095
         End
         Begin VB.OptionButton optBG 
            Caption         =   "Background"
            Height          =   255
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   16
            Tag             =   "68"
            Top             =   0
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optSprite 
            Caption         =   "Sprite"
            Height          =   255
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   15
            Tag             =   "69"
            Top             =   0
            Width           =   1215
         End
         Begin VB.Frame fraChannels 
            Caption         =   "Color Channels"
            Height          =   615
            Left            =   120
            TabIndex        =   14
            Tag             =   "70"
            Top             =   1920
            Width           =   4095
            Begin VB.PictureBox pic7 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   120
               ScaleHeight     =   255
               ScaleWidth      =   3855
               TabIndex        =   86
               Top             =   240
               Width           =   3855
               Begin VB.CheckBox chkRed 
                  Caption         =   "Red"
                  Height          =   255
                  Left            =   480
                  TabIndex        =   89
                  Tag             =   "71"
                  Top             =   0
                  Width           =   975
               End
               Begin VB.CheckBox chkGreen 
                  Caption         =   "Green"
                  Height          =   255
                  Left            =   1560
                  TabIndex        =   88
                  Tag             =   "72"
                  Top             =   0
                  Width           =   975
               End
               Begin VB.CheckBox chkBlue 
                  Caption         =   "Blue"
                  Height          =   255
                  Left            =   2640
                  TabIndex        =   87
                  Tag             =   "73"
                  Top             =   0
                  Width           =   975
               End
            End
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 31"
            Height          =   255
            Index           =   15
            Left            =   3240
            TabIndex        =   52
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 30"
            Height          =   255
            Index           =   14
            Left            =   3240
            TabIndex        =   51
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 29"
            Height          =   255
            Index           =   13
            Left            =   3240
            TabIndex        =   50
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 28"
            Height          =   255
            Index           =   12
            Left            =   3240
            TabIndex        =   49
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 27"
            Height          =   255
            Index           =   11
            Left            =   2280
            TabIndex        =   48
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 26"
            Height          =   255
            Index           =   10
            Left            =   2280
            TabIndex        =   47
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 25"
            Height          =   255
            Index           =   9
            Left            =   2280
            TabIndex        =   46
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 24"
            Height          =   255
            Index           =   8
            Left            =   2280
            TabIndex        =   45
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 23"
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   44
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 22"
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   43
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 21"
            Height          =   255
            Index           =   5
            Left            =   1320
            TabIndex        =   42
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 20"
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   41
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 19"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   40
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 18"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   39
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 17"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   38
            Top             =   720
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPalOBJ 
            Caption         =   "Pal 16"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   37
            Top             =   480
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 15"
            Height          =   255
            Index           =   15
            Left            =   3240
            TabIndex        =   33
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 14"
            Height          =   255
            Index           =   14
            Left            =   3240
            TabIndex        =   32
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 13"
            Height          =   255
            Index           =   13
            Left            =   3240
            TabIndex        =   31
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 12"
            Height          =   255
            Index           =   12
            Left            =   3240
            TabIndex        =   30
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 11"
            Height          =   255
            Index           =   11
            Left            =   2280
            TabIndex        =   29
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 10"
            Height          =   255
            Index           =   10
            Left            =   2280
            TabIndex        =   28
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 9"
            Height          =   255
            Index           =   9
            Left            =   2280
            TabIndex        =   27
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 8"
            Height          =   255
            Index           =   8
            Left            =   2280
            TabIndex        =   26
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 7"
            Height          =   255
            Index           =   7
            Left            =   1320
            TabIndex        =   25
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 6"
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   24
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 5"
            Height          =   255
            Index           =   5
            Left            =   1320
            TabIndex        =   23
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 4"
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   22
            Top             =   480
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 3"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   21
            Top             =   1200
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 2"
            Height          =   255
            Index           =   2
            Left            =   360
            TabIndex        =   20
            Top             =   960
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 1"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   19
            Top             =   720
            Width           =   855
         End
         Begin VB.CheckBox chkPal 
            Caption         =   "Pal 0"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   18
            Top             =   480
            Width           =   855
         End
      End
      Begin DayAndNight.PaletteIndex palOBJ 
         Height          =   1935
         Left            =   2400
         TabIndex        =   55
         Top             =   600
         Width           =   1935
         _extentx        =   3413
         _extenty        =   3413
      End
      Begin DayAndNight.PaletteIndex palBG 
         Height          =   1935
         Left            =   240
         TabIndex        =   56
         Top             =   600
         Width           =   1935
         _extentx        =   3413
         _extenty        =   3413
      End
      Begin VB.Label lblSpritePalette 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite"
         Height          =   195
         Left            =   2400
         TabIndex        =   58
         Tag             =   "69"
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblBackgroundPalette 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Background"
         Height          =   195
         Left            =   240
         TabIndex        =   57
         Tag             =   "68"
         Top             =   360
         Width           =   840
      End
   End
   Begin DayAndNight.xpWellsStatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      Top             =   8055
      Width           =   9000
      _extentx        =   15875
      _extenty        =   529
      backcolor       =   15790320
      forecolor       =   0
      forecolordissabled=   -2147483631
      font            =   "frmMain.frx":04E7
      numberofpanels  =   3
      maskcolor       =   0
      pwidth1         =   300
      ptttext1        =   ""
      ptext1          =   "Welcome! ^^"
      penabled1       =   -1  'True
      pwidth2         =   130
      ptttext2        =   ""
      ptext2          =   "Copyright Mastermind_X"
      penabled2       =   0   'False
      pwidth3         =   200
      ptttext3        =   ""
      ptext3          =   "By ZodiacDaGreat + Interdpth"
      penabled3       =   -1  'True
   End
   Begin VB.PictureBox picDay 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      Picture         =   "frmMain.frx":050F
      ScaleHeight     =   1455
      ScaleWidth      =   9000
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox picBanner 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   9000
      TabIndex        =   2
      Top             =   0
      Width           =   9000
      Begin VB.Timer tmrBanner 
         Interval        =   30
         Left            =   8400
         Top             =   360
      End
      Begin VB.TextBox txtOpacity 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   7680
         TabIndex        =   3
         Text            =   "0"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblSwitch 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         Caption         =   "0"
         Height          =   255
         Left            =   7680
         TabIndex        =   4
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.PictureBox picNight 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      Picture         =   "frmMain.frx":9990
      ScaleHeight     =   1455
      ScaleWidth      =   9000
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      HelpContextID   =   1
      Begin VB.Menu mnuOpen 
         Caption         =   "Open ROM"
         HelpContextID   =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
         HelpContextID   =   3
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuRealTimeClock 
         Caption         =   "Real Time Clock"
         Begin VB.Menu mnuRTC 
            Caption         =   "By Interdpth"
            Index           =   0
         End
         Begin VB.Menu mnuRTC 
            Caption         =   "By ZodiacDaGreat"
            Index           =   1
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   13
      Begin VB.Menu mnuReadme 
         Caption         =   "Readme"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         HelpContextID   =   15
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BLENDFUNCTION
    BlendOp As Byte
    BlendFlags As Byte
    SourceConstantAlpha As Byte
    AlphaFormat As Byte
End Type

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function AlphaBlend Lib "MSIMG32.dll" (ByVal hdcDest As Long, ByVal nXOriginDest As Long, ByVal nYOriginDest As Long, ByVal nWidthDest As Long, ByVal nHeightDest As Long, ByVal hdcSrc As Long, ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal lBlendFunction As Long) As Long

Private Sub AdjustTime()
    ' Adds 60 and 24 to the minutes and hours if less then 0
    If Val(lblHours.Caption) < 0 Then lblHours.Caption = Val(lblHours.Caption + 24)
    If Val(lblMinutes.Caption) < 0 Then lblMinutes.Caption = Val(lblMinutes.Caption + 60)
End Sub

Private Sub chkBlue_Click()
Dim i As Integer
    If optMorning.Value = True Then
        MorningEffect.Blue = chkBlue.Value
    ElseIf optAfternoon.Value = True Then
        AfternoonEffect.Blue = chkBlue.Value
    ElseIf optEvening.Value = True Then
        EveningEffect.Blue = chkBlue.Value
    ElseIf optNight.Value = True Then
        NightEffect.Blue = chkBlue.Value
    End If
    
    For i = 0 To 15
        Call chkPal_Click(i)
        Call chkPalobj_Click(i)
    Next i
End Sub

Private Sub chkBlue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar.PanelCaption(1) = LoadResString(16)
End Sub

Private Sub chkContrast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar.PanelCaption(1) = LoadResString(19)
End Sub

Private Sub chkGreen_Click()
Dim i As Integer
    If optMorning.Value = True Then
        MorningEffect.Green = chkGreen.Value
    ElseIf optAfternoon.Value = True Then
        AfternoonEffect.Green = chkGreen.Value
    ElseIf optEvening.Value = True Then
        EveningEffect.Green = chkGreen.Value
    ElseIf optNight.Value = True Then
        NightEffect.Green = chkGreen.Value
    End If
    
    For i = 0 To 15
        Call chkPal_Click(i)
        Call chkPalobj_Click(i)
    Next i
End Sub

Private Sub chkGreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar.PanelCaption(1) = LoadResString(17)
End Sub

Private Sub chkPal_Click(Index As Integer)
Dim i As Integer
Dim Flag As Long
    palBG.ControlArrayIndex = Index
    If chkPal(Index).Value = 1 Then
        palBG.DoEffect = 1
    Else
        palBG.DoEffect = 0
    End If
    
    For i = 0 To 15
        Flag = Flag Or (chkPal(i).Value * 2 ^ i)
    Next i
    
    txtBGFlag.Text = Right("0000" & Hex(Flag), 4)
End Sub

Private Sub chkPalobj_Click(Index As Integer)
Dim i As Integer
Dim Flag As Long
    palOBJ.ControlArrayIndex = Index
    If chkPalOBJ(Index).Value = 1 Then
        palOBJ.DoEffect = 1
    Else
        palOBJ.DoEffect = 0
    End If
    
    For i = 0 To 15
        Flag = Flag Or (chkPalOBJ(i).Value * 2 ^ i)
    Next i
    
    txtOBJFlag.Text = Right("0000" & Hex(Flag), 4)
End Sub

Private Sub chkRed_Click()
Dim i As Integer
    If optMorning.Value = True Then
        MorningEffect.Red = chkRed.Value
    ElseIf optAfternoon.Value = True Then
        AfternoonEffect.Red = chkRed.Value
    ElseIf optEvening.Value = True Then
        EveningEffect.Red = chkRed.Value
    ElseIf optNight.Value = True Then
        NightEffect.Red = chkRed.Value
    End If

    For i = 0 To 15 ' Applying the effects
        Call chkPal_Click(i)
        Call chkPalobj_Click(i)
    Next i
End Sub

Private Sub chkRed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar.PanelCaption(1) = LoadResString(18)
End Sub

Private Sub cmbHours_Click()
    If optMorning.Value = True Then
        tmMorning.Hours = cmbHours.ListIndex
    ElseIf optDay.Value = True Then
        tmDay.Hours = cmbHours.ListIndex
    ElseIf optAfternoon.Value = True Then
        tmAfternoon.Hours = cmbHours.ListIndex
    ElseIf optEvening.Value = True Then
        tmEvening.Hours = cmbHours.ListIndex
    ElseIf optNight.Value = True Then
        tmNight.Hours = cmbHours.ListIndex
    End If
End Sub

Private Sub cmbMinutes_Click()
    If optMorning.Value = True Then
        tmMorning.Minutes = cmbMinutes.ListIndex
    ElseIf optDay.Value = True Then
        tmDay.Minutes = cmbMinutes.ListIndex
    ElseIf optAfternoon.Value = True Then
        tmAfternoon.Minutes = cmbMinutes.ListIndex
    ElseIf optEvening.Value = True Then
        tmEvening.Minutes = cmbMinutes.ListIndex
    ElseIf optNight.Value = True Then
        tmNight.Minutes = cmbMinutes.ListIndex
    End If
End Sub

Private Sub cmdI_Click()
Dim X As Integer
If optBG.Value = True Then
    For X = 0 To 15
        chkPal(X).Value = (chkPal(X).Value) Xor 1
    Next X
Else
    For X = 0 To 15
        chkPalOBJ(X).Value = (chkPalOBJ(X).Value) Xor 1
    Next X
End If
End Sub

Private Sub cmdImplement_Click()
Dim Contrast As Byte
Dim arrTemp() As Byte
    If txtOffset.Text = vbNullString Then Exit Sub
    If CLng("&H" & txtOffset.Text) Mod 4 <> 0 Then Exit Sub
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        ' Applying Nintendo Fix
        If NintendoFix <> 0 Then
            arrTemp = LoadResData("NINTENDOFIX", "FIXES")
            Put #iFileNum, NintendoFix + 1, arrTemp
            Erase arrTemp
        End If
        
        ' Applying SWI01 patch
        Put #iFileNum, SWI01Pos + 1, CByte(&HFE)
        If Left(sHeader, 2) = "AX" Then
            Put #iFileNum, (SWI01Pos + &H78) + 1, &HF0550000
        End If
        
        ' Patching Hook 1
        If Left(sHeader, 2) = "AX" Then
            PutFiller sFilePath, Hook1, 3
            Put #iFileNum, Hook1 + 1 + 3, &H470849
            Put #iFileNum, Hook1 + 1 + 6, CLng("&H" & txtOffset.Text) + 1 + &H8000000
            PutFiller sFilePath, Hook1 + 10, 14
        ElseIf Left(sHeader, 3) = "BPE" Then
            PutFiller sFilePath, Hook1, 50
            Put #iFileNum, Hook1 + 1 + 3, &H470849
            Put #iFileNum, Hook1 + 1 + 38, &H4708BC02
            Put #iFileNum, Hook1 + 1 + 6, CLng("&H" & txtOffset.Text) + 1 + &H8000000
        End If
        
        ' Patching Hook 2
        Put #iFileNum, Hook2 + 1, &H470A4901
        PutFiller sFilePath, Hook2 + 4, 14
        If Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then
            Put #iFileNum, Hook2 + 1 + 8, CLng("&H" & txtOffset.Text) + 1 + 1052 + &H8000000
        Else
            Put #iFileNum, Hook2 + 1 + 8, CLng("&H" & txtOffset.Text) + 1 + 232 + &H8000000
        End If
        
        ' Inserting Check Time Routine
        If Left(sHeader, 2) = "AX" Or Left(sHeader, 3) = "BPE" Then
            arrTemp = LoadResData("RSECHECKTIME", "ROUTINES")
            If Left(sHeader, 2) = "AX" Then arrTemp(100) = &H37
            
            arrTemp(108) = cmbExceptions1.ListIndex ' Applying the exceptions
            arrTemp(112) = cmbExceptions2.ListIndex
            arrTemp(116) = cmbExceptions3.ListIndex
            
            EditArray arrTemp, 164, RTC, 4
            EditArray arrTemp, 168, RTCReturnAddress, 4
            EditArray arrTemp, 192, IndoorFlag, 4
            EditArray arrTemp, 196, MenuFlag, 4
            EditArray arrTemp, 200, BattleFlag, 4
            
            EditArray arrTemp, 208, LeftShift(tmMorning.Hours, &H10) Or LeftShift(tmMorning.Minutes, &H8), 4
            EditArray arrTemp, 212, LeftShift(tmDay.Hours, &H10) Or LeftShift(tmDay.Minutes, &H8), 4
            EditArray arrTemp, 216, LeftShift(tmAfternoon.Hours, &H10) Or LeftShift(tmAfternoon.Minutes, &H8), 4
            EditArray arrTemp, 220, LeftShift(tmEvening.Hours, &H10) Or LeftShift(tmEvening.Minutes, &H8), 4
            EditArray arrTemp, 224, LeftShift(tmNight.Hours, &H10) Or LeftShift(tmNight.Minutes, &H8), 4
        Else
            arrTemp = LoadResData("RTC", "ROUTINES")
            arrTemp(934) = cmbExceptions1.ListIndex ' Applying the exceptions
            arrTemp(938) = cmbExceptions2.ListIndex
            arrTemp(942) = cmbExceptions3.ListIndex
            
            EditArray arrTemp, 1012, IndoorFlag, 4
            EditArray arrTemp, 1016, MenuFlag, 4
            EditArray arrTemp, 1020, BattleFlag, 4
            
            EditArray arrTemp, 1028, LeftShift(tmMorning.Hours, &H10) Or LeftShift(tmMorning.Minutes, &H8), 4
            EditArray arrTemp, 1032, LeftShift(tmDay.Hours, &H10) Or LeftShift(tmDay.Minutes, &H8), 4
            EditArray arrTemp, 1036, LeftShift(tmAfternoon.Hours, &H10) Or LeftShift(tmAfternoon.Minutes, &H8), 4
            EditArray arrTemp, 1040, LeftShift(tmEvening.Hours, &H10) Or LeftShift(tmEvening.Minutes, &H8), 4
            EditArray arrTemp, 1044, LeftShift(tmNight.Hours, &H10) Or LeftShift(tmNight.Minutes, &H8), 4
            
            Put #iFileNum, HookRTC + 1, &H469F4B01
            Put #iFileNum, HookRTC + 1 + 4, CInt(&H0)
            Put #iFileNum, HookRTC + 1 + 6, CLng("&H" & txtOffset.Text) + &H8000000
        End If
        Put #iFileNum, CLng("&H" & txtOffset.Text) + 1, arrTemp
        Erase arrTemp
        
        ' Inserting DMA3 Routine
        arrTemp = LoadResData("DMA3", "ROUTINES")
        If chkContrast.Value = 1 Then
            Contrast = &H5E
        Else
            Contrast = &H1E
        End If
        
        arrTemp(76) = Contrast ' Applying the bytes to change the palette change contrast
        arrTemp(150) = Contrast
        arrTemp(224) = Contrast
        arrTemp(298) = Contrast
        
        If MorningEffect.Red = 1 Then ' Applying Morning Palette Shades
            EditArray arrTemp, 88, 0, 2
        End If
        If MorningEffect.Green = 1 Then
            EditArray arrTemp, 90, 0, 2
        End If
        If MorningEffect.Blue = 1 Then
            EditArray arrTemp, 92, 0, 2
        End If
        
        If AfternoonEffect.Red = 1 Then ' Applying Afternoon Palette Shades
            EditArray arrTemp, 162, 0, 2
        End If
        If AfternoonEffect.Green = 1 Then
            EditArray arrTemp, 164, 0, 2
        End If
        If AfternoonEffect.Blue = 1 Then
            EditArray arrTemp, 166, 0, 2
        End If
        
        If EveningEffect.Red = 1 Then ' Applying Evening Palette Shades
            EditArray arrTemp, 236, 0, 2
        End If
        If EveningEffect.Green = 1 Then
            EditArray arrTemp, 238, 0, 2
        End If
        If EveningEffect.Blue = 1 Then
            EditArray arrTemp, 240, 0, 2
        End If
        
        If NightEffect.Red = 1 Then ' Applying Night Palette Shades
            EditArray arrTemp, 310, 0, 2
        End If
        If NightEffect.Green = 1 Then
            EditArray arrTemp, 312, 0, 2
        End If
        If NightEffect.Blue = 1 Then
            EditArray arrTemp, 314, 0, 2
        End If
        
        EditArray arrTemp, 368, PaletteOriginal, 4
        EditArray arrTemp, 380, DMA3ReturnAddress, 4
        EditArray arrTemp, 384, CLng("&H" & frmMain.txtBGFlag.Text), 2
        EditArray arrTemp, 386, CLng("&H" & frmMain.txtOBJFlag.Text), 2
        
        If Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then
            Put #iFileNum, CLng("&H" & txtOffset.Text) + 1 + 1052, arrTemp
        Else
            Put #iFileNum, CLng("&H" & txtOffset.Text) + 1 + 232, arrTemp
        End If
        
        MsgBox LoadResString(33), vbInformation
        Erase arrTemp
    Close #iFileNum
End Sub

Private Sub cmdImplement_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar.PanelCaption(1) = LoadResString(20)
End Sub

Private Sub cmdS_Click()
Dim X As Integer
If optBG.Value = True Then
    For X = 0 To 15
        chkPal(X).Value = 1
    Next X
Else
    For X = 0 To 15
        chkPalOBJ(X).Value = 1
    Next X
End If
End Sub

Private Sub cmdU_Click()
Dim X As Integer
If optBG.Value = True Then
    For X = 0 To 15
        chkPal(X).Value = 0
    Next X
Else
    For X = 0 To 15
        chkPalOBJ(X).Value = 0
    Next X
End If
End Sub

Private Sub DefaultValues()
    cmbExceptions1.ListIndex = 8 ' Exceptions
    cmbExceptions2.ListIndex = 4
    cmbExceptions3.ListIndex = 9
    
    tmMorning.Hours = 4 ' Time
    tmDay.Hours = 6
    tmAfternoon.Hours = 17
    tmEvening.Hours = 18
    tmEvening.Minutes = 30
    tmNight.Hours = 22
    
    MorningEffect.Red = 1     ' Palette Shades
    MorningEffect.Green = 1
    AfternoonEffect.Red = 1
    EveningEffect.Blue = 1
End Sub

Private Sub Form_Load()
    Localize Me
    SetIcon Me.hWnd, "AAA"
    DefaultValues
    optMorning_Click
    LoadPalette sHeader
    mnuRTC(Val(GetFromINI("Settings", "RTC", 0, App.Path & "\Settings.ini"))).Checked = True

    StatusBar.PanelCaption(1) = LoadResString(23)
End Sub

Private Sub Form_Paint()
Dim lBlend As Long
Dim BlendFunc As BLENDFUNCTION
Const AC_SRC_OVER = &H0 ' BlendOp
Const AC_SRC_ALPHA = &H1 ' AlphaFormat
    picDay.AutoRedraw = True
    picNight.AutoRedraw = True
    picBanner.AutoRedraw = True

    BlendFunc.AlphaFormat = 0
    BlendFunc.BlendFlags = 0
    BlendFunc.BlendOp = AC_SRC_OVER
    BlendFunc.SourceConstantAlpha = 255
    CopyMemory lBlend, BlendFunc, 4
    
    AlphaBlend picBanner.hDC, 0, 0, picDay.ScaleWidth \ Screen.TwipsPerPixelX, picDay.ScaleHeight \ Screen.TwipsPerPixelY, picDay.hDC, 0, 0, picDay.ScaleWidth \ Screen.TwipsPerPixelX, picDay.ScaleHeight \ Screen.TwipsPerPixelY, lBlend
    BlendFunc.SourceConstantAlpha = Val(txtOpacity.Text)
    CopyMemory lBlend, BlendFunc, 4
    
    AlphaBlend picBanner.hDC, 0, 0, picNight.ScaleWidth \ Screen.TwipsPerPixelX, picNight.ScaleHeight \ Screen.TwipsPerPixelY, picNight.hDC, 0, 0, picNight.ScaleWidth \ Screen.TwipsPerPixelX, picNight.ScaleHeight \ Screen.TwipsPerPixelY, lBlend
    picBanner.Refresh
    
    picDay.AutoRedraw = False
    picNight.AutoRedraw = False
    picBanner.AutoRedraw = False
End Sub

Private Sub fraExceptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar.PanelCaption(1) = LoadResString(21)
End Sub

Private Sub fraTimePeriods_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StatusBar.PanelCaption(1) = LoadResString(22)
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuOpen_Click()
Dim sResult As String
Dim cdgOpen As clsCommonDialog
    Set cdgOpen = New clsCommonDialog
    sResult = cdgOpen.ShowOpen(Me.hWnd, LoadResString(2) & "...", , "Gameboy Advance ROMs (*.gba,*.agb,*.bin)|*.gba;*.agb;*.bin")
    If LenB(sResult) = 0 Then GoTo Endme
    sFilePath = sResult
    
    iFileNum = FreeFile
    Open sFilePath For Binary As #iFileNum
        Get #iFileNum, &HAC + 1, sHeader
        
        If LOF(iFileNum) > 16777216 Then
            EnlargedROM = 1
            txtOffset.MaxLength = 7
        Else
            EnlargedROM = 0
            txtOffset.MaxLength = 6
        End If
    
        If OpenROM(sHeader) = False Then GoTo UnSupportedROM
        CheckROMLanguage sHeader, lblLanguage
        LoadPalette sHeader
    Close #iFileNum
    
    With Me
        .lblROM.ToolTipText = sFilePath
        .lblHeader = sHeader
        .cmdImplement.Enabled = True
    End With
    
    If Left(sHeader, 3) = "BPR" Or Left(sHeader, 3) = "BPG" Then
        lblInfo.Caption = LoadResString(74)
    Else
        lblInfo.Caption = LoadResString(64)
    End If
    GoTo Endme
    
UnSupportedROM:
    sFilePath = vbNullString

    With Me
        .lblROM.Caption = "???"
        .lblHeader.Caption = "???"
        .lblLanguage.Caption = "???"
        .lblROM.ToolTipText = vbNullString
        .cmdImplement.Enabled = False
    End With
    
    DefaultValues
Endme:
    Set cdgOpen = Nothing
End Sub

Private Sub mnuQuit_Click()
    Unload Me
End Sub

Private Sub mnuReadme_Click()
Dim arrTemp() As Byte
    arrTemp = LoadResData("README", 100)
    WriteByteArray App.Path & "\Readme.txt", arrTemp, 0
    Shell "notepad.exe " & App.Path & "\Readme.txt", vbNormalFocus
    'Kill App.Path & "\Readme.txt"
    Erase arrTemp
End Sub

Private Sub optAfternoon_Click()
    cmbHours.ListIndex = tmAfternoon.Hours
    cmbMinutes.ListIndex = tmAfternoon.Minutes
    
    If tmEvening.Minutes = 0 Then
        lblHours.Caption = Str(tmEvening.Hours - 1)
        lblMinutes.Caption = Str(tmEvening.Minutes - 1)
    Else
        lblHours.Caption = Str(tmEvening.Hours)
        lblMinutes.Caption = Str(tmEvening.Minutes - 1)
    End If
    AdjustTime
    
    chkRed.Enabled = True
    chkGreen.Enabled = True
    chkBlue.Enabled = True
    
    chkRed.Value = AfternoonEffect.Red
    chkGreen.Value = AfternoonEffect.Green
    chkBlue.Value = AfternoonEffect.Blue
End Sub

Private Sub optBG_Click()
Dim i As Integer
    For i = 0 To 15
        chkPal(i).Visible = True
    Next i
    For i = 0 To 15
        chkPalOBJ(i).Visible = False
    Next i
    txtBGFlag.Visible = True
    txtOBJFlag.Visible = False
End Sub

Private Sub optDay_Click()
    cmbHours.ListIndex = tmDay.Hours
    cmbMinutes.ListIndex = tmDay.Minutes
    
    If tmAfternoon.Minutes = 0 Then
        lblHours.Caption = Str(tmAfternoon.Hours - 1)
        lblMinutes.Caption = Str(tmAfternoon.Minutes - 1)
    Else
        lblHours.Caption = Str(tmAfternoon.Hours)
        lblMinutes.Caption = Str(tmAfternoon.Minutes - 1)
    End If
    AdjustTime
    
    chkRed.Enabled = False ' Disabled only for Day
    chkGreen.Enabled = False
    chkBlue.Enabled = False
    
    chkRed.Value = 1
    chkGreen.Value = 1
    chkBlue.Value = 1
End Sub

Private Sub optEvening_Click()
    cmbHours.ListIndex = tmEvening.Hours
    cmbMinutes.ListIndex = tmEvening.Minutes
    
    If tmNight.Minutes = 0 Then
        lblHours.Caption = Str(tmNight.Hours - 1)
        lblMinutes.Caption = Str(tmNight.Minutes - 1)
    Else
        lblHours.Caption = Str(tmNight.Hours)
        lblMinutes.Caption = Str(tmNight.Minutes - 1)
    End If
    AdjustTime
    
    chkRed.Enabled = True
    chkGreen.Enabled = True
    chkBlue.Enabled = True
    
    chkRed.Value = EveningEffect.Red
    chkGreen.Value = EveningEffect.Green
    chkBlue.Value = EveningEffect.Blue
End Sub

Private Sub optMorning_Click()
    cmbHours.ListIndex = tmMorning.Hours
    cmbMinutes.ListIndex = tmMorning.Minutes
    
    If tmDay.Minutes = 0 Then
        lblHours.Caption = Str(tmDay.Hours - 1)
        lblMinutes.Caption = Str(tmDay.Minutes - 1)
    Else
        lblHours.Caption = Str(tmDay.Hours)
        lblMinutes.Caption = Str(tmDay.Minutes - 1)
    End If
    AdjustTime
    
    chkRed.Enabled = True
    chkGreen.Enabled = True
    chkBlue.Enabled = True
    
    chkRed.Value = MorningEffect.Red
    chkGreen.Value = MorningEffect.Green
    chkBlue.Value = MorningEffect.Blue
End Sub

Private Sub optNight_Click()
    cmbHours.ListIndex = tmNight.Hours
    cmbMinutes.ListIndex = tmNight.Minutes
    
    If tmMorning.Minutes = 0 Then
        lblHours.Caption = Str(tmMorning.Hours - 1)
        lblMinutes.Caption = Str(tmMorning.Minutes - 1)
    Else
        lblHours.Caption = Str(tmMorning.Hours)
        lblMinutes.Caption = Str(tmMorning.Minutes - 1)
    End If
    AdjustTime
    
    chkRed.Enabled = True
    chkGreen.Enabled = True
    chkBlue.Enabled = True
    
    chkRed.Value = NightEffect.Red
    chkGreen.Value = NightEffect.Green
    chkBlue.Value = NightEffect.Blue
End Sub

Private Sub optSprite_Click()
Dim i As Integer
    For i = 0 To 15
        chkPal(i).Visible = False
    Next i
    For i = 0 To 15
        chkPalOBJ(i).Visible = True
    Next i
    txtBGFlag.Visible = False
    txtOBJFlag.Visible = True
End Sub

Private Sub tmrBanner_Timer()
    If lblSwitch.Caption = 0 Then
        txtOpacity.Text = txtOpacity.Text + 1
    Else
        txtOpacity.Text = txtOpacity.Text - 1
    End If

    If txtOpacity.Text = 255 Then
        lblSwitch.Caption = 1
    End If

    If txtOpacity.Text = 0 Then
        lblSwitch.Caption = 0
    End If

    If txtOpacity.Text < 255 Then
        tmrBanner.Interval = 30
        tmrBanner.Enabled = True
        Form_Paint
    End If
End Sub
