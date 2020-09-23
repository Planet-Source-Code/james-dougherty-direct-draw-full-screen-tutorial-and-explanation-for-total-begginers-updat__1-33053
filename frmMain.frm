VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   5670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   378
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   692
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   3
      Left            =   10320
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   20
      Top             =   360
      Width           =   10035
      Begin VB.PictureBox picTerm 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   4200
         Picture         =   "frmMain.frx":1521
         ScaleHeight     =   90
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   53
         Top             =   2640
         Visible         =   0   'False
         Width           =   3000
         Begin VB.Label lblTermDesc 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   855
            Left            =   240
            TabIndex        =   55
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label lblTermTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDSCL_ALLOWMODEX - "
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   1965
         End
      End
      Begin VB.PictureBox picTerms 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   7200
         ScaleHeight     =   3585
         ScaleWidth      =   2385
         TabIndex        =   47
         Top             =   600
         Width           =   2415
         Begin VB.Label lblTerm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDSCL_NORMAL"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   3
            Left            =   240
            TabIndex        =   52
            Top             =   2040
            Width           =   1380
         End
         Begin VB.Label lblTerm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDSCL_FULLSCREEN"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   2
            Left            =   240
            TabIndex        =   51
            Top             =   1560
            Width           =   1725
         End
         Begin VB.Label lblTerm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDSCL_EXCLUSIVE"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   1
            Left            =   240
            TabIndex        =   50
            Top             =   1080
            Width           =   1545
         End
         Begin VB.Label lblTerm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDSCL_ALLOWMODEX"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   0
            Left            =   240
            TabIndex        =   49
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Useful Terms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   48
            Top             =   120
            Width           =   1125
         End
      End
      Begin VB.Image Image3 
         Height          =   1500
         Left            =   240
         Picture         =   "frmMain.frx":1D17
         Top             =   2880
         Width           =   6000
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   3960
         Width           =   45
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2 - Initializing Direct Draw 1 of 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   3885
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":5023
         Height          =   2250
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   6255
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   0
      Left            =   10320
      Picture         =   "frmMain.frx":5355
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   5
      Top             =   480
      Width           =   10035
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "When enabled click the ""Preview"" button to run a sample Direct Draw application to follow along easier."
         Height          =   195
         Left            =   480
         TabIndex        =   19
         Top             =   3840
         Width           =   7350
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOTE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   360
         TabIndex        =   18
         Top             =   3600
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tutorial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3840
         TabIndex        =   10
         Top             =   2400
         Width           =   2040
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direct Draw Fullscreen"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1920
         TabIndex        =   9
         Top             =   1560
         Width           =   6045
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome To The"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2640
         TabIndex        =   8
         Top             =   720
         Width           =   4455
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   6
      Left            =   240
      Picture         =   "frmMain.frx":656C
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   34
      Top             =   120
      Width           =   10035
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":7783
         Height          =   450
         Left            =   240
         TabIndex        =   40
         Top             =   3000
         Width           =   9375
      End
      Begin VB.Image Image5 
         Height          =   1800
         Left            =   120
         Picture         =   "frmMain.frx":7838
         Top             =   1080
         Width           =   4500
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   39
         Top             =   3960
         Width           =   45
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 3 - Cleaning Up DirectX"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   3030
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "All we need to do is release all surfaces and kill the objects. Heres how its done."
         Height          =   210
         Left            =   240
         TabIndex        =   37
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Note:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   36
         Top             =   3720
         Width           =   1020
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Any surfaces you add need to be released just like the others are."
         Height          =   315
         Left            =   360
         TabIndex        =   35
         Top             =   3960
         Width           =   4620
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Index           =   7
      Left            =   10320
      Picture         =   "frmMain.frx":A267
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   656
      TabIndex        =   41
      Top             =   0
      Width           =   9840
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is all for this tutorial, now"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   720
         TabIndex        =   43
         Top             =   600
         Width           =   8010
      End
      Begin VB.Label lblDemo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Run Demo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4380
         TabIndex        =   46
         Top             =   3540
         Width           =   840
      End
      Begin VB.Image imgDemo 
         Height          =   360
         Left            =   4200
         Picture         =   "frmMain.frx":B47E
         Stretch         =   -1  'True
         Top             =   3480
         Width           =   1170
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "to use them. I hope this helped at least a little."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   45
         Top             =   2880
         Width           =   6915
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "More soon to come"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2160
         TabIndex        =   44
         Top             =   1440
         Width           =   5130
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I have provided other functions and a demo on how"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   840
         TabIndex        =   42
         Top             =   2520
         Width           =   7815
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   5
      Left            =   10320
      Picture         =   "frmMain.frx":BA2C
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   30
      Top             =   240
      Width           =   10035
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Now lets learn to ""Cleanup"" DX to free the memory"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   33
         Top             =   2760
         Width           =   7605
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You have now Initialized"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1680
         TabIndex        =   32
         Top             =   720
         Width           =   6435
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direct Draw In Fullscreen!!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1440
         TabIndex        =   31
         Top             =   1440
         Width           =   7005
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   4
      Left            =   10320
      Picture         =   "frmMain.frx":CC43
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   24
      Top             =   0
      Width           =   10035
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Under Dim Caps as DDSCAPS2 please put  -  Caps.lCaps = DDSCAPS_BACKBUFFER - or you will get an automation error."
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   3960
         Width           =   9060
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Note:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   28
         Top             =   3720
         Width           =   1020
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":DE5A
         Height          =   690
         Left            =   240
         TabIndex        =   27
         Top             =   720
         Width           =   8775
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 2 - Initializing Direct Draw 2 of 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   3885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   25
         Top             =   3960
         Width           =   45
      End
      Begin VB.Image Image4 
         Height          =   2250
         Left            =   120
         Picture         =   "frmMain.frx":DF81
         Top             =   1320
         Width           =   8250
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   1
      Left            =   10320
      Picture         =   "frmMain.frx":14AE1
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   6
      Top             =   240
      Width           =   10035
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":15CF8
         BeginProperty Font 
            Name            =   "ASI_Mono"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   9450
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "First we create our main DX object so we have access to the librarys or DX's API functions."
         Height          =   570
         Left            =   3840
         TabIndex        =   12
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   11
         Top             =   3960
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1 - Creating The DirectX Objects 1 of 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   4635
      End
      Begin VB.Image Image1 
         Height          =   1395
         Left            =   120
         Picture         =   "frmMain.frx":15E99
         Stretch         =   -1  'True
         Top             =   480
         Width           =   3690
      End
      Begin VB.Image Image2 
         Height          =   1500
         Left            =   120
         Picture         =   "frmMain.frx":17E86
         Top             =   2880
         Width           =   5625
      End
   End
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4815
      Index           =   2
      Left            =   240
      Picture         =   "frmMain.frx":1AC34
      ScaleHeight     =   321
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   669
      TabIndex        =   14
      Top             =   120
      Width           =   10035
      Begin VB.PictureBox picTerms2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   6960
         ScaleHeight     =   3585
         ScaleWidth      =   2505
         TabIndex        =   59
         Top             =   480
         Width           =   2535
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Useful Terms"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label lblTerm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDS_CAPS"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   7
            Left            =   240
            TabIndex        =   63
            Top             =   600
            Width           =   915
         End
         Begin VB.Label lblTerm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDS_HEIGHT"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   6
            Left            =   240
            TabIndex        =   62
            Top             =   1080
            Width           =   1065
         End
         Begin VB.Label lblTerm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDS_WIDTH"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   5
            Left            =   240
            TabIndex        =   61
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label lblTerm 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDS_BACKBUFFERCOUNT"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   165
            Index           =   4
            Left            =   240
            TabIndex        =   60
            Top             =   2040
            Width           =   2145
         End
      End
      Begin VB.PictureBox picTerm2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   3960
         Picture         =   "frmMain.frx":1BE4B
         ScaleHeight     =   90
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   200
         TabIndex        =   56
         Top             =   2520
         Visible         =   0   'False
         Width           =   3000
         Begin VB.Label lblTermTitle2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DDSCL_ALLOWMODEX - "
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   120
            TabIndex        =   58
            Top             =   120
            Width           =   1965
         End
         Begin VB.Label lblTermDesc2 
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   855
            Left            =   240
            TabIndex        =   57
            Top             =   360
            Width           =   2655
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":1C641
         BeginProperty Font 
            Name            =   "ASI_Mono"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   6135
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Step 1 - Creating The DirectX Objects 2 of 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   4635
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   15
         Top             =   3960
         Width           =   45
      End
   End
   Begin VB.Image imgPreviewOver 
      Height          =   360
      Left            =   0
      Picture         =   "frmMain.frx":1C7C8
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   1170
   End
   Begin VB.Image imgPreviewDown 
      Height          =   360
      Left            =   0
      Picture         =   "frmMain.frx":1CDF7
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1170
   End
   Begin VB.Image imgPreviewNormal 
      Height          =   360
      Left            =   0
      Picture         =   "frmMain.frx":1D4D5
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1170
   End
   Begin VB.Image imgStepNormal 
      Height          =   375
      Left            =   1200
      Picture         =   "frmMain.frx":1DA83
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   375
   End
   Begin VB.Image imgStepOver 
      Height          =   375
      Left            =   1200
      Picture         =   "frmMain.frx":1DEEF
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   375
   End
   Begin VB.Image imgStepDown 
      Height          =   375
      Left            =   1200
      Picture         =   "frmMain.frx":1E328
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   375
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   9840
      TabIndex        =   4
      Top             =   4995
      Width           =   180
   End
   Begin VB.Image imgStep 
      Height          =   375
      Index           =   2
      Left            =   9735
      Picture         =   "frmMain.frx":1E7A4
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¥"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   9450
      TabIndex        =   3
      Top             =   4995
      Width           =   90
   End
   Begin VB.Label lblStep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   8985
      TabIndex        =   2
      Top             =   4995
      Width           =   180
   End
   Begin VB.Image imgStep 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   9315
      Picture         =   "frmMain.frx":1EC10
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Image imgStep 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   8895
      Picture         =   "frmMain.frx":1F07C
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   375
   End
   Begin VB.Label lblClose 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   9990
      TabIndex        =   1
      Top             =   120
      Width           =   105
   End
   Begin VB.Label lblPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   555
      TabIndex        =   0
      Top             =   4980
      Width           =   675
   End
   Begin VB.Image imgPreview 
      Enabled         =   0   'False
      Height          =   360
      Left            =   315
      Picture         =   "frmMain.frx":1F4E8
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1170
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   9915
      Picture         =   "frmMain.frx":1FA96
      Top             =   105
      Width           =   225
   End
   Begin VB.Image imgBack 
      Height          =   5670
      Left            =   0
      Picture         =   "frmMain.frx":1FD7A
      Top             =   0
      Width           =   10380
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private HoldIndex As Integer

Private Sub Form_Load()
Dim i As Integer

imgBack.Move 0, 0
For i = 0 To 7
 picFrame(i).Move 20, 23, 656, 300
 picFrame(i).Visible = False
Next
picFrame(0).Visible = True
End Sub

Private Sub imgDemo_Click()
On Local Error Resume Next
PreviewIndex = 1
frmDX.Show
End Sub

Private Sub imgPreview_Click()
frmDX.Show
End Sub

Private Sub imgStep_Click(Index As Integer)
On Local Error Resume Next
Dim i As Integer

Select Case Index
 Case 0
  HoldIndex = HoldIndex - 1
  If HoldIndex = 4 Or HoldIndex = 5 Then
   imgPreview.Enabled = True: lblPreview.Enabled = True
   PreviewIndex = 0
  End If
  If HoldIndex = 3 Then imgPreview.Enabled = False: lblPreview.Enabled = False
  If imgStep(2).Enabled = False Then imgStep(2).Enabled = True
  If lblStep(2).Enabled = False Then lblStep(2).Enabled = True
  If HoldIndex < 0 Then
   imgStep(0).Enabled = False
   lblStep(0).Enabled = False
   imgStep(1).Enabled = False
   lblStep(1).Enabled = False
   If imgPreview.Enabled Then imgPreview.Enabled = False
   If lblPreview.Enabled Then lblPreview.Enabled = False
   Exit Sub
  End If
  picFrame(HoldIndex + 1).Visible = False
  picFrame(HoldIndex).Visible = True
 Case 1
  imgStep(0).Enabled = False
  lblStep(0).Enabled = False
  imgStep(1).Enabled = False
  lblStep(1).Enabled = False
  imgStep(2).Enabled = True
  lblStep(2).Enabled = True
  If imgPreview.Enabled Then imgPreview.Enabled = False
  If lblPreview.Enabled Then lblPreview.Enabled = False
  For i = 0 To 7
   picFrame(i).Visible = False
  Next
  picFrame(0).Visible = True
  HoldIndex = 0
 Case 2
  HoldIndex = HoldIndex + 1
  If HoldIndex = 4 Then
   imgPreview.Enabled = True: lblPreview.Enabled = True
   PreviewIndex = 0
  End If
  If HoldIndex = 6 Then imgPreview.Enabled = False: lblPreview.Enabled = False
  If imgStep(0).Enabled = False Then imgStep(0).Enabled = True
  If lblStep(0).Enabled = False Then lblStep(0).Enabled = True
  If imgStep(1).Enabled = False Then imgStep(1).Enabled = True
  If lblStep(1).Enabled = False Then lblStep(1).Enabled = True
  If HoldIndex > 7 Then
   imgStep(2).Enabled = False
   lblStep(2).Enabled = False
   Exit Sub
  End If
  picFrame(HoldIndex - 1).Visible = False
  picFrame(HoldIndex).Visible = True
End Select
End Sub

Private Sub imgClose_Click()
Unload frmDX
Unload Me
End
End Sub

'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|
'|¶¶    The code below is all for the skinning,  not    ¶¶|
'|¶¶    very important for this app.                    ¶¶|
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|

Private Sub imgPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewDown.Picture
End Sub

Private Sub imgPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewOver.Picture
End Sub

Private Sub imgPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewNormal.Picture
End Sub

Private Sub imgStep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepDown.Picture
End Sub

Private Sub imgStep_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepOver.Picture
End Sub

Private Sub imgStep_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepNormal.Picture
End Sub

Private Sub lblClose_Click()
imgClose_Click
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &HFF&
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &H0
End Sub

Private Sub lblDemo_Click()
imgDemo_Click
End Sub

Private Sub lblDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewDown.Picture
End Sub

Private Sub lblDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewOver.Picture
End Sub

Private Sub lblDemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewNormal.Picture
End Sub

Private Sub lblPreview_Click()
imgPreview_Click
End Sub

Private Sub lblPreview_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewDown.Picture
End Sub

Private Sub lblPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewOver.Picture
End Sub

Private Sub lblPreview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgPreview.Picture = imgPreviewNormal.Picture
End Sub

Private Sub lblStep_Click(Index As Integer)
imgStep_Click Index
End Sub

Private Sub lblStep_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepDown.Picture
End Sub

Private Sub lblStep_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepOver.Picture
End Sub

Private Sub lblStep_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgStep(Index).Picture = imgStepNormal.Picture
End Sub

Private Sub imgDemo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewDown.Picture
End Sub

Private Sub imgDemo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewOver.Picture
End Sub

Private Sub imgDemo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDemo.Picture = imgPreviewNormal.Picture
End Sub

Private Sub imgBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &H0
imgStep(0).Picture = imgStepNormal.Picture
imgStep(1).Picture = imgStepNormal.Picture
imgStep(2).Picture = imgStepNormal.Picture
imgPreview.Picture = imgPreviewNormal.Picture
imgDemo.Picture = imgPreviewNormal.Picture
End Sub

Private Sub lblTerm_Click(Index As Integer)
lblTermTitle.Caption = lblTerm(Index).Caption & " - "
lblTermTitle2.Caption = lblTerm(Index).Caption & " - "

Select Case Index
 Case 0
 picTerm.Top = 80
 lblTermDesc.Caption = "tells DX to use ModeX. A VGA Display Hybrid (Mode 13) with 256 kb of display memory."
 picTerm.Visible = True
 Case 1
 picTerm.Top = 112
 lblTermDesc.Caption = "Requests the exclusive level, MUST be used with DDSCL_FULLSCREEN."
 picTerm.Visible = True
 Case 2
 picTerm.Top = 144
 lblTermDesc.Caption = "tells DX that your in control of the Primary Surface, the entire screen. Your also responsible for the entire Primry Surface."
 picTerm.Visible = True
 Case 3
 picTerm.Top = 176
 lblTermDesc.Caption = "lets DX function as a normal windows application."
 picTerm.Visible = True
 Case 4
 picTerm2.Top = 168
 lblTermDesc2.Caption = "lets DX know the number of Back Buffers are to be created."
 picTerm2.Visible = True
 Case 5
 picTerm2.Top = 136
 lblTermDesc2.Caption = "lets you define the width of the surface you are creating."
 picTerm2.Visible = True
 Case 6
 picTerm2.Top = 104
 lblTermDesc2.Caption = "lets you define the height of the surface you are creating."
 picTerm2.Visible = True
 Case 7
 picTerm2.Top = 72
 lblTermDesc2.Caption = "tells DX to use the capabilities of the video card."
 picTerm2.Visible = True
End Select
End Sub

Private Sub lblTerm_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblTerm(Index).ForeColor = &H80FF&
End Sub

Private Sub picFrame_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
lblClose.ForeColor = &H0
imgStep(0).Picture = imgStepNormal.Picture
imgStep(1).Picture = imgStepNormal.Picture
imgStep(2).Picture = imgStepNormal.Picture
imgPreview.Picture = imgPreviewNormal.Picture
imgDemo.Picture = imgPreviewNormal.Picture
End Sub

Private Sub picTerms_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 0 To 3
 lblTerm(i).ForeColor = &HFF&
Next
picTerm.Visible = False
End Sub

Private Sub picTerms2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Integer
For i = 4 To 7
 lblTerm(i).ForeColor = &HFF&
Next
picTerm2.Visible = False
End Sub
