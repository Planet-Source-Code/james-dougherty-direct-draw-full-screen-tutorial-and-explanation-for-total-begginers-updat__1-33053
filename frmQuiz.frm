VERSION 5.00
Begin VB.Form frmQuiz 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quiz 1"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ok"
      Height          =   495
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCheck 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Check Answers"
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CheckBox chkAnswer5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "False"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   17
         Top             =   4200
         Width           =   735
      End
      Begin VB.CheckBox chkAnswer5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "True"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   16
         Top             =   4200
         Width           =   735
      End
      Begin VB.CheckBox chkAnswer4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "False"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   15
         Top             =   3360
         Width           =   735
      End
      Begin VB.CheckBox chkAnswer4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "True"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   14
         Top             =   3360
         Width           =   735
      End
      Begin VB.CheckBox chkAnswer3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "False"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   13
         Top             =   2400
         Width           =   735
      End
      Begin VB.CheckBox chkAnswer3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "True"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   12
         Top             =   2400
         Width           =   735
      End
      Begin VB.CheckBox chkAnswer2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "False"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   11
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkAnswer2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "True"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox chkAnswer1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "False"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox chkAnswer1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "True"
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5.) DDS_HEIGHT lets you specify the height of a surface"
         Height          =   195
         Left            =   45
         TabIndex        =   7
         Top             =   3840
         Width           =   4035
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "4.) The DirectX objects, DirectX7 and DirectDraw7, give you access       to the DX API's"
         Height          =   435
         Left            =   45
         TabIndex        =   6
         Top             =   2880
         Width           =   4830
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3.) The primary surface resides in memory"
         Height          =   195
         Left            =   45
         TabIndex        =   5
         Top             =   2040
         Width           =   2910
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2.) DDS_WIDTH reads the width of the surface"
         Height          =   195
         Left            =   45
         TabIndex        =   4
         Top             =   1200
         Width           =   3360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.) DDS_CAPS tells DX to use the capabilities of the users video card"
         Height          =   195
         Left            =   40
         TabIndex        =   3
         Top             =   360
         Width           =   4905
      End
   End
End
Attribute VB_Name = "frmQuiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()

If chkAnswer1(0).Value = 1 And _
   chkAnswer1(1).Value = 0 And _
   chkAnswer2(0).Value = 0 And _
   chkAnswer2(1).Value = 1 And _
   chkAnswer3(0).Value = 0 And _
   chkAnswer3(1).Value = 1 And _
   chkAnswer4(0).Value = 1 And _
   chkAnswer4(1).Value = 0 And _
   chkAnswer5(0).Value = 1 And _
   chkAnswer5(1).Value = 0 Then
 MsgBox "Great Job, you got every answer correct!", vbOKOnly, "Quiz Result"
Else
 MsgBox "You have missed 1 or more answers.", vbOKOnly, "Quiz Result"
End If

End Sub

Private Sub cmdOk_Click()
Unload frmQuiz
End Sub
