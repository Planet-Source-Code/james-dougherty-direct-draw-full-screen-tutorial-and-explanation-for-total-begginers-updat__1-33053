VERSION 5.00
Begin VB.Form frmDX 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   5985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   ControlBox      =   0   'False
   Icon            =   "frmDX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   533
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblInit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Direct Draw Initialized!!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   660
      Left            =   1800
      TabIndex        =   1
      Top             =   3120
      Width           =   6075
   End
   Begin VB.Label lblNFO 
      BackStyle       =   0  'Transparent
      Caption         =   "Press a button or click on the form to exit."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3480
   End
End
Attribute VB_Name = "frmDX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Lets try blitting a picture(Dont forget to release it)
'We will call it test pic
Private TestPic As DirectDrawSurface7
'Do we exit now?
Private EndLoop As Boolean

Private Sub Form_Load()
'not yet
EndLoop = False

If PreviewIndex = 0 Then
 InitializeDX
 lblInit.Visible = True
 lblNFO.Visible = True
ElseIf PreviewIndex = 1 Then
 InitializeDX
 lblInit.Visible = False
 lblNFO.Visible = False
 RunDemo
Else
 Unload frmDX
End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If PreviewIndex = 0 Then
 Unload frmDX
ElseIf PreviewIndex = 1 Then
 EndLoop = True
Else
 Unload frmDX
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If PreviewIndex = 0 Then
 Unload frmDX
ElseIf PreviewIndex = 1 Then
 EndLoop = True
Else
 Unload frmDX
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CleanupDX
End Sub

Private Sub RunDemo()
'lets create our surface
Set TestPic = LoadSurfaceFromFile(App.Path & "\TestPic.jpg", 640, 480, False)
'Lets setup a normal game loop
Do
 DoEvents
 'First thing we do clear the old image by blitting color
 ClearBackBuffer frmDX.hWnd
 'then we blit our picture
 BlitSurface TestPic, 0, 0, 640, 480
 'update our FPS to see what we are pushing
 UpdateFPS
 'and what is the FPS? Lets show it
 DrawText 5, 5, "FPS - " & Display_FPS
 'Display Information to the user
 DrawText 5, 25, "Press a button or click on the form to exit."
 'Then as I described we "Flip" the page like a notebook
 PrimarySurface.Flip Nothing, DDFLIP_WAIT
 'Thats it, simple?
Loop Until EndLoop = True
Unload frmDX
End Sub
