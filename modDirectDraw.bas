Attribute VB_Name = "modDirectDraw"
Option Explicit

Public DX As New DirectX7
Public DDraw As DirectDraw7

Public PrimarySurface As DirectDrawSurface7
Public PrimarySurfaceDescription As DDSURFACEDESC2

Public BackBuffer As DirectDrawSurface7
Public BackBufferDescription As DDSURFACEDESC2

'This is just to enable the preview button at certain parts
Public PreviewIndex As Integer

'This will hold our fps
Private FPS As Single
'And the FPS from the last cycle
Private LastFPS As Long

'This is our first preview
Public Sub InitializeDX()
Set DDraw = DX.DirectDrawCreate("")
DDraw.SetCooperativeLevel frmDX.hWnd, DDSCL_FULLSCREEN Or _
                          DDSCL_EXCLUSIVE Or DDSCL_ALLOWMODEX
DDraw.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT

PrimarySurfaceDescription.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
PrimarySurfaceDescription.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or _
                                   DDSCAPS_COMPLEX Or DDSCAPS_FLIP
PrimarySurfaceDescription.lBackBufferCount = 1
Set PrimarySurface = DDraw.CreateSurface(PrimarySurfaceDescription)

Dim Caps As DDSCAPS2
Caps.lCaps = DDSCAPS_BACKBUFFER
Set BackBuffer = PrimarySurface.GetAttachedSurface(Caps)
BackBuffer.GetSurfaceDesc BackBufferDescription
End Sub

Public Sub CleanupDX()
DDraw.RestoreDisplayMode
DDraw.SetCooperativeLevel frmDX.hWnd, DDSCL_NORMAL
Set BackBuffer = Nothing
Set PrimarySurface = Nothing
Set DDraw = Nothing
Set DX = Nothing
End Sub

'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|
'|¶¶                Rendering Functions                 ¶¶|
'|¶¶                -------------------                 ¶¶|
'|¶¶                                                    ¶¶|
'|¶¶ ClearBackBuffer:                                   ¶¶|
'|¶¶  Optional Color - What color you want the          ¶¶|
'|¶¶  background after we clear it?                     ¶¶|
'|¶¶                                                    ¶¶|
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|

Public Function ClearBackBuffer(hWnd As Long, Optional Color As Long = vbBlack)
Dim WindowRect As RECT
'get the dimensions of the window agian
DX.GetWindowRect hWnd, WindowRect
'Then we fill it with a solid color
BackBuffer.BltColorFill WindowRect, Color
End Function

'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|
'|¶¶                   FPS Functions                    ¶¶|
'|¶¶                   -------------                    ¶¶|
'|¶¶                                                    ¶¶|
'|¶¶ UpdateFPS:                                         ¶¶|
'|¶¶ (No Parameters)                                    ¶¶|
'|¶¶                                                    ¶¶|
'|¶¶ Display_FPS:                                       ¶¶|
'|¶¶ Optional sFormat - If true the FPS will be         ¶¶|
'|¶¶ displayed with the last digit only, like 150.6,    ¶¶|
'|¶¶ if it is false it will be displayed like 150.6945  ¶¶|
'|¶¶                                                    ¶¶|
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|

Public Sub UpdateFPS()
On Local Error GoTo ErrOut
Dim T As Long
Dim Delta As Single
Static fCount As Long
    
'In our main rendering loop we add 1 to fCount
fCount = fCount + 1
'if we hit 30 loops calculate the FPS
If fCount = 30 Then
 'Ok say T is 169064
 'And from our last check say LastFPS equals 120094
 T = DX.TickCount()
 'So (T - LastFPS) equals 48970
 '30000 / 48970 will give us 0.6
 'FPS will equal 0.6
 FPS = 30000 / (T - LastFPS)
 'LastFPS will equal 169064 now
 LastFPS = T
 'Reset this so we dont re-enter our loop every cycle
 fCount = 0
End If
     
'0.6 FPS...LOL I think I'm right on this,
'If not let me know
ErrOut:
End Sub

Public Property Get Display_FPS(Optional sFormat As Boolean = True) As Single
On Local Error Resume Next
If sFormat = True Then
  'Format the FPS to read like 150.6
  Display_FPS = Format$(FPS, "####.0")
Else
  'Default
  Display_FPS = FPS
End If
End Property

'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|
'|¶¶                   Text Functions                   ¶¶|
'|¶¶                   --------------                   ¶¶|
'|¶¶                                                    ¶¶|
'|¶¶ DrawText:                                          ¶¶|
'|¶¶  Left - Where you wabt the text to be placed going ¶¶|
'|¶¶         from left to right                         ¶¶|
'|¶¶  Top - Where you want the text to be placed going  ¶¶|
'|¶¶        from top to bottom                          ¶¶|
'|¶¶  Text - What you want to display or say            ¶¶|
'|¶¶  ForeColor - The color you want the text           ¶¶|
'|¶¶                                                    ¶¶|
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|

Public Sub DrawText(Left As Long, Top As Long, Text As String, Optional ForeColor As Long = vbBlack)
'Pretty simple, just define where you want it and what text.
BackBuffer.DrawText Left, Top, Text, False
'The Color of our text
BackBuffer.SetForeColor ForeColor
End Sub

'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|
'|¶¶                Surface Functions                   ¶¶|
'|¶¶                -----------------                   ¶¶|
'|¶¶ BlitSurface:                                       ¶¶|
'|¶¶  Surface - The surface to be displayed.            ¶¶|
'|¶¶  Left & Top - Where to place the surface.          ¶¶|
'|¶¶  Width & Height - The Dimension of the surface.    ¶¶|
'|¶¶ LoadSurfaceFromFile:                               ¶¶|
'|¶¶  Filename - Wheres the image located?              ¶¶|
'|¶¶  SurfaceWidth & SurfaceHeight - How big do you     ¶¶|
'|¶¶                                 want it to be?     ¶¶|
'|¶¶                                                    ¶¶|
'|¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶¶|

Public Sub BlitSurface(Surface As DirectDrawSurface7, Left As Long, Top As Long, Width As Long, Height As Long)
On Local Error GoTo ErrOut 'Just in case
Dim TempRect As RECT
'Define the dimensions
TempRect.Top = 0: TempRect.Left = 0
TempRect.Right = Width: TempRect.Bottom = Height
'blit the surface
BackBuffer.BltFast Left, Top, Surface, TempRect, DDBLTFAST_WAIT
ErrOut:
End Sub

Public Function LoadSurfaceFromFile(Filename As String, SurfaceWidth As Long, _
                                    SurfaceHeight As Long, Optional Transparent As Boolean) As DirectDrawSurface7
On Local Error Resume Next
Dim TempSurfaceDescription As DDSURFACEDESC2
Dim ColorKey As DDCOLORKEY
Dim tmpPic As IPictureDisp

'The .bmp was 900k so I threw this in
'Open the .jpg and resave it .bmp (I think this is Simons)
If UCase(Right(Filename, 3)) = "JPG" Then
 Set tmpPic = LoadPicture(Filename)
 Filename = Left(Filename, Len(Filename) - 3) & "bmp"
 SavePicture tmpPic, Filename
End If

'Initialize the surface
Set LoadSurfaceFromFile = Nothing

'Lets make it offscreen(in memory or invisable) and lets set
'the height and width
TempSurfaceDescription.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
TempSurfaceDescription.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
TempSurfaceDescription.lWidth = SurfaceWidth
TempSurfaceDescription.lHeight = SurfaceHeight
Set LoadSurfaceFromFile = DDraw.CreateSurfaceFromFile(Filename, TempSurfaceDescription)

If Transparent Then
 ColorKey.low = 0
 ColorKey.high = 0 'Black - Black wide range huh
 'And just set the color key
 LoadSurfaceFromFile.SetColorKey DDCKEY_SRCBLT, ColorKey
End If
End Function

