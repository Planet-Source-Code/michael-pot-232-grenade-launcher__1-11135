VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5700
   ClientLeft      =   1920
   ClientTop       =   1560
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   380
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============== GRENADE LAUNCHER by Michael Pote (michaelpote@worldonline.co.za)=================

'This is just a simple example of game programming using DirectSound and DirectDraw.
'(please vote...)


Dim Pddsd As DDSURFACEDESC2, ScrDdsd As DDSURFACEDESC2
Dim Primary As DirectDrawSurface7 'these are direct draw surfaces and their respective descriptions.
Dim Backbuffer As DirectDrawSurface7
Dim Surf As DirectDrawSurface7, SDDSD As DDSURFACEDESC2
Dim Ending As Boolean, bRestore As Boolean
Dim MX As Long, MY As Long, Gc As Long, I As Long
Dim G(0 To 5) As Grenade 'The grenades.
Dim P(0 To 1100) As Particle    'The shrapnel pieces.
Dim sndLau As DirectSoundBuffer 'All the direct sound buffers holding the various sounds.
Dim sndHit1 As DirectSoundBuffer
Dim sndHit2 As DirectSoundBuffer
Dim sndExp As DirectSoundBuffer, Pp As Long
Dim Shr(1 To 5) As Thing 'These hold the information on where in the art.bmp file
                         'the program must find the shrapnel pieces.
Dim T As Long, ButX As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then EndIt
End Sub

Private Sub Form_Load()
Init  'This runs a sub which sets up directx.
InitSurfaces 'this loads the sounds and the art.bmp file into memory.
FillShrs 'This fills in the coordinates for the shrapnel pieces.


'------------
bRestore = False
Do Until ExModeActive
DoEvents
bRestore = True
Loop
DoEvents
If bRestore Then
bRestore = False
DD.RestoreAllSurfaces
InitSurfaces
End If
'/\ --- all this does is if the program looses focus, it refreshes the surfaces.
'-----------

Dim ScreenRect As RECT 'Describes a rectangle of the whole screen
With ScreenRect
.Bottom = 600
.Right = 800
.Left = 0
.Top = 0
End With

Ending = False


'================================ Main Loop
Do While Ending = False
'--------------
DoEvents
bRestore = False
Do Until ExModeActive
DoEvents
bRestore = True
Loop
DoEvents
If bRestore Then
bRestore = False
DD.RestoreAllSurfaces
InitSurfaces
End If
'-------------- See above note.

Backbuffer.BltColorFill ScreenRect, 0   'Clear the screen in black

BltFast 615 + ButX, MY - 70, 0, 159, 185, 141, Surf, True 'Draw the grenade launcher.

If ButX >= 1 Then ButX = ButX - ((ButX / 20) + 1)
'This makes the launcher kick back
'after you fire it.


For I = 0 To 5    'Loop through all grenades
If G(I).Used = True Then  'It the grenade is 'used' (bouncing around the screen)
With G(I)
.V = .V + 1   'Add gravity

If .X <= 0 Or .X >= 800 Then  'Check if its hitting the walls
.SV = -.SV

sndHit2.Stop  'Stop the hit sound
sndHit2.SetCurrentPosition 0 'Rewind to the beginig
sndHit2.Play DSBPLAY_DEFAULT 'Play it again.

End If

If .Y >= 600 Then 'The grenade is hitting the floor
.Y = 599
.V = -(.V * 0.9) 'Make it bounce
.NB = .NB + 1 'Increace the Number of Bounces '.NB'
sndHit1.Stop 'Stop the hit sound
sndHit1.SetCurrentPosition 0 'Rewind
sndHit1.Play DSBPLAY_DEFAULT 'Play
End If

If .NB = 5 Then 'If the grenade has bounced 5 times

'Setup the shrapnel pieces
For Pp = (I * 100) + 1 To (I * 100) + 100
P(Pp).Life = (Rnd * 300) + 300
P(Pp).X = .X
P(Pp).Y = .Y - 2
P(Pp).V = -(Rnd * 20) - 5
P(Pp).SV = -(Rnd * 20) + 10
Next

sndExp.Stop 'Stop the explosion sound
sndExp.SetCurrentPosition 0 'Rewind
sndExp.Play DSBPLAY_DEFAULT 'Play
End If

.X = .X + .SV 'Add the velocities onto the x & y coordinates.
.Y = .Y + .V

If .NB >= 5 Then 'If the grenade has bounced more than five times, show the explosion
BltFast .X - 125, .Y - 75, 0, 0, 249, 150, Surf, True
.NB = .NB + 1
If .NB >= 9 Then G(I).Used = False 'if the explosion has been shown for 4 more frames, then kill the grenade
Else
BltFast .X - 32, .Y - 23, 186, 152, 64, 46, Surf, True 'Otherwise just draw a simple grenade
End If

End With

ElseIf G(I).Used = False Then 'If the grenade is not in use, then
                              'it must have exploded so do it's shrapnel

For Pp = (I * 100) + 1 To (I * 100) + 100
With P(Pp) 'go through all shrapnel pieces
If .Life <= 0 Then GoTo Don 'Dont display the particle if it has no more life left
.V = .V + 1
If .X <= 0 Or .X >= 800 Then .SV = -.SV 'bounce of side wall

.X = .X + .SV
.Y = .Y + .V

If .Y >= 600 Then 'bounce off floor
.Y = 601
.V = -(.V * (Rnd * 0.4) + 0.1)
End If

T = Pp - (I * 100)
T = (T Mod 5) + 1 'Which piece of shrapnel should be displayed
BltFast .X, .Y, CSng(Shr(T).X), CSng(Shr(T).Y), CSng(Shr(T).Wid), CSng(Shr(T).Hgt), Surf, True 'draw that piece

.Life = .Life - 1 'Take away some life

Don:
End With
Next
End If
Next
Primary.Flip Nothing, DDFLIP_WAIT 'Expose the back buffer to the screen
Loop
EndIt

End Sub
Public Sub Init()
'This is a directX set-up routine
Dim Ret As Boolean
On Local Error GoTo ErrorOut
Set DD = DirectX.DirectDrawCreate("")
Set DS = DirectX.DirectSoundCreate("")

Me.Show
DS.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL 'Allow direct sound to play sounds

DD.SetCooperativeLevel Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
DD.SetDisplayMode 800, 600, 16, 0, DDSDM_DEFAULT 'change the resolution.

     
'-----------------Creating a Primary and Back Buffer--------------
     Pddsd.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
     Pddsd.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
     Pddsd.lBackBufferCount = 1
     DoEvents
     Set Primary = DD.CreateSurface(Pddsd)
     Dim Caps As DDSCAPS2
     Caps.lCaps = DDSCAPS_BACKBUFFER
     Set Backbuffer = Primary.GetAttachedSurface(Caps)
     Backbuffer.GetSurfaceDesc ScrDdsd
     Backbuffer.SetFontTransparency True
     Backbuffer.SetForeColor vbWhite

'--------------------------------------------------------------

InitFlag = True 'all's well that ends well
Exit Sub
ErrorOut: 'Ouch
MsgBox "Sorry but an error occured while trying to setup directX." & Chr(13) & "Make sure you have the latest version of DirectX.", vbCritical
EndIt
End Sub

Public Sub EndIt()
'Restore the old resolution and quit out of the program
Ending = True
DD.RestoreDisplayMode
DD.SetCooperativeLevel Me.hWnd, DDSCL_NORMAL
If Err.Description <> "" Then MsgBox Err.Description, vbCritical
DoEvents
Unload Me
End
End Sub

Public Sub BltFast(dX As Long, dY As Long, SrcX As Single, SrcY As Single, SrcWid As Single, SrcHgt As Single, ByRef Surf As DirectDrawSurface7, ColourKey As Boolean)
'This sub draws something to the screen
'it is quite complicated.

On Local Error GoTo Errot

If dX > 800 Or dY > 600 Then Exit Sub
'The thing you are drawing to the screen is outside the screen so dont even bother.



Dim SrcRect As RECT, Retval 'with directdraw, you have to use RECT's

'DirectDraw has the annoying habit:
'if your picture is slightly cut off, the whole thing disapears!
'This code solves that and changes width and height into RECT coordinates
'------------------------------------------------------\/\/\/
With SrcRect

If dY <= 0 Then
.Top = -dY + SrcY
dY = 0
Else
.Top = SrcY
End If

If dX <= 0 Then
.Left = -(X) + SrcX
dX = 0
Else
.Left = SrcX
End If

If dY + SrcHgt > 600 Then
.Bottom = SrcY + SrcHgt - ((dY + SrcHgt) - 600)
Else
.Bottom = SrcY + SrcHgt
End If

If dX + SrcWid > 800 Then
.Right = SrcX + SrcWid - ((dX + SrcWid) - 800)
Else
.Right = SrcX + SrcWid
End If

End With
'----------------------------------------/\/\/\


If ColourKey Then 'If it needs to be transparent:
Retval = Backbuffer.BltFast(dX, dY, Surf, SrcRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
Retval = Backbuffer.BltFast(dX, dY, Surf, SrcRect, DDBLTFAST_WAIT)
End If
Exit Sub
Errot:
EndIt
End Sub

Private Sub CreateSurface(ByRef Surf As DirectDrawSurface7, ByRef Ddsd As DDSURFACEDESC2, Filename As String, Wid As Integer, Hgt As Integer, ColourKey As Boolean)
'Load a Bitmap into memory
On Local Error GoTo SIE
Set Surf = Nothing                                   'Clear the surface
Ddsd.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
Ddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
Ddsd.lWidth = Wid
Ddsd.lHeight = Hgt
Set Surf = DD.CreateSurfaceFromFile(Filename, Ddsd) 'Load the bitmap
If ColourKey = True Then
Dim key As DDCOLORKEY
key.low = 0
key.high = 0
Surf.SetColorKey DDCKEY_SRCBLT, key
End If
Exit Sub
SIE:
EndIt
End Sub

Sub InitSurfaces()
'Load the surface and sounds.
CreateSurface Surf, SDDSD, App.Path & "\Art.bmp", 300, 300, True

LoadWave App.Path & "\Launch.wav", sndLau
LoadWave App.Path & "\Hit1.wav", sndHit1
LoadWave App.Path & "\Hit2.wav", sndHit2
LoadWave App.Path & "\Explode.wav", sndExp
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Look through all grenades, if one is not used, then use it...

For Gc = 0 To 5
If G(Gc).Used = False Then
ButX = 80                        'Kick the launcher back
G(Gc).Used = True
G(Gc).X = 615                    'Setup the grenades velocity and momentums...
G(Gc).Y = MY - 80
G(Gc).SV = -10
G(Gc).V = -10
G(Gc).NB = 0
sndLau.Stop   'Stop the launch sound
sndLau.SetCurrentPosition 0 'Rewind
sndLau.Play DSBPLAY_DEFAULT 'Play the launchsound
Exit Sub
End If
Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MX = X 'Record the mouse coordinates
MY = Y
End Sub


Sub FillShrs()
'These tell the program where to find the shrapnel pieces in art.bmp.
Shr(1).X = 253
Shr(2).X = 226
Shr(3).X = 253
Shr(4).X = 272
Shr(5).X = 272

Shr(1).Y = 66
Shr(2).Y = 66
Shr(3).Y = 90
Shr(4).Y = 90
Shr(5).Y = 113

Shr(1).Wid = 21
Shr(2).Wid = 21
Shr(3).Wid = 17
Shr(4).Wid = 26
Shr(5).Wid = 22

Shr(1).Hgt = 23
Shr(2).Hgt = 22
Shr(3).Hgt = 28
Shr(4).Hgt = 20
Shr(5).Hgt = 20
End Sub
