Attribute VB_Name = "Engine"
Option Explicit
'Support.bas -- Contains constants and other sundry support
'functions and subroutines.
'Author: Devin Watson

Global Const PI = 3.14159                  'For the rotation calculations
Global Const D3DADAPTER_DEFAULT = 0        'Taken from one of the C++ headers. It isn't
                                           'defined in the Type Library for VB.

'Another custom constant, which tells Direct3D
'that we're using the Flexible Vertext Format (FVF) with some
'diffuse colors so the triangle will be "self-illuminating"
'when lighting is turned off.
Global Const D3DFVF_CUSTOMVERTEX As Long = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

'We're going to use this for some arbitrary angle calculation.
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

'This is needed to construct our timing loop, instead of using
'the Timer control. This allows us a higher level of precision.
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Global Const TIMERRATE = 10         'Synchronization rate (10ms)
Global D3DUtil As D3DX8             'Just a little helper class we use all
                                    'over the place.

Global RenderFont As D3DXFont       'This is the actual font we'll be using
                                    'See InitD3D in frmMain for actual setup.

Global TextRect As RECT             'We need this to tell DirectX where
                                    'to render text on the screen.

'Sleep() is needed for the frame-limiting loop, and
'ShowCursor is used to toggle the mouse cursor on and off
'during fullscreen mode.
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
