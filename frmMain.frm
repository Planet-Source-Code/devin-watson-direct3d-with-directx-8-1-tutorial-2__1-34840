VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Direct3D w/ DirectX 8.1: Tutorial 2"
   ClientHeight    =   4260
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   331
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRender 
         Caption         =   "&Render"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Direct3D with DirectX 8.1 Tutorial 2
'Author: Devin Watson
'Original C++ code from "Special Effects Game Programming with DirectX"
'by Mason McCuskey

'If you have any questions about this, you can e-mail me
'at dwatson@erols.com

'This tutorial builds on the previous tutorial, which was simply a
'spinning triangle in a window. This tutorial introduces several
'new things, including:
'   1. Fullscreen mode
'   2. Rendering text to the screen using a font in DirectX
'   3. Frame-limiting timer loop instead of
'      the generic Timer control from Tutorial 1.

'Some form-level globals
Private mVertBuff As Direct3DVertexBuffer8  'Vertex Buffer object (need this to hold our triangle)

Private Running As Boolean                  'For our game loop


'Need this to calculate size when creating the Vertex
'Buffer and also when rendering.
Private testVert As CustomVertex

'Our triangle, which is composed of
'3 custom vertices (0->2)
Private MyTriangle(2) As CustomVertex

'Our custom vertex for the FVF
Private Type CustomVertex
    X As Single
    Y As Single
    Z As Single
    Color As Long
End Type

'This is what we display on the screen to let people
'know how to get out of the program or turn rendering on or off.
Private Const DISPLAY_TEXT = "HIT CTRL-R TO TOGGLE RENDERING. CTRL-X TO EXIT"

'These are important here, as they
'are used for startup of DirectX and Direct3D
Private mDX As DirectX8                     'DirectX object (ALWAYS NEED THIS)
Private mDX3D As Direct3D8                  'Direct3D object (need this for 3D)
Private mDX3DDevice As Direct3DDevice8      'Direct3D Device object (for output to video card)
Private CanRun As Boolean                   'Flag that prevents some bad things from happening
Private Sub Cleanup()
    'This takes care of cleaning up
    'everything and returning the system
    'to normal.
    On Local Error Resume Next
    
    'Generally, it
    'doesn't hurt you to check for
    'these things on exit explicitly.
    Running = False
    'Destroys the Direct3D device,
    'relinquishing control back
    'to Windows for its window.
    If Not mDX3DDevice Is Nothing Then
        Set mDX3DDevice = Nothing
    End If
    
    'And last but not least, we
    'take out Direct3D itself.
    If Not mDX3D Is Nothing Then
        Set mDX3D = Nothing
    End If
    
    'And turn on the mouse cursor
    ShowCursor True
End Sub

Private Sub InitD3D(ByVal hWndSurface As Long, FullScreen As Boolean, Lighting As Boolean, Culling As Boolean)
    'Initializes Direct3D and
    'gathers some information. If
    'you wanted to be paranoid, you
    'could check for video card features
    'here and manually set things
    'according to what you need.
    
    On Local Error Resume Next
    
    Dim DisplayMode As D3DDISPLAYMODE
    Dim d3dpp As D3DPRESENT_PARAMETERS
    Dim tmpFont As IFont
    
    'First, we create a Direct3D object.
    'This provides interfaces to all
    'of our other objects we can use
    'to generate our scene.
    Set mDX3D = mDX.Direct3DCreate
    
    'Since there is no such thing as the FAILED() macro
    'like in the C++ library, we have to be a little
    'more verbose in our error checking.
    If mDX3D Is Nothing Then
        'When you move to class modules (and more
        'importantly, to ActiveX DLLs), it is
        'generally a better idea to use the Err.Raise()
        'method, instead of flashing a message box.
        'This makes it easier for other programmers
        'to catch your errors.
        Err.Raise 10000, "Engine.InitD3D()", "Could not create base Direct3D system."
        CanRun = False
        Exit Sub
    End If
    
    
    'Get the default video card device information
    mDX3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DisplayMode

    'Now, Windowed's value is
    'passed in. This allows us
    'for more flexibility outside the
    'class module, while retaining
    'the same underlying feature.
    d3dpp.BackBufferHeight = DisplayMode.Height
    d3dpp.BackBufferWidth = DisplayMode.Width
    d3dpp.FullScreen_PresentationInterval = D3DPRESENT_INTERVAL_DEFAULT
    d3dpp.FullScreen_RefreshRateInHz = D3DPRESENT_RATE_DEFAULT
    'd3dpp.flags = 0
    d3dpp.Windowed = Not FullScreen

    'I set this just to make sure that
    'even if I muck up Present(), DirectX
    'has something to fall back on.
    d3dpp.hDeviceWindow = Me.hWnd
    
    'We're not really going to do anything
    'advanced, so let's just just get rid of
    'what we swap out of the back buffer.
    d3dpp.SwapEffect = D3DSWAPEFFECT_DISCARD
    
    'And make sure the back buffer pixel format is compatible
    'with the video card's screen pixel format.
    d3dpp.BackBufferFormat = DisplayMode.Format
    'd3dpp.BackBufferCount = 1
    'Now, let's create our Direct3D device, now that
    'we've got all of the information!
    Set mDX3DDevice = mDX3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, _
            D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    
    'If it failed, we have to exit the subroutine, and set a flag,
    'so that future processing can decide if it needs
    'to run or not.
    If mDX3DDevice Is Nothing Then
        Err.Raise 10001, "Engine.InitD3D()", "Could not create Direct3D Device from DirectX 8!"
        CanRun = False
        Exit Sub
    End If
    

    'Turn on/off culling.
    'No culling lets us see both
    'the front and back of an object.
    If Culling Then
        'We'll keep with the idea that
        'since DirectX renders vertices
        'in clockwise motion, then
        'we'll also turn on culling for clockwise
        'vertex processing.
        mDX3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
    Else
        mDX3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
    End If
    
    
    'Again, this option is passed into the
    'method, instead of being hard-coded. We
    'can now turn lighting on and off. We'll need
    'this later anyway.
    mDX3DDevice.SetRenderState D3DRS_LIGHTING, Lighting
    
    
    'Make sure our helper utility class is ready to go.
    'We'll need it to make our on-screen font.
    Set D3DUtil = New D3DX8
    
    'tmpFont is an IFont, which is part of OLE
    'Automation, which is by default, set up as
    'a reference in Visual Basic Standard EXE projects.
    'It does not appear directly in declaration lists,
    'but it is there. Trust me. :)
    
    'I just set tmpFont equal to the form's Font because it
    'can be a little faster just to set the values from the GUI.
    'You can also create a StdFont object, and use it
    'to set properties, then you just set it equal to an IFont.
    Set tmpFont = Me.Font
    
    'Now, we can set up the D3DFont class using the
    'helper utility class. This will put it in a format
    'that DirectX will understand.
    Set RenderFont = D3DUtil.CreateFont(mDX3DDevice, tmpFont.hFont)
    
    'And since we don't need tmpFont anymore, we free
    'up the memory for it.
    Set tmpFont = Nothing
    If Err Then
        MsgBox "Error Making Font: " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Running = False
    End If

    'This tells DirectX where on the
    'screen we can render our text.
    TextRect.Top = 0
    TextRect.Left = 0
    TextRect.bottom = 20
    TextRect.Right = 300
    
    'Turning off cursor for this
    'fullscreen application.
    ShowCursor False
    CanRun = True
End Sub






Public Sub InitDX()
    'Starts up DirectX
    CanRun = True
    Set mDX = New DirectX8
    If Err Then
        MsgBox "Error starting DirectX: " & Err.Description, vbOKOnly + vbCritical, "Error"
        CanRun = False
    End If
End Sub
Private Sub InitGeometry()
    'Creates the vertex buffer we
    'will be displaying. Basically,
    'the easiest way to communicate
    'with DirectX as to what
    'you want it to display is
    'to talk in terms of triangles, otherwise
    'known as an array or collection of
    '3 vertices (3-D coordinates)
    
    On Local Error Resume Next
    'If InitD3D failed, we can just exit now.
    If CanRun = False Then Exit Sub

    Dim RC As Long
    
    'NOTE: I am using the helper function D3DColorRGBA
    'to produce color values for each vertex, but,
    'you could easily use the regular RGB() function
    'built into VB. I don't because I like to use
    'the "DX-native functions". If this doesn't work
    'on your video card, try using D3DColorXRGB or
    'D3DColorARGB.
    
    'Also, try changing the Z component
    'to a different value for some interesting warping
    'during the rotation.
    
    
    'First vertex: Lower left-hand corner
    MyTriangle(0).X = -1
    MyTriangle(0).Y = -1
    MyTriangle(0).Z = 0
    MyTriangle(0).Color = D3DColorRGBA(0, 255, 0, 0)
    
    'Second vertex: Lower right-hand corner
    MyTriangle(1).X = 1
    MyTriangle(1).Y = -1
    MyTriangle(1).Z = 0
    MyTriangle(1).Color = D3DColorRGBA(255, 0, 255, 0)
    
    'Third vertex: Top
    MyTriangle(2).X = 0
    MyTriangle(2).Y = 1
    MyTriangle(2).Z = 0
    MyTriangle(2).Color = D3DColorRGBA(255, 255, 255, 0)
    
    'Now that we've defined our vertex, we need to
    'set up the vertex processing pipeline
    'to accept this custom format, using the
    'Flexible Vertex Format (FVF)
    Set mVertBuff = mDX3DDevice.CreateVertexBuffer(3 * Len(testVert), 0, D3DFVF_CUSTOMVERTEX, D3DPOOL_DEFAULT)
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error creating Vertex Buffer: Invalid Call", vbOKOnly + vbCritical, "ERROR"
        Exit Sub
    End If
    
    'Well, now that we've created the vertex buffer,
    'based on our own custom FVF, we
    'need to fill it.
    'We can use this helper function,
    'provided by Direct3D, to lock,
    'fill, and unlock the vertex
    'buffer all in one line. Neat, eh?
    RC = D3DVertexBuffer8SetData(mVertBuff, 0, 3 * LenB(testVert), 0, MyTriangle(0))
    
    'Since this function does not set Err.Number,
    'we need to check against this known constant
    'to make sure it executed correctly.
    If RC = D3DERR_INVALIDCALL Then
        MsgBox "Invalid call to D3DVertextBuffer8SetData()!", vbOKOnly + vbCritical, "ERROR"
        Exit Sub
    End If
    
End Sub

Private Sub Render()
    'The main render routine.
    'This is now a framelimiter
    'loop, which ensures a constant
    'render rate at 10 milliseconds,
    'which is better precision than
    'the standard Timer control used
    'in Tutorial 1.
    On Local Error Resume Next
    Dim CurTime As Long
    Dim PrevTime As Long
    Dim DestRect As RECT
    
    If CanRun = False Then
        MsgBox "Cannot run: Failure in CanRun flag!", vbOKOnly + vbCritical, "ERROR"
        Exit Sub
    End If
 
    DestRect.Top = 0
    DestRect.Left = 0
    DestRect.bottom = 20
    DestRect.Right = Me.ScaleWidth
    
    Do
        'First, we check to see if we should render at all
        If Not Running Then Exit Sub
        
        CurTime = GetTickCount
        If CurTime - PrevTime > TIMERRATE Then
        
            'Clear the back buffer to a light blue
            'color.
            mDX3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorXRGB(0, 0, 128), 1, 0
        
            If Err.Number = D3DERR_INVALIDCALL Then
                Running = False
                Exit Sub
            End If
            
            'Begin rendering the scene.
            mDX3DDevice.BeginScene
            
            SetupMatrices
    
            'Render the vertex buffer contents
            mDX3DDevice.SetStreamSource 0, mVertBuff, LenB(testVert)
            mDX3DDevice.SetVertexShader D3DFVF_CUSTOMVERTEX
            mDX3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, 1
            
            'The D3DFont we created back in InitD3D.
            'The Begin and End functions act just like
            'BeginScene and EndScene, where the memory is
            'locked so that the text is displayed properly.
            
            'NOTE: Begin and End blocks MUST be completely inside
            '      a BeginScene and EndScene block! This is important,
            '      otherwise you may end up with anomalies on the
            '      screen, strange artifacts, or the application
            '      simply won't work.
            RenderFont.Begin
            
            RenderFont.DrawTextW DISPLAY_TEXT, Len(DISPLAY_TEXT), TextRect, DT_LEFT Or DT_NOCLIP, D3DColorARGB(255, 255, 255, 255)
            RenderFont.End
            
            mDX3DDevice.EndScene
            'This is the equivalent to a Blt()
            'operation in the good ol' days of
            'DX 7.0
            mDX3DDevice.Present ByVal 0, ByVal 0, ByVal 0, ByVal 0
            'If we don't set this, then
            'funny things can happen.
            PrevTime = CurTime
            DoEvents
        Else
            DoEvents
            Sleep 2
        End If
    Loop
End Sub

Private Sub SetupMatrices()
    'This sets up the World, View,
    'and Projection transform matrices.
    On Local Error Resume Next
    
    'For the World, we'll rotate along
    'the Y-axis. We're using timeGetTime()
    'from the Win32 API to derive an
    'arbitrary angle from which to rotate to.
    Dim matWorld As D3DMATRIX
    Dim matView As D3DMATRIX
    Dim matProj As D3DMATRIX
    
    'These vectors are needed to calculate a view,
    'as they show where we are, what we
    'are looking at, and which way is up.
    Dim vecEye As D3DVECTOR
    Dim vecAt As D3DVECTOR
    Dim vecUp As D3DVECTOR
    
    D3DXMatrixRotationY matWorld, (timeGetTime / 150#)
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error rotating World Matrix: " & Err.Description, vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    'Now that we've got it rotated,
    'we apply the transformation to the World
    mDX3DDevice.SetTransform D3DTS_WORLD, matWorld
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error setting transform on world: Invalid procedure call.", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    'Now we set up the View Matrix. This one
    'is a little trickier. First, we need
    'to set the 3 Vectors for positioning everything.
    
    'The first one defines where our position is.
    vecEye.X = 0#
    vecEye.Y = 3#
    vecEye.Z = -5#
    
    'The second defines what we are looking at
    'in the 3D World. In this case, it is (0,0,0),
    'or, the origin of the entire 3D World.
    vecAt.X = 0#
    vecAt.Y = 0#
    vecAt.Z = 0#
    
    'The third vector is our normal, which tells
    'us which way is up.
    vecUp.X = 0#
    vecUp.Y = 1#
    vecUp.Z = 0#
    
    'And we make the View Matrix!
    D3DXMatrixLookAtLH matView, vecEye, vecAt, vecUp
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error calling D3DXMatrixLookAtLH()!", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    
   mDX3DDevice.SetTransform D3DTS_VIEW, matView
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error setting transform on view: Invalid procedure call.", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    'Whew! Now for the Projection Matrix. This one
    'isn't nearly as rough.
    D3DXMatrixPerspectiveFovLH matProj, PI / 4, 1#, 1#, 100#
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error calling D3DXMatrixPerspectiveFovLH()!", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    
    
    mDX3DDevice.SetTransform D3DTS_PROJECTION, matProj
    If Err.Number = D3DERR_INVALIDCALL Then
        MsgBox "Error setting transform on projection: Invalid procedure call.", vbOKOnly + vbCritical, "ERROR"
        CanRun = False
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()
    InitDX
    InitD3D Me.hWnd, True, False, False
    If Err Then
        MsgBox "Error: " & Err.Source & ": " & Err.Description, vbOKOnly + vbCritical, "OOPS!"
    End If
    
    InitGeometry
    Me.Visible = True
    Me.ZOrder 0
    Running = False
    mnuRender_Click
End Sub



Private Sub Form_Unload(Cancel As Integer)
    'We're exiting, so we should
    'gracefully leave by cleaning
    'up memory.
    'Set myEngine = Nothing
    Cleanup
    End
End Sub


Private Sub mnuExit_Click()
    'Calls Form_Unload()
    Unload Me
End Sub


Private Sub mnuRender_Click()
    'Okay, now this just turns
    'the Boolean switch for
    'rendering on and off.
    'I've put this now into a
    'framelocked loop, instead of
    'using a Timer control, which
    'doesn't have the kind of
    'precision we need for a
    'good 3D application or game.
    mnuRender.Checked = Not mnuRender.Checked
    Running = mnuRender.Checked
    If Running Then
        Render
    End If
End Sub


