VERSION 5.00
Begin VB.Form Stage 
   BackColor       =   &H00800000&
   Caption         =   "Stencil Buffer App"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   330
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   StartUpPosition =   3  'Windows Default
   Tag             =   "101"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4800
      Top             =   4320
   End
End
Attribute VB_Name = "Stage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Author        : Jan Tosovsky
'Email         : j.tosovsky@tiscali.cz
'Website       : http://nio.astronomy.cz
'Date          : 13 September 2003
'Version       : 1.0
'Description   : Uses the stencil buffer to clip drawing area of images

'This code was converted from Delphi source http://www.sulaco.co.za/

'Program requires OpenGL Type Library from Patrice Scribe
'http://is6.pacific.net.hk/~edx/tlb.htm
'With library you needn't declare any used OpenGL functions or constants
'Copy library to system directory and then register it:
'regsvr32 "C:\Windows\System\vbogl.tlb" where path may vary
'In Project>References... in VB menu check item VB OpenGL API 1.2 (ANSI)

Option Explicit

Dim FPSCount As Single, DemoStart As Long, ElapsedTime As Long
Dim xcoord As Long, ycoord As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then End
End Sub

Private Sub Form_Load()
    EnableOpenGL Stage.hDC
    DrawInit
End Sub

Private Sub DrawInit()
    glClearColor 0, 0, 0, 0
    glShadeModel GL_SMOOTH
    glClearDepth 1#
    glEnable GL_DEPTH_TEST
    glDepthFunc GL_LESS
    glEnable GL_STENCIL_TEST
    glHint GL_PERSPECTIVE_CORRECTION_HINT, GL_NICEST
    createKnot 16, 64, 2, 4#, 1#
    'createKnot 16, 64, 3, 3#, 1#
    'createKnot 16, 64, 4, 2#, 1#
End Sub

Sub MainLoop()
    Dim LastTime As Long
    DemoStart = GetTickCount
    Do
        LastTime = ElapsedTime
        ElapsedTime = GetTickCount() - DemoStart
        ElapsedTime = (LastTime + ElapsedTime) \ 2
        DoEvents
        Render
        FPSCount = FPSCount + 1
    Loop
End Sub

Sub Render()
    
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT Or GL_STENCIL_BUFFER_BIT
    glLoadIdentity
    glTranslatef 0#, 0#, -20
    glPolygonMode GL_FRONT, GL_FILL
    glPolygonMode GL_BACK, GL_FILL
    glStencilFunc GL_ALWAYS, 1, 1
    glStencilOp GL_REPLACE, GL_REPLACE, GL_REPLACE
    glBegin GL_QUADS
        glVertex3f -1# + xcoord / 100, -1# - ycoord / 100, 13#
        glVertex3f 1# + xcoord / 100, -1# - ycoord / 100, 13#
        glVertex3f 1# + xcoord / 100, 1# - ycoord / 100, 13#
        glVertex3f -1# + xcoord / 100, 1# - ycoord / 100, 13#
    glEnd
    
    ' clear buffers and draw screen and draw only in that square area
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    glStencilFunc GL_EQUAL, 1, 1
    glStencilOp GL_KEEP, GL_KEEP, GL_KEEP
    
    ' set wireframe mode
    glPolygonMode GL_FRONT, GL_LINE
    glPolygonMode GL_BACK, GL_LINE
    
    'draw a square border sothat you can see where the square is
    glLineWidth 2
    glColor3f 1, 1, 1
    glBegin GL_QUADS ' draw the square borders
        glVertex3f -1# + xcoord / 100, -1# - ycoord / 100, 13#
        glVertex3f 1# + xcoord / 100, -1# - ycoord / 100, 13#
        glVertex3f 1# + xcoord / 100, 1# - ycoord / 100, 13#
        glVertex3f -1# + xcoord / 100, 1# - ycoord / 100, 13#
    glEnd
    glLineWidth 1
    
    'Rotate and draw the wireframe knot
    glRotatef ElapsedTime / 10, 1, 0, 0
    glRotatef ElapsedTime / 8, 0, 1, 0
    
    glCallList KnotDL
    
    'now draw in the area that is not in the square (Not equal to sqaure area)
    glStencilFunc GL_NOTEQUAL, 1, 1
    glStencilOp GL_KEEP, GL_KEEP, GL_KEEP
    glPolygonMode GL_FRONT, GL_FILL
    glPolygonMode GL_BACK, GL_FILL
    
    glCallList KnotDL
    
    DoEvents
    SwapBuffers Stage.hDC
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xcoord = X - Me.ScaleWidth / 2
    ycoord = Y - Me.ScaleHeight / 2
End Sub

Private Sub Form_Resize()
    Dim w As Long, h As Long
    w = Me.ScaleWidth: h = Me.ScaleHeight
    If h = 0 Then h = 1
    glViewport 0, 0, w, h
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    gluPerspective 45#, w / h, 1#, 100#
    glMatrixMode GL_MODELVIEW
    glLoadIdentity
    DoEvents
    MainLoop
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DisableOpenGL
End Sub

Private Sub Timer1_Timer()
    DoEvents
    Me.Caption = "Stencil Buffer App [" & FPSCount & " FPS]"
    FPSCount = 0
End Sub
