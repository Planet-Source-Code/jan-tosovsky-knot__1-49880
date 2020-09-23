Attribute VB_Name = "OpenGL"
Option Explicit
Dim hRC As Long

Sub EnableOpenGL(ghDC As Long)
Dim pfd As PIXELFORMATDESCRIPTOR, PixFormat As Long
    
    ZeroMemory pfd, Len(pfd)
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER Or PFD_GENERIC_FORMAT
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 24
    pfd.cDepthBits = 16
    pfd.cStencilBits = 1
    pfd.iLayerType = PFD_MAIN_PLANE
    
    PixFormat = ChoosePixelFormat(ghDC, pfd)
    If PixFormat = 0 Then GoTo ee
    SetPixelFormat ghDC, PixFormat, pfd
    hRC = wglCreateContext(ghDC)
    wglMakeCurrent ghDC, hRC
    
Exit Sub
ee: MsgBox "Can't create OpenGL context!", vbCritical, "Errror"
    End
End Sub

Sub DisableOpenGL()
    wglMakeCurrent 0, 0
    wglDeleteContext hRC
End Sub
