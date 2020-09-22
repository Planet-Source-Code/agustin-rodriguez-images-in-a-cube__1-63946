VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   7995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   ControlBox      =   0   'False
   Icon            =   "Cube.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   533
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   533
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function SetCapture Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Pt As POINTAPI
Private xx As Long
Private yy As Long
Private capture As Integer
Public Img1 As cTexture
Public Img2 As cTexture
Public Img3 As cTexture
Public Img4 As cTexture
Public Img5 As cTexture
Public Img6 As cTexture

Public Function DrawGLScene() As Boolean

  Static xrot As GLfloat
  Static yrot As GLfloat
  Static zrot As GLfloat

    ' Clear the Img1buffer and the depth buffer
    glClear clrColorBufferBit Or clrDepthBufferBit
    ' Reset the modelview matrix
    glLoadIdentity
    
    glTranslatef 0#, 0#, gflZ
    ' Rotate the scene along the x and y axis
    glRotatef xrot, 1#, 0#, 0#
    glRotatef yrot, 0#, 1#, 0#
        
    ' Build a cube using quads
    Img1.useTexture
    
    glBegin GL_QUADS
    ' Img2 Face
    glNormal3f 0#, 0#, 1#                              'Set the surface normal
    glTexCoord2f 0#, 0#: glVertex3f -1#, -1#, 1#       'Bottom Left Of The Texture and Quad
    glTexCoord2f 1#, 0#: glVertex3f 1#, -1#, 1#        'Bottom Right Of The Texture and Quad
    glTexCoord2f 1#, 1#: glVertex3f 1#, 1#, 1#         'Top Right Of The Texture and Quad
    glTexCoord2f 0#, 1#: glVertex3f -1#, 1#, 1#        'Top Left Of The Texture and Quad
    
    glEnd

    Img2.useTexture
    
    glBegin GL_QUADS
    ' Img1 Face
    glNormal3f 0#, 0#, -1#                             'Set the surface normal
    glTexCoord2f 1#, 0#: glVertex3f -1#, -1#, -1#      'Bottom Right Of The Texture and Quad
    glTexCoord2f 1#, 1#: glVertex3f -1#, 1#, -1#       'Top Right Of The Texture and Quad
    glTexCoord2f 0#, 1#: glVertex3f 1#, 1#, -1#        'Top Left Of The Texture and Quad
    glTexCoord2f 0#, 0#: glVertex3f 1#, -1#, -1#       'Bottom Left Of The Texture and Quad
    
    glEnd
    
    Img3.useTexture
    
    glBegin GL_QUADS
    ' Top Face
    glNormal3f 0#, 1#, 0#                              'Set the surface normal
    glTexCoord2f 0#, 1#: glVertex3f -1#, 1#, -1#       'Top Left Of The Texture and Quad
    glTexCoord2f 0#, 0#: glVertex3f -1#, 1#, 1#        'Bottom Left Of The Texture and Quad
    glTexCoord2f 1#, 0#: glVertex3f 1#, 1#, 1#         'Bottom Right Of The Texture and Quad
    glTexCoord2f 1#, 1#: glVertex3f 1#, 1#, -1#        'Top Right Of The Texture and Quad
    
    glEnd

    Img4.useTexture
    
    glBegin GL_QUADS
    
    ' Bottom Face
    glNormal3f 0#, -1#, 0#                             'Set the surface normal
    glTexCoord2f 1#, 1#: glVertex3f -1#, -1#, -1#      'Top Right Of The Texture and Quad
    glTexCoord2f 0#, 1#: glVertex3f 1#, -1#, -1#       'Top Left Of The Texture and Quad
    glTexCoord2f 0#, 0#: glVertex3f 1#, -1#, 1#        'Bottom Left Of The Texture and Quad
    glTexCoord2f 1#, 0#: glVertex3f -1#, -1#, 1#       'Bottom Right Of The Texture and Quad
    
    glEnd

    Img5.useTexture
    
    glBegin GL_QUADS
    ' Right face
    glNormal3f 1#, 0#, 0#                              'Set the surface normal
    glTexCoord2f 1#, 0#: glVertex3f 1#, -1#, -1#       'Bottom Right Of The Texture and Quad
    glTexCoord2f 1#, 1#: glVertex3f 1#, 1#, -1#        'Top Right Of The Texture and Quad
    glTexCoord2f 0#, 1#: glVertex3f 1#, 1#, 1#         'Top Left Of The Texture and Quad
    glTexCoord2f 0#, 0#: glVertex3f 1#, -1#, 1#        'Bottom Left Of The Texture and Quad
    
    glEnd
    
    Img6.useTexture
    
    glBegin GL_QUADS
        
    ' Left Face
    glNormal3f -1#, 0#, 0#                             'Set the surface normal
    glTexCoord2f 0#, 0#: glVertex3f -1#, -1#, -1#      'Bottom Left Of The Texture and Quad
    glTexCoord2f 1#, 0#: glVertex3f -1#, -1#, 1#       'Bottom Right Of The Texture and Quad
    glTexCoord2f 1#, 1#: glVertex3f -1#, 1#, 1#        'Top Right Of The Texture and Quad
    glTexCoord2f 0#, 1#: glVertex3f -1#, 1#, -1#       'Top Left Of The Texture and Quad
    glEnd
    
    xrot = xrot + gflXSpeed
    yrot = yrot + gflYSpeed
    
    DrawGLScene = True

End Function

Public Function InitGL() As Boolean

  Dim aflLightAmbient(4) As GLfloat
  Dim aflLightDiffuse(4) As GLfloat
  Dim aflLightPosition(4) As GLfloat
    
    ' Create new texture
    Set Img1 = New cTexture
    Set Img2 = New cTexture
    Set Img3 = New cTexture
    Set Img4 = New cTexture
    Set Img5 = New cTexture
    Set Img6 = New cTexture
    
    '    Img1.loadTexture App.Path & "\Data\Crate.tga", FILETYPE_TGA
    Img1.loadTexture App.Path & "\Side 1.bmp", 0 'FILETYPE_TGA
    Img2.loadTexture App.Path & "\Side 2.bmp", 0 'FILETYPE_TGA
    Img3.loadTexture App.Path & "\Side 3.bmp", 0 'FILETYPE_TGA
    Img4.loadTexture App.Path & "\Side 4.bmp", 0 'FILETYPE_TGA
    Img5.loadTexture App.Path & "\Side 5.bmp", 0 'FILETYPE_TGA
    Img6.loadTexture App.Path & "\Side 6.bmp", 0 'FILETYPE_TGA
 
    ' Enable texture mapping
    glEnable glcTexture2D
    ' Smooth shading
    glShadeModel smSmooth
    
    ' Set the clear colour
    glClearColor 0#, 0#, 0#, 0#
    ' Set the clear depth
    glClearDepth 1#
    
    ' Enable Z-buffer
    glEnable glcDepthTest
    ' Set test type
    glDepthFunc cfLEqual
    ' Best perspective correction
    glHint htPerspectiveCorrectionHint, hmNicest
      
    ' Ambient light settings
    aflLightAmbient(0) = 0.5
    aflLightAmbient(1) = 0.5
    aflLightAmbient(2) = 0.5
    aflLightAmbient(3) = 1#
    ' Diffuse light settings
    aflLightDiffuse(0) = 1#
    aflLightDiffuse(1) = 1#
    aflLightDiffuse(2) = 1#
    aflLightDiffuse(3) = 1#
    ' Light position settings
    aflLightPosition(0) = 0#
    aflLightPosition(1) = 0#
    aflLightPosition(2) = 2#
    aflLightPosition(3) = 1#
      
    ' Set the light's ambient and diffuse levels and its position
    glLightfv ltLight1, lpmAmbient, aflLightAmbient(0)
    glLightfv ltLight1, lpmDiffuse, aflLightDiffuse(0)
    glLightfv ltLight1, lpmPosition, aflLightPosition(0)
    
    ' Enable light1
    glEnable glcLight1
    
    InitGL = True

End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  ' Set the key to be pressed

    gbKeys(KeyCode) = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

  ' Set the key to be not pressed

    gbKeys(KeyCode) = False

End Sub

Private Sub Form_Load()

  Dim bFullscreen As Boolean
  Dim frm As frmMain
  Dim bLightSwitched As Boolean
  Dim bFilterSwitched As Boolean
  Dim bLightOn As Boolean
  Dim giCurrFilter As Integer
  Dim ret As Long

    Erase gbKeys

    gflXSpeed = 0.05
    gflYSpeed = 0.05

    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    
    SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_COLORKEY Or LWA_ALPHA

    ' Put us into fullscreen automatically
    'bFullscreen = True
    bLightSwitched = False
    bFilterSwitched = False
    bLightOn = False
    gflZ = -7#

    ' Save the current display settings
    SaveDisplaySettings

    ' Show this form
    Me.Show
    ' Attempt to create a compatible GL window and set the display mode
    If (CreateGLWindow(Me, 640, 480, 32, bFullscreen) = False) Then
        Unload Me
    End If
    ' Attempt to set up OpenGL
    If (InitGL() = False) Then
        Unload Me
    End If
  
    ' Loop until the form is unloaded, process windows events every time we're not rendering
    Do While DoEvents()
        ' Render the scene, if it failed or the user has pressed the escape key then exit the program
        If (DrawGLScene() = False) Or (gbKeys(vbKeyEscape)) Then
            Exit Do '>---> Loop
          Else 'NOT (DRAWGLSCENE()...
            ' Swap the Img2 and Img1 buffers to display what we've just rendered
            SwapBuffers Me.hDC
      
            ' Toggle lighting
            If (gbKeys(vbKeyL)) And (bLightSwitched = False) Then
                bLightOn = Not (bLightOn)
                If (bLightOn) Then
                    glEnable glcLighting
                  Else '(BLIGHTON) = 0
                    glDisable glcLighting
                End If
              
                bLightSwitched = True
            End If
      
            If (gbKeys(vbKeyL) = False) Then
                bLightSwitched = False
            End If
      
            ' Toggle filtering
            If (gbKeys(vbKeyF)) And (bFilterSwitched = False) Then
                giCurrFilter = Img1.getFilter
                giCurrFilter = giCurrFilter + 1
                If giCurrFilter > 2 Then
                    giCurrFilter = 0
                End If
                    
                Select Case giCurrFilter
                  Case 0:
                    Img1.setFilter FILTER_NEAREST
                  Case 1:
                    Img1.setFilter FILTER_LINEAR
                  Case 2:
                    Img1.setFilter FILTER_MIPMAPPED
                End Select
            
                bFilterSwitched = True
            End If
      
            If (gbKeys(vbKeyF) = False) Then
                bFilterSwitched = False
            End If
        
            ' Zoom in and out
            If (gbKeys(vbKeyPageUp)) Then
                gflZ = gflZ - 0.02
                
            End If
            
            If (gbKeys(vbKeyPageDown)) Then
                gflZ = gflZ + 0.02
                If gflZ > -4.44 Then
                    gflZ = -4.44
                End If
            End If
            
            ' Increase / decrease cube's rotation amount
            If gbKeys(vbKeyUp) Then
                gflXSpeed = gflXSpeed - 0.01
            End If
            
            If gbKeys(vbKeyDown) Then
                gflXSpeed = gflXSpeed + 0.01
            End If
            
            If gbKeys(vbKeyLeft) Then
                gflYSpeed = gflYSpeed - 0.01
            End If
            
            If gbKeys(vbKeyRight) Then
                gflYSpeed = gflYSpeed + 0.01
            End If
            
            ' Key escape has been pressed, so exit the program!
            If gbKeys(vbKeyEscape) Then
                Exit Do '>---> Loop
            End If
        End If
        DoEvents
    Loop
    
    Terminou = True
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        
        xx = X * Screen.TwipsPerPixelX: yy = Y * Screen.TwipsPerPixelY
        capture = True
        ReleaseCapture
        SetCapture Me.hwnd
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If capture Then
        GetCursorPos Pt
        Move Pt.X * Screen.TwipsPerPixelX - xx, Pt.Y * Screen.TwipsPerPixelY - yy
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    capture = False

End Sub

Private Sub Form_Resize()

  ' When the user resizes the form, tell OpenGL to update so that it renders to the right place!
  ' Primarily used when in windowed mode

    ReSizeGLScene ScaleWidth, ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)

  ' Shut down OpenGL

    KillGLWindow Me
    
End Sub


