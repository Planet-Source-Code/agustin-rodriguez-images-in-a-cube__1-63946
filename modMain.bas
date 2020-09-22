Attribute VB_Name = "modMain"
Option Explicit
Public Terminou As Integer
Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const LWA_ALPHA As Long = &H2&
Public Const LWA_COLORKEY As Integer = &H1

Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Global gbKeys(256) As Boolean       'Indicates which keys are currently pressed

Global gflXRot As GLfloat           'X rotation
Global gflYRot As GLfloat           'Y rotation
Global gflXSpeed As GLfloat         'X rotation speed
Global gflYSpeed As GLfloat         'Y rotation speed
Global gflZ As GLfloat              'Z position


