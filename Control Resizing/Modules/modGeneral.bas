Attribute VB_Name = "modGeneral"
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_DRAWFRAME = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_FLAGS = SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
Private Const GWL_STYLE = (-16)
Private Const WS_THICKFRAME = &H40000

Dim initBoxStyle As Long


Public Sub InitialState(frmSource As Form, objSource As Object)
    initBoxStyle = GetWindowLong(objSource.hwnd, GWL_STYLE)
    
    SetControlStyle frmSource, objSource, initBoxStyle
End Sub

Public Sub CanResize(frmSource As Form, objSource As Object, blnState As Boolean)
    Dim style As Long
    
    Select Case blnState
        Case True
            style = GetWindowLong(objSource.hwnd, GWL_STYLE)
            style = style Or WS_THICKFRAME
            SetControlStyle frmSource, objSource, style
        Case False
            SetControlStyle frmSource, objSource, initBoxStyle
    End Select
End Sub

Private Sub SetControlStyle(frmSource As Form, objSource As Control, style)
    Dim r
    
    If style Then
        Call SetWindowLong(objSource.hwnd, GWL_STYLE, style)
        Call SetWindowPos(objSource.hwnd, frmSource.hwnd, 0, 0, 0, 0, SWP_FLAGS)
    End If
End Sub
