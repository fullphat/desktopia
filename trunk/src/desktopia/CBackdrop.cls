VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CBackdrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

    '/*********************************************************************************************
    '/
    '/  File:           CBackDrop.cls
    '/
    '/  Description:    Manages the backdrop window
    '/
    '/  Created:
    '/
    '/  © 2011 k23 productions
    '/
    '/  This file may be used under the terms of the Simplified BSD Licence
    '/
    '*********************************************************************************************/

    ' /* Notes:
    '
    '
    ' */

Private Const CLASS_NAME = "w>desktopia"

Dim mhWnd As Long
Dim mView As mfxView

Dim mConfig As CConfFile3

Implements BWndProcSink

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static pps As PAINTSTRUCT
Static hDC As Long
Static pwp As WINDOWPOS
Dim pmi As OMMenuItem

    Select Case uMsg
    Case WM_WINDOWPOSCHANGING
        CopyMemory pwp, ByVal lParam, Len(pwp)
        pwp.hWndInsertAfter = HWND_BOTTOM
        CopyMemory ByVal lParam, pwp, Len(pwp)
        ReturnValue = 0
        BWndProcSink_WndProc = True


    Case WM_RBUTTONDOWN
        With New OMMenu
            .AddItem .CreateItem("cfg", "Settings...")
            .AddSeparator
            .AddItem .CreateItem("bye", "Close")

            Set pmi = .Track(hWnd)
            If NOTNULL(pmi) Then
                Select Case pmi.Name
                Case "cfg"
                    frmAbout.Show
                
                Case "bye"
                    PostQuitMessage 0

                End Select

            End If

        End With

    Case WM_ERASEBKGND
        ReturnValue = -1
        BWndProcSink_WndProc = True

    Case WM_PAINT
        hDC = BeginPaint(hWnd, pps)
        If NOTNULL(mView) Then _
            draw_view mView, hDC

        EndPaint hWnd, pps
        ReturnValue = 0
        BWndProcSink_WndProc = True

    End Select

End Function

Private Sub uRedraw(ByRef Config As CConfSection)
Dim pb As mfxBitmap
Dim dw As Long
Dim sz As String

'    MsgBox Command$

    If Command$ <> "" Then
        Set pb = load_image_obj(g_RemoveQuotes(Command$))
        If NOTNULL(pb) Then
            uDrawToFit mView, pb
            Exit Sub

        End If
    End If

    sz = Config.GetValueWithDefault("type", "solid")
    dw = g_RGBAFromString(Config.GetValueWithDefault("colour", "#2a2a2a"))

    With mView
        Select Case sz
        Case "solid"
            .SetHighColour dw
            .FillRect .Bounds
            
        Case "image"
            sz = Config.GetValueWithDefault("image", g_MakePath(App.Path) & "sample.jpg")
            Set pb = load_image_obj(sz)
            uDrawToFit mView, pb

        Case Else
            .SetHighColour rgba(0, 0, 0)
            .FillRect .Bounds
            .SetHighColour rgba(255, 0, 0)
            .StrokeLine new_BRect(0, 100, .Bounds.Right, 100)

        End Select

'Dim x As Long
'
'        For x = 0 To .Width Step 8
'            .StrokeLine new_BRect(x, 0, x, .Height - 1)
'
'        Next x

    End With

'    apply_view_to_window mView, mhWnd

End Sub

Private Sub Class_Initialize()
Dim ps As CConfSection

    Set mConfig = New CConfFile3
    With mConfig
        .SetFile g_MakePath(App.Path) & "desktopia.conf"
        .Load

        If Not .Find("general", ps) Then
            Set ps = New CConfSection
            ps.SetName "general"
            .Add ps
            .Save

        End If

    End With

    EZRegisterClass CLASS_NAME
    mhWnd = EZ4AddWindow(CLASS_NAME, Me, , WS_POPUPWINDOW, WS_EX_TOOLWINDOW)

Dim rc As RECT

    g_GetWorkArea rc

    Set mView = New mfxView
    With mView
        .SizeTo g_ScreenWidth, g_ScreenHeight '(rc.Right - rc.Left), (rc.Bottom - rc.Top)
        .EnableSmoothing False

    End With

    uRedraw ps

    g_SizeWindow mhWnd, mView.Width, mView.Height
    g_MoveWindow rc.Left, rc.Top
    g_ShowWindow mhWnd, True, False

End Sub

Private Sub Class_Terminate()

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

End Sub

Private Sub uDrawToFit(ByRef aView As mfxView, ByRef aPic As mfxBitmap)
Dim pv As mfxView
Dim d As Double
Dim p As BPoint

    With aView
        .SetHighColour rgba(0, 0, 0)
        .FillRect aView.Bounds

    End With

'    If (aPic.Width > aView.Width) Or (aPic.Height > aView.Height) Then
        ' /* shrink */
        If (aPic.Width / aPic.Height) > (aView.Width / aView.Height) Then
            d = aView.Height / aPic.Height

        Else
            d = aView.Width / aPic.Width

        End If

        Set p = new_BPoint(aPic.Width * d, aPic.Height * d)
        aView.DrawScaledImage aPic, new_BPoint(Fix((aView.Width - p.x) / 2), Fix((aView.Height - p.y) / 2)), p

'    Else
'        aView.DrawScaledImage aPic, new_BPoint(Fix((aView.Width - aPic.Width) / 2), Fix((aView.Height - aPic.Height) / 2))
'
'    End If

End Sub

' // obscure taskbar
' // pick random picture from folder
' // letterbox colour
' // watermark
' //




