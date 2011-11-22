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

Dim mConfig As CConfFile3
Dim mGeneral As CConfSection

Dim mView As mfxView
Dim mhWnd As Long

Dim mImageToUse As String

Dim thePanel As BPrefsPanel
Attribute thePanel.VB_VarHelpID = -1

Implements BWndProcSink
Implements KPrefsPanel
Implements KPrefsPage

Private Sub Class_Initialize()

    uLoadConfig

    EZRegisterClass CLASS_NAME
    mhWnd = EZ4AddWindow(CLASS_NAME, Me, CLASS_NAME, WS_POPUPWINDOW, WS_EX_TOOLWINDOW)

Dim rc As RECT

    g_GetWorkArea rc

    Set mView = New mfxView
    With mView
        .SizeTo g_ScreenWidth, g_ScreenHeight '(rc.Right - rc.Left), (rc.Bottom - rc.Top)
        .EnableSmoothing False

    End With

    uRedraw

    g_SizeWindow mhWnd, mView.Width, mView.Height
    g_MoveWindow rc.Left, rc.Top
    g_ShowWindow mhWnd, True, False

    If g_IsIDE() Then _
        uDoPrefs

End Sub

Private Sub Class_Terminate()

    EZ4RemoveWindow mhWnd
    EZUnregisterClass CLASS_NAME

    If NOTNULL(thePanel) Then _
        thePanel.Quit

End Sub

Private Function BWndProcSink_WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal PrevWndProc As Long, ReturnValue As Long) As Boolean
Static pps As PAINTSTRUCT
Static hDC As Long
Static pwp As WINDOWPOS
Dim pmi As OMMenuItem

    Select Case uMsg

    Case WM_CLOSE
        PostQuitMessage 0
        ReturnValue = 0
        BWndProcSink_WndProc = True


    Case WM_WINDOWPOSCHANGING
        CopyMemory pwp, ByVal lParam, Len(pwp)
        pwp.hWndInsertAfter = HWND_BOTTOM
        CopyMemory ByVal lParam, pwp, Len(pwp)
        ReturnValue = 0
        BWndProcSink_WndProc = True


    Case WM_RBUTTONDOWN
        With New OMMenu
'            .AddItem .CreateItem("", "Desktopia", , False)
'            .AddSeparator
            .AddItem .CreateItem("cfg", "Settings...")
            .AddSeparator
            .AddItem .CreateItem("bye", "Close Desktopia")

            Set pmi = .Track(hWnd)
            If NOTNULL(pmi) Then
                Select Case pmi.Name
                Case "cfg"
                    uDoPrefs

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

Private Sub uRedraw()
Dim pb As mfxBitmap
Dim dw As Long
Dim sz As String

    If Command$ <> "" Then
        Set pb = load_image_obj(g_RemoveQuotes(Command$))
        If NOTNULL(pb) Then
            uDrawToFit mView, pb
            Exit Sub

        End If
    End If

    sz = mGeneral.GetValueWithDefault("type")

    With mView
        Select Case sz
        Case "1"
            ' /* solid */
            .SetHighColour g_SafeLong(mGeneral.GetValueWithDefault("col1"))
            .FillRect .Bounds
            
        Case "2"
            ' /* gradient */
            .SetHighColour g_SafeLong(mGeneral.GetValueWithDefault("col1"))
            .SetLowColour g_SafeLong(mGeneral.GetValueWithDefault("col2"))
            .FillRect .Bounds, MFX_VERT_GRADIENT

        Case "3"
            ' /* image */
            sz = mGeneral.GetValueWithDefault("image-file")

            If mGeneral.GetValueWithDefault("use_random_pic") = "1" Then _
                sz = mImageToUse

            Set pb = load_image_obj(sz)
            If ISNULL(pb) Then
                ' /* invalid */
                .SetHighColour rgba(0, 0, 0)
                .FillRect .Bounds
                .SetHighColour rgba(255, 192, 0)
                .StrokeLine new_BRect(0, 100, .Bounds.Right, 100), 2
                .SetFont "Courier New", 10, True
                .DrawString "Image '" & sz & "' could not be loaded", .Bounds.InsetByCopy(16, 48)

            Else
                Select Case mGeneral.GetValueWithDefault("image-pos")
                Case "1"
                    '/ * centre */
                    .DrawScaledImage pb, new_BPoint(Fix((.Width - pb.Width) / 2), Fix((.Height - pb.Height) / 2))

                Case "2"
                    ' /* tile */

                Case "3"
                    ' /* scale */
                    uDrawToFit mView, pb

                Case "4"
                    ' /* fit */
                    uDrawToFit mView, pb

                End Select

            End If

        Case Else
            ' /* invalid */
            .SetHighColour rgba(0, 0, 0)
            .FillRect .Bounds
            .SetHighColour rgba(255, 0, 0)
            .StrokeLine new_BRect(0, 100, .Bounds.Right, 100), 2

        End Select

        ' /* watermark */

Dim pp As BPoint

        If mGeneral.GetValueWithDefault("use_watermark") = "1" Then
            Set pb = load_image_obj(mGeneral.GetValueWithDefault("watermark-file"))
            If NOTNULL(pb) Then
                Select Case mGeneral.GetValueWithDefault("watermark-pos")
                Case "2"
                    Set pp = new_BPoint(.Width - pb.Width, 0).OffsetByCopy(-48, 48)

                Case "3"
                    Set pp = new_BPoint(0, .Height - pb.Height).OffsetByCopy(48, -48)

                Case "4"
                    Set pp = new_BPoint(.Width - pb.Width, .Height - pb.Height).OffsetByCopy(-48, -48)
                
                Case Else
                    Set pp = new_BPoint(0, 0).OffsetByCopy(48, 48)

                End Select

                .DrawScaledImage pb, pp, , 48
            
            End If
        End If

    End With

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

Private Sub KPrefsPage_AllAttached()
End Sub

Private Sub KPrefsPage_Attached()
End Sub

Private Sub KPrefsPage_ControlChanged(Control As prefs_kit_d2.BControl, ByVal Value As String)

    Debug.Print Control.GetName

    Select Case Control.GetName
    Case "go_site", "change_pic"
        Exit Sub

    End Select

    mGeneral.Update Control.GetName, Value
    mConfig.Save
    uSync


    Select Case Control.GetName
    Case "type"
        
        prefskit_SafeEnable thePanel, "col2", (Value = "2")
        prefskit_SafeEnable thePanel, "image-file", (Value = "3")
        prefskit_SafeEnable thePanel, "use_random_pic", (Value = "3")
        prefskit_SafeEnable thePanel, "image-pos", (Value = "3")
        
    End Select

End Sub

Private Sub KPrefsPage_ControlInvoked(Control As prefs_kit_d2.BControl)

    Select Case Control.GetName()
    Case "go_site"
        ShellExecute mhWnd, "open", "http://www.fullphat.net", vbNullString, vbNullString, SW_SHOW

    Case "change_pic"
        mImageToUse = uGetRandomFile(mGeneral.GetValueWithDefault("image-file"))
        uSync

    End Select

End Sub

Private Sub KPrefsPage_ControlNotify(Control As prefs_kit_d2.BControl, ByVal Notification As String, Data As melon.MMessage)
End Sub

Private Sub KPrefsPage_Create(Page As prefs_kit_d2.BPrefsPage)
End Sub

Private Sub KPrefsPage_Destroy()
End Sub

Private Sub KPrefsPage_Detached()
End Sub

Private Function KPrefsPage_hwnd() As Long
End Function

Private Sub KPrefsPage_PanelResized(ByVal Width As Long, ByVal Height As Long)
End Sub

Private Sub KPrefsPanel_PageChanged(ByVal NewPage As Long)
End Sub

Private Sub KPrefsPanel_Quit()

    Set thePanel = Nothing

End Sub

Private Sub KPrefsPanel_Ready()
End Sub

Private Sub KPrefsPanel_Selected(ByVal Command As String)
End Sub

Public Sub uDoPrefs()
Dim pp As BPrefsPage
Dim pm As CTempMsg

    If ISNULL(thePanel) Then

        g_Debug "frmAbout.uDoPrefs(): creating panel..."

        Set thePanel = New BPrefsPanel
        With thePanel
            .SetHandler Me

            thePanel.SetTitle "Desktopia Preferences"
            thePanel.SetWidth 400

            ' /* general page */

            g_Debug "frmAbout.NewDoPrefs(): general page..."
            Set pp = new_BPrefsPage("Background", load_image_obj(g_MakePath(App.Path) & "ico\prefs-bgnd.png"), Me)
            With pp
                .SetMargin 128

                ' /* launch options */

                .Add new_BPrefsControl("banner", "", "Background")
                .Add new_BPrefsControl("fancycycle", "type", "Solid|Gradient|Image", "Type:", mGeneral.GetValueWithDefault("type"))
                .Add new_BPrefsControl("fancycolourpicker", "col1", , "Primary colour:", mGeneral.GetValueWithDefault("col1"))
                .Add new_BPrefsControl("fancycolourpicker", "col2", , "Secondary colour:", mGeneral.GetValueWithDefault("col2"), , (mGeneral.GetValueWithDefault("type") = "2"))
                .Add new_BPrefsControl("fancyfilepicker", "image-file", mGeneral.GetValueWithDefault("image-file"), "Image:", , , (mGeneral.GetValueWithDefault("type") = "3"))
                .Add new_BPrefsControl("fancycycle", "image-pos", "Centre|Tile|Scale|Fit", "Position:", mGeneral.GetValueWithDefault("image-pos"), , (mGeneral.GetValueWithDefault("type") = "3"))
                .Add new_BPrefsControl("banner", "", "Options")
                .Add new_BPrefsControl("fancytoggle2", "use_random_pic", "Pick random picture from folder?", , mGeneral.GetValueWithDefault("use_random_pic"), , (mGeneral.GetValueWithDefault("type") = "3"))
                .Add new_BPrefsControl("fancybutton2", "change_pic", "Change", "", , , (mGeneral.GetValueWithDefault("type") = "3") And (mGeneral.GetValueWithDefault("use_random_pic") = "1"))

            End With
            .AddPage pp

            Set pp = new_BPrefsPage("Watermark", load_image_obj(g_MakePath(App.Path) & "ico\prefs-mark.png"), Me)
            With pp
                .SetMargin 128
                .Add new_BPrefsControl("banner", "", "Watermark")
                .Add new_BPrefsControl("fancytoggle2", "use_watermark", "Draw watermark?", , mGeneral.GetValueWithDefault("use_watermark"))
                .Add new_BPrefsControl("fancyfilepicker", "watermark-file", mGeneral.GetValueWithDefault("watermark-file"), "Image:")
                .Add new_BPrefsControl("fancycycle", "watermark-pos", "Top left|Top right|Bottom left|Bottom right", "Position:", mGeneral.GetValueWithDefault("watermark-pos"))

            End With
            .AddPage pp


            Set pp = new_BPrefsPage("Advanced", load_image_obj(g_MakePath(App.Path) & "ico\prefs-extr.png"), Me)
            With pp
                .SetMargin 24
                .Add new_BPrefsControl("banner", "", "Other Settings")
                .Add new_BPrefsControl("fancytoggle2", "obscure_taskbar", "Obscure Taskbar?", , mGeneral.GetValueWithDefault("obscure_taskbar"), , False)

            End With
            .AddPage pp


            Set pp = new_BPrefsPage("About", load_image_obj(g_MakePath(App.Path) & "ico\prefs-about.png"), Me)
            With pp
                .SetMargin 24
                .Add new_BPrefsControl("banner", "", "Desktopia " & App.Major & "." & App.Minor & " " & App.Comments)
                
                Set pm = New CTempMsg
                pm.Add "image-file", g_MakePath(App.Path) & "icon.png"
                pm.Add "image-height", 48
                pm.Add "valign", "centre"
                .Add new_BPrefsControl("labelex", "", "Build " & App.Revision, , , pm)

                .Add new_BPrefsControl("label", "", "© 2011 full phat products")
                .Add new_BPrefsControl("label", "", "http://www.fullphat.net/")
                .Add new_BPrefsControl("label", "", "")
                .Add new_BPrefsControl("label", "", "Released under the Simplified BSD Licence")
                .Add new_BPrefsControl("fancybutton2", "go_site", "Web site")

            End With
            .AddPage pp


'            ' /* apps */
'
'            g_Debug "frmAbout.NewDoPrefs(): apps page..."
'            Set mAppsPage = New TAppsPage
'            .AddPage new_BPrefsPage("Apps", load_image_obj(g_MakePath(App.Path) & "etc\icons\apps.png"), mAppsPage)

'            Set pp = new_BPrefsPage("Styles", load_image_obj(g_MakePath(App.Path) & "etc\icons\styles.png"), Me)
'            With pp
'                .SetMargin 0
'                Set pm = New CTempMsg
'                pm.Add "height", 412
'                Set pc = new_BPrefsControl("tabstrip", "", , , , pm)
'                BTabStrip_AddPage pc, "Display Styles", new_BPrefsPage("sty-display", , New TNetSubPage)
'                BTabStrip_AddPage pc, "Redirect Styles", new_BPrefsPage("sty-redirect", , New TNetSubPage)
'                .Add pc
'
'            End With
'            .AddPage pp
'
'                .Add new_BPrefsControl("fancytoolbar", "ftb>web_stuff", "Snarl Website|Discussion Group|Blog")
'
'
            g_Debug "frmAbout.NewDoPrefs(): displaying..."
            .Go
            g_SetWindowIconToAppResourceIcon .hWnd

            g_Debug "frmAbout.NewDoPrefs(): done"

        End With

    End If

    g_ShowWindow thePanel.hWnd, True, True
    SetForegroundWindow thePanel.hWnd

End Sub

Private Sub uLoadConfig()

    Set mConfig = New CConfFile3
    With mConfig
        .SetFile g_MakePath(App.Path) & "desktopia.conf"
        .Load

        If Not .Find("general", mGeneral) Then
            Set mGeneral = New CConfSection
            mGeneral.SetName "general"
            .Add mGeneral

        End If

        With mGeneral
            .AddIfMissing "type", "1"
            .AddIfMissing "col1", CStr(rgba(0, 48, 48))
            .AddIfMissing "col2", CStr(rgba(0, 0, 48))
            .AddIfMissing "image-file", ""
            .AddIfMissing "use_random_pic", "0"
            .AddIfMissing "image-pos", "1"
            .AddIfMissing "use_watermark", "0"
            .AddIfMissing "watermark-file", ""
            .AddIfMissing "watermark-pos", "2"

            If (.GetValueWithDefault("type") = "3") And (.GetValueWithDefault("use_random_pic") = "1") Then _
                mImageToUse = uGetRandomFile(.GetValueWithDefault("image-file"))

        End With

        .Save

    End With



End Sub

Private Function uGetRandomFile(ByVal Path As String) As String

    Path = g_GetPathParent(Path)
    If Path = "" Then _
        Exit Function

Dim pt As BTagList
Dim sz As String

    Set pt = new_BTagList()

    With New CFolderContent2
        If .SetTo(Path) Then
            .Rewind
            Do While .GetNextFile(sz)
                If (sz <> ".") And (sz <> "..") Then
                    Select Case g_GetExtension(sz, True)
                    Case "jpg", "jpeg", "gif", "png", "bmp", "tiff"
                        pt.Add new_BTagItem(sz, "")
                        Debug.Print sz

                    End Select

                End If
            Loop
        End If
    End With

Dim n As Long

    If pt.CountItems Then
        Randomize Timer
        n = (Rnd * (pt.CountItems - 1)) + 1
        uGetRandomFile = pt.TagAt(n).Name

    End If

'    MsgBox uGetRandomFile

End Function

Private Sub uSync()

    uRedraw
    g_UpdateWindow mhWnd, False

End Sub
