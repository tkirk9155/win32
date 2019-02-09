Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Text

Public Class Main

    ' Function declarations placed first for readability

    Public Delegate Function KbdMouseProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Public Delegate Function ShellProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Public Delegate Function CBTProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

    <DllImport("user32.dll")>
    Public Shared Function SendInput(ByVal nInputs As Integer,
                                      ByVal pInputs() As INPUT,
                                      ByVal cbSize As Integer) As Integer
    End Function


    <DllImport("User32.dll")>
    Public Shared Sub mouse_event(dwFlags As UInteger,
                                  dx As UInteger,
                                  dy As UInteger,
                                  dwData As UInteger,
                                  dwExtraInfo As Integer)
    End Sub

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function SetWindowsHookEx(ByVal idHook As Integer,
                                                      ByVal HookProc As KbdMouseProc,
                                                      ByVal hInstance As IntPtr,
                                                      ByVal wParam As Integer) As Integer
    End Function

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function SetWindowsHookEx(ByVal idHook As Integer,
                                                      ByVal HookProc As ShellProc,
                                                      ByVal hInstance As IntPtr,
                                                      ByVal wParam As Integer) As Integer
    End Function

    ' example:
    ' SetWindowsHookEx(WH_KEYBOARD_LL, MOUSEDLG, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function CallNextHookEx(ByVal idHook As Integer,
                                                    ByVal nCode As Integer,
                                                    ByVal wParam As IntPtr,
                                                    ByVal lParam As IntPtr) As Integer
    End Function

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)>
    Public Overloads Shared Function UnhookWindowsHookEx(ByVal idHook As Integer) As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function WindowFromPoint(ByVal p As Point) As IntPtr
    End Function

    <DllImport("user32.dll", ExactSpelling:=True, SetLastError:=True)>
    Public Shared Function GetCursorPos(ByRef lpPoint As Point) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowThreadProcessId(ByVal hwnd As IntPtr,
                                                    ByRef lpdwProcessId As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function GetClientRect(ByVal hWnd As System.IntPtr,
                                         ByRef lpRECT As Rectangle) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function ShowWindow(hWnd As IntPtr,
                                      <MarshalAs(UnmanagedType.I4)> nCmdShow As ShowWindowCommands) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindow(ByVal lpClassName As String,
                                       ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindowByClass(ByVal lpClassName As String,
                                              ByVal zero As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="FindWindow", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindowByCaption(ByVal zero As IntPtr,
                                                ByVal lpWindowName As String) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function FindWindowEx(ByVal parentHandle As IntPtr,
                                         ByVal childAfter As IntPtr,
                                         ByVal lclassName As String,
                                         ByVal windowTitle As String) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr,
                                      ByVal lParam As IntPtr) As Boolean

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function GetWindowText(ByVal hWnd As IntPtr,
                                         ByVal strText As StringBuilder,
                                         ByVal maxCount As Integer) As Integer
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer
    End Function

    <DllImport("user32.dll")>
    Public Shared Function EnumWindows(ByVal enumProc As EnumWindowsProc,
                                       ByVal lParam As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Public Shared Function IsWindowVisible(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindow(ByVal hWnd As IntPtr,
                                     ByVal uCmd As GetWindow_Cmd) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Public Shared Function GetWindowRect(ByVal hWnd As IntPtr,
                                         ByRef lpRect As RECT) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr,
                                       ByVal Msg As Integer,
                                       ByVal wParam As IntPtr,
                                       ByVal lParam As StringBuilder) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr,
                                       ByVal Msg As Integer,
                                       ByVal wParam As IntPtr,
                                       <MarshalAs(UnmanagedType.LPWStr)> ByVal lParam As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr,
                                       ByVal Msg As Integer,
                                       ByVal wParam As Integer,
                                       <MarshalAs(UnmanagedType.LPWStr)> ByVal lParam As String) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr,
                                       ByVal Msg As Integer,
                                       ByVal wParam As IntPtr,
                                       ByRef lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr,
                                       ByVal Msg As Integer,
                                       ByVal wParam As Integer,
                                       ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("shell32.dll")>
    Public Shared Sub SHGetSetSettings(ByRef lpShellState As SHELLSTATE, 
                                       ByVal dwMask As SSF, 
                                       ByVal bSet As Boolean)
    End Sub

    <StructLayout(LayoutKind.Sequential)>
    Public Structure KBDLLHOOKSTRUCT
        Public vkCode As UInt32
        Public scanCode As UInt32
        Public flags As KBDLLHOOKSTRUCTFlags
        Public time As UInt32
        Public dwExtraInfo As UIntPtr
    End Structure

    <Flags()>
    Public Enum KBDLLHOOKSTRUCTFlags As UInt32
        LLKHF_EXTENDED = &H1
        LLKHF_INJECTED = &H10
        LLKHF_ALTDOWN = &H20
        LLKHF_UP = &H80
    End Enum

    Public Enum GetWindow_Cmd As UInteger
        GW_HWNDFIRST = 0
        GW_HWNDLAST = 1
        GW_HWNDNEXT = 2
        GW_HWNDPREV = 3
        GW_OWNER = 4
        GW_CHILD = 5
        GW_ENABLEDPOPUP = 6
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RECT
        Public _Left As Integer, _Top As Integer, _Right As Integer, _Bottom As Integer

        Public Sub New(ByVal Rectangle As Rectangle)
            Me.New(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom)
        End Sub
        Public Sub New(ByVal Left As Integer, ByVal Top As Integer, ByVal Right As Integer, ByVal Bottom As Integer)
            _Left = Left
            _Top = Top
            _Right = Right
            _Bottom = Bottom
        End Sub

        Public Property X As Integer
            Get
                Return _Left
            End Get
            Set(ByVal value As Integer)
                _Right = _Right - _Left + value
                _Left = value
            End Set
        End Property
        Public Property Y As Integer
            Get
                Return _Top
            End Get
            Set(ByVal value As Integer)
                _Bottom = _Bottom - _Top + value
                _Top = value
            End Set
        End Property
        Public Property Left As Integer
            Get
                Return _Left
            End Get
            Set(ByVal value As Integer)
                _Left = value
            End Set
        End Property
        Public Property Top As Integer
            Get
                Return _Top
            End Get
            Set(ByVal value As Integer)
                _Top = value
            End Set
        End Property
        Public Property Right As Integer
            Get
                Return _Right
            End Get
            Set(ByVal value As Integer)
                _Right = value
            End Set
        End Property
        Public Property Bottom As Integer
            Get
                Return _Bottom
            End Get
            Set(ByVal value As Integer)
                _Bottom = value
            End Set
        End Property
        Public Property Height() As Integer
            Get
                Return _Bottom - _Top
            End Get
            Set(ByVal value As Integer)
                _Bottom = value + _Top
            End Set
        End Property
        Public Property Width() As Integer
            Get
                Return _Right - _Left
            End Get
            Set(ByVal value As Integer)
                _Right = value + _Left
            End Set
        End Property
        Public Property Location() As Point
            Get
                Return New Point(Left, Top)
            End Get
            Set(ByVal value As Point)
                _Right = _Right - _Left + value.X
                _Bottom = _Bottom - _Top + value.Y
                _Left = value.X
                _Top = value.Y
            End Set
        End Property
        Public Property Size() As Size
            Get
                Return New Size(Width, Height)
            End Get
            Set(ByVal value As Size)
                _Right = value.Width + _Left
                _Bottom = value.Height + _Top
            End Set
        End Property

        Public Shared Widening Operator CType(ByVal Rectangle As RECT) As Rectangle
            Return New Rectangle(Rectangle.Left, Rectangle.Top, Rectangle.Width, Rectangle.Height)
        End Operator
        Public Shared Widening Operator CType(ByVal Rectangle As Rectangle) As RECT
            Return New RECT(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom)
        End Operator
        Public Shared Operator =(ByVal Rectangle1 As RECT, ByVal Rectangle2 As RECT) As Boolean
            Return Rectangle1.Equals(Rectangle2)
        End Operator
        Public Shared Operator <>(ByVal Rectangle1 As RECT, ByVal Rectangle2 As RECT) As Boolean
            Return Not Rectangle1.Equals(Rectangle2)
        End Operator

        Public Overrides Function ToString() As String
            Return "{Left: " & _Left & "; " & "Top: " & _Top & "; Right: " & _Right & "; Bottom: " & _Bottom & "}"
        End Function

        Public Overloads Function Equals(ByVal Rectangle As RECT) As Boolean
            Return Rectangle.Left = _Left AndAlso Rectangle.Top = _Top AndAlso Rectangle.Right = _Right AndAlso Rectangle.Bottom = _Bottom
        End Function
        Public Overloads Overrides Function Equals(ByVal [Object] As Object) As Boolean
            If TypeOf [Object] Is RECT Then
                Return Equals(DirectCast([Object], RECT))
            ElseIf TypeOf [Object] Is Rectangle Then
                Return Equals(New RECT(DirectCast([Object], Rectangle)))
            End If

            Return False
        End Function
    End Structure

    <StructLayout(LayoutKind.Sequential)> Public Structure SHELLSTATE
        Public flags_1 As UInteger
        Public dwWin95Unused As UInteger
        Public uWin95Unused As UInteger
        Public lParamSort As Integer
        Public iSortDirection As Integer
        Public version As UInteger
        Public uNotUsed As UInteger
        Public flags_2 As UInteger

        Public Property fShowAllObjects() As Boolean
            Get
                Return (flags_1 And &H1UI) = &H1UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H1UI, flags_1 And Not &H1UI)
            End Set
        End Property
        Public Property fShowExtensions() As Boolean
            Get
                Return (flags_1 And &H2UI) = &H2UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H2UI, flags_1 And Not &H2UI)
            End Set
        End Property
        Public Property fNoConfirmRecycle() As Boolean
            Get
                Return (flags_1 And &H4UI) = &H4UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H4UI, flags_1 And Not &H4UI)
            End Set
        End Property
        Public Property fShowSysFiles() As Boolean
            Get
                Return (flags_1 And &H8UI) = &H8UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H8UI, flags_1 And Not &H8UI)
            End Set
        End Property
        Public Property fShowCompColor() As Boolean
            Get
                Return (flags_1 And &H10UI) = &H10UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H10UI, flags_1 And Not &H10UI)
            End Set
        End Property
        Public Property fDoubleClickInWebView() As Boolean
            Get
                Return (flags_1 And &H20UI) = &H20UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H20UI, flags_1 And Not &H20UI)
            End Set
        End Property
        Public Property fDesktopHTML() As Boolean
            Get
                Return (flags_1 And &H40UI) = &H40UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H40UI, flags_1 And Not &H40UI)
            End Set
        End Property
        Public Property fWin95Classic() As Boolean
            Get
                Return (flags_1 And &H80UI) = &H80UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H80UI, flags_1 And Not &H80UI)
            End Set
        End Property
        Public Property fDontPrettyPath() As Boolean
            Get
                Return (flags_1 And &H100UI) = &H100UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H100UI, flags_1 And Not &H100UI)
            End Set
        End Property
        Public Property fShowAttribCol() As Boolean
            Get
                Return (flags_1 And &H200UI) = &H200UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H200UI, flags_1 And Not &H200UI)
            End Set
        End Property
        Public Property fMapNetDrvBtn() As Boolean
            Get
                Return (flags_1 And &H400UI) = &H400UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H400UI, flags_1 And Not &H400UI)
            End Set
        End Property
        Public Property fShowInfoTip() As Boolean
            Get
                Return (flags_1 And &H800UI) = &H800UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H800UI, flags_1 And Not &H800UI)
            End Set
        End Property
        Public Property fHideIcons() As Boolean
            Get
                Return (flags_1 And &H1000UI) = &H1000UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H1000UI, flags_1 And Not &H1000UI)
            End Set
        End Property
        Public Property fWebView() As Boolean
            Get
                Return (flags_1 And &H2000UI) = &H2000UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H2000UI, flags_1 And Not &H2000UI)
            End Set
        End Property
        Public Property fFilter() As Boolean
            Get
                Return (flags_1 And &H4000UI) = &H4000UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H4000UI, flags_1 And Not &H4000UI)
            End Set
        End Property
        Public Property fShowSuperHidden() As Boolean
            Get
                Return (flags_1 And &H8000UI) = &H8000UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H8000UI, flags_1 And Not &H8000UI)
            End Set
        End Property
        Public Property fNoNetCrawling() As Boolean
            Get
                Return (flags_1 And &H10000UI) = &H10000UI
            End Get
            Set(ByVal value As Boolean)
                flags_1 = IIf(value, flags_1 Or &H10000UI, flags_1 And Not &H10000UI)
            End Set
        End Property

        Public Property fSepProcess() As Boolean
            Get
                Return (flags_2 And &H1UI) = &H1UI
            End Get
            Set(ByVal value As Boolean)
                flags_2 = IIf(value, flags_2 Or &H1UI, flags_2 And Not &H1UI)
            End Set
        End Property
        Public Property fStartPanelOn() As Boolean
            Get
                Return (flags_2 And &H2UI) = &H2UI
            End Get
            Set(ByVal value As Boolean)
                flags_2 = IIf(value, flags_2 Or &H2UI, flags_2 And Not &H2UI)
            End Set
        End Property
        Public Property fShowStartPage() As Boolean
            Get
                Return (flags_2 And &H4UI) = &H4UI
            End Get
            Set(ByVal value As Boolean)
                flags_2 = IIf(value, flags_2 Or &H4UI, flags_2 And Not &H4UI)
            End Set
        End Property
        Public Property fAutoCheckSelect() As Boolean
            Get
                Return (flags_2 And &H8UI) = &H8UI
            End Get
            Set(ByVal value As Boolean)
                flags_2 = IIf(value, flags_2 Or &H8UI, flags_2 And Not &H8UI)
            End Set
        End Property
        Public Property fIconsOnly() As Boolean
            Get
                Return (flags_2 And &H10UI) = &H10UI
            End Get
            Set(ByVal value As Boolean)
                flags_2 = IIf(value, flags_2 Or &H10UI, flags_2 And Not &H10UI)
            End Set
        End Property
        Public Property fShowTypeOverlay() As Boolean
            Get
                Return (flags_2 And &H20UI) = &H20UI
            End Get
            Set(ByVal value As Boolean)
                flags_2 = IIf(value, flags_2 Or &H20UI, flags_2 And Not &H20UI)
            End Set
        End Property
        Public Property fShowStatusBar() As Boolean
            Get
                Return (flags_2 And &H40UI) = &H40UI
            End Get
            Set(ByVal value As Boolean)
                flags_2 = IIf(value, flags_2 Or &H40UI, flags_2 And Not &H40UI)
            End Set
        End Property
    End Structure

    <Flags()> Public Enum SSF As UInteger
        SSF_SHOWALLOBJECTS = &H1
        SSF_SHOWEXTENSIONS = &H2
        SSF_HIDDENFILEEXTS = &H4
        SSF_SERVERADMINUI = &H4
        SSF_SHOWCOMPCOLOR = &H8
        SSF_SORTCOLUMNS = &H10
        SSF_SHOWSYSFILES = &H20
        SSF_DOUBLECLICKINWEBVIEW = &H80
        SSF_SHOWATTRIBCOL = &H100
        SSF_DESKTOPHTML = &H200
        SSF_WIN95CLASSIC = &H400
        SSF_DONTPRETTYPATH = &H800
        SSF_MAPNETDRVBUTTON = &H1000
        SSF_SHOWINFOTIP = &H2000
        SSF_HIDEICONS = &H4000
        SSF_NOCONFIRMRECYCLE = &H8000
        SSF_FILTER = &H10000
        SSF_WEBVIEW = &H20000
        SSF_SHOWSUPERHIDDEN = &H40000
        SSF_SEPPROCESS = &H80000
        SSF_NONETCRAWLING = &H100000
        SSF_STARTPANELON = &H200000
        SSF_SHOWSTARTPAGE = &H400000
        SSF_AUTOCHECKSELECT = &H800000
        SSF_ICONSONLY = &H1000000
        SSF_SHOWTYPEOVERLAY = &H2000000
        SSF_SHOWSTATUSBAR = &H4000000
    End Enum

    Public Enum ShowWindowCommands As Integer
        Hide = 0
        Normal = 1
        ShowMinimized = 2
        Maximize = 3
        ShowMaximized = 3
        ShowNoActivate = 4
        Show = 5
        Minimize = 6
        ShowMinNoActive = 7
        ShowNA = 8
        Restore = 9
        ShowDefault = 10
        ForceMinimize = 11
    End Enum

    Public Enum INPUT_TYPE As Integer
        KeyDown = &H0
        KeyUp = &H2
    End Enum

    <StructLayout(LayoutKind.Explicit)>
    Public Structure INPUT
        'Field offset 32 bit machine 4
        '64 bit machine 8
        <FieldOffset(0)>
        Public type As Integer
        <FieldOffset(8)>
        Public mi As MOUSEINPUT
        <FieldOffset(8)>
        Public ki As KEYBDINPUT
        <FieldOffset(8)>
        Public hi As HARDWAREINPUT
    End Structure
    Public Structure MOUSEINPUT
        Public dx As Integer
        Public dy As Integer
        Public mouseData As Integer
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure
    Public Structure KEYBDINPUT
        Public wVk As Short
        Public wScan As Short
        Public dwFlags As Integer
        Public time As Integer
        Public dwExtraInfo As IntPtr
    End Structure
    Public Structure HARDWAREINPUT
        Public uMsg As Integer
        Public wParamL As Short
        Public wParamH As Short
    End Structure

    Public Enum MouseEventFlags As UInteger
        LEFTDOWN = &H2
        LEFTUP = &H4
        MIDDLEDOWN = &H20
        MIDDLEUP = &H40
        MOVE = &H1
        ABSOLUTE = &H8000
        RIGHTDOWN = &H8
        RIGHTUP = &H10
    End Enum

    Public Const WM_RBUTTONDOWN = &H204
    Public Const WM_RBUTTONUP = &H205
    Public Const WM_LBUTTONDOWN = &H201
    Public Const WM_LBUTTONUP = &H202

    Public Const WM_KEYDOWN = &H100
    Public Const WM_KEYUP = &H101
    Public Const WM_SYSKEYDOWN = &H104
    Public Const WM_SYSKEYUP = &H105
    Public Const WM_HOTKEY = &H312

    Public Const WM_COMMAND = &H111

    Public Const HC_ACTION = 0

    Public Const HSHELL_ACCESSIBILITYSTATE As Integer = 11
    Public Const HSHELL_ACTIVATESHELLWINDOW As Integer = 3
    Public Const HSHELL_APPCOMMAND As Integer = 12
    Public Const HSHELL_GETMINRECT As Integer = 5
    Public Const HSHELL_LANGUAGE As Integer = 8
    Public Const HSHELL_REDRAW As Integer = 6
    Public Const HSHELL_TASKMAN As Integer = 7
    Public Const HSHELL_WINDOWACTIVATED = 4
    Public Const HSHELL_WINDOWCREATED = 1
    Public Const HSHELL_WINDOWDESTROYED = 2
    Public Const HSHELL_WINDOWREPLACED = 13

    Public Const WH_MOUSE_LL As Integer = 14
    Public Const WH_KEYBOARD_LL As Integer = 13
    Public Const WH_CALLWNDPROC As Integer = 4
    Public Const WH_CALLWNDPROCRET As Integer = 12
    Public Const WH_CBT As Integer = 5
    Public Const WH_DEBUG As Integer = 9
    Public Const WH_FOREGROUNDIDLE As Integer = 11
    Public Const WH_GETMESSAGE As Integer = 3
    Public Const WH_JOURNALPLAYBACK As Integer = 1
    Public Const WH_JOURNALRECORD As Integer = 0
    Public Const WH_KEYBOARD As Integer = 2
    Public Const WH_MOUSE As Integer = 7
    Public Const WH_MSGFILTER As Integer = -1
    Public Const WH_SHELL As Integer = 10
    Public Const WH_SYSMSGFILTER As Integer = 6

    Public Const HCBT_ACTIVATE As Integer = 5
    Public Const HCBT_CLICKSKIPPED As Integer = 6
    Public Const HCBT_CREATEWND As Integer = 3
    Public Const HCBT_DESTROYWND As Integer = 4
    Public Const HCBT_KEYSKIPPED As Integer = 7
    Public Const HCBT_MINMAX As Integer = 1
    Public Const HCBT_MOVESIZE As Integer = 0
    Public Const HCBT_QS As Integer = 2
    Public Const HCBT_SETFOCUS As Integer = 9
    Public Const HCBT_SYSCOMMAND As Integer = 8



    ' not PInvoke, used for Strings.Chr(CharCode As Integer)
    '
    <Flags()> Public Enum ASCII As Integer
        ExclamationMark = 33
        DoubleQuotes = 34
        Pound = 35
        Dollar = 36
        Percent = 37
        Ampersand = 38
        Apostrophe = 39
        OpenParentheses = 40
        CloseParentheses = 41
        Asterisk = 42
        Plus = 43
        SingleQuote = 44
        Minus = 45
        Period = 46
        ForwardSlash = 47
        Zero = 48
        One = 49
        Two = 50
        Three = 51
        Four = 52
        Five = 53
        Six = 54
        Seven = 55
        Eight = 56
        Nine = 57
        Colon = 58
        Semicolon = 59
        LessThan = 60
        Equal = 61
        GreaterThan = 62
        QuestionMark = 63
        At = 64
        A_Upper = 65
        B_Upper = 66
        C_Upper = 67
        D_Upper = 68
        E_Upper = 69
        F_Upper = 70
        G_Upper = 71
        H_Upper = 72
        I_Upper = 73
        J_Upper = 74
        K_Upper = 75
        L_Upper = 76
        M_Upper = 77
        N_Upper = 78
        O_Upper = 79
        P_Upper = 80
        Q_Upper = 81
        R_Upper = 82
        S_Upper = 83
        T_Upper = 84
        U_Upper = 85
        V_Upper = 86
        W_Upper = 87
        X_Upper = 88
        Y_Upper = 89
        Z_Upper = 90
        OpenBracket = 91
        BackSlash = 92
        CloseBracket = 93
        Carrot = 94
        Underscore = 95
        Tilde = 96
        a_Lower = 97
        b_Lower = 98
        c_Lower = 99
        d_Lower = 100
        e_Lower = 101
        f_Lower = 102
        g_Lower = 103
        h_Lower = 104
        i_Lower = 105
        j_Lower = 106
        k_Lower = 107
        l_Lower = 108
        m_Lower = 109
        n_Lower = 110
        o_Lower = 111
        p_Lower = 112
        q_Lower = 113
        r_Lower = 114
        s_Lower = 115
        t_Lower = 116
        u_Lower = 117
        v_Lower = 118
        w_Lower = 119
        x_Lower = 120
        y_Lower = 121
        z_Lower = 122
        OpenBrace = 123
        VertBar = 124
        CloseBrace = 125
        SquigglyTilde = 126
        Space = 127
    End Enum

End Class
