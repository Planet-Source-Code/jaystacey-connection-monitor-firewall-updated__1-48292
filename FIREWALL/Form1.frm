VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Connections Monitor"
   ClientHeight    =   4500
   ClientLeft      =   450
   ClientTop       =   1140
   ClientWidth     =   11445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   11445
   Begin MSComctlLib.ProgressBar Progbar 
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   4245
      Width           =   11445
      _ExtentX        =   20188
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5760
      Top             =   3360
   End
   Begin VB.PictureBox pic16 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2820
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox pic32 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   2280
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   1
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   3360
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   4560
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "iml32"
      SmallIcons      =   "iml16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   3529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Direction"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Local Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Remote Host"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Remote Port"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "File Path"
         Object.Width           =   7056
      EndProperty
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Begin VB.Menu Refreshlist 
         Caption         =   "Refresh"
      End
      Begin VB.Menu seperator1 
         Caption         =   "-"
      End
      Begin VB.Menu AutoFresh 
         Caption         =   "Automatic Refresh"
         Checked         =   -1  'True
      End
      Begin VB.Menu ChangeInterval 
         Caption         =   "Change Refresh Interval"
      End
      Begin VB.Menu seperator2 
         Caption         =   "-"
      End
      Begin VB.Menu ToTray 
         Caption         =   "Minimise to system tray"
      End
      Begin VB.Menu ExitProg 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu MenuPop 
      Caption         =   "MenuPop"
      Visible         =   0   'False
      Begin VB.Menu KillProcess 
         Caption         =   "Kill Process"
      End
   End
   Begin VB.Menu MenuSysPop 
      Caption         =   "MenuSysPop"
      Visible         =   0   'False
      Begin VB.Menu OpenConMon 
         Caption         =   "Open Connections Monitor"
      End
      Begin VB.Menu Seperator3 
         Caption         =   "-"
      End
      Begin VB.Menu ExitProg1 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 0
Option Explicit

DefLng A-N, P-Z
DefBool O

'*********************************************

'This UDT will help us handle the Tray Icon
Private Type NOTIFYICONDATA
    cbSize As Long              ' Its size
    hWnd As Long                ' The handle of the window which will receive Windows' _
                                 messages
    uID As Long                 ' Its Unique Identification
    uFlags As Long              ' Its style
    uCallbackMessage As Long    ' The message Windows will send us to process
    hIcon As Long               ' Its Icon
    szTip As String * 64        ' Its tip. NOTE: MUST BE FINISHED WITH NULL
End Type

' This API lets us receive any window's messages
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' This API handles the Tray Icon
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" _
        (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Boolean

' dwMessage Constants
Private Const NIM_ADD = &H0             'Flag : "ALL NEW nid"
Private Const NIM_MODIFY = &H1          'Flag : "ONLY MODIFYING nid"
Private Const NIM_DELETE = &H2          'Flag : "DELETE THE CURRENT nid"
Private Const NIF_MESSAGE = &H1         'Flag : "Message in nid is valid"
Private Const NIF_ICON = &H2            'Flag : "Icon in nid is valid"
Private Const NIF_TIP = &H4             'Flag : "Tip in nid is valid"

' Subclassing constant
Private Const GWL_WNDPROC = (-4)

' Our class' events
Event LeftButtonUp()                'Raised when user stops clicking with left button
Event LeftButtonDown()              'Raised when user clicks with left button
Event LeftButtonDoubleClick()       'Raised when user double-clicks with left button

Event RightButtonUp()               'Raised when user stops clicking with right button
Event RightButtonDown()             'Raised when user clicks with right button
Event RightButtonDoubleClick()      'Raised when user double-clicks with right button

Event MiddleButtonUp()              'Raised when user stops clicking with middle button
Event MiddleButtonDown()            'Raised when user clicks with middle button
Event MiddleButtonDoubleClick()     'Raised when user double-clicks with middle button

Private nidTrayIcon         As NOTIFYICONDATA   ' Variable which contains our Tray Icon
Private m_lHWnd             As Long             ' Variable which contains the handle _






'********************************************



'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME _
        Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX _
        Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type

'----------------------------------------------------------
'Functions & Procedures
'----------------------------------------------------------
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
    (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" _
    (ByVal himl&, ByVal i&, ByVal hDCDest&, _
    ByVal x&, ByVal y&, ByVal flags&) As Long


'----------------------------------------------------------
'Private variables
'----------------------------------------------------------
Private ShInfo As SHFILEINFO
Public Tablenum As Long
Private pTcpTable As MIB_TCPTABLE
Public Processing As Boolean

Public Sub RefreshView()
  Dim i As Integer, O As Integer
  Dim fileNum As String
  Dim Item As ListItem
  Dim strRetVal As String
  
  Processing = True
  
  StatusBar1.Panels(2).Text = "Loading..."
  
   Me.MousePointer = vbHourglass

  ListView1.ListItems.Clear
   
    ListView1.Icons = Nothing
    ListView1.SmallIcons = Nothing
    iml32.ListImages.Clear
    iml16.ListImages.Clear
        
   Progbar.Visible = True
   Progbar.Value = 2

DoEvents
strRetVal = Execute("Netstat - o")
Progbar.Value = 4

DoEvents
Parse strRetVal
Progbar.Value = 6

DoEvents
MdlLoadProcess.LoadNTProcess
Progbar.Value = 8
  
 On Error Resume Next
  For i = 0 To StatsLen - 1
  
  If Connection(i).FileName <> "" Then Set Item = ListView1.ListItems.Add(, , Right(Connection(i).FileName, Len(Connection(i).FileName) - InStrRev(Connection(i).FileName, "\"))) Else Set Item = ListView1.ListItems.Add(, , "Unknown")
    
    If Connection(i).LocalPort = Connection(i).RemotePort And Connection(i).LocalPort <> "" Then Item.SubItems(1) = "Incomming" Else Item.SubItems(1) = "Outgoing"
    Item.SubItems(2) = Connection(i).LocalPort
    Item.SubItems(3) = Connection(i).RemoteHost
    Item.SubItems(4) = Connection(i).RemotePort
    Item.SubItems(5) = Connection(i).State
    Item.SubItems(6) = Connection(i).FileName
    ConProcessID(ListView1.ListItems.Count).FileName = Right(Connection(i).FileName, Len(Connection(i).FileName) - InStrRev(Connection(i).FileName, "\"))
    ConProcessID(ListView1.ListItems.Count).ProcessId = Connection(i).ProcessId
    'Item.EnsureVisible
    DoEvents
  Next i
  
  GetAllIcons
  ShowIcons
  
Progbar.Value = 10
Progbar.Value = 0

    Me.MousePointer = vbNormal
    
'SetForegroundWindow (Me.hWnd) <-- sometimes windows jumps out of focus when loading this ill bring it to front window

StatusBar1.Panels(2).Text = "Loaded"
StatusBar1.Panels(1).Text = "Last Refresh - " & Time
Progbar.Visible = False
Processing = False
End Sub

Private Sub AutoFresh_Click()
If AutoFresh.Checked = True Then
Timer1.Enabled = False
AutoFresh.Checked = False
Else
AutoFresh.Checked = True
Timer1.Enabled = True
End If
End Sub

Private Sub ChangeInterval_Click()
Dim InputValue As String

InputValue = InputBox("Input time between interval in seconds:", "Set Interval Timer", Timer1.Interval / 1000)

If InputValue = "" Then Exit Sub

If IsNumeric(InputValue) = True Then Timer1.Interval = InputValue * 1000 Else MsgBox "Not a valid numeracy" & vbNewLine & "Interval not changed!", vbInformation + vbOKOnly, "Interval Error"

End Sub

Private Sub ExitProg_Click()
Unload Me
End Sub

Private Sub Form_Load()
ListView1.ColumnHeaders(1).Width = 1300 'ListView1.Width \ 4 - 1500
ListView1.ColumnHeaders(2).Width = 1100
ListView1.ColumnHeaders(3).Width = 1100
ListView1.ColumnHeaders(4).Width = ListView1.Width \ 4 - 1000
ListView1.ColumnHeaders(5).Width = 1100
ListView1.ColumnHeaders(6).Width = 1300
ListView1.ColumnHeaders(7).Width = ListView1.Width \ 2 + 1000

If IsNetConnectOnline() = True Then StatusBar1.Panels(3).Text = "Online"

Me.Show
DoEvents
Timer1_Timer

'DoEvents
'RefreshView
End Sub

Private Sub Form_Resize()

On Error Resume Next
ListView1.Width = Me.Width - 150
ListView1.Left = 0
ListView1.Height = Me.Height - 1050

ListView1.ColumnHeaders(1).Width = 1300 'ListView1.Width \ 4 - 1500
ListView1.ColumnHeaders(2).Width = 1100
ListView1.ColumnHeaders(3).Width = 1100
ListView1.ColumnHeaders(4).Width = ListView1.Width \ 4 - 1000
ListView1.ColumnHeaders(5).Width = 1100
ListView1.ColumnHeaders(6).Width = 1300
ListView1.ColumnHeaders(7).Width = ListView1.Width \ 2 + 1000

StatusBar1.Panels(1).Width = Me.Width / 4
StatusBar1.Panels(2).Width = Me.Width / 2
StatusBar1.Panels(3).Width = Me.Width / 4

Progbar.Width = StatusBar1.Panels(1).Width
Progbar.Height = StatusBar1.Height - 80
Progbar.Top = StatusBar1.Top + 50
Progbar.Left = StatusBar1.Left
Progbar.Max = 10

End Sub

Private Sub Form_Unload(Cancel As Integer)
RemoveTrayIcon
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Button

    Case vbLeftButton
    'MsgBox ListView1.SelectedItem.Key

    Case vbRightButton
    PopupMenu MenuPop

End Select
End Sub
Private Sub KillProcess_Click()
Dim MessageAnswer
MessageAnswer = MsgBox("Terminate: " & ConProcessID(ListView1.SelectedItem.Index).FileName, vbYesNo + vbExclamation, "Terminate Process")

If MessageAnswer = vbYes Then
KillProcessById (ConProcessID(ListView1.SelectedItem.Index).ProcessId)
RefreshView
StatusBar1.Panels(2).Text = "Process Terminated: " & ConProcessID(ListView1.SelectedItem.Index).FileName
End If

End Sub

Private Sub Refreshlist_Click()
RefreshView
DoEvents
End Sub

Private Sub ShowIcons()
'-----------------------------------------
'Show the icons in the lvw
'-----------------------------------------
On Error Resume Next

Dim Item As ListItem
With ListView1
  '.ListItems.Clear
  .Icons = iml32        'Large
  .SmallIcons = iml16   'Small
  For Each Item In .ListItems
    Item.Icon = Item.Index
    Item.SmallIcon = Item.Index
  Next
End With

End Sub

Private Sub GetAllIcons()
'--------------------------------------------------
'Extract all icons
'--------------------------------------------------
Dim Item As ListItem
Dim FileName As String

On Local Error Resume Next
For Each Item In ListView1.ListItems
  FileName = Item.SubItems(Item.ListSubItems.Count) ' & Item.Text
  GetIcon FileName, Item.Index
Next

End Sub

Private Function GetIcon(FileName As String, Index As Long) As Long
'---------------------------------------------------------------------
'Extract an individual icon
'---------------------------------------------------------------------
Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
Dim r As Long

'Get a handle to the small icon
hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
'Get a handle to the large icon
hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), _
         BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)

'If the handle(s) exists, load it into the picture box(es)
If hLIcon <> 0 Then
  'Large Icon
  
  With pic32
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  'Small Icon
  With pic16
    Set .Picture = LoadPicture("")
    .AutoRedraw = True
    r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hDC, 0, 0, ILD_TRANSPARENT)
    .Refresh
  End With
  Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
  Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
End If
End Function

Private Sub Timer1_Timer()
'RefreshView
 
Dim pdwSize As Long
Dim bOrder As Long
Dim nRet As Long
Dim TableLen As Long
Dim i, Act, NoAct As Integer

If Processing = True Then Exit Sub

If IsNetConnectOnline() = False Then
Menu.Enabled = False
StatusBar1.Panels(3).Text = "No Connection Present"
ListView1.Enabled = False
Exit Sub
End If

If ListView1.Enabled = False Then
Menu.Enabled = True
StatusBar1.Panels(3).Text = "Online"
ListView1.Enabled = True
End If

Act = 0
NoAct = 0

nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)
nRet = GetTcpTable(pTcpTable, pdwSize, bOrder)


For i = 0 To pTcpTable.dwNumEntries - 1
If pTcpTable.table(i).dwState - 1 <> 0 Then
    If pTcpTable.table(i).dwState - 1 <> 1 Then
        If pTcpTable.table(i).dwState - 1 <> 10 Then
            Act = Act + 1
        End If
    End If
End If
Next i

If Act <> StatsLen Then
DoEvents
RefreshView
End If

End Sub


Public Function ChangeIcon(ByVal lNewIcon As Long) As Long
On Error GoTo ErrHandler
    Dim nidNewTray As NOTIFYICONDATA        ' We declare a "dummy" Tray Icon _
                                             just to modify the actual one
    
    With nidNewTray
        .cbSize = Len(nidNewTray)           ' We change its size
        .hWnd = m_lHWnd                     ' We set the window we first chose to _
                                             capture every message
        .uID = 1                            ' The same uID as the original one
        .uFlags = NIF_ICON                  ' This flag makes the Shell_NotifyIcon API _
                                             change the icon
        .hIcon = lNewIcon                   ' The new icon
    End With
    
    Shell_NotifyIcon NIM_MODIFY, nidNewTray ' We call the API to do the work
    
    ChangeIcon = 0                          ' Everything went OK, return 0
    Exit Function
ErrHandler:
    ChangeIcon = Err.Number                 ' Oops, something went wrong, _
                                             return the error number
End Function

Public Function ChangeTip(ByVal nNewTip As String) As Long
On Error GoTo ErrHandler
    Dim nidNewTray As NOTIFYICONDATA            ' We declare a "dummy" Tray Icon _
                                             just to modify the actual one
    
    With nidNewTray
        .cbSize = Len(nidNewTray)           ' We change its size
        .hWnd = m_lHWnd                     ' We set the window we first chose to _
                                             capture every message
        .uID = 1                            ' The same uID as the original one
        .uFlags = NIF_TIP                   ' This flag makes the Shell_NotifyIcon API _
                                             change the tip
        .szTip = nNewTip & Chr$(0)          ' The new tip
    End With
    
    Shell_NotifyIcon NIM_MODIFY, nidNewTray ' We call the API to do the work
    
    ChangeTip = 0                           ' Everything went OK, return 0
    Exit Function
ErrHandler:
    ChangeTip = Err.Number                  ' Oops, something went wrong, _
                                             return the error number
End Function

Public Function RemoveTrayIcon() As Long
On Error GoTo ErrHandler
    
    SetWindowLong m_lHWnd, GWL_WNDPROC, lPreviousProcess   ' Stops subclassing the window
    Shell_NotifyIcon NIM_DELETE, nidTrayIcon        ' Removes the Tray Icon
    
    RemoveTrayIcon = 0                              ' Everything went OK, return 0
    Exit Function
ErrHandler:
    RemoveTrayIcon = Err.Number                     ' Oops, something went wrong, _
                                                     return the error number
End Function

Public Function ShowTrayIcon(ByVal lHWnd As Long, ByVal sTip As String, _
                            ByVal lIcon As Long) As Long
On Error GoTo ErrHandler
    ' Save the window's handle
    m_lHWnd = lHWnd
    
    ' Subclass the window, but save its process just to make Windows _
     handle the messages we don't want to
    lPreviousProcess = SetWindowLong(m_lHWnd, GWL_WNDPROC, AddressOf WndProc)
    
    With nidTrayIcon
        .cbSize = Len(nidTrayIcon)                      ' Its size
        .hIcon = lIcon                                  ' Its icon
        .hWnd = m_lHWnd                                 ' The handle of the window which _
                                                         will capture its messages
        .szTip = sTip & Chr$(0)                         ' The tip
        .uCallbackMessage = WM_USER + 1                 ' The CallBack value we will _
                                                         use to determine the events
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uID = 1                                        ' Its unique ID
    End With
    
    Shell_NotifyIcon NIM_ADD, nidTrayIcon               ' We call the API to do the work
    
    ShowTrayIcon = 0                                    ' Everything went OK, return 0
    Exit Function
ErrHandler:
    ShowTrayIcon = Err.Number                           'ERROR: we return its number
End Function

Friend Sub ShowEvent(ByVal lIdEvent As Long)
    Select Case lIdEvent
        Case WM_LBUTTONDOWN: RaiseEvent LeftButtonDown              'Left click down
        Case WM_LBUTTONUP: RaiseEvent LeftButtonUp                  'Left click up
        Case WM_LBUTTONDBLCLK: RaiseEvent LeftButtonDoubleClick     'Left double click
        Case WM_RBUTTONDOWN: RaiseEvent RightButtonDown             'Right click down
        Case WM_RBUTTONUP: 'RaiseEvent RightButtonUp                'Right click up
        PopupMenu MenuSysPop
        Case WM_RBUTTONDBLCLK: RaiseEvent RightButtonDoubleClick    'Right double click
        Case WM_MBUTTONDOWN: RaiseEvent MiddleButtonDown            'Middle click down
        Case WM_MBUTTONUP: RaiseEvent MiddleButtonUp                'Middle click up
        Case WM_MBUTTONDBLCLK: RaiseEvent MiddleButtonDoubleClick   'Middle double click
    End Select
End Sub

Private Sub Class_Initialize()
    Set modTrayIcon.cTray = Me  ' Sets the variable of the module _
                                 to let it handle this class' events
End Sub

Private Sub ToTray_Click()
Class_Initialize
ShowTrayIcon Me.hWnd, "Connections Monitor (Monitoring)", Me.Icon
Me.Hide
End Sub

Private Sub OpenConMon_click()
Me.Show
RemoveTrayIcon
End Sub

Private Sub ExitProg1_click()
Unload Me
End Sub

'************************************Additions***********************

Function TransparentObject(ByVal hWnd As Long)
'Example:
'Private Sub Form_Load() To make a flat-edged textbox...
'TransparentObject Text1.Hwnd
'End Sub
SetWindowLong hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT
SetWindowPos hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME
End Function
