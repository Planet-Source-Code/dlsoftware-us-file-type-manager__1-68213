VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   Caption         =   "File Types Manager"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8385
   Icon            =   "fMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Scan for file errors"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3240
      Width           =   8175
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   6360
      TabIndex        =   6
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3360
      TabIndex        =   5
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   2040
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.PictureBox pSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7920
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7440
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6960
      Top             =   840
   End
   Begin MSComctlLib.ListView lvwFile 
      Height          =   1650
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2910
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileExt"
         Text            =   "File Extension"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "FileType"
         Text            =   "File Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "FileErrors"
         Text            =   "File Error(s)"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFile2 
      Height          =   1650
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3840
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   2910
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Action"
         Text            =   "Action(s)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Command"
         Text            =   "Action Command"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   5640
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   8175
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   5700
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Default Icon File"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Default Action"
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Context Type MIME"
      Height          =   255
      Left            =   3360
      TabIndex        =   12
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Registry Key Value"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "File Type"
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "File Extension"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Private Const HKEY_CLASSES_ROOT As Long = &H80000000
    Private Const GOOD_RETURN_CODE As Long = 0
    Private Const STARTS_WITH_A_PERIOD As Long = 46
    Private Const MAX_PATH_LENGTH As Long = 260
    Private Const REG_SZ = (1)
    Private Const REG_EXPAND_SZ = (2)
    Private Const ILD_TRANSPARENT = &H1
    Private Const STANDARD_RIGHTS_READ As Long = &H20000
    Private Const KEY_QUERY_VALUE As Long = &H1
    Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
    Private Const KEY_NOTIFY As Long = &H10
    Private Const SYNCHRONIZE As Long = &H100000
    Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    Private Const SH_USEFILEATTRIBUTES As Long = &H10
    Private Const SH_TYPENAME As Long = &H400
    Private Const SH_DISPLAYNAME = &H200
    Private Const SH_EXETYPE = &H2000
    Private Const SH_SYSICONINDEX = &H4000
    Private Const SH_LARGEICON = &H0
    Private Const SH_SMALLICON = &H1
    Private Const SH_SHELLICONSIZE = &H4
    Private Const FILE_ATTRIBUTE_NORMAL = &H80
    Private Const BASIC_SH_FLAGS = SH_TYPENAME Or SH_SHELLICONSIZE Or SH_SYSICONINDEX Or SH_DISPLAYNAME Or SH_EXETYPE
    Private Type FILETIME: dwLowDateTime As Long: dwHighDateTime As Long: End Type
    Private Type SHFILEINFO: hIcon As Long: iIcon As Long: dwAttributes As Long: szDisplayName As String * MAX_PATH_LENGTH: szTypeName As String * 80: End Type
    Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
    Private Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    Private Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As Long, lpcbData As Long) As Long
    Private Declare Function RegEnumKeyEx Lib "advapi32" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long
    Private Declare Function SHGetFileInfo Lib "shell32" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
    Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
    Private shinfo          As SHFILEINFO
    Private lvi             As ListItem
    Private iSmall          As ListImage
    Private ftime           As FILETIME
    Private iPos            As Integer
    Private lSortCol        As Long
    Private lIcon           As Long
    Private lBarValue       As Long
    Private lCount          As Long
    Private lErrors         As Long
    Private lResult         As Long
    Private lResultEnumKey  As Long
    Private lrc0            As Long
    Private lrc1            As Long
    Private lrc2            As Long
    Private lrc3            As Long
    Private lrc4            As Long
    Private lcch            As Long
    Private lIndex          As Long
    Private lType           As Long
    Private lRegKeyIndex    As Long
    Private sActionValue    As String
    Private sActionCommand  As String
    Private sActionKey      As String
    Private sAction         As String
    Private sTitle          As String
    Private sValue          As String
    Private sKey            As String
    Private sAlreadyAdded   As String
    Private sImageList1Key  As String
    Private sFileTypeName   As String
    Private sFileExtension  As String
    Private sRegSubkey      As String * MAX_PATH_LENGTH
    Private sRegKeyClass    As String * MAX_PATH_LENGTH
    Private sSource         As String
    Private vValue          As Variant
    
Private Sub Command1_Click()
    On Error Resume Next
    Unload fMain
    Set fMain = Nothing
End Sub

Private Sub Command2_Click()
    On Error Resume Next
    Screen.MousePointer = vbArrowHourglass
    pBar1.Value = 0
    Label8.Caption = " "
    pBar1.Visible = True
    Label8.Visible = True
    LockControl lvwFile, True
    lCount = 0
    lErrors = 0
    For Each lvi In lvwFile.ListItems
        sAlreadyAdded = "N"
        lCount = lCount + 1
        lBarValue = ((lCount / 700) * 100)
        If lBarValue > 100 Then
           lBarValue = 100
        End If
        pBar1.Value = lBarValue
        Label8 = lBarValue & " %"
        Label8.Refresh
        sFileExtension = "." & lvi.Text
        lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sFileExtension, 0, KEY_READ, lResult)
        If lrc1 = 0 Then
           lrc2 = QueryValueEx(lResult, vbNullString, vValue)
           TrimValue
           sKey = sValue
        Else
           sKey = sFileExtension
        End If
        RegCloseKey HKEY_CLASSES_ROOT
        lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKey & "\DefaultIcon", 0, KEY_READ, lResult)
        If lrc1 = 0 Then
           lrc2 = QueryValueEx(lResult, vbNullString, vValue)
           TrimValue
           sSource = sValue
        Else
           sSource = ""
        End If
        RegCloseKey HKEY_CLASSES_ROOT
        If sSource <> "" Then
           If Not FileExistsDefaultIconFile(sSource) Then
              lErrors = lErrors + 1
              lvi.SubItems(2) = "File Error(s)"
              lrc4 = ModifyForeColor(lvwFile, lCount, vbRed)
              sAlreadyAdded = "Y"
           End If
        End If
        lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKey & "\shell", 0, KEY_READ, lResult)
        lRegKeyIndex = 0
        Do While RegEnumKeyEx(lResult, lRegKeyIndex, sRegSubkey, MAX_PATH_LENGTH, 0, sRegKeyClass, MAX_PATH_LENGTH, ftime) = GOOD_RETURN_CODE
           sAction = TrimNull(sRegSubkey)
           sActionKey = sKey & "\shell\" & sAction
           lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sActionKey, 0, KEY_READ, lResultEnumKey)
           If lrc1 = 0 Then
              lrc2 = QueryValueEx(lResultEnumKey, vbNullString, vValue)
              TrimValue
              sActionValue = sValue
           Else
              sActionValue = ""
           End If
           If sActionValue = "" Then
              sActionValue = sAction
           End If
           sActionKey = sKey & "\shell\" & sAction & "\command"
           lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sActionKey, 0, KEY_READ, lResultEnumKey)
           If lrc1 = 0 Then
              lrc2 = QueryValueEx(lResultEnumKey, vbNullString, vValue)
              TrimValue
              sActionCommand = sValue
           Else
              sActionCommand = ""
           End If
           If sActionCommand <> "" Then
              If Not FileExistsActionFile(sActionCommand) Then
                 If sAlreadyAdded = "N" Then
                    lErrors = lErrors + 1
                 End If
                 lvi.SubItems(2) = "File Error(s)"
                 lrc4 = ModifyForeColor(lvwFile, lCount, vbRed)
                 sAlreadyAdded = "Y"
              End If
           End If
          lRegKeyIndex = lRegKeyIndex + 1
    Loop
    RegCloseKey HKEY_CLASSES_ROOT
    Next lvi
    Screen.MousePointer = vbDefault
    sTitle = "File Types Manager   -   " & CStr(lErrors) & " File Error(s)"
    fMain.Caption = sTitle
    LockControl lvwFile, False
    pBar1.Visible = False
    Label8.Visible = False
    lvwFile.SetFocus
    SortlvwFile
    Call lvwFile_ItemClick(lvwFile.ListItems(1))
    Set lvwFile.SelectedItem = lvwFile.ListItems(1)
    DoEvents
End Sub

Private Sub Form_Load()
    On Error Resume Next
    With lvwFile
        .SmallIcons = ImageList1
        .ColumnHeaders("FileExt").Width = .Width * 0.2
        .ColumnHeaders("FileType").Width = .Width * 0.6
        .ColumnHeaders("FileError(s)").Width = .Width * 0.2
    End With
End Sub

Private Sub lvwFile_ColumnClick(ByVal ColumnHeader As ColumnHeader)
    On Error Resume Next
    lSortCol = ColumnHeader.Index - 1
    With lvwFile
        If .SortKey = lSortCol Then
            If .SortOrder = lvwAscending Then
               .SortOrder = lvwDescending
            Else
               .SortOrder = lvwAscending
            End If
        Else
            .SortKey = lSortCol
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
End Sub

Private Sub SortlvwFile()
    On Error Resume Next
    lSortCol = 2
    With lvwFile
        .SortKey = lSortCol
        .SortOrder = lvwDescending
        .Sorted = True
    End With
End Sub

Private Sub TrimValue()
    On Error Resume Next
    If vValue = "" Then
       sValue = ""
    Else
       sValue = vValue
       sValue = TrimNull(sValue)
    End If
End Sub

Private Sub GetValues()
    On Error Resume Next
    sFileExtension = "." & sFileExtension
    Text1.Text = sFileExtension
    Text6.Text = sFileTypeName
    lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sFileExtension, 0, KEY_READ, lResult)
    If lrc1 = 0 Then
       lrc2 = QueryValueEx(lResult, vbNullString, vValue)
       TrimValue
       Text2.Text = sValue
       sKey = sValue
       lrc2 = QueryValueEx(lResult, "content type", vValue)
       TrimValue
       Text3.Text = sValue
    Else
       Text2.Text = ""
       Text3.Text = ""
    End If
    RegCloseKey HKEY_CLASSES_ROOT
    If sKey = "" Then
       sKey = sFileExtension
    End If
    lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKey & "\DefaultIcon", 0, KEY_READ, lResult)
    If lrc1 = 0 Then
       lrc2 = QueryValueEx(lResult, vbNullString, vValue)
       TrimValue
       Text4.Text = sValue
    Else
       Text4.Text = ""
    End If
    RegCloseKey HKEY_CLASSES_ROOT
    sSource = Text4.Text
    If sSource <> "" Then
       If Not FileExistsDefaultIconFile(sSource) Then
          Text4.ForeColor = vbRed
       Else
          Text4.ForeColor = vbBlack
       End If
    End If
    With lvwFile2
        .ColumnHeaders("Action").Width = .Width * 0.25
        .ColumnHeaders("Command").Width = .Width * 1.3
    End With
    lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKey & "\shell", 0, KEY_READ, lResult)
    If lrc1 = 0 Then
       lrc2 = QueryValueEx(lResult, vbNullString, vValue)
       TrimValue
       Text5.Text = sValue
    Else
       Text5.Text = ""
    End If
    lRegKeyIndex = 0
    Do While RegEnumKeyEx(lResult, lRegKeyIndex, sRegSubkey, MAX_PATH_LENGTH, 0, sRegKeyClass, MAX_PATH_LENGTH, ftime) = GOOD_RETURN_CODE
       sAction = TrimNull(sRegSubkey)
       sActionKey = sKey & "\shell\" & sAction
       lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sActionKey, 0, KEY_READ, lResultEnumKey)
       If lrc1 = 0 Then
          lrc2 = QueryValueEx(lResultEnumKey, vbNullString, vValue)
          TrimValue
          sActionValue = sValue
       Else
          sActionValue = ""
       End If
       If sActionValue = "" Then
          sActionValue = sAction
       End If
       sActionKey = sKey & "\shell\" & sAction & "\command"
       lrc1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, sActionKey, 0, KEY_READ, lResultEnumKey)
       If lrc1 = 0 Then
          lrc2 = QueryValueEx(lResultEnumKey, vbNullString, vValue)
          TrimValue
          sActionCommand = sValue
       Else
          sActionCommand = ""
       End If
       lIcon = SHGetFileInfo(Text1.Text, FILE_ATTRIBUTE_NORMAL, shinfo, Len(shinfo), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_SMALLICON)
       sFileExtension = Text1.Text
       sFileExtension = Right(sFileExtension, Len(sFileExtension) - 1)
       pSmall.Picture = LoadPicture()
       Call ImageList_Draw(lIcon, shinfo.iIcon, pSmall.hDC, 0, 0, ILD_TRANSPARENT)
       pSmall.Picture = pSmall.Image
       sImageList1Key = "#" & sFileExtension & "#"
       Set iSmall = ImageList1.ListImages.Add(, sImageList1Key, pSmall.Picture)
       Set lvi = lvwFile2.ListItems.Add(, , sActionValue)
           lvi.SmallIcon = ImageList1.ListImages(sImageList1Key).Key
           lvi.SubItems(1) = sActionCommand
       lRegKeyIndex = lRegKeyIndex + 1
    Loop
    RegCloseKey HKEY_CLASSES_ROOT
    lIndex = 0
    For Each lvi In lvwFile2.ListItems
        sSource = lvi.SubItems(1)
        lIndex = lIndex + 1
        If Not FileExistsActionFile(sSource) Then
           lrc4 = ModifyForeColor(lvwFile2, lIndex, vbRed)
        End If
    Next lvi
    lvwFile.SetFocus
    Refresh
End Sub

Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
      On Error GoTo QueryValueExError
      lrc3 = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, lcch)
      If lrc3 <> 0 Then Error 5
      Select Case lType
             Case REG_SZ, REG_EXPAND_SZ:
                  sValue = String(lcch, 0)
                  lrc0 = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, lcch)
                  If lrc0 = 0 Then
                     vValue = Left(sValue, lcch)
                  Else
                     vValue = Empty
                  End If
             Case Else
                  vValue = Empty
                  lrc0 = -1
      End Select
QueryValueExExit:
    QueryValueEx = lrc0
    Exit Function
QueryValueExError:
      vValue = Empty
      Resume QueryValueExExit
End Function

Private Sub lvwFile_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    sFileExtension = Item.Text
    sFileTypeName = lvwFile.SelectedItem.SubItems(1)
    lvwFile2.ListItems.Clear
    GetValues
End Sub

Private Sub Timer1_Timer()
    On Error Resume Next
    Timer1.Enabled = False
    Screen.MousePointer = vbArrowHourglass
    LockControl lvwFile, True
    pBar1.Value = 0
    Do While RegEnumKeyEx(HKEY_CLASSES_ROOT, lRegKeyIndex, sRegSubkey, MAX_PATH_LENGTH, 0, sRegKeyClass, MAX_PATH_LENGTH, ftime) = GOOD_RETURN_CODE
       If Asc(sRegSubkey) = STARTS_WITH_A_PERIOD Then
          lCount = lCount + 1
          lBarValue = ((lCount / 700) * 100)
          If lBarValue > 100 Then
             lBarValue = 100
          End If
          pBar1.Value = lBarValue
          Label8 = lBarValue & " %"
          Label8.Refresh
          lIcon = SHGetFileInfo(sRegSubkey, FILE_ATTRIBUTE_NORMAL, shinfo, Len(shinfo), SH_USEFILEATTRIBUTES Or BASIC_SH_FLAGS Or SH_SMALLICON)
          sFileTypeName = TrimNull(shinfo.szTypeName)
          sFileExtension = TrimNull(sRegSubkey)
          sFileExtension = Right(sFileExtension, Len(sFileExtension) - 1)
          pSmall.Picture = LoadPicture()
          Call ImageList_Draw(lIcon, shinfo.iIcon, pSmall.hDC, 0, 0, ILD_TRANSPARENT)
          pSmall.Picture = pSmall.Image
          sImageList1Key = "#" & sFileExtension & "#"
          Set iSmall = ImageList1.ListImages.Add(, sImageList1Key, pSmall.Picture)
          Set lvi = lvwFile.ListItems.Add(, , sFileExtension)
                    lvi.SmallIcon = ImageList1.ListImages(sImageList1Key).Key
                    lvi.SubItems(1) = sFileTypeName
                    lvi.SubItems(2) = " "
       End If
       lRegKeyIndex = lRegKeyIndex + 1
    Loop
    sTitle = "File Types Manager   -   " & CStr(lCount) & " Registered File Extensions"
    fMain.Caption = sTitle
    Screen.MousePointer = vbDefault
    LockControl lvwFile, False
    lvwFile.SetFocus
    Call lvwFile_ItemClick(lvwFile.ListItems(1))
    pBar1.Visible = False
    Label8.Visible = False
    DoEvents
End Sub

Private Function TrimNull(StartStr As String) As String
    On Error Resume Next
    iPos = InStr(StartStr, Chr$(0))
    If iPos Then
       TrimNull = Left$(StartStr, iPos - 1)
       Exit Function
    End If
    TrimNull = StartStr
End Function

Private Function FileExistsDefaultIconFile(FullFileName As String) As Boolean
    On Error GoTo NotFound
    lrc4 = InStrRev(FullFileName, ",")
    If lrc4 > 0 Then
       FullFileName = Left(FullFileName, lrc4 - 1)
    End If
    FullFileName = Replace(FullFileName, Chr(34), " ")
    FullFileName = Trim(FullFileName)
    lrc4 = InStrRev(FullFileName, ":\")
    If lrc4 < 4 And lrc4 <> 0 Then
        Open FullFileName For Input As #1
        Close #1
        FileExistsDefaultIconFile = True
    Else
        FileExistsDefaultIconFile = True
    End If
    Exit Function
NotFound:
        FileExistsDefaultIconFile = False
    Exit Function
End Function

Private Function FileExistsActionFile(FullFileName As String) As Boolean
    On Error GoTo NotFound
    FullFileName = Replace(FullFileName, Chr(34), " ")
    lrc4 = InStr(FullFileName, "-")
    If lrc4 > 0 Then
       FullFileName = Left(FullFileName, lrc4 - 1)
    End If
    lrc4 = InStr(FullFileName, "/")
    If lrc4 > 0 Then
       FullFileName = Left(FullFileName, lrc4 - 1)
    End If
    lrc4 = InStr(FullFileName, "%")
    If lrc4 > 0 Then
       FullFileName = Left(FullFileName, lrc4 - 1)
    End If
    lrc4 = InStr(FullFileName, "Macro=")
    If lrc4 > 0 Then
       FullFileName = Left(FullFileName, lrc4 - 1)
    End If
    lrc4 = InStr(FullFileName, "QueueFile=")
    If lrc4 > 0 Then
       FullFileName = Left(FullFileName, lrc4 - 1)
    End If
    FullFileName = Trim(FullFileName)
    lrc4 = InStrRev(FullFileName, ":\")
    If lrc4 < 4 And lrc4 <> 0 Then
        Open FullFileName For Input As #1
        Close #1
        FileExistsActionFile = True
    Else
        FileExistsActionFile = True
    End If
    Exit Function
NotFound:
        FileExistsActionFile = False
    Exit Function
End Function

Private Sub LockControl(objX As Object, bLock As Boolean)
    On Error Resume Next
    If bLock Then
        LockWindowUpdate objX.hwnd
    Else
        LockWindowUpdate 0
        objX.Refresh
    End If
End Sub

Private Function ModifyForeColor(lvwListView As ListView, lngindex As Long, strForeColor As String) As Long
    On Error GoTo err_ModifyForeColor
    With lvwListView
        If .ListItems.Count < lngindex Then
            ModifyForeColor = 1
            Exit Function
        End If
        With .ListItems.Item(lngindex)
             .ForeColor = strForeColor
        End With
        If .ColumnHeaders.Count < 1 Then
            ModifyForeColor = 2
            Exit Function
        End If
        For iPos = 1 To .ColumnHeaders.Count - 1
            With .ListItems.Item(lngindex).ListSubItems.Item(iPos)
                .ForeColor = strForeColor
            End With
        Next iPos
    End With
    ModifyForeColor = 0
    Exit Function
err_ModifyForeColor:
    ModifyForeColor = 3
    Err.Clear
End Function
