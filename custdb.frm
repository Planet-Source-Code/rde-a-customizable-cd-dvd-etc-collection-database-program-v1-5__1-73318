VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCustDB 
   Caption         =   " Custom Database"
   ClientHeight    =   5715
   ClientLeft      =   4980
   ClientTop       =   4320
   ClientWidth     =   8955
   ClipControls    =   0   'False
   Icon            =   "custdb.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8955
   Begin VB.TextBox Headers 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   300
      Width           =   1680
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   0
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   8970
      TabIndex        =   14
      Top             =   660
      Width           =   8970
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Selected Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   6660
         TabIndex        =   15
         Top             =   60
         Width           =   2170
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update Selected Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   4500
         TabIndex        =   16
         Top             =   60
         Width           =   2170
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Selected Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   2340
         TabIndex        =   17
         Top             =   60
         Width           =   2170
      End
      Begin VB.CommandButton cmdAddItem 
         Caption         =   "Add &New Item"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   180
         TabIndex        =   18
         Top             =   60
         Width           =   2170
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1845
      Index           =   1
      Left            =   0
      ScaleHeight     =   1845
      ScaleWidth      =   8970
      TabIndex        =   2
      Top             =   4365
      Width           =   8970
      Begin VB.CommandButton cmdCust 
         Caption         =   "&Customise DB"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   6630
         TabIndex        =   5
         Top             =   915
         Width           =   2170
      End
      Begin VB.CommandButton cmdSaveList 
         Caption         =   "Save List to D&B"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   4470
         TabIndex        =   6
         Top             =   915
         Width           =   2170
      End
      Begin VB.CommandButton cmdReload 
         Caption         =   "Reload Saved &List"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   2310
         TabIndex        =   7
         Top             =   915
         Width           =   2170
      End
      Begin VB.CommandButton cmdSaveSel 
         Caption         =   "&Save Selected Item(s)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   6630
         TabIndex        =   4
         Top             =   90
         Width           =   2170
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "E&mail Selected Item(s)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   4470
         TabIndex        =   9
         Top             =   90
         Width           =   2170
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print Selected Item(s)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   2310
         TabIndex        =   10
         Top             =   90
         Width           =   2170
      End
      Begin VB.CommandButton cmdDupes 
         Caption         =   "&Remove Duplicates"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   150
         TabIndex        =   11
         Top             =   90
         Width           =   2170
      End
      Begin RichTextLib.RichTextBox rtf 
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   1380
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   556
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"custdb.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ProgressBar pgbProgress 
         Height          =   180
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "List Keys: ALT-I, Letter keys, Arrow keys, ENTER to edit, DEL to delete, CTRL-DEL to delete duplicates"
         Top             =   585
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   318
         _Version        =   393216
         Appearance      =   0
         Max             =   500
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Entries :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -330
         TabIndex        =   13
         Top             =   975
         Width           =   1785
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   975
         Width           =   795
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   8820
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   120
         X2              =   8820
         Y1              =   780
         Y2              =   780
      End
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   3360
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSMAPI.MAPIMessages MAPIMess 
      Left            =   4590
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISess 
      Left            =   5490
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3510
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "custdb.frx":038A
            Key             =   "ascend"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "custdb.frx":04E4
            Key             =   "descend"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   180
      Left            =   2190
      Top             =   2580
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3105
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   5477
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Artist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   20
      ToolTipText     =   "List Keys: ALT-I, Letter keys, Arrow keys, ENTER to edit, DEL to delete, CTRL-DEL to delete duplicates"
      Top             =   60
      Width           =   1665
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&i"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   4
      Left            =   30
      TabIndex        =   0
      Top             =   1140
      Width           =   555
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "MainMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit This Item"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete This Item"
      End
      Begin VB.Menu mnuDiv 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Selected Item(s)"
      End
      Begin VB.Menu mnuEmail 
         Caption         =   "Email Selected Item(s)"
      End
      Begin VB.Menu mnuSaveSel 
         Caption         =   "Save Selected Item(s)"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCust 
         Caption         =   "Customise Database"
      End
   End
   Begin VB.Menu mnuCustom 
      Caption         =   "CustomMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuCustomDB 
         Caption         =   "&Customise Database"
      End
   End
End
Attribute VB_Name = "frmCustDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Added to Toolbox:
' "Microsoft Windows Common Controls 6.0"
' "Microsoft Common Dialog Control 6.0"
' "Microsoft Rich Textbox Control 6.0"
' "Microsoft MAPI Controls 6.0"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' API which initializes XP Styles on common controls by Amer Tahir
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Long
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Const ICC_USEREX_CLASSES = &H200

' These APIs prevent shutdown crashes
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private m_hMod As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function GetVersionExA Lib "kernel32" (lpVerInfo As OSVERSIONINFO) As Long
Private Const OFS_MAXPATHNAME = 128
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumberLo As Integer
    dwBuildNumberHi As Integer
    dwPlatformId As Long
    szCSDVersion As String * OFS_MAXPATHNAME
End Type
Private osvi As OSVERSIONINFO

'Private Const dwMask3x = &H0&
'Private Const dwMask9x = &H1&
Private Const dwMaskNx = &H2&
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const DB_BUF = 500&
Private db_Max As Long

Private BuildString As cBuildStr
Private cSSorter As cStableSorter
Implements ICompareClient

Private Type tCDB
   Data(9) As String
End Type

'Private Type tCDB
'   Artist As String
'   Album As String
'   Genre As String
'   Year As String
'End Type

Private arrCDB() As tCDB
Private aIdx() As Long
Private lSortFlag As Long
Private lProgress As Long

Private bEditFlag As Boolean
Private lEditIdx As Long
Private mIsWinXP As Boolean

' Is running on Windows XP or higher
Private Function IsWinXP() As Boolean
    osvi.dwOSVersionInfoSize = Len(osvi)
    If (GetVersionExA(osvi)) Then
        If (osvi.dwPlatformId = dwMaskNx) Then
            ' Is Windows XP or higher
            IsWinXP = (osvi.dwMinorVersion > 0&)
    End If: End If
End Function

' This is the main sub where the XP Styles are initialized
Private Sub Form_Initialize() ' Amer Tahir
    On Error GoTo Abort
    Dim iccex As tagInitCommonControlsEx
    mIsWinXP = IsWinXP
    If mIsWinXP Then
        iccex.lngSize = Len(iccex)
        iccex.lngICC = ICC_USEREX_CLASSES
        InitCommonControlsEx iccex
        ' This is to prevent crash
        m_hMod = LoadLibrary("shell32.dll")
    End If
    Exit Sub
Abort:
    On Error Resume Next
    FreeLibrary m_hMod
    mIsWinXP = False
End Sub

Private Sub Form_Load()
   Dim s As String
   If mIsWinXP Then DoCmdResize
   sINIFile = App.Path & "\CustomDB.ini"
   Load frmCustom
   Set BuildString = New cBuildStr
   Set ListView1.ColumnHeaderIcons = ImageList1
   lColFlag = 1
   Set cSSorter = New cStableSorter
   cSSorter.Attach Me
   cSSorter.Order = Ascending
   pgbProgress.Min = 0
   dlgCommon.InitDir = App.Path
   db_Max = DB_BUF
   ReDim arrCDB(0 To db_Max) As tCDB
   Call OpenDB(True)
   Me.Move CSng(GetINIKey(sINIFile, "WindPos", "Left", "4000")), _
           CSng(GetINIKey(sINIFile, "WindPos", "Top", "3500")), _
           CSng(GetINIKey(sINIFile, "WindPos", "Width", "9080")), _
           CSng(GetINIKey(sINIFile, "WindPos", "Height", "6120"))
   LoadSounds ' Requires res file
   PlaySound complete
End Sub

Private Sub Form_Resize()
   On Error GoTo ErrH
   Dim x As Long
   If (Me.WindowState <> vbMinimized) Then
      Picture1(0).Left = (Me.ScaleWidth - Picture1(0).Width) * 0.5
      Picture1(1).Top = Me.ScaleHeight - 1335
      Picture1(1).Left = Picture1(0).Left
      ListView1.Height = Picture1(1).Top - 1270
      ListView1.Width = Me.ScaleWidth - 210
      Labels(0).Width = (Me.ScaleWidth - Labels(0).Left) / lColCnt - 90
      Headers(0).Width = (Me.ScaleWidth - Headers(0).Left) / lColCnt - 90
      For x = 1 To lColCnt - 1
          Labels(x).Width = Labels(0).Width
          Labels(x).Left = x * (Labels(0).Width + 90) + Labels(0).Left
          Headers(x).Width = Headers(0).Width
          Headers(x).Left = x * (Headers(0).Width + 90) + Headers(0).Left
      Next
   End If
ErrH:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If (Me.WindowState = vbNormal) Then
       SetINIKey sINIFile, "WindPos", "Left", CStr(Me.Left)
       SetINIKey sINIFile, "WindPos", "Top", CStr(Me.Top)
       SetINIKey sINIFile, "WindPos", "Width", CStr(Me.Width)
       SetINIKey sINIFile, "WindPos", "Height", CStr(Me.Height)
    End If
    Unload frmCustom
    Timer1.Enabled = False
    Set BuildString = Nothing
    cSSorter.Detach
    Set cSSorter = Nothing
    If mIsWinXP Then FreeLibrary m_hMod
End Sub

Private Sub ICompareClient_Compare(ByVal ThisIdx As Long, ByVal ThanIdx As Long, Result As eiCompare, ByVal Percent As Long, Cancel As Boolean)
    pgbProgress.Value = Percent + lProgress
    Result = StrComp(arrCDB(ThisIdx).Data(lSortFlag - 1), arrCDB(ThanIdx).Data(lSortFlag - 1), vbTextCompare)
   'Cancel = False
End Sub

Private Sub cmdReload_Click()
   Call OpenDB
End Sub

Private Sub cmdAddItem_Click()
   Call AddNewItem
End Sub

Private Sub cmdUpdate_Click()
   Call UpdateItem
End Sub

Private Sub cmdCust_Click()
    mnuCustomDB_Click
End Sub

Private Sub cmdEdit_Click()
   Call EditItem
End Sub

Private Sub mnuEdit_Click()
   Call EditItem
End Sub

Private Sub cmdDelete_Click()
   Call DeleteItem
End Sub

Private Sub cmdSaveList_Click()
   Call SaveList
End Sub

Private Sub cmdDupes_Click()
    Call RemoveDupes
End Sub

Private Sub cmdEmail_Click()
    Call EmailList
End Sub

Private Sub mnuEmail_Click()
    Call EmailList
End Sub

Private Sub cmdPrint_Click()
    Call PrintList
End Sub

Private Sub mnuPrint_Click()
    Call PrintList
End Sub

Private Sub cmdSaveSel_Click()
   Call SaveItems
End Sub

Private Sub mnuSaveSel_Click()
   Call SaveItems
End Sub

Private Sub mnuDelete_Click()
   Call DeleteItem
End Sub

Private Sub ListView1_DblClick()
   Call EditItem
End Sub

Private Sub mnuCust_Click()
    mnuCustomDB_Click
End Sub

Private Sub mnuCustomDB_Click()
    Dim i As Long
    If lColCnt > 2 Then i = lColCnt - 2
    With frmCustom
        .cboCols.ListIndex = i
        .Headers(0).Text = Trim$(Me.Caption)
        For i = 1 To lColCnt
            .Headers(i).Text = ListView1.ColumnHeaders.Item(i).Text
        Next i
        .Show vbModal
    End With
    ' On return
    For i = 1 To lColCnt
        arrCDB(0).Data(i - 1) = ListView1.ColumnHeaders.Item(i).Text
    Next i
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   Dim lIdx As Long
   Dim lDir As Long
   lIdx = ColumnHeader.Index
   lDir = cSSorter.Order
   If (lColFlag = lIdx) Then
      lDir = -(lDir)
      cSSorter.Order = lDir
   Else
      ListView1.ColumnHeaders(lColFlag).Icon = 0
      lColFlag = lIdx
   End If
   If (lDir = -1) Then lDir = 2
   ListView1.ColumnHeaders(lIdx).Icon = lDir
   Call SortArray
   Call FillListView
   If Not (ListView1.SelectedItem Is Nothing) Then
      ListView1.SelectedItem.Selected = False
      Set ListView1.SelectedItem = Nothing
   End If
   lEditIdx = 0
   bEditFlag = False
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If (Button = vbRightButton) Then
      PopupMenu mnuMenu
   End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If (Button = vbRightButton) Then
      PopupMenu mnuCustom
   End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If (Button = vbRightButton) Then
      PopupMenu mnuCustom
   End If
End Sub

Private Sub Labels_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If (Button = vbRightButton) Then
      PopupMenu mnuCustom
   End If
End Sub

Private Sub Headers_GotFocus(Index As Integer)
   Headers(Index).SelStart = 0
   Headers(Index).SelLength = Len(Headers(Index))
End Sub

Private Sub Headers_LostFocus(Index As Integer)
   If (Index = 0) Then bEditFlag = False
End Sub

Private Sub Headers_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim li As ListItem, iSel As Long
    If (Index = 0) And (Not bEditFlag) Then
        If Headers(0) = vbNullString Then Exit Sub
        Set li = ListView1.FindItem(Headers(0), , , lvwPartial)
        If Not (li Is Nothing) Then
           iSel = ListView1.SelectedItem.Index
           If (iSel > -1) Then ListView1.ListItems(iSel).Selected = False
           ListView1.ListItems(li.Index).Selected = True
           EnsureVisible li.Index
        End If
    End If
End Sub

Private Sub Headers_KeyPress(Index As Integer, KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Then
        If (Index < 3) Then
            Headers(Index + 1).SetFocus
        Else
            Call AddNewItem
            Headers(0).SetFocus
        End If
        KeyAscii = 0
    End If
End Sub

Private Sub ListView1_KeyUp(KeyCode As Integer, Shift As Integer)
   If (KeyCode = vbKeyReturn) Then
      Call EditItem
   ElseIf (KeyCode = vbKeyDelete) Then
      If (Shift And vbKeyControl = vbKeyControl) Then
         If MsgBox("This will remove all duplicates. " & _
                   vbCr & "Do you wish to continue?", _
                   vbOKCancel, Me.Caption) = vbCancel Then Exit Sub
         Call RemoveDupes
      Else
         Call DeleteItem
      End If
   End If
End Sub

Private Sub EnsureVisible(ByVal lIdx As Long)
   Dim lVisibleCount As Long
   Dim lTopIdx As Long
   Dim lMid As Long
   lTopIdx = ListView1.GetFirstVisible.Index
   lVisibleCount = LV_GetVisibleCount(ListView1.hWnd)
   lMid = lVisibleCount \ 2
   If (lIdx < lTopIdx) Then
       If Not (lIdx < lMid) Then lTopIdx = lIdx - lMid Else lTopIdx = 1

   ElseIf Not (lIdx < lTopIdx + lVisibleCount) Then
       If Not (lIdx > ListView1.ListItems.Count - lMid) Then
           lTopIdx = lIdx + lMid - 1
       Else
           lTopIdx = ListView1.ListItems.Count
       End If
   End If
   ListView1.ListItems(lTopIdx).EnsureVisible
End Sub

Private Sub AddNewItem()
   On Error GoTo ErrH
   Dim x As Long, y As Long
   For x = 0 To lColCnt - 1
      If LenB(Headers(x)) Then Exit For
   Next
   If (x = lColCnt) Then
      Beep ' Must enter at least one attribute
      Exit Sub
   End If
   lItemCnt = lItemCnt + 1
   For x = 0 To lColCnt - 1
      arrCDB(lItemCnt).Data(x) = Headers(x)
   Next
   Call SortArray
   Call FillListView
   x = ListView1.FindItem("Key" & lItemCnt, lvwTag).Index
   y = ListView1.SelectedItem.Index
   If (y > -1) Then ListView1.ListItems(y).Selected = False
   ListView1.ListItems(x).Selected = True
   EnsureVisible x
   Exit Sub
ErrH:
   If (Err.Number = 9) Then
      db_Max = db_Max + DB_BUF
      ReDim Preserve arrCDB(0 To db_Max) As tCDB
      Resume
   End If
End Sub

Private Sub UpdateItem()
   Dim x As Long, y As Long, lItem As Long
   For x = 0 To lColCnt - 1
      If LenB(Headers(x)) Then Exit For
   Next
   If (x = lColCnt) Then
      Beep ' Must enter at least one attribute
      Exit Sub
   End If
   If Not (ListView1.SelectedItem Is Nothing) Then
      x = ListView1.SelectedItem.Index
      y = CLng(Mid$(ListView1.ListItems(x).Tag, 4))
      lItem = aIdx(x)
      If Not (y = lEditIdx) Then
          If MsgBox("Update '" & arrCDB(lItem).Data(0) & " - " & arrCDB(lItem).Data(1) & "'" & vbCrLf & _
                    "with '" & Headers(0) & " - " & Headers(1) & "' ?", vbYesNo, Me.Caption) = vbNo Then Exit Sub
      End If
      For x = 0 To lColCnt - 1
         arrCDB(lItem).Data(x) = Headers(x)
      Next

      Call SortArray
      Call FillListView
      x = ListView1.FindItem("Key" & lItem, lvwTag).Index
      y = ListView1.SelectedItem.Index
      If (y > -1) Then ListView1.ListItems(y).Selected = False
      ListView1.ListItems(x).Selected = True
      EnsureVisible x

      ListView1.SetFocus
   End If
End Sub

Private Sub EditItem()
   Dim x As Long, y As Long
   If Not (ListView1.SelectedItem Is Nothing) Then
      x = ListView1.SelectedItem.Index
      lEditIdx = aIdx(x)
      For y = 0 To lColCnt - 1
         Headers(y) = arrCDB(lEditIdx).Data(y)
      Next
'      Headers(0) = arrCDB(aIdx(x)).Artist
'      Headers(1) = arrCDB(aIdx(x)).Album
'      Headers(2) = arrCDB(aIdx(x)).Genre
'      Headers(3) = arrCDB(aIdx(x)).Year
      Headers(0).SetFocus
      bEditFlag = True
   End If
End Sub

Private Sub DeleteItem()
   Dim li As ListItem
   Dim x As Long, y As Long
   Dim s As String
   If Not (ListView1.SelectedItem Is Nothing) Then
      x = ListView1.SelectedItem.Index
      s = ListView1.ListItems(x).Text
      If MsgBox("Delete '" & arrCDB(aIdx(x)).Data(0) & " - " & _
                             arrCDB(aIdx(x)).Data(1) & "' ?", _
                             vbOKCancel, Me.Caption) = vbCancel Then Exit Sub
      If (aIdx(x) <> lItemCnt) Then
          For y = 0 To lColCnt - 1
             arrCDB(aIdx(x)).Data(y) = arrCDB(lItemCnt).Data(y)
          Next
      End If
      lItemCnt = lItemCnt - 1
      Call SortArray
      Call FillListView
      x = Len(s)
      Do
        Set li = ListView1.FindItem(Left$(s, x), , , lvwPartial)
        If Not (li Is Nothing) Then
           ListView1.ListItems(li.Index).Selected = True
           EnsureVisible li.Index
           Exit Do
        ElseIf (x = 1) Then Exit Do
        End If
        x = x - 1
      Loop
      ListView1.SetFocus
   End If
   lEditIdx = 0
   bEditFlag = False
End Sub

Private Sub RemoveDupes()
   Dim x As Long
   Dim y As Long
   Dim z As Long
   If (lItemCnt > 0) Then
      On Error Resume Next
      x = 1
      Do Until (x > lItemCnt)
         If (pgbProgress.Max <> lItemCnt) Then pgbProgress.Max = lItemCnt
         For y = x + 1 To lItemCnt
             For z = 0 To lColCnt - 1
                If (arrCDB(x).Data(z) <> arrCDB(y).Data(z)) Then Exit For
             Next
             If (z = lColCnt) Then
                If (y < lItemCnt) Then
                    For z = 0 To lColCnt - 1
                       arrCDB(y).Data(z) = arrCDB(lItemCnt).Data(z)
                    Next
'                   arrCDB(y).Artist = arrCDB(lItemCnt).Artist
'                   arrCDB(y).Album = arrCDB(lItemCnt).Album
'                   arrCDB(y).Genre = arrCDB(lItemCnt).Genre
'                   arrCDB(y).Year = arrCDB(lItemCnt).Year
                End If
                lItemCnt = lItemCnt - 1
             End If
         Next y
         pgbProgress.Value = x
         x = x + 1
      Loop
      Call SortArray
      Call FillListView
      If Not (ListView1.SelectedItem Is Nothing) Then
         ListView1.SelectedItem.Selected = False
         Set ListView1.SelectedItem = Nothing
      End If
      lEditIdx = 0
      bEditFlag = False
      Timer1.Enabled = True
   End If
End Sub

Private Sub SaveList()
   Dim sFile As String
   Dim x As Long, y As Long
   If lItemCnt = 0 Then Beep: Exit Sub
   sFile = SaveDialog
   If (LenB(sFile) = 0) Then Exit Sub
   On Error Resume Next
   pgbProgress.Max = lItemCnt
   BuildString.Reset
   BuildString.Appends Trim$(Me.Caption), vbCrLf
   For x = 0 To lItemCnt
      For y = 0 To lColCnt - 2
          BuildString.Appends arrCDB(aIdx(x)).Data(y), vbTab
      Next
      BuildString.Appends arrCDB(aIdx(x)).Data(y), vbCrLf

      pgbProgress.Value = x
   Next x
   Open sFile For Output As #1
     Print #1, BuildString.Value;
   Close #1
   BuildString.Reset
   Timer1.Enabled = True
End Sub

Private Sub OpenDB(Optional ByVal OnLoad As Boolean)
    Dim sFile As String
    sFile = OpenDialog
    If (LenB(sFile) = 0) Then
        If OnLoad Then Call mnuCustomDB_Click
    Else
      If FillArray(sFile) Then
          Call SortArray
          Call ListViewHeaders
          Call FillListView
      Else
          MsgBox "An error occured loading your database!", vbExclamation
      End If
    End If
End Sub

Private Sub SortArray()
    If (lItemCnt > 1) Then
        ReDim aIdx(0 To lItemCnt)
        With cSSorter
            If (lColFlag > 2) Then
               pgbProgress.Max = 300
            Else
               pgbProgress.Max = 200
            End If
            lProgress = 0
            If (lColFlag <> 2) Then
               lSortFlag = 2
               .Sort aIdx, 1, lItemCnt ' Sort col 2
               lProgress = 100
            End If
            lSortFlag = 1
            .Sort aIdx, 1, lItemCnt    ' Sort col 1
            lProgress = lProgress + 100
            If (lColFlag > 1) Then
               lSortFlag = lColFlag
               .Sort aIdx, 1, lItemCnt ' Sort col x
            End If
            Timer1.Enabled = True
        End With
    ElseIf (lItemCnt > 0) Then ' Single Item
        ReDim aIdx(0 To 1)
       'aIdx(0) = 0 ' Header captions
        aIdx(1) = 1
    End If
End Sub

Public Sub FillListView()
    Dim li As ListItem
    Dim s As String
    Dim x As Long, y As Long
    With ListView1
        .ListItems.Clear
        If (lItemCnt > 0) Then
             For x = 1 To lItemCnt
                  Set li = .ListItems.Add(x, , arrCDB(aIdx(x)).Data(0))
                  li.Tag = "Key" & aIdx(x) ' Actual array idx
                  For y = 1 To lColCnt - 1
                      li.SubItems(y) = arrCDB(aIdx(x)).Data(y)
                  Next y
             Next
             s = .ColumnHeaders.Item(lColFlag).Text
             .ColumnHeaders.Item(lColFlag).Text = Space$(10) & s
             For x = 1 To .ColumnHeaders.Count
                 LV_SetColumnWidth .hWnd, x
             Next x
             .ColumnHeaders.Item(lColFlag).Text = s
        End If
    End With
    Label2 = lItemCnt
End Sub

Public Sub ListViewHeaders()
    On Error GoTo ErrHandler
    Dim s As String
    Dim x As Long, y As Long
    With ListView1
        .ColumnHeaders.Clear
        x = Labels.Count - 1
        Do While x
            Unload Headers(x)
            Unload Labels(x)
            x = x - 1
        Loop
        s = arrCDB(0).Data(0)
        .ColumnHeaders.Add , , s, 1000
        Labels(0).Caption = "&" & s
        Labels(0).Width = (Me.ScaleWidth - Labels(0).Left) / lColCnt - 90
        Headers(0).Text = ""
        Headers(0).Width = (Me.ScaleWidth - Headers(0).Left) / lColCnt - 90
        For x = 1 To lColCnt - 1
            s = arrCDB(0).Data(x)
            .ColumnHeaders.Add , , s, 1000
            Load Labels(x)
            Labels(x).Caption = "&" & s
            Labels(x).Width = Labels(0).Width
            Labels(x).Left = x * (Labels(0).Width + 90) + Labels(0).Left
            Labels(x).Visible = True
            Load Headers(x)
            Headers(x).Width = Headers(0).Width
            Headers(x).Left = x * (Headers(0).Width + 90) + Headers(0).Left
            Headers(x).Visible = True
        Next
        y = cSSorter.Order
        If (y = -1) Then y = 2
        If (lColFlag > lColCnt) Then lColFlag = lColCnt
        .ColumnHeaders(lColFlag).Icon = y
    End With
ErrHandler:
End Sub

Private Function FillArray(sFile As String) As Boolean
    On Error GoTo ErrH
    Dim sLine As String
    Dim iNum As Long
    Dim cCols As Long
    Dim x As Long, y As Long, z As Long
    iNum = FreeFile
    Open sFile For Input As #iNum
      If Not EOF(iNum) Then
           Line Input #iNum, sLine
           If LenB(sLine) Then
              If InStr(1, sLine, vbTab) Then ' Compatability with 1.0
                  Me.Caption = " " & GetINIKey(sINIFile, sFile, "Caption", "Custom Database")
                  Close #iNum
                  iNum = FreeFile
                  Open sFile For Input As #iNum
              Else
                  Me.Caption = " " & sLine   ' Version 1.5
              End If
           End If
      End If
      lItemCnt = -1  ' 0 = Header captions
      Do Until EOF(iNum)
           Line Input #iNum, sLine
           If (LenB(sLine) <> 0) Then
                lItemCnt = lItemCnt + 1
                x = 1: z = 0
                y = InStr(x, sLine, vbTab)
                Do While (y > 0)
                   arrCDB(lItemCnt).Data(z) = Mid$(sLine, x, y - x)
                   z = z + 1
                   x = y + 1
                   y = InStr(x, sLine, vbTab)
                Loop
                arrCDB(lItemCnt).Data(z) = Mid$(sLine, x)
                z = z + 1
                If (z > cCols) Then cCols = z
                Do While (z < lColCnt) ' Clear unused cols
                   arrCDB(lItemCnt).Data(z) = vbNullString
                   z = z + 1
                Loop
           End If
      Loop
      lColCnt = cCols
      FillArray = True
    Close #iNum
    Exit Function
ErrH:
    If (Err.Number = 9) Then
       db_Max = db_Max + DB_BUF
       ReDim Preserve arrCDB(0 To db_Max) As tCDB
       Resume
    End If
    Close #iNum
End Function

Private Sub Timer1_Timer()
   'Used to delay resetting the progbar on completion
   Timer1.Enabled = False
   pgbProgress.Value = 0
End Sub

'FILE ROUTINES

Private Function OpenDialog() As String
    ' Returns a valid file name, or ""
    On Error GoTo ErrH
    With dlgCommon
        .DialogTitle = " Open CDB file..."
        .Filter = "Custom Database (*.cdb)|*.cdb"
        .FileName = "custdb.cdb"
        .DefaultExt = "cdb"
        ' File must exist, use Win95 style dialog, hide read-only checkbox ' , allow multi-select
        .flags = cdlOFNFileMustExist + cdlOFNExplorer + cdlOFNHideReadOnly ' + cdlOFNAllowMultiselect
        .CancelError = True
        On Error GoTo Canceled
        .ShowOpen
Canceled:
        If (Err = cdlCancel) Then Exit Function
        On Error GoTo ErrH
        OpenDialog = .FileName
    End With
ErrH:
End Function

Private Function SaveDialog(Optional sName As String) As String
    On Error GoTo ErrH
    With dlgCommon
        .DialogTitle = " Save CDB file As..."
        .Filter = "Custom Database (*.cdb)|*.cdb"
        If LenB(sName) Then .FileName = sName
        .DefaultExt = "cdb"
        ' Prompt user if file exists, require valid path, hide read-only checkbox
        .flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNHideReadOnly
        .CancelError = True
        On Error GoTo Canceled
        .ShowSave
Canceled:
        If (Err = cdlCancel) Then Exit Function
        On Error GoTo 0
        On Error GoTo ErrH
        .InitDir = CurDir$
        SaveDialog = .FileName
    End With
ErrH:
End Function

Private Function BuildSelectedItems() As Long
    Dim x As Long, y As Long
    With ListView1
       For x = 1 To .ListItems.Count
          If .ListItems(x).Selected Then
             BuildString.Append .ListItems(x).Text
             For y = 1 To .ListItems(x).ListSubItems.Count 'lColCnt-1
                 BuildString.Appends vbTab, .ListItems(x).SubItems(y)
             Next y
             BuildString.Append vbCrLf
             BuildSelectedItems = BuildSelectedItems + 1
          End If
       Next x
    End With
End Function

Private Sub SaveItems()
    On Error GoTo ErrH
    Dim sFile As String
    sFile = SaveDialog("custdb0.cdb")
    If (LenB(sFile) = 0) Then Exit Sub
    BuildString.Reset
    If (BuildSelectedItems <> 0) Then
       Open sFile For Output As #1
          Print #1, BuildString.Value;
       Close #1
       BuildString.Reset
    End If
ErrH:
End Sub

'CONTACT ROUTINES

Private Sub PrintList()
    On Error GoTo Abort
    If FillRTF Then
        With dlgCommon
            .flags = cdlPDReturnDC + cdlPDAllPages
            .CancelError = True
            On Error GoTo Canceled
            .ShowPrinter
Canceled:
            If (Err = cdlCancel) Then Exit Sub
            On Error GoTo 0
            On Error GoTo Abort
    
            rtf.SelPrint .hDC
        End With
    End If
Abort:
End Sub

Private Function FillRTF() As Boolean
    On Error GoTo Abort
    rtf.Text = vbNullString
    BuildString.Reset
    If (BuildSelectedItems <> 0) Then
        FillRTF = True
        rtf.Text = BuildString.Value
        BuildString.Reset
    End If
Abort:
End Function

Private Sub EmailList()
    On Error GoTo Abort
    BuildString.Reset
    If (BuildSelectedItems = 0) Then Exit Sub
    If (SendEmail = False) Then
       Clipboard.Clear
       Clipboard.SetText BuildString.Value
       BuildString.Reset
       BuildString.Appends "The list has been copied to the clipboard.  ", vbCr, _
                           "Simply paste it into your email program."
       MsgBox BuildString.Value, vbInformation, Me.Caption
    End If
    BuildString.Reset
Abort:
End Sub

Private Function SendEmail() As Boolean
    On Error GoTo Abort
    Me.Visible = False
    MAPISess.SignOn
    MAPIMess.SessionID = MAPISess.SessionID
    MAPIMess.Compose
    MAPIMess.MsgSubject = "Thank you for your enquiry"
    MAPIMess.MsgNoteText = BuildString.Value
    MAPIMess.RecipIndex = 0           ' Add recipient
    MAPIMess.RecipType = mapToList    ' Recipient in TO line
    MAPIMess.RecipAddress = "client"  ' E-mail address
    MAPIMess.AddressResolveUI = True
    MAPIMess.ResolveName
    MAPIMess.Send True
    MAPISess.SignOff
    SendEmail = True
Abort:
'    ' If user cancelled return True
'    If (Err = mapUserAbort) Then SendEmail = True
    Me.Visible = True
End Function


Private Sub DoCmdResize()
    cmdAddItem.Width = 2110
    cmdAddItem.BackColor = &HFFFFFF
    
    cmdEdit.Width = 2110
    cmdEdit.Left = cmdEdit.Left + 20
    cmdEdit.BackColor = &HFFFFFF
    
    cmdUpdate.Width = 2110
    cmdUpdate.Left = cmdUpdate.Left + 40
    cmdUpdate.BackColor = &HFFFFFF
    
    cmdDelete.Width = 2110
    cmdDelete.Left = cmdDelete.Left + 60
    cmdDelete.BackColor = &HFFFFFF

    cmdDupes.Width = 2110
    cmdDupes.BackColor = &HFFFFFF
    
    cmdPrint.Width = 2110
    cmdPrint.Left = cmdPrint.Left + 20
    cmdPrint.BackColor = &HFFFFFF
    
    cmdEmail.Width = 2110
    cmdEmail.Left = cmdEmail.Left + 40
    cmdEmail.BackColor = &HFFFFFF
    
    cmdSaveSel.Width = 2110
    cmdSaveSel.Left = cmdSaveSel.Left + 60
    cmdSaveSel.BackColor = &HFFFFFF

    cmdReload.Width = 2110
    cmdReload.BackColor = &HFFFFFF
    
    cmdSaveList.Width = 2110
    cmdSaveList.Left = cmdSaveList.Left + 20
    cmdSaveList.BackColor = &HFFFFFF
    
    cmdCust.Width = 2110
    cmdCust.Left = cmdCust.Left + 40
    cmdCust.BackColor = &HFFFFFF
End Sub

