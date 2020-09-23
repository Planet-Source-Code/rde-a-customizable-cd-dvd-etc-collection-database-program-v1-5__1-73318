Attribute VB_Name = "modLV"
Option Explicit

Private Const LVM_FIRST = &H1000&
Private Const LVM_GETITEMCOUNT = LVM_FIRST + 4&
Private Const LVM_GETNEXTITEM = LVM_FIRST + 12&
Private Const LVM_GETSTRINGWIDTHA = LVM_FIRST + 17&
Private Const LVM_GETCOLUMNWIDTH = LVM_FIRST + 29&
Private Const LVM_SETCOLUMNWIDTH = LVM_FIRST + 30&
Private Const LVM_GETHEADER = LVM_FIRST + 31&
Private Const LVM_GETTOPINDEX = LVM_FIRST + 39&
Private Const LVM_GETCOUNTPERPAGE = LVM_FIRST + 40&
Private Const LVM_SETITEMSTATE = LVM_FIRST + 43&
Private Const LVM_GETSELECTEDCOUNT = LVM_FIRST + 50&
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54&
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55&
Private Const LVM_SUBITEMHITTEST = LVM_FIRST + 57&
Private Const LVM_SETHOTITEM = LVM_FIRST + 60&
Private Const LVM_GETHOTITEM = LVM_FIRST + 61&
Private Const LVM_GETSTRINGWIDTHW = LVM_FIRST + 87&

#If Unicode Then
 Private Const LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHW
#Else
 Private Const LVM_GETSTRINGWIDTH = LVM_GETSTRINGWIDTHA
#End If

Private Type LVITEM
    mask       As Long
    iItem      As Long
    iSubItem   As Long
    state      As Long
    stateMask  As Long
    pszText    As String
    cchTextMax As Long
    iImage     As Long
    lParam     As Long
    iIndent    As Long
End Type

Private Const LVIF_STATE = &H8&
Private Const LVNI_SELECTED = &H2&
Private Const LVIS_SELECTED = &H2&

Private Const LVSCW_AUTOSIZE = -1&
Private Const LVSCW_AUTOSIZE_USEHEADER = -2&
Private Const LVS_EX_FULLROWSELECT = &H20&

Private Type HDITEM    ' Header item
    mask As Long       ' Indicates which structure members are valid
    cxy As Long        ' Width or height of the item
    pszText As String  ' Item string
    hbm As Long        ' Handle to the item bitmap HBITMAP
    cchTextMax As Long ' Length of the item string, in characters
    fmt As Long        ' Bit flags that specify the item's format
    lParam As Long     ' Application-defined item data
    iImage As Long     ' Zero-based index within the image list
    iOrder As Long     ' Zero-based order of member from left to right
End Type

Private Const HDI_FORMAT = &H4&
Private Const HDI_IMAGE = &H20&

Private Const HDF_IMAGE = &H800&
Private Const HDF_STRING = &H4000&

Private Const HDM_SETITEM = &H1204&
Private Const HDM_SETIMAGELIST = &H1208&

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type LVHITTESTINFO
    pt       As POINTAPI
    flags    As Long
    iItem    As Long
    iSubItem As Long
End Type

Private hWndHdr As Long

Sub LV_SetColumnWidth(ByVal LV_hWnd As Long, ByVal iCol As Long, Optional ByVal iSize As Long = LVSCW_AUTOSIZE_USEHEADER)
   ' Sets one based column width for List and Report. Specify iSize
   ' as LVSCW_AUTOSIZE, LVSCW_AUTOSIZE_USEHEADER or positive Twips.
   If iSize > 0 Then iSize = iSize \ Screen.TwipsPerPixelX
   SendMessageLong LV_hWnd, LVM_SETCOLUMNWIDTH, iCol - 1&, iSize
End Sub

Function LV_GetColumnWidth(ByVal LV_hWnd As Long, ByVal iCol As Long) As Single
    ' Returns the one based column width in Twips, or zero otherwise.
    LV_GetColumnWidth = SendMessageLong(LV_hWnd, LVM_GETCOLUMNWIDTH, iCol - 1&, 0&) * Screen.TwipsPerPixelX
End Function

Function LV_GetVisibleCount(ByVal LV_hWnd As Long) As Long
   ' Returns the number of items that are fully visible vertically in a list view
   ' control when in list or report view. If the current view is icon or small icon
   ' view, the return value is the total number of items in the list view control.
   LV_GetVisibleCount = SendMessageLong(LV_hWnd, LVM_GETCOUNTPERPAGE, 0&, 0&)
End Function

Function LV_GetHotItem(ByVal LV_hWnd As Long) As Long
   ' Returns the one based index of the item that is hot.
   LV_GetHotItem = SendMessageLong(LV_hWnd, LVM_GETHOTITEM, 0&, 0&) + 1& 'api zero-based
End Function

Function LV_SetHotItem(ByVal LV_hWnd As Long, ByVal lNewIndex As Long) As Long
   ' Sets the one based index of the item to be set as the hot item.
   ' Returns the one based index of the item that was previously hot.
   LV_SetHotItem = SendMessageLong(LV_hWnd, LVM_SETHOTITEM, lNewIndex - 1&, 0&) + 1& 'one-based
End Function

Sub LV_SelectEntireRow(ByVal LV_hWnd As Long, ByVal bEntire As Boolean)
'   Dim lStyle As Long
'   lStyle = SendMessage(ListView.hWnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
'   lStyle = lStyle Or LVS_EX_FULLROWSELECT
'   Call SendMessage(ListView.hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal lStyle)
   SendMessageLong LV_hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, CLng(bEntire)
End Sub

Function LV_GetTopIndex(ByVal LV_hWnd As Long) As Long
   ' Retrieves the index of the topmost visible item when in list or report view.
   ' If the list view control is in icon or small icon view returns zero.
   LV_GetTopIndex = SendMessageLong(LV_hWnd, LVM_GETTOPINDEX, 0&, 0&)
End Function

Function LV_GetSelectedCount(ByVal LV_hWnd As Long) As Long
   ' Returns the number of items that are selected.
   LV_GetSelectedCount = SendMessageLong(LV_hWnd, LVM_GETSELECTEDCOUNT, 0&, 0&)
End Function

Function LV_GetSelectedNext(ByVal LV_hWnd As Long, Optional ByVal iStart As Long) As Long
   ' Returns the one based index of the next selected item, or 0 otherwise.
   ' The iStart parameter specifies the item to begin the search after, or 0 to
   ' find the first item. The specified item itself is excluded from the search.
   LV_GetSelectedNext = SendMessageLong(LV_hWnd, LVM_GETNEXTITEM, iStart - 1&, LVNI_SELECTED) + 1& 'one-based
End Function

Function LV_GetListCount(ByVal LV_hWnd As Long) As Long
   LV_GetListCount = SendMessageLong(LV_hWnd, LVM_GETITEMCOUNT, 0&, 0&)
End Function

Sub LV_SelectAll(ByVal LV_hWnd As Long)
    Dim lvi As LVITEM
    With lvi ' Select all list items
        .mask = LVIF_STATE
        .state = LVIS_SELECTED
        .stateMask = LVIS_SELECTED
    End With
    SendMessageLong LV_hWnd, LVM_SETITEMSTATE, -1&, VarPtr(lvi)
End Sub

Sub LV_SetHeaderImageList(ByVal LV_hWnd As Long, ImgList As ImageList)
   ' Assigns an image list to an existing header control.
   hWndHdr = SendMessageLong(LV_hWnd, LVM_GETHEADER, 0&, 0&)
   SendMessageLong hWndHdr, HDM_SETIMAGELIST, 0&, ImgList.hImageList
End Sub

Sub LV_SetHeaderIcon(ByVal iCol As Long, ByVal iIcon As Long)
    If hWndHdr = 0 Then Exit Sub ' Requires calling LV_SetHeaderImageList first.
    Dim hdi As HDITEM
    hdi.mask = HDI_FORMAT                  ' The fmt member is valid
    hdi.fmt = HDF_STRING                   ' Must specify the string flag
    If (iIcon) Then                        ' Remove icon if HDF_IMAGE is not set
        hdi.mask = hdi.mask Or HDI_IMAGE   ' The iImage member is also valid
        hdi.fmt = hdi.fmt Or HDF_IMAGE     ' Use an image from the image list
        hdi.iImage = iIcon - 1&            ' Zero-based index within the image list
    End If                                 ' Zero-based column header
    SendMessageLong hWndHdr, HDM_SETITEM, iCol - 1&, VarPtr(hdi)
End Sub

Function LV_HitTest(ByVal LV_hWnd As Long, ByVal x As Single, ByVal y As Single) As Long
    Dim tHitTestInfo As LVHITTESTINFO
    ' Determines which list view item or subitem is at a given position.
    ' Returns the one based index of the item or subitem tested, or 0 otherwise.
    ' If an item or subitem is at the given coordinates, the fields of the
    ' LVHITTESTINFO structure will be filled with the applicable hit information.
    With tHitTestInfo
        .pt.x = x / Screen.TwipsPerPixelX
        .pt.y = y / Screen.TwipsPerPixelY
    End With
    LV_HitTest = SendMessageLong(LV_hWnd, LVM_SUBITEMHITTEST, 0&, VarPtr(tHitTestInfo)) + 1& 'one-based
End Function

' LVM_GETSTRINGWIDTH
'    wParam = 0;
'    lParam = psz;
' Determines the width of a specified string using the specified list view
' control's current font. Returns the string width if successful, or zero.
' psz - Address of a null-terminated string.
' The LVM_GETSTRINGWIDTH message returns the exact width, in pixels, of the
' specified string. If you use the returned string width as the column width
' in the LVM_SETCOLUMNWIDTH message, the string will be truncated.
' To get the column width that can contain the string without truncating it,
' you must add padding to the returned string width.

