Attribute VB_Name = "modListStyle"
Option Explicit

Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const LVS_EX_GRIDLINES As Long = &H1&
Public Const LVS_EX_CHECKBOXES As Long = &H4&
Public Const LVS_EX_FULLROWSELECT As Long = &H20&
Const LVIS_STATEIMAGEMASK As Long = &HF000
Const LVIF_STATE As Long = &H8

Public Enum LISTVIEW_MESSAGES
    LVM_FIRST = &H1000
    LVM_SETITEMCOUNT = (LVM_FIRST + 47)
    LVM_GETITEMRECT = (LVM_FIRST + 14)
    LVM_SETITEMSTATE = (LVM_FIRST + 43)
    LVM_GETITEMSTATE = (LVM_FIRST + 44)
    LVM_SCROLL = (LVM_FIRST + 20)
    LVM_GETTOPINDEX = (LVM_FIRST + 39)
    LVM_HITTEST = (LVM_FIRST + 18)
    LVM_DELETEALLITEMS = (LVM_FIRST + 9)
    LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54)
    LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
End Enum
