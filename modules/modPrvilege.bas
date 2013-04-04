Attribute VB_Name = "modPrvilege"
Option Explicit

Private Declare Function RtlAdjustPrivilege Lib "ntdll.dll" (ByVal Privilege As Long, ByVal Enable As Long, ByVal CurrentThread As Long, Enabled As Long) As Long

Sub MiniEnablePrivilege()
    RtlAdjustPrivilege 20&, 1, 0, 0
End Sub
