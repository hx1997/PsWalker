Attribute VB_Name = "modPsWalker"
Option Explicit

Private Declare Function GetLogicalDriveStrings Lib "kernel32.dll" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function QueryDosDevice Lib "kernel32.dll" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long

Const BUFSIZE = 512

Function NTPathToDosPath(ByVal szNTPath As String, ByRef szDosPath As String) As Boolean
    Dim szTemp As String * BUFSIZE, n As Integer
    
    n = 1
    
    If GetLogicalDriveStrings(260, szTemp) Then
        Dim szName As String * 260
        Dim szDrive As String
        Dim bFound As Boolean
        Dim p As String
        
        szDrive = " :"
        p = szTemp
        
        Do
            Mid(szDrive, 1, 1) = Mid(p, n, 1)
            
            If QueryDosDevice(szDrive, szName, 260) Then
                Dim uNameLen As Long
                
                uNameLen = InStr(1, szName, Chr$(0)) - 1
                
                If (uNameLen < 260) Then
                    bFound = (Left$(szNTPath, uNameLen) = Left$(szName, uNameLen))
                    
                    If (bFound) Then
                        Dim szTempFile As String
                        
                        szTempFile = szDrive & Right$(szNTPath, Len(szNTPath) - uNameLen)
                        szDosPath = szTempFile
                        
                        NTPathToDosPath = True
                        
                        Exit Do
                    Else
                        szName = String(260, Chr$(0))
                        n = n + 4
                    End If
                End If
            Else
                Exit Do
            End If
        Loop
    End If
End Function

Function StandardizeProcessPath(ByVal szPath As String) As String
    Dim szTemp As String
    
    szTemp = szPath
    
    If Left$(szPath, 4) = "\??\" Then
        szTemp = Replace(szTemp, "\??\", "")
    End If
    
    If Left$(UCase(szPath), 12) = "\SYSTEMROOT\" Then
        szTemp = Replace(UCase(szTemp), "\SYSTEMROOT\", Environ("SystemRoot") & "\")
    End If
    
    If Left$(UCase(szPath), 8) = "\DEVICE\" Then
        Dim DosPath As String
        
        If NTPathToDosPath(szTemp, DosPath) Then
            szTemp = DosPath
        End If
    End If
    
    StandardizeProcessPath = szTemp
End Function
