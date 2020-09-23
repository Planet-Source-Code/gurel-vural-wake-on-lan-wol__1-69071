Attribute VB_Name = "Module1"
Public Declare Function GetTickCount Lib "kernel32" () As Long

Private Const IP_SUCCESS As Long = 0
Private Const MAX_WSADescription As Long = 256
Private Const MAX_WSASYSStatus As Long = 128
Private Const WS_VERSION_REQD As Long = &H101
Private Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD As Long = 1
Private Const SOCKET_ERROR As Long = -1
Private Const ERROR_SUCCESS As Long = 0

Private Type WSADATA
wVersion As Integer
wHighVersion As Integer
szDescription(0 To MAX_WSADescription) As Byte
szSystemStatus(0 To MAX_WSASYSStatus) As Byte
wMaxSockets As Long
wMaxUDPDG As Long
dwVendorInfo As Long
End Type

Private Declare Function gethostbyname Lib "wsock32.dll" _
(ByVal hostname As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
Alias "RtlMoveMemory" _
(xDest As Any, _
xSource As Any, _
ByVal nbytes As Long)

Private Declare Function lstrlenA Lib "kernel32" _
(lpString As Any) As Long

Private Declare Function WSAStartup Lib "wsock32.dll" _
(ByVal wVersionRequired As Long, _
lpWSADATA As WSADATA) As Long

Private Declare Function WSACleanup Lib "wsock32.dll" () As Long

Private Declare Function inet_ntoa Lib "wsock32.dll" _
(ByVal addr As Long) As Long

Private Declare Function lstrcpyA Lib "kernel32" _
(ByVal RetVal As String, _
ByVal Ptr As Long) As Long

Private Declare Function gethostname Lib "wsock32.dll" _
(ByVal szHost As String, _
ByVal dwHostLen As Long) As Long

Public AbortThis As Boolean
Public PCs


Public Function hex2ascii(ByVal hextext As String) As String
    For Y = 1 To Len(hextext)
    num = Mid(hextext, Y, 2)
    Value = Value & Chr(Val("&h" & num))
    Y = Y + 1
    Next Y
    
    hex2ascii = Value
End Function

Function GetIPFromHostName(ByVal sHostName As String) As String

    'converts a host name to an IP address
    
    Dim nbytes As Long
    Dim ptrHosent As Long 'address of HOSENT structure
    Dim ptrName As Long 'address of name pointer
    Dim ptrAddress As Long 'address of address pointer
    Dim ptrIPAddress As Long
    Dim ptrIPAddress2 As Long
    
    ptrHosent = gethostbyname(sHostName & vbNullChar)
    
    If ptrHosent <> 0 Then
        ptrAddress = ptrHosent + 12
        'get the IP address
        CopyMemory ptrAddress, ByVal ptrAddress, 4
        CopyMemory ptrIPAddress, ByVal ptrAddress, 4
        CopyMemory ptrIPAddress2, ByVal ptrIPAddress, 4
        
        GetIPFromHostName = GetInetStrFromPtr(ptrIPAddress2)
    End If

End Function


Function GetStrFromPtrA(ByVal lpszA As Long) As String

    GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
    Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)

End Function


Function GetInetStrFromPtr(Address As Long) As String

    GetInetStrFromPtr = GetStrFromPtrA(inet_ntoa(Address))

End Function

Function ReadMacs(ByVal FIlename As String) As String
On Error GoTo ERR_Control
Set FS = CreateObject("Scripting.FileSystemObject")
If FS.FileExists(FIlename) Then
    whichfile = (FIlename)
    Set thisfile = FS.OpenTextFile(whichfile, 1, False)
    While Not thisfile.AtEndOfStream
        thisline = Trim(thisfile.ReadLine)
        If Len(thisline) > 11 Then ReadMacs = ReadMacs & thisline & ","
    Wend
    If Len(ReadMacs) > 1 Then ReadMacs = Left(ReadMacs, Len(ReadMacs) - 1)
Else
    MsgBox ("MACS.TXT does not exist!")
End If

ERR_Control:
If Err <> 0 Then MsgBox (Err.Description)

End Function

