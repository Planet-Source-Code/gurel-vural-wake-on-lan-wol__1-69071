VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "WAKE ON LAN"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5085
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   5085
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   1560
      TabIndex        =   12
      Top             =   960
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   3600
      Picture         =   "Form1.frx":08CA
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Text            =   "idle"
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Wake All"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   4
      Top             =   1440
      Width           =   975
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      ItemData        =   "Form1.frx":0FB2
      Left            =   120
      List            =   "Form1.frx":0FB4
      TabIndex        =   3
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "2"
      Top             =   3840
      Width           =   375
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   4200
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtmac 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "00123F34FE27"
      Top             =   480
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Wake"
      Height          =   375
      Index           =   0
      Left            =   3600
      MaskColor       =   &H0080FF80&
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Computer Name : "
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MAC"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MACS"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Number of Packets to send"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   4320
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub picmov()

Select Case Picture1.Left
    Case 3600
        Picture1.Left = 3650
        wait 500
    Case 3650
        Picture1.Left = 3600
        wait 250
End Select

DoEvents
End Sub

Private Sub Command1_Click(Index As Integer)
AbortThis = False
NumOfCounts = Text1.Text
Picture1.Visible = True
Adet = Text1.Text
    Command1(Index).Enabled = False
    Dim sHostName As String
    Dim data As String
    Dim MacAdd As String
    Dim IPAddress As String
    
    

    Winsock2.Protocol = sckUDPProtocol
    Winsock2.RemoteHost = "255.255.255.255"
    Winsock2.RemotePort = "4000"
 
If Index = 1 Then
    For say = 0 To List1.ListCount - 1
        
        List1.ListIndex = say
        
        MacAdd = List1.List(say)
        
        data = "FFFFFFFFFFFF"
        For i = 1 To 16
            data = data & MacAdd
        Next i
        
        data = hex2ascii(data)
    
    
        
        For t = 1 To Adet
            Text2.Text = "Wake up Call for " & MacAdd
            'Cancel button pressed
            If AbortThis Then GoTo Aborted
            Winsock2.SendData data
            picmov
            Winsock2.Close
            picmov
            Text1.Text = t
            DoEvents
        Next
    Next
Else
    
    MacAdd = txtmac.Text
    
    data = "FFFFFFFFFFFF"
    
    For i = 1 To 16
        data = data & MacAdd
    Next i
    
    data = hex2ascii(data)
    
    For t = 1 To Adet
        Text2.Text = "Wake up Call for " & MacAdd
        'Cancel button pressed
        If AbortThis Then GoTo Aborted
        Winsock2.SendData data
        picmov
        Winsock2.Close
        picmov
        Text1.Text = t
        DoEvents
    Next
    
End If

Aborted:
If AbortThis Then
    Picture1.Visible = False
    'MsgBox "Cancelled by user!", , "Cancelled"
    Text2.Text = "Aborted!"
    Text1.Text = NumOfCounts
End If

Command1(Index).Enabled = True
Picture1.Visible = False
End Sub


Public Sub wait(ByVal dblMilliseconds As Double)
    Dim dblStart As Double
    Dim dblEnd As Double
    Dim dblTickCount As Double
    
    dblTickCount = GetTickCount()
    dblStart = GetTickCount()
    dblEnd = GetTickCount + dblMilliseconds
    
    Do
    DoEvents
    dblTickCount = GetTickCount()
    Loop Until dblTickCount > dblEnd Or dblTickCount < dblStart
       
    
End Sub



Private Sub Command2_Click()
AbortThis = True
End Sub

Private Sub Form_Load()
Dim Mac

allmacs = ReadMacs(App.Path & "\Macs.txt")
Mac = Split(allmacs, ",")
ReDim PCs(UBound(Mac))

For t = 0 To UBound(Mac)
    SpltP = (InStr(Mac(t), ":"))
    If SpltP > 0 Then
        PCs(t) = Right(Mac(t), Len(Mac(t)) - SpltP)
        Mac(t) = Mid(Mac(t), 1, SpltP - 1)
    Else
        PCs(t) = "N/A"
    End If

List1.AddItem Mac(t)

Next
End Sub

Private Sub List1_Click()
txtmac.Text = List1.List(List1.ListIndex)
Text3.Text = PCs(List1.ListIndex)
End Sub
