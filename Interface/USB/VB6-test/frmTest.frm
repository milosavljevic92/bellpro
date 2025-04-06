VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Interface test"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Otvori Port"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox txtComm 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test Connection"
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   2895
   End
   Begin MSCommLib.MSComm comm 
      Left            =   120
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Relay 1 - OFF"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Relay 1 - ON"
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Comm Number"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If comm.PortOpen Then
        comm.Output = "R11"  ' Šaljemo komandu za paljenje prvog releja
    End If
End Sub

Private Sub Command2_Click()
    If comm.PortOpen Then
        comm.Output = "R10"  ' Šaljemo komandu za gašenje prvog releja
    End If
End Sub
Private Sub Command3_Click()
    TestConnection
End Sub
Private Sub Command4_Click()
initSerial (txtComm.Text)
End Sub
Private Sub initSerial(commNumber As Integer)
    comm.CommPort = commNumber      ' Postavite COM port koji koristite (npr. COM1)
    comm.Settings = "9600,N,8,1"    ' Podešavanje baud rate-a i parametara
    comm.InputLen = 0               ' Èita sve podatke
    comm.PortOpen = True            ' Otvaramo port za komunikaciju
End Sub
Private Sub TestConnection()
    If comm.PortOpen Then
        comm.Output = "HELLO"  ' Šaljemo komandu za testiranje veze
        ' Èekamo odgovor sa serijskog porta
        Do
            DoEvents
        Loop While comm.InputLen = 0
        
        Dim response As String
        response = comm.Input
        response = Trim(response)  ' Uklanjamo praznine sa poèetka i kraja
        
        If response = "OK" Then
            MsgBox "Interface is connected and ready!"
        Else
            MsgBox "Interface not connected!"
        End If
    End If
End Sub
Private Sub comm_OnComm()
    Select Case comm.CommEvent
        Case comEvReceive
            ' Ovdje æemo obraditi podatke ako ih primimo
            Dim strData As String
            strData = comm.Input
            If strData = "OK" Then
                MsgBox "Arduino is connected and responding!"
            End If
    End Select
End Sub

Private Sub usbRelay(relayNum As Integer, relayState As Integer)
 If comm.PortOpen = False Then Return
 comm.Output = "R10"
End Sub

Private Sub Form_Load()

End Sub
