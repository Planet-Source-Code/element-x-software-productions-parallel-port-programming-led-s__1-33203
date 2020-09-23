VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   7965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Off"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Pin 9"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Pin 8"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pin 7"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pin 6"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Pin 5"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Pin 4"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pin 3"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pin 2"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' These two declarations for the inpout32.dll are for getting data from the
' parallel port, Inp, and sending data to the parallel port, Out.

Private Declare Function Inp Lib "inpout32.dll" _
Alias "Inp32" (ByVal PortAddress As Integer) As Integer

Private Declare Sub Out Lib "inpout32.dll" _
Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)

' Declare the veriables for the Port Address and the Value to send to the port
Dim PortAddress As String

Private Sub Command1_Click()
' This is the same for all of the out commands.  To send data to pin 2 you
' use Out PortAddress, 1  where PortAddress is the address to your
' parallel port and 1 is which data port you want to turn on.

Out PortAddress, 1
End Sub

Private Sub Command2_Click()
Out PortAddress, 2
End Sub

Private Sub Command3_Click()
Out PortAddress, 3
End Sub

Private Sub Command4_Click()
Out PortAddress, 4
End Sub

Private Sub Command5_Click()
Out PortAddress, 5
End Sub

Private Sub Command6_Click()
Out PortAddress, 6
End Sub

Private Sub Command7_Click()
Out PortAddress, 7
End Sub

Private Sub Command8_Click()
Out PortAddress, 8
End Sub

Private Sub Command9_Click()
' This will send data to the I/O port or pin 1 on the parallel port.  Sending data
' here will turn off all of the other ports you had turned on.

Out PortAddress, 0
End Sub

Private Sub Form_Load()
' The Port Address is the HEX address of your parallel port.  You find this
' by going to your device manager and clicking on the Resources tab for your
' parallel port.  &H378 is the default but if yours isn't change it below.

PortAddress = &H378
End Sub
