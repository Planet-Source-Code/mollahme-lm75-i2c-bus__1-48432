VERSION 5.00
Object = "{9F3B4DE1-AA29-11D1-A3D9-FDA4E35D1D25}#1.0#0"; "Io.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   ScaleHeight     =   1935
   ScaleWidth      =   3555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Text            =   "0"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "0"
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   840
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin IOLib.IO IO1 
      Left            =   1920
      Top             =   120
      _Version        =   65536
      _ExtentX        =   1270
      _ExtentY        =   1270
      _StockProps     =   0
   End
   Begin VB.Label Label2 
      Caption         =   "Abfragen:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Fehler:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()

Timer1.Enabled = True

End Sub

Private Sub Command2_Click()

Timer1.Enabled = False

End Sub

Private Sub Timer1_Timer()

DoEvents

'Vcc On
'Spannungsversorgung EIN:
Out = IO1.Out(&H378, 128)

'Grundstellung:
Out = IO1.Out(&H37A, 2)     'SCL high
Out = IO1.Out(&H37A, 0)     'SDA high

'Wait exactly 90ms (do not change, or I2C Bus won't initialize!)
'Exakt 90ms warten (nicht ändern, oder I2C BUS initialisiert nicht richtig!)
DoEvents
Sleep (90)

'Start:
Out = IO1.Out(&H37A, 2)     'SDA low
Out = IO1.Out(&H37A, 1)     'SCL low

'Adressbyte senden:
Out = IO1.Out(&H37A, 1)     'Bit7 = 1  'Bit7/Bit6/Bit5/Bit4: feste (static) Adresse des LM75 = 1001
Out = IO1.Out(&H37A, 0)     'SCL high
Out = IO1.Out(&H37A, 1)     'SCL low
Out = IO1.Out(&H37A, 3)     'Bit6 = 0
Out = IO1.Out(&H37A, 2)     'SCL high
Out = IO1.Out(&H37A, 3)     'SCL low
Out = IO1.Out(&H37A, 3)     'Bit5 = 0
Out = IO1.Out(&H37A, 2)     'SCL high
Out = IO1.Out(&H37A, 3)     'SCL low
Out = IO1.Out(&H37A, 1)     'Bit4 = 1
Out = IO1.Out(&H37A, 0)     'SCL high
Out = IO1.Out(&H37A, 1)     'SCL low
Out = IO1.Out(&H37A, 3)     'Bit3 = 1
Out = IO1.Out(&H37A, 2)     'SCL high
Out = IO1.Out(&H37A, 3)     'SCL low
Out = IO1.Out(&H37A, 3)     'Bit2 = 0
Out = IO1.Out(&H37A, 2)     'SCL high
Out = IO1.Out(&H37A, 3)     'SCL low
Out = IO1.Out(&H37A, 3)     'Bit1 = 0
Out = IO1.Out(&H37A, 2)     'SCL high
Out = IO1.Out(&H37A, 3)     'SCL low
Out = IO1.Out(&H37A, 1)     'Bit0 = 1   'Bit0:  0 = Master schreibt anschließend (Master writes)
Out = IO1.Out(&H37A, 0)     'SCL high   '       1 = Master liest anschließend (Master reads)
Out = IO1.Out(&H37A, 1)     'SCL low
 
'LM75 sends Ack
'LM75 sendet Acknowledge (hier nicht ausgewertet):
Out = IO1.Out(&H37A, 3)     'SDA low
Out = IO1.Out(&H37A, 2)     'SCL high
Out = IO1.Out(&H37A, 3)     'SCL low

DoEvents

'Datenbyte 1 lesen(read):
Out = IO1.Out(&H37A, 1)     'SDA high
Out = IO1.Out(&H37A, 0)     'SCL high

'Input
Status = IO1.In(&H379)
If Status And 64 Then ByteW = 128

Out = IO1.Out(&H37A, 1)     'SCL low
Out = IO1.Out(&H37A, 0)     'SCL high

Status = IO1.In(&H379)
If Status And 64 Then ByteW = ByteW + 64

Out = IO1.Out(&H37A, 1)     'SCL low
Out = IO1.Out(&H37A, 0)     'SCL high

Status = IO1.In(&H379)
If Status And 64 Then ByteW = ByteW + 32

Out = IO1.Out(&H37A, 1)     'SCL low
Out = IO1.Out(&H37A, 0)     'SCL high

Status = IO1.In(&H379)
If Status And 64 Then ByteW = ByteW + 16

Out = IO1.Out(&H37A, 1)     'SCL low
Out = IO1.Out(&H37A, 0)     'SCL high

Status = IO1.In(&H379)
If Status And 64 Then ByteW = ByteW + 8

Out = IO1.Out(&H37A, 1)     'SCL low
Out = IO1.Out(&H37A, 0)     'SCL high

Status = IO1.In(&H379)
If Status And 64 Then ByteW = ByteW + 4

Out = IO1.Out(&H37A, 1)     'SCL low
Out = IO1.Out(&H37A, 0)     'SCL high

Status = IO1.In(&H379)
If Status And 64 Then ByteW = ByteW + 2

Out = IO1.Out(&H37A, 1)     'SCL low
Out = IO1.Out(&H37A, 0)     'SCL high

Status = IO1.In(&H379)
If Status And 64 Then ByteW = ByteW + 1

Out = IO1.Out(&H37A, 1)     'SCL low
 
'Send Ack
'Acknowledge vom Master senden:
Out = IO1.Out(&H37A, 3)     'SDA low
Out = IO1.Out(&H37A, 2)     'SCL high
Out = IO1.Out(&H37A, 3)     'SCL low
Out = IO1.Out(&H37A, 1)     'SDA high

'Read highest bit of the second databyte
'Hoechstwertiges Bit vom 2. Datenbyte lesen:
Out = IO1.Out(&H37A, 1)     'SCL high

Status = IO1.In(&H379)
If Status And 64 Then BitW = 1

'Brutal: Vcc Off (we got all data)
'Brutalversion; Spannung abklemmen (wir haben alle Daten):
Out = IO1.Out(&H378, 0)

'Make Temp-value in °C
'Temperaturwert in °C erzeugen:
If (ByteW And 128) Then        'Vorzeichen-Bit (Byte1,Bit7) auswerten
  Tmp! = ByteW - 256           'Temperatur unter Null
  Text2.Text = Text2.Text + 1
 Else
  Tmp! = ByteW                'Temperatur Null oder drueber
  Text3.Text = Text3.Text + 1
End If
If BitW Then Tmp! = Tmp! + 0.5 'Halbgrad-Bit (Byte2,Bit7) auswerten

Text1.Text = Tmp!

End Sub
