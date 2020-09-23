VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Note Frequency Calculator"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   315
      Left            =   7425
      TabIndex        =   8
      Top             =   450
      Width           =   915
   End
   Begin VB.ComboBox cmboDec 
      Height          =   315
      Left            =   4950
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   75
      Width           =   990
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   75
      Width           =   2340
   End
   Begin VB.ComboBox cmboOct 
      Height          =   315
      Left            =   2625
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   75
      Width           =   990
   End
   Begin VB.ComboBox cmboNote 
      Height          =   315
      Left            =   750
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   75
      Width           =   990
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Written By Jolyon Bloomfield, January 2001. Jolyon_B@Hotmail.Com"
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   1650
      TabIndex        =   9
      Top             =   525
      Width           =   4965
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Octave:"
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   150
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note:"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Decimal Places:"
      Height          =   195
      Left            =   3675
      TabIndex        =   4
      Top             =   150
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C4 is Middle C"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   525
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Written By Jolyon Bloomfield January 2001
' Frequency Calculator - Used to calculate frequencies for notes of a set pitch
' If you find this useful, thank somebody, long ago, for telling me the formula.
'
' Use this as you wish, just don't claim it as yours, which it is not.
'
' I hopy you enjoy it.
' Jolyon Bloomfield
' ICQ UIN: 11084041           E-mail: Jolyon_B@Hotmail.Com
'

Private Const A1 = 55

Private Sub cmboDec_Click()
Calculate
End Sub

Private Sub cmboNote_Click()
Calculate
End Sub

Private Sub cmboOct_Click()
Calculate
End Sub

Private Sub Form_Load()
Dim I As Integer
For I = 1 To 8
  cmboOct.AddItem I
Next I
cmboOct.ListIndex = 3
For I = 0 To 10
  cmboDec.AddItem I
Next I
cmboDec.ListIndex = 10
With cmboNote
  .AddItem "A"
  .AddItem "Bb"
  .AddItem "B"
  .AddItem "C"
  .AddItem "Db"
  .AddItem "D"
  .AddItem "Eb"
  .AddItem "E"
  .AddItem "F"
  .AddItem "F#"
  .AddItem "G"
  .AddItem "Ab"
  .ListIndex = 0
End With
End Sub

Private Sub Calculate()
If cmboOct.ListIndex = -1 Or cmboNote.ListIndex = -1 Or cmboDec.ListIndex = -1 Then Exit Sub
Dim Temp As Double
Dim Temp2 As String
Temp = A1 * (2 ^ (cmboOct.ListIndex))
Temp = Temp * (2 ^ ((cmboNote.ListIndex) / 12))
Temp2 = Format(Temp, "#." & String(cmboDec.ListIndex, "#"))
If Right(Temp2, 1) = "." Then Temp2 = Left(Temp2, Len(Temp2) - 1)
Text1.Text = Temp2 & " Hertz"
End Sub
