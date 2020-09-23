VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Open"
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   9840
      TabIndex        =   2
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   495
      Left            =   4440
      TabIndex        =   0
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   7575
      Left            =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Image1.Picture = LoadPicture(App.Path + "\sunset.jpg")
End Sub

Private Sub Command2_Click()
a = MsgBox("After you have opend an erotic image, just click change if you don't want someone to see the file you are viewing. Created by Luke Rhodes of Rhodes Programming Inc.", vbOKOnly, "About")
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim sFile As String


    With CommonDialog1
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    Image1.Picture = LoadPicture(CommonDialog1.FileName)

    Resume
End Sub

Private Sub Form_Load()
Image1.Picture = LoadPicture(App.Path + "\sunset.jpg")
End Sub
