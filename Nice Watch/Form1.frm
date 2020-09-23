VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sweet DeskTop Clock"
   ClientHeight    =   3120
   ClientLeft      =   6645
   ClientTop       =   4470
   ClientWidth     =   2685
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   2685
   Begin VB.CheckBox chkTopMost 
      BackColor       =   &H80000007&
      Caption         =   "Make it on Top"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   2820
      Width           =   1635
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      ExtentX         =   4789
      ExtentY         =   4789
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'10/31/2002
'Rafael Bonventi
'rafael.bonventi@systemmarketing.com.br
'----------------------------------------------
'this is a nice little tool to add to your app
'i think its just sweet to just stare at it
'and watch the time go by :)
'i hope you all like it
'and please dont forget to vote for it
'-----------------------------------------------
Private FormOnTop As New Class1

Private Sub chkTopMost_Click()
    If chkTopMost = 1 Then
        FormOnTop.MakeTopMost Form1.hWnd
    Else
        FormOnTop.MakeNormal Form1.hWnd
    End If
End Sub

Private Sub Form_Load()

    MsgBox "A very SWEET desktop clock by Rafael Bonventi", vbInformation
    WebBrowser1.Navigate App.Path & "\relogio.html"
    
End Sub

