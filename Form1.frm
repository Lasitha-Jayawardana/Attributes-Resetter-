VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{01646141-065C-11D4-8ED3-00E07D815373}#1.0#0"; "MBBrowse.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin MBBrowse.BrowseFF m 
      Left            =   600
      Top             =   3360
      _ExtentX        =   1085
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "folder"
      Height          =   855
      Left            =   4680
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "file"
      Height          =   735
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog c 
      Left            =   240
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
c.ShowOpen
Label1.Caption = c.FileName
Dim fso As New FileSystemObject
Dim f As File
Dim fo As Folder
Set f = fso.GetFile(Label1.Caption)
f.Attributes = Normal
End Sub

Private Sub Command2_Click()
m.Browse
Label2.Caption = m.SelectedItem
Dim fso As New FileSystemObject
Dim f As File
Dim fo As Folder
Set fo = fso.GetFolder(Label2.Caption)
fo.Attributes = Normal

End Sub

