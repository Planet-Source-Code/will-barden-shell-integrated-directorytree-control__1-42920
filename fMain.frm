VERSION 5.00
Begin VB.Form fMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ucDirectoryTree test container"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   513
   StartUpPosition =   3  'Windows Default
   Begin prjExplorerControls.ucDirectoryTree ucDirectoryTree1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   6376
      _ExtentY        =   13150
      BorderType      =   2
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   '
   On Error GoTo Err
   '
   ucDirectoryTree1.Initialize
   Exit Sub
   '
Err:
   MsgBox Err.Description
   '
End Sub
