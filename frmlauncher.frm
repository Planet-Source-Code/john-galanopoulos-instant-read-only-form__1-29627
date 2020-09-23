VERSION 5.00
Begin VB.Form frmlauncher 
   Caption         =   "Form2"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form2"
   ScaleHeight     =   5505
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReadOnly 
      Caption         =   "Read only form"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdNormal 
      Caption         =   "Normal Form"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Original Idea : Ian Ippolito, https://www.planet-source-code.com/xq/ASP/txtCodeId.29562/lngWId.1/qx/vb/scripts/ShowCode.htm "
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   4800
      Width           =   7575
   End
   Begin VB.Label Label4 
      Caption         =   "This sample works for every control that supports hWnd"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label Label3 
      Caption         =   $"frmlauncher.frx":0000
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6975
   End
   Begin VB.Label Label2 
      Caption         =   "John Galanopoulos - GreekThought@yahoo.gr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   4440
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Once again, the Windows API prooved to be the best problem solver... Use freely. Peace :)"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
   End
End
Attribute VB_Name = "frmlauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNormal_Click()
Dim objForm As New frmTowerForm
    objForm.Show
End Sub

Private Sub cmdReadOnly_Click()
Dim objForm As New frmTowerForm
    Load objForm
     SwitchReadOnly objForm, objForm.Command1
     'By passing objForm.Command1 we exclude all command buttons from getting read only
     
         
    'objForm.MakeReadOnly
    objForm.Show
End Sub


