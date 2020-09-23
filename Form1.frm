VERSION 5.00
Begin VB.Form frmTowerForm 
   Caption         =   "Tower Form"
   ClientHeight    =   5880
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1455
      Left            =   5040
      TabIndex        =   17
      Top             =   2760
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   7320
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   675
      TabIndex        =   16
      Top             =   2760
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   840
      Left            =   5040
      TabIndex        =   14
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Text            =   "Other stuff"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4695
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   1575
         Left            =   1680
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
         Begin VB.OptionButton optTall 
            Caption         =   "Tall"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optMedium 
            Caption         =   "Medium"
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton optSmall 
            Caption         =   "Small"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.ComboBox cmbTowerType 
         Height          =   315
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Text            =   "My Tower"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tower Options:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Tower Type:"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tower Name:"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   735
      Left            =   6120
      Top             =   3840
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6120
      Picture         =   "Form1.frx":0912
      Top             =   2880
      Width           =   480
   End
   Begin VB.Line Line1 
      X1              =   600
      X2              =   240
      Y1              =   240
      Y2              =   120
   End
   Begin VB.Label Label5 
      Caption         =   "Hey!! Use this dudes to switch the Enable attribute of form's controls :)"
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   120
      Width           =   6495
   End
   Begin VB.Label Label4 
      Caption         =   "Project Name:"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuReadOnly 
         Caption         =   "&Read Only"
      End
   End
End
Attribute VB_Name = "frmTowerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    cmbTowerType.AddItem "--Unknown"
    cmbTowerType.AddItem "Guyed"
    cmbTowerType.AddItem "Monopole"
    
    cmbTowerType.ListIndex = 1
    
    optMedium.Value = True
    
    With List1
     .AddItem "John Doe"
     .AddItem "Ian Ippolito"
     .AddItem "John Galanopoulos"
     .AddItem "Jennifer Lopez"
     .AddItem "Samantha Fox"
     .AddItem "Britney Spears"
    End With
   
End Sub

'Public Sub MakeReadOnly()
'    mMakeFormReadOnly Me
'End Sub
'Private Sub mnuReadOnly1_Click()
'    mMakeFormReadOnly1 Me
'End Sub

'Private Sub mnuReadOnly_Click()
'    mMakeFormReadOnly Me
'End Sub
'Sub mMakeFormReadOnly1(ByRef robjForm As Form)
'Dim objControl As Control
'
'    'disable all controls
'    For Each objControl In robjForm.Controls
'
'        If InStr(LCase$(TypeName(objControl)), "frame") > 0 Then
'            'do nothing
'        ElseIf InStr(LCase$(TypeName(objControl)), "textbox") > 0 Then
'        ElseIf InStr(LCase$(TypeName(objControl)), "combobox") > 0 Then
'        ElseIf InStr(LCase$(TypeName(objControl)), "optionbutton") > 0 Then
'        ElseIf InStr(LCase$(TypeName(objControl)), "label") > 0 Then
'        Else
'            'disable
'            On Error Resume Next
'            objControl.Enabled = False
'            On Error GoTo 0
'        End If
'    Next
'End Sub
'Sub mMakeFormReadOnly(ByRef robjForm As Form)
    'disable all controls
'Dim objControl As Control

    'For Each objControl In robjForm.Controls
        
    '    If InStr(LCase$(TypeName(objControl)), "textbox") > 0 Then
            'create label for textbox
    '        mCreateLabelOverControl objControl, robjForm
    '    ElseIf InStr(LCase$(TypeName(objControl)), "combobox") > 0 Then
            'create label for combobox
    '        mCreateLabelOverControl objControl, robjForm
    '    ElseIf InStr(LCase$(TypeName(objControl)), "optionbutton") > 0 Then
            'create label
    '        If objControl.Value = True Then
    '            objControl.Enabled = True
    '        Else
    '            objControl.Enabled = False
    '        End If
    '    Else
            'do nothing
    '    End If
        
    'Next
'End Sub

'Sub mCreateLabelOverControl(ByRef robjControl As Control, _
'    ByRef robjForm As Form)
'creates label over control
'Dim strValue As String
    'get value
        'If InStr(LCase$(TypeName(robjControl)), "textbox") > 0 Then
            'create label
        '    strValue = robjControl.Text
        'ElseIf InStr(LCase$(TypeName(robjControl)), "combobox") > 0 Then
            'create combobox
        '    strValue = robjControl.Text
        'ElseIf InStr(LCase$(TypeName(robjControl)), "optionbutton") > 0 Then
            'create combobox
        '    strValue = robjControl.Caption
        'End If

'Dim objControl As Control
'    Set objControl = mCreateControl(robjForm, "vb.label", "lblReadOnly_" & robjControl.Name, _
'        strValue, _
'        robjControl.Container, robjControl.Left, robjControl.Top, robjControl.Width, _
'        robjControl.Height)
   
   ' objControl.ZOrder 0
   
    
'    robjControl.Visible = False

'End Sub

'Private Function mCreateControl( _
    ByRef robjForm As Form, _
    ByVal vstrControlType As String, _
    ByVal vstrControlName As String, _
    ByVal vstrValue As String, _
    ByRef rconContainer As Object, _
    ByVal vlngLeft As Long, _
    ByVal vlngTop As Long, _
    ByVal vlngWidth As Long, _
    ByVal vlngHeight As Long _
    ) As Control
'*********************************************************************
'purpose:dynamically create a GUI control
'inputs:vstrControlType--type of control: ex: "VB.Label"
'       vstrControlName--name to give control
'       vstrValue--value to assign to it
'       rconContainer--container of control
'       vlngTop,vlngLeft,vlngWidth,vlngHeight=Control placement info
'
'*********************************************************************

    'set properties
'    robjForm.Controls.Add vstrControlType, vstrControlName, rconContainer 'rconContainer
    
'    With robjForm.Controls.Item(vstrControlName)
'        Select Case (LCase$(vstrControlType))
'            Case "vb.label"
'                .Caption = vstrValue
'            Case "vb.textbox"
'                .Text = vstrValue
'            Case "vb.combobox"
                '.Sorted = True
'            Case Else
'                'Stop
'        End Select
'        .Top = vlngTop
'        .Left = vlngLeft
'        .Width = vlngWidth
    
        'set height (can't do it for comboboxes)
'        If (LCase$(vstrControlType) <> "vb.combobox") Then
'            .Height = vlngHeight
'        End If
        
        'set font
'        .FontName = "Tahoma"
        
        'set alignment (if label)
'        If (LCase$(vstrControlType) = "vb.label") Then
            'determine label type
'            If (Left$(vstrControlName, Len(mstrGuiControlPrefix)) = mstrGuiControlPrefix) Then
'                'substitue for input control (i.e. id value)
'                .Alignment = 0
'            Else
'                'robjForm.Caption control
'                .Alignment = 1
'            End If
'        End If
        
        'set style if combobox
    '    If (LCase$(vstrControlType) = "vb.combobox") Then
    '        .Style = 2
    '    End If
       
        
        'show it
 '       .Visible = True
 '   End With
    'return control
 '   Set mCreateControl = robjForm.Controls.Item(vstrControlName)
    
'End Function

Private Sub mnuReadOnly_Click()
SwitchReadOnly Me, Command1
'By passing Command1 we exclude all command buttons from getting read only
End Sub
