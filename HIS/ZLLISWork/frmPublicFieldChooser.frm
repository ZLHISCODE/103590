VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPublicFieldChooser 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "列选择"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   2565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ListView lvwChooser 
      Height          =   4305
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   7594
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "列名选择"
         Object.Width           =   3704
      EndProperty
   End
End
Attribute VB_Name = "frmPublicFieldChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrFiled As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnTrue As Boolean
    Dim intLoop As Integer
    If KeyCode = 65 And Shift = 2 Then
        With Me.lvwChooser
            blnTrue = Not .ListItems(1).Checked
            For intLoop = 1 To .ListItems.Count
                .ListItems(intLoop).Checked = blnTrue
            Next
        End With
    End If

End Sub

Private Sub Form_Resize()
    lvwChooser.Move 4, 4, ScaleWidth - 8, ScaleHeight - 8
End Sub

Public Function ShowMe(objfrm As Object, Cols As ReportColumns) As String
    Dim intLoop As Integer
    Dim Item As ListItem
    Dim strFiled As String
    With Me.lvwChooser
        For intLoop = 0 To Cols.Count - 1
            If Cols(intLoop).ShowInFieldChooser = True Then
                Set Item = .ListItems.Add(, "A" & intLoop, Cols(intLoop).Caption)
                Item.Checked = Cols(intLoop).Visible
            End If
        Next
    End With
    Me.Show vbModal, objfrm
    With Me.lvwChooser
        For intLoop = 0 To .ListItems.Count - 1
            If .ListItems(intLoop).Checked = True Then
                strFiled = strFiled & ";" & .ListItems(intLoop).Text
            End If
        Next
    End With
    ShowMe = mstrFiled
End Function

Private Sub Form_Unload(Cancel As Integer)
    Dim intLoop As Integer
    Dim strFiled As String
    With Me.lvwChooser
        For intLoop = 1 To .ListItems.Count
            If .ListItems(intLoop).Checked = True Then
                strFiled = strFiled & ";" & .ListItems(intLoop).Text
            End If
        Next
    End With
    mstrFiled = strFiled
End Sub

