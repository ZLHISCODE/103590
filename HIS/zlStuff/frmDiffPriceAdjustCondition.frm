VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiffPriceAdjustCondition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Զ���������"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   270
      TabIndex        =   14
      Top             =   2580
      Width           =   1100
   End
   Begin VB.Frame fraRangeSelect 
      Caption         =   "������Χ"
      Height          =   2310
      Left            =   120
      TabIndex        =   8
      Top             =   90
      Width           =   5850
      Begin MSComCtl2.UpDown updRate 
         Height          =   300
         Left            =   3015
         TabIndex        =   11
         Top             =   1260
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtRate"
         BuddyDispid     =   196611
         OrigLeft        =   3720
         OrigTop         =   4200
         OrigRight       =   3960
         OrigBottom      =   4575
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtRate 
         Height          =   300
         Left            =   1110
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "10"
         Top             =   1260
         Width           =   1935
      End
      Begin VB.CommandButton Cmd��; 
         Caption         =   "��"
         Height          =   300
         Left            =   5520
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   345
         Width           =   270
      End
      Begin VB.TextBox Txt���� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   1
         Top             =   345
         Width           =   4500
      End
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   780
         Width           =   4665
      End
      Begin VB.Label Label2 
         Caption         =   "  ˵����ʵ�ʲ����ʵ�ʽ��֮�ȴ��ڻ�С��ָ������ʵİٷֵ�Ϊ������ʵ���Щ���ϲų�����"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   5550
      End
      Begin VB.Label Label1 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Lbl�̵㷽ʽ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   135
         TabIndex        =   4
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ϸ���"
         Height          =   180
         Left            =   315
         TabIndex        =   0
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lbl�ⷿ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ⷿ"
         Height          =   180
         Left            =   675
         TabIndex        =   2
         Top             =   840
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4440
      TabIndex        =   6
      Top             =   2565
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3225
      TabIndex        =   5
      Top             =   2550
      Width           =   1100
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   5280
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCondition.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCondition.frx":0E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiffPriceAdjustCondition.frx":2B5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tvw���Ϸ��� 
      Height          =   2400
      Left            =   1215
      TabIndex        =   9
      Top             =   3255
      Visible         =   0   'False
      Width           =   4860
      _ExtentX        =   8573
      _ExtentY        =   4233
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmDiffPriceAdjustCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnSelect As Boolean
Public BlnBootUp As Boolean
Private blnFirstUp As Boolean

Private mstr����ID As String
Private mlng�ⷿID As Long
Private mintRate As Integer
Private Const mlngModule = 1715

Private mfrmMain As Form

Public Function GetCondition(frmMain As Form, ByRef str����ID, ByRef lng�ⷿID As Long, _
        ByRef int������ As Integer) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:��ȡ����۵������������
    '--�����:
    '--������:
    '--��  ��:�����ڹ��ܷ���true,���򷵻�False
    '-----------------------------------------------------------------------------------------------------------

    
    mstr����ID = ""
    mlng�ⷿID = 0
    mintRate = int������
    mblnSelect = False
    
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain
    
    GetCondition = mblnSelect
    str����ID = mstr����ID
    lng�ⷿID = mlng�ⷿID
    int������ = mintRate
    
End Function

Private Sub cbo�ⷿ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtRate.SetFocus
    End If
End Sub

Private Sub CmdCancel_Click()
    mblnSelect = False
    Unload Me
End Sub

Private Sub CmdOK_Click()
    Dim intIndex As Integer
    
    If BlnBootUp = True Then
        If Tvw���Ϸ���.SelectedItem.Key <> "R" Then
                mstr����ID = "" & _
                    "   Select ID From ���Ʒ���Ŀ¼  " & _
                    "   where ����=7" & _
                    "   Start With ID=" & Mid(Tvw���Ϸ���.SelectedItem.Key, 3) & _
                    "   Connect by Prior ID=�ϼ�ID"
        End If
        
        mlng�ⷿID = cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)
        mintRate = Val(txtRate.Text)
        
        mblnSelect = True
        frmDiffPriceAdjustCard.txtStock.Caption = cbo�ⷿ.Text
        frmDiffPriceAdjustCard.txtStock.Tag = mlng�ⷿID
        
        frmDiffPriceAdjustCard.CmdSave.Enabled = False
        frmDiffPriceAdjustCard.cmdCancel.Enabled = False
    End If
    Hide
    Unload Me
End Sub

Private Sub Cmd��;_Click()
    '�Ѳ�����;����װ��TREEVIEW
    Tvw���Ϸ���.Visible = Tvw���Ϸ���.Visible Xor True
    If Tvw���Ϸ���.Visible Then
        Tvw���Ϸ���.Top = Txt����.Top + Txt����.Height + fraRangeSelect.Top
        Tvw���Ϸ���.Left = Txt����.Left + fraRangeSelect.Left
        Tvw���Ϸ���.ZOrder 0
        Tvw���Ϸ���.SetFocus
    End If
End Sub
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub Form_Click()
    If Tvw���Ϸ���.Visible = True Then
        Tvw���Ϸ���.Visible = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim rsTemp As New Recordset
    Dim str���� As String
    Dim strSelectStock As String
    
    On Error GoTo errHandle
    strSelectStock = IIf(Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModule, "0")) = 1, 1, 0)
    
    BlnBootUp = False
    blnFirstUp = True
    
    With mfrmMain.cboStock
        cbo�ⷿ.Clear
        For i = 0 To .ListCount - 1
            cbo�ⷿ.AddItem .List(i)
            cbo�ⷿ.ItemData(cbo�ⷿ.NewIndex) = .ItemData(i)
        Next
        cbo�ⷿ.ListIndex = .ListIndex
    End With
        
    If InStr(1, gstrPrivs, "���пⷿ") <> 0 Then
        If strSelectStock = "0" Then
            cbo�ⷿ.Enabled = False
        Else
            cbo�ⷿ.Enabled = True
        End If
    Else
        cbo�ⷿ.Enabled = False
    End If
    
    With rsTemp
        gstrSQL = "" & _
            "   Select id,�ϼ�id,����,0 as ĩ�� " & _
            "   From  ���Ʒ���Ŀ¼ " & _
            "   Where ����=7" & _
            "   Start with �ϼ�ID is null connect by prior ID =�ϼ�ID " & _
            "   Order by level,ID "
        
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        If .EOF Then
            ShowMsgBox "���Ϸ��಻����,���ڲ���Ŀ¼���������ã�"
            Exit Sub
        End If
        
        With rsTemp
            Tvw���Ϸ���.Nodes.Clear
            Tvw���Ϸ���.Nodes.Add , , "R", "���в��Ϸ���", 1, 1
            Txt����.Text = "���в��Ϸ���"
            Do While Not .EOF
                If IsNull(!�ϼ�ID) Then
                    If !ĩ�� = 1 Then
                        Tvw���Ϸ���.Nodes.Add "R", tvwChild, "K_" & !Id, !����, 3, 3
                    Else
                        Tvw���Ϸ���.Nodes.Add "R", tvwChild, "K_" & !Id, !����, 2, 2
                    End If
                Else
                    If !ĩ�� = 1 Then
                        Tvw���Ϸ���.Nodes.Add "K_" & !�ϼ�ID, tvwChild, "K_" & !Id, !����, 3, 3
                    Else
                        Tvw���Ϸ���.Nodes.Add "K_" & !�ϼ�ID, tvwChild, "K_" & !Id, !����, 2, 2
                    End If
                End If
                Tvw���Ϸ���.Nodes("K_" & !Id).Tag = !ĩ��
                .MoveNext
            Loop
        End With
    
        Tvw���Ϸ���.Nodes("R").Selected = True
        Tvw���Ϸ���.Nodes("R").Expanded = True
    End With
    BlnBootUp = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Select Case UnloadMode
        Case vbFormControlMenu, vbAppWindows, vbAppTaskManager, vbFormOwner
            Me.Hide
        Case vbFormCode
            If Tvw���Ϸ���.Visible Then
                Tvw���Ϸ���.Visible = False
                Cmd��;.SetFocus
                Cancel = 1
                Exit Sub
            End If
    End Select
End Sub

Private Sub fraRangeSelect_Click()
    If Tvw���Ϸ���.Visible = True Then
        Tvw���Ϸ���.Visible = False
    End If
End Sub

Private Sub Tvw���Ϸ���_DblClick()
    Me.Txt����.Text = Tvw���Ϸ���.SelectedItem.Text
    Tvw���Ϸ���.Visible = False
    On Error Resume Next
End Sub

Private Sub Tvw���Ϸ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Tvw���Ϸ���_DblClick
    End If
End Sub

Private Sub Tvw���Ϸ���_LostFocus()
    Tvw���Ϸ���.Visible = False
End Sub

Private Sub txtRate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyAdd
            
            If Val(txtRate.Text) < 100 Then
                txtRate.Text = Val(txtRate.Text) + 1
            End If
        Case vbKeySubtract
            
            If Val(txtRate.Text) > 1 Then
                txtRate.Text = Val(txtRate.Text) - 1
            End If
        
    End Select
    
End Sub

Private Sub txtRate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
        
        Case 48 To 57
            If IsNumeric(txtRate.Text) Then
                If txtRate.SelLength <> Len(txtRate.Text) Then
                    If Val(txtRate.Text & Chr(KeyAscii)) > 100 Then
                        KeyAscii = 0
                    End If
                End If
            End If
        Case 8          '�˸��
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtRate_Validate(Cancel As Boolean)
    If Trim(txtRate.Text) = "" Or Trim(txtRate.Text) = "0" Then
        Cancel = True
    End If
End Sub

