VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�û���Ϣ�༭"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12600
   Icon            =   "frmUserEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.TreeView tvwPerson 
      Height          =   6210
      Left            =   0
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5520
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10954
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImgСͼ��"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.PictureBox picProcess 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   12600
      TabIndex        =   34
      Top             =   7155
      Visible         =   0   'False
      Width           =   12600
      Begin MSComctlLib.ProgressBar prg 
         Height          =   240
         Left            =   60
         TabIndex        =   35
         Top             =   285
         Width           =   12270
         _ExtentX        =   21643
         _ExtentY        =   423
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblProssCaption 
         AutoSize        =   -1  'True
         Caption         =   "#"
         Height          =   180
         Left            =   180
         TabIndex        =   37
         Top             =   60
         Width           =   90
      End
      Begin VB.Label lblStep 
         Alignment       =   1  'Right Justify
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10695
         TabIndex        =   36
         Top             =   60
         Width           =   420
      End
   End
   Begin VB.ComboBox cboWorkRange 
      Height          =   300
      Left            =   6315
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3975
      Width           =   1320
   End
   Begin VB.ComboBox cboManPros 
      Height          =   300
      Left            =   7725
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3975
      Width           =   1530
   End
   Begin VB.Frame fraApplyTo 
      Caption         =   "Ӧ������δ�����û�����Ա����Ϊ                                  ��"
      Height          =   870
      Left            =   3450
      TabIndex        =   23
      Top             =   4035
      Width           =   7890
      Begin VB.OptionButton optApplyTo 
         Caption         =   "��ǰ��Ա(&3)"
         Height          =   240
         Index           =   2
         Left            =   675
         TabIndex        =   26
         Top             =   405
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton optApplyTo 
         Caption         =   "������Ա(&5)"
         Height          =   240
         Index           =   1
         Left            =   3885
         TabIndex        =   28
         Top             =   405
         Width           =   1395
      End
      Begin VB.OptionButton optApplyTo 
         Caption         =   "��ǰ������Ա(&4)"
         Height          =   240
         Index           =   0
         Left            =   2145
         TabIndex        =   27
         Top             =   405
         Width           =   1785
      End
   End
   Begin VB.OptionButton optMode 
      Caption         =   "����Ա��������û���(&2)"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   15
      ToolTipText     =   "���������ÿһ��Ϊ�ַ���,��ֱ���ñ��빹���û�,����ΪU+������ʽ�����û�"
      Top             =   4500
      Width           =   2595
   End
   Begin VB.OptionButton optMode 
      Caption         =   "���������������û���(&1)"
      Height          =   180
      Index           =   0
      Left            =   450
      TabIndex        =   14
      ToolTipText     =   "��Ӧ����������Աʱ,����д�����ͬ�ļ���,���Լ���+��ŵ���ʽ�����û�"
      Top             =   4200
      Value           =   -1  'True
      Width           =   2595
   End
   Begin MSComctlLib.ImageList ImgСͼ�� 
      Left            =   5295
      Top             =   2070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEdit.frx":000C
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEdit.frx":0326
            Key             =   "User"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEdit.frx":0640
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserEdit.frx":0F1A
            Key             =   "Module"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   11400
      TabIndex        =   31
      Top             =   4410
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   11400
      TabIndex        =   30
      Top             =   705
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   11400
      TabIndex        =   29
      Top             =   330
      Width           =   1100
   End
   Begin VB.Frame fraPerson 
      Caption         =   "��Ӧ��Ա"
      Height          =   2850
      Left            =   135
      TabIndex        =   7
      Tag             =   "0"
      Top             =   2070
      Width           =   3210
      Begin VB.TextBox txtno 
         Height          =   300
         Left            =   675
         MaxLength       =   20
         TabIndex        =   38
         Top             =   292
         Width           =   1125
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   300
         Left            =   1800
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   285
         Width           =   315
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   675
         TabIndex        =   13
         Top             =   1110
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   675
         TabIndex        =   11
         Top             =   690
         Width           =   2175
      End
      Begin VB.Label lblDept 
         Caption         =   "����"
         Height          =   180
         Left            =   240
         TabIndex        =   12
         Top             =   1155
         Width           =   585
      End
      Begin VB.Label lblName 
         Caption         =   "����"
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lblNo 
         Caption         =   "����"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   345
         Width           =   540
      End
   End
   Begin VB.Frame fraUser 
      Caption         =   "�û���Ϣ"
      Height          =   1755
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   3210
      Begin VB.TextBox txtVerify 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1200
         Width           =   1860
      End
      Begin VB.TextBox txtPasswd 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   795
         Width           =   1860
      End
      Begin VB.TextBox txtUserName 
         Height          =   300
         Left            =   1005
         MaxLength       =   20
         TabIndex        =   2
         Top             =   405
         Width           =   1860
      End
      Begin VB.Label lblExamPwd 
         Caption         =   "ȷ������"
         Height          =   195
         Left            =   225
         TabIndex        =   5
         Top             =   1260
         Width           =   900
      End
      Begin VB.Label lblPasswd 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   585
         TabIndex        =   3
         Top             =   870
         Width           =   360
      End
      Begin VB.Label lblUserName 
         Caption         =   "�û���"
         Height          =   180
         Left            =   405
         TabIndex        =   1
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.Frame fraLine 
      Caption         =   "  ��������Ȩ��  "
      Height          =   3030
      Left            =   -15
      TabIndex        =   32
      Top             =   5085
      Width           =   12525
      Begin VSFlex8Ctl.VSFlexGrid vsfModule 
         Height          =   2370
         Left            =   150
         TabIndex        =   40
         Top             =   210
         Width           =   12270
         _cx             =   21643
         _cy             =   4180
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmUserEdit.frx":14B4
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame fraRole 
      Caption         =   "��ɫ��Ȩ:(ѡ����û��ɳ䵱�Ľ�ɫ)"
      Height          =   3705
      Left            =   3450
      TabIndex        =   16
      Top             =   105
      Width           =   7875
      Begin VB.CheckBox chkGranted 
         Caption         =   "ֻ������Ȩ��ɫ(&O)"
         Height          =   285
         Left            =   4185
         TabIndex        =   19
         Top             =   278
         Width           =   1860
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   990
         TabIndex        =   21
         Top             =   645
         Width           =   2805
      End
      Begin VB.ComboBox cboRoleGroups 
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   270
         Width           =   2805
      End
      Begin VSFlex8Ctl.VSFlexGrid vsRole 
         Height          =   2580
         Left            =   105
         TabIndex        =   22
         Top             =   1020
         Width           =   7740
         _cx             =   13652
         _cy             =   4551
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483626
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmUserEdit.frx":1545
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "��  ��(&S)"
         Height          =   180
         Left            =   150
         TabIndex        =   20
         Top             =   705
         Width           =   810
      End
      Begin VB.Label lblRoleGroups 
         AutoSize        =   -1  'True
         Caption         =   "��ɫ��(&R)"
         Height          =   180
         Left            =   150
         TabIndex        =   17
         Top             =   345
         Width           =   810
      End
   End
   Begin VB.Label lblRole 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   4620
      TabIndex        =   33
      Top             =   1320
      Width           =   90
   End
End
Attribute VB_Name = "frmUserEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==ģ�����
'==============================================================
Private Enum ApplyToEnum
    ATE_��ǰ���� = 0
    ATE_������Ա = 1
    ATE_��ǰ��Ա = 2
End Enum

Private Enum VsfModuleTitle
    VMT_ϵͳ = 0
    VMT_��� = 1
    VMT_���� = 2
    VMT_���� = 3
    VMT_˵�� = 4
End Enum

Private mstr������ As String
Private mrsModule As New ADODB.Recordset '�����ɫ�Ĺ���ϸ��
Private mrsRole As New ADODB.Recordset   '�����ɫ
Private mstrUser As String
Private mblnSucceed As Boolean
Private mstrItem As String
Private mblnLoad As Boolean
Private mblnRISMsg As Boolean
Private mstrCreateUserList As String     '��¼�������û����б�

'==============================================================
'==�����ӿ�
'==============================================================
Public Function UserEdit(ByVal strOwner As String, Optional ByVal strUser As String, Optional ByRef strItem As String) As Boolean
'����: strOwner ���ڱ༭��ϵͳ����������
'      strUser  ��ǰ�༭���û��������Ϊ�գ���ʾ����
'����:strItem  �޸ĺ�ķ���ֵ�����ڸ��½�����ʾ��
'���أ�������ӻ��޸ĳɹ�,����true,����False
    mstr������ = strOwner
    mstrUser = strUser
    mblnSucceed = False: mstrItem = ""
    frmUserEdit.Show vbModal, frmMDIMain
    UserEdit = mblnSucceed
    strItem = mstrItem
End Function

'==============================================================
'==�ؼ��¼�
'==============================================================
Private Sub cboRoleGroups_Click()
    Call FillRole
End Sub

Private Sub chkGranted_Click()
    Call FillRole
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp Me.hwnd, "ZL9Svrtools\" & Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim blnChangeMen As Boolean, strUser As String
    Dim rsPerson As ADODB.Recordset
    Dim strPre���� As String
    Dim objPercent As clsPercent
    Dim blnHaveRis As Boolean
    Dim strNote As String, strCheck As String, strNoCheck As String
    
    '����У��
    If Not VilidateData() Then Exit Sub
    mblnRISMsg = False
    If UCase(gstrSTOwner) = UCase(mstr������) And gblnMustRIS Then  '�Ǳ�׼���������
        blnHaveRis = gblnRIS
    End If
    On Error Resume Next
    '�������ݿ��û��Լ���ӦӦ��ϵͳ��Ա
    strUser = UCase(Trim(txtUserName.Text))
    '�޸��û��Ҷ�Ӧ��Ա�����仯
    blnChangeMen = txtName.Tag <> txtNO.Tag And txtNO.Tag <> ""
    If blnChangeMen And mstrUser = mstr������ Then 'ֻ���޸Ķ�Ӧ��Ա����ֱ���˳�
        gcnOracle.Execute "Delete From " & mstr������ & ".�ϻ���Ա�� Where �û���='" & strUser & "'"
        gcnOracle.Execute "Insert Into " & mstr������ & ".�ϻ���Ա��(�û���,��Աid) Values ('" & strUser & "'," & txtNO.Tag & ")"
        If err.Number <> 0 Then
            MsgBox "���ڹ�����ԭ���û������Ϣ(���š�����)����ʧ�ܡ�" & vbNewLine & err.Description, vbExclamation, gstrSysName
            err.Clear
        ElseIf blnHaveRis Then '֪ͨRIS���û����޸�
            If Not gobjRIS.UserEdit(2, strUser) Then
                mblnRISMsg = True
            End If
        End If
    Else
        If blnChangeMen Then '����Ӧ��ϵͳ�û�ʧ�����˳�
            If Not CreateApptionUser(Val(txtNO.Tag), strUser, txtPasswd.Text, mstrUser = "", True) Then
                txtUserName = "": txtPasswd = "": txtVerify = ""
                txtUserName.SetFocus: Exit Sub
            ElseIf blnHaveRis Then '֪ͨRIS���û����޸�
                If Not gobjRIS.UserEdit(IIf(mstrUser = "", 1, 2), strUser) Then
                    mblnRISMsg = True
                End If
            End If
        End If
        If mstrUser <> mstr������ Then '���Խ�����Ȩ����
            If Not gblnDBA Then '��ǰ�û���DBA���Ҵ��������û������Ľ�ɫ������Ҫʹ��System����
                mrsRole.Filter = "Grantee <> '" & gstrUserName & "'"
                If mrsRole.RecordCount > 0 Then '��ȡSysTem�û�����
                    Set gcnSystem = GetConnection("SYSTEM")
                    If gcnSystem Is Nothing Then Exit Sub 'ʧ�ܾ��˳�
                End If
            End If
            mrsRole.Filter = ""
            Do While Not mrsRole.EOF
                If mrsRole!Granted <> mrsRole!��ѡ Or mrsRole!Admin <> mrsRole!ת�� Then
                    mrsRole.Update "�ı�", 1
                End If
                mrsRole.MoveNext
            Loop
            mrsRole.Filter = "�ı�=1"
            If Not optApplyTo(ATE_��ǰ��Ա).value Then Set rsPerson = GetOtherPerson
            '��ֻӦ���뵱ǰ��Ա���ҽ�ɫû�з����仯�����Զ��˳�.����������Ӧ����û�в�ѯ��
            If Not (mrsRole.RecordCount = 0 And optApplyTo(ATE_��ǰ��Ա).value Or Not optApplyTo(ATE_��ǰ��Ա).value And rsPerson Is Nothing And mrsRole.RecordCount = 0) Then
                If mrsRole.RecordCount <> 0 Then
                    mrsRole.Filter = "�ı�=1"
                    If Not ApplyOnePerson(strUser, True, True) Then
                        If vsRole.Enabled Then vsRole.SetFocus
                        Exit Sub
                    End If
                    mrsRole.Filter = "��ѡ=1": mrsRole.Sort = "Role" 'ȡ�����ˣ�����Ȩʹ��
                Else
                    mrsRole.Filter = "��ѡ=1": mrsRole.Sort = "Role" 'ȡ�����ˣ�����Ȩʹ��
                End If
                If Not rsPerson Is Nothing Then 'Ӧ�õ�������Ա
                    Set objPercent = New clsPercent
                    Call objPercent.InitPercent(prg, rsPerson.RecordCount)
                    lblProssCaption.Caption = "": lblStep.Caption = "": picProcess.Visible = True
                    Do While Not rsPerson.EOF
                        strUser = GetUserName(rsPerson!��� & "", rsPerson!���� & "", strPre����)
                        lblProssCaption.Caption = "���ڴ����û�:" & strUser & "(" & rsPerson!���� & ")"
                        If CreateApptionUser(Val(rsPerson!id & ""), strUser, strUser, True) Then
                            If blnHaveRis Then '֪ͨRIS�û�������
                                If Not gobjRIS.UserEdit(1, strUser) Then
                                    mblnRISMsg = True
                                End If
                            End If
                            Call ApplyOnePerson(strUser)
                        End If
                        strPre���� = rsPerson!���� & ""
                        objPercent.LoopPercent
                        lblStep.Caption = prg.value & "%"
                        rsPerson.MoveNext
                    Loop
                    lblProssCaption.Caption = "": lblStep.Caption = "": picProcess.Visible = False
                End If
            End If
        End If
    End If
    If mblnRISMsg Then
        MsgBox "��ǰ������Ӱ����Ϣϵͳ�ӿڣ� ������Ӱ����Ϣϵͳ�ӿ�(UserEdit)δ���óɹ�������ϵ����Ա��", vbInformation, gstrSysName
    End If
    mblnSucceed = True
    If txtUserName.Enabled = False Then '�޸��û�
        strNote = "": strCheck = "": strNoCheck = ""
        If txtName.Tag <> txtNO.Tag And txtNO.Tag <> "" Then
            strNote = "��Ӧ��Ա�ɡ�" & txtUserName.Tag & "���޸�Ϊ��" & txtName.Text & "����"
            txtUserName.Tag = ""
        End If
        mrsRole.Filter = "�ı�=1"
        If mrsRole.RecordCount > 0 Then
            Do While Not mrsRole.EOF
                If mrsRole!��ѡ = 0 Then  'ȡ����ѡ�Ľ�ɫ
                    strNoCheck = IIf(strNoCheck = "", "ȡ����ѡ�Ľ�ɫ�У�", strNoCheck & "��") & mrsRole!RoleName
                Else  '��ӹ�ѡ�Ľ�ɫ
                    strCheck = IIf(strCheck = "", "��ӹ�ѡ�Ľ�ɫ�У�", strCheck & "��") & mrsRole!RoleName
                End If
                mrsRole.MoveNext
            Loop
        End If
        strNote = strNote & IIf(strCheck = "", "", strCheck & "��") & IIf(strNoCheck = "", "", strNoCheck & "��")
        If mstrCreateUserList <> "" Then
            strNote = strNote & "������ӵ��û��У�" & mstrCreateUserList
            mstrCreateUserList = ""
        End If
                
        '������Ҫ������־
        If strNote <> "" Then
            Call SaveAuditLog(2, "�޸��û�", txtUserName.Text & "��" & strNote)
        End If
        If optApplyTo(ATE_��ǰ��Ա).value Then
            mstrItem = txtNO.Text & "|" & txtName.Text & "|" & txtDept.Text
        End If
        Unload Me
    Else '����
        mrsRole.Filter = "��ѡ=1"
        If mrsRole.RecordCount > 0 Then
            Do While Not mrsRole.EOF
                strNote = IIf(strNote = "", "", strNote & "��") & mrsRole!RoleName
                mrsRole.MoveNext
            Loop
        End If
        '������Ҫ������־
        If mstrCreateUserList <> "" Then
            Call SaveAuditLog(1, "�����û�", mstrCreateUserList & "����ӵĽ�ɫΪ��" & strNote)
        End If
        
        txtUserName.Text = "": txtPasswd.Text = "": txtVerify.Text = ""
        txtNO.Text = "": txtNO.Tag = "": txtName.Text = ""
        txtDept.Text = "": mstrCreateUserList = ""
        cmdSelect.SetFocus
    End If
End Sub

Private Sub cmdSelect_Click()
    If Not LoadPerson Then
        MsgBox "δ�ҵ�" & IIf(mstrUser = "", "��δ�����û�����Ա��", "������Ա��"), vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNO Is ActiveControl Then
            KeyAscii = 0
        Else
            PressKey vbKeyTab
        End If
    ElseIf KeyAscii = Asc("'") Or Chr(KeyAscii) = "@" Or Chr(KeyAscii) = " " Or Chr(KeyAscii) = "\" Or Chr(KeyAscii) = """" Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    If Not InitBaseData Then GoTo errEnd
    '�����н�ɫ�Ķ�Ӧ��ģ��
    If Not GetModuleAndRole Then GoTo errEnd
    If Not InitUser() Then GoTo errEnd
    'ѡ�����н�ɫ
    cboRoleGroups.ListIndex = 0
    Exit Sub
errEnd: 'ǿ�ƹرմ���
    On Error Resume Next
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsModule Is Nothing Then Set mrsModule = Nothing
    Call SaveSetting("ZLSOFT", "�û�����", "��ʽ", IIf(optMode(0).value, "0", "1"))
End Sub

Private Sub tvwPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call tvwPerson_DblClick
End Sub

Private Sub tvwPerson_LostFocus()
    DoEvents
    If cmdSelect Is ActiveControl Or txtNO Is ActiveControl Or tvwPerson Is ActiveControl Then Exit Sub
    If txtNO.Tag = "" And lblNO.Tag <> "" Then
        txtNO.Tag = Split(lblNO.Tag, ",")(0)
        txtNO.Text = Split(lblNO.Tag, ",")(1)
        txtName.Text = lblName.Tag
    End If
    tvwPerson.Visible = False
    tvwPerson.Nodes.Clear
End Sub

Private Sub tvwPerson_DblClick()
    If tvwPerson.SelectedItem Is Nothing Then Exit Sub
    If tvwPerson.SelectedItem.Tag <> 2 Then Exit Sub
    'ѡ�������Ա�ڵ�
    Call ChangePerson(tvwPerson.SelectedItem)
End Sub

Private Sub txtNO_GotFocus()
    SelAll txtNO
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    txtNO.Tag = ""
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lblNO.Tag <> "" Then
        If txtNO.Text = Split(lblNO.Tag, ",")(1) Then
            txtNO.Tag = Split(lblNO.Tag, ",")(0)
            Exit Sub
        End If
    End If
    LoadPerson (txtNO.Text)
    If txtNO.Tag = "" And txtNO.Text <> "" And Not tvwPerson.Visible Then
        MsgBox "û���ҵ��κ�" & IIf(mstrUser = "", "��δ�����û���", "") & "��Ա��Ϣ��������������룡"
        If lblNO.Tag <> "" Then
            txtNO.Tag = Split(lblNO.Tag, ",")(0)
            txtNO.Text = Split(lblNO.Tag, ",")(1)
        End If
        SelAll txtNO
    End If
End Sub

Private Sub txtno_LostFocus()
    If tvwPerson.Visible And Not cmdSelect Is ActiveControl And Not txtNO Is ActiveControl And Not tvwPerson Is ActiveControl Then
        tvwPerson.Visible = False
        txtNO.SetFocus
        txtNO.SelStart = 0
        Me.txtNO.SelLength = Len(txtNO.Text)
    ElseIf tvwPerson.Visible Then
        tvwPerson.SetFocus
    End If
End Sub

Private Sub txtno_Validate(Cancel As Boolean)
    If cmdSelect Is ActiveControl Or tvwPerson.Visible Then Exit Sub
     '�г����ű�Ͷ�Ӧ��Ա
    If txtNO.Tag = "" Then
        If lblNO.Tag <> "" Then
            If txtNO.Text = Split(lblNO.Tag, ",")(1) Then
                Me.txtNO.Tag = Split(lblNO.Tag, ",")(0)
                Exit Sub
            End If
        End If
        LoadPerson (txtNO.Text)
    End If
    
    If txtNO.Tag = "" And txtNO.Text <> "" And Not tvwPerson.Visible Then
        MsgBox "û���ҵ��κ�" & IIf(mstrUser = "", "��δ�����û���", "") & "��Ա��Ϣ��������������룡"
        If lblNO.Tag <> "" Then
            txtNO.Tag = Split(lblNO.Tag, ",")(0)
            txtNO.Text = Split(lblNO.Tag, ",")(1)
        End If
        Cancel = True
        SelAll txtNO
    End If
End Sub

Private Sub txtPasswd_GotFocus()
    SelAll txtPasswd
End Sub

Private Sub txtSearch_Change()
    Call FillRole
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("*") Or KeyAscii = Asc("_") Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUserName_Change()
    txtUserName = UCase(txtUserName)
    txtUserName.SelStart = Len(txtUserName)
End Sub

Private Sub txtUserName_GotFocus()
    SelAll txtUserName
End Sub

Private Sub txtVerify_GotFocus()
    SelAll txtVerify
End Sub

Private Sub vsRole_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngNameCol As Long
    With vsRole
        lngNameCol = (Col \ 2) * 2
        If Col = lngNameCol Then
            If .Cell(flexcpChecked, Row, Col) = flexUnchecked Then
                .Cell(flexcpChecked, Row, Col + 1) = flexUnchecked
            End If
        Else
            If .Cell(flexcpChecked, Row, Col) = flexChecked Then
                .Cell(flexcpChecked, Row, Col - 1) = flexChecked
            End If
        End If
        Call RecUpdate(mrsRole, "RoleName='" & .TextMatrix(Row, lngNameCol) & "'", "��ѡ", IIf(.Cell(flexcpChecked, Row, lngNameCol) = flexUnchecked, 0, 1), "ת��", IIf(.Cell(flexcpChecked, Row, lngNameCol + 1) = flexUnchecked, 0, 1))
        '��ѡת�ڣ��Լ���ѡ��Ȩ�����µ���ģ��չʾ
        If Col = lngNameCol Or Col <> lngNameCol And .Cell(flexcpChecked, Row, lngNameCol + 1) = flexChecked Then
            Call FillModule
        End If
    End With
End Sub

Private Sub vsRole_BeforeSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long, Cancel As Boolean)
    If vsRole.TextMatrix(NewRowSel, NewColSel - (NewColSel Mod 2)) = "" Then Cancel = True
End Sub

Private Sub vsRole_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsRole.BackColor = vsRole.BackColorFixed Then
        Cancel = True
    ElseIf vsRole.TextMatrix(Row, Col - (Col Mod 2)) = "" Then
        Cancel = True
    End If
End Sub

'==============================================================
'==˽�з���
'==============================================================
Private Function InitBaseData() As Boolean
'���ܣ���ʼ��������Χ����Ա����,�û����÷�ʽ
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strInfo As String
    '�û����÷�ʽ
    optMode(0).value = Val(GetSetting("ZLSOFT", "�û�����", "��ʽ", "0")) = 0
    optMode(1).value = Not Me.optMode(0).value
    '��ʼ��������Χ
    With cboWorkRange
        .addItem "      ����": .ItemData(.NewIndex) = 1
        .addItem "      סԺ": .ItemData(.NewIndex) = 2
        .addItem "������סԺ": .ItemData(.NewIndex) = 3
    End With
    '��ʼ����Ա����
    On Error GoTo errh:
    strInfo = "��Ա����"
RUMMan:
    strSQL = "Select ����, ���� From " & mstr������ & ".��Ա���ʷ���"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    cboManPros.Clear
    With rsTmp
        Do While Not .EOF
            cboManPros.addItem !����
            .MoveNext
        Loop
    End With
    strInfo = "��ɫ����"
RUMGroup:
    strSQL = "Select '���з���' ����, Count(1) ����, 2 ��ʶ" & vbNewLine & _
            "From Zlroles" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select Nvl(b.����,'δ����') ����, Count(1) ����, Decode(b.����, Null, 1, 0) ��ʶ" & vbNewLine & _
            "From Zltools.Zlroles a, Zltools.Zlrolegroups b" & vbNewLine & _
            "Where a.���� = b.��ɫ(+)" & vbNewLine & _
            "Group By b.����" & vbNewLine & _
            "Order By ����"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    With cboRoleGroups
        .Clear
        rsTmp.Filter = "���� = '���з���' And ��ʶ = 2"
        .addItem "���н�ɫ" & "(" & rsTmp!���� & ")"
        rsTmp.Filter = "���� = 'δ����'"
        If rsTmp.RecordCount <> 0 Then
            .addItem "δ����" & "(" & rsTmp!���� & ")"
        End If
        rsTmp.Filter = "��ʶ = 0"
        Do While Not rsTmp.EOF
            .addItem rsTmp!���� & "(" & rsTmp!���� & ")"
            rsTmp.MoveNext
        Loop
    End With
    InitBaseData = True
    Exit Function
errh:
    If strInfo <> "" Then
        If MsgBox("װ��" & strInfo & "ʱ�������´���" & vbCrLf & vbCrLf & _
                    err.Description & vbCrLf & vbCrLf & "��Ҫ����һ����", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
            err.Clear
            If strInfo = "��Ա����" Then
                GoTo RUMMan
            Else
                GoTo RUMGroup
            End If
        End If
    Else
        MsgBox "  �����:" & err.Number & "  ��������:" & err.Description, vbInformation, gstrSysName
    End If
End Function

Private Function InitUser() As Boolean
'���ܣ���ʼ���û�����
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    '�Ȳ���DBA��Ҳ���ǵ�ǰϵͳ�������ߣ������޸Ķ�Ӧ��Ա
    If Not (gblnDBA Or mstr������ = UCase(gstrUserName)) Then
        cmdSelect.Enabled = False
    End If
    
    If mstrUser <> "" Then
        '��ǰ�û��������ߣ��Ͳ�����Խ�ɫ�����޸�
        If mstr������ = mstrUser Then
            cboRoleGroups.Enabled = False
            txtSearch.Enabled = False
            chkGranted.Enabled = False
            vsRole.BackColor = vsRole.BackColorFixed: vsRole.BackColorBkg = vsRole.BackColorFixed
            cboWorkRange.Enabled = False
            cboManPros.Enabled = False
            optApplyTo(ATE_��ǰ����).Enabled = False: optApplyTo(ATE_������Ա).Enabled = False: optApplyTo(ATE_��ǰ��Ա).Enabled = False
        End If
        '�����û������Ϣ
        txtUserName.Text = mstrUser: txtPasswd.Text = "12345678": txtVerify.Text = "12345678"
        txtUserName.Enabled = False: txtPasswd.Enabled = False: txtVerify.Enabled = False
        optMode(0).Enabled = False: optMode(1).Enabled = False
        On Error GoTo errh
ReLoad:
        strSQL = "Select c.Id, c.���, c.����, a.����, a.Id As ����id" & vbNewLine & _
                    "From " & mstr������ & ".���ű� a, " & mstr������ & ".������Ա b, " & mstr������ & ".��Ա�� c, " & mstr������ & ".�ϻ���Ա�� d" & vbNewLine & _
                    "Where a.Id = b.����id And (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And b.ȱʡ = 1 And b.��Աid = c.Id And" & vbNewLine & _
                    "      c.Id = d.��Աid And d.�û��� = '" & mstrUser & "'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        
        If rsTmp.RecordCount > 0 Then
            txtNO.Text = rsTmp!���
            txtName.Text = rsTmp!����
            txtUserName.Tag = rsTmp!����
            txtDept.Text = rsTmp!����
            txtDept.Tag = rsTmp!����id
            txtNO.Tag = rsTmp!id: txtName.Tag = rsTmp!id
            Call LocateManPros(rsTmp!id)
            Call LocateWorkRange(rsTmp!����id)
        End If
    End If
    InitUser = True
    Exit Function
errh:
    If MsgBox("װ����Աʱ�������´���" & vbCrLf & vbCrLf & _
                err.Description & vbCrLf & vbCrLf & "��Ҫ����һ����", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        err.Clear
        GoTo ReLoad
    End If
End Function


Private Function GetModuleAndRole() As Boolean
'���ܣ���ȡ��ɫ�����Լ���ɫ��Ӧ��ģ������
    Dim strInfo As String, strSQL As String
    
    strSQL = "Select User As Grantee, g.���� Role, g.ϵͳ, Decode(n.Granted_Role, Null, 0, 1) As Granted," & vbNewLine & _
            "       Decode(n.Admin_Option, 'YES', 1, 0) As Admin, b.����, Zlspellcode(Substr(g.����, 3)) As ����, Substr(g.����, 4) Rolename," & vbNewLine & _
            "       Decode(n.Granted_Role, Null, 0, 1) ��ѡ, Decode(n.Admin_Option, 'YES', 1, 0) ת��, 1 չʾ, 0 �ı�" & vbNewLine & _
            "From (Select ����, ϵͳ From Zlroles) g," & vbNewLine & _
            "     (Select Granted_Role, Admin_Option From Dba_Role_Privs Where Grantee = [1] And Granted_Role Like 'ZL_%') n," & vbNewLine & _
            "     Zlrolegroups b" & vbNewLine & _
            "Where g.���� = n.Granted_Role(+) And g.���� = b.��ɫ(+)" & vbNewLine & _
            "Order By g.����"

    strInfo = "��ɫ����"
    On Error GoTo errh
RUMRole:
    Set mrsRole = Nothing
    Set mrsRole = CopyNewRec(gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption, mstrUser))
    strSQL = "Select *" & vbNewLine & _
            "From (" & vbNewLine & _
            "       Select c.����, c.���, a.��ɫ, a.���, a.����, b.����, b.˵��" & vbNewLine & _
            "From Zlrolegrant a, Zlprograms b, Zlsystems c" & vbNewLine & _
            "Where a.��� = b.��� And Nvl(a.ϵͳ, 0) = Nvl(b.ϵͳ, 0) And b.ϵͳ = c.���(+)" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select t.����, t.���, r.Grantee ��ɫ, Null ���, Null ����, t.���� ����, t.˵��" & vbNewLine & _
            "From (Select s.����, s.���, s.������, b.����, b.˵�� From Zlsystems s, Zlbasecode b Where b.ϵͳ = s.���) t," & vbNewLine & _
            "     (Select Grantee, Owner, Table_Name" & vbNewLine & _
            "       From User_Tab_Privs" & vbNewLine & _
            "       Where Privilege In ('SELECT', 'INSERT', 'UPDATE', 'DELETE')" & vbNewLine & _
            "       Group By Grantee, Owner, Table_Name" & vbNewLine & _
            "       Having Count(Privilege) = 4) r" & vbNewLine & _
            "Where t.������ = r.Owner And t.���� = r.Table_Name" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select s.����, s.���, r.Grantee ��ɫ, Null ���, Null ����, f.������ || '(' || f.������ || ')' ����, f.˵��" & vbNewLine & _
            "From Zlsystems s, Zlfunctions f, User_Tab_Privs r" & vbNewLine & _
            "       Where f.ϵͳ = s.��� And s.������ = r.Owner And Upper(f.������) = r.Table_Name And r.Privilege = 'EXECUTE')" & vbNewLine & _
            "       Order By ���, ���"

    strInfo = "��ɫģ������"
RUMModule:
    Set mrsModule = Nothing
    Set mrsModule = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    GetModuleAndRole = True
    Exit Function
errh:
    If MsgBox("װ��" & strInfo & "ʱ�������´���" & vbCrLf & vbCrLf & _
                err.Description & vbCrLf & vbCrLf & "��Ҫ����һ����", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        err.Clear
        If strInfo = "��ɫ����" Then
            GoTo RUMRole
        Else
            GoTo RUMModule
        End If
    End If
End Function

Private Sub LocateManPros(ByVal lng��Աid As Long)
'���ܣ���λ��ǰ��Ա����Ա����
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Integer
    
    If mstr������ = mstrUser Then Exit Sub
    strSQL = "Select ��Ա���� From ��Ա����˵�� Where ��Աid= " & lng��Աid
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)

    With cboManPros
        .ListIndex = -1
        If rsTmp.EOF Then Exit Sub
        For i = 0 To .ListCount
            If .List(i) = Nvl(rsTmp!��Ա����) Then
                .ListIndex = i
                Exit For
            End If
        Next
    End With
End Sub

Private Sub LocateWorkRange(ByVal lng����ID As Long)
'���ܣ���λ������Χ
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim blnLock As Boolean
    If mstr������ = mstrUser Then Exit Sub
    If lng����ID <> 0 Then
        strSQL = "Select 1 From " & mstr������ & ".��������˵�� Where ����id = " & lng����ID & " And �������� = '�ٴ�'"
        Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
        blnLock = rsTmp.EOF
    Else
        blnLock = True
    End If
    If blnLock Then
        cboWorkRange.ListIndex = -1
        cboWorkRange.Enabled = False
    Else
        cboWorkRange.ListIndex = 2
        cboWorkRange.Enabled = True
    End If
End Sub

Private Function FillRole() As Boolean
'����:��������䵽Role�б���
'����:strFilter-��������
'����:���سɹ�,����true,���򷵻�False
    Dim lngRow As Long, lngCol As Long, i As Long
    
    Call RecUpdate(mrsRole, "", "չʾ", 0)
    mrsRole.Filter = GetRoleFilter
    mrsRole.Sort = "Rolename"
    With vsRole
        .Redraw = flexRDNone: .Rows = .FixedRows
        .Rows = -1 * Int(-1 * Val(mrsRole.RecordCount / 3)) + .FixedRows
        .Row = 0: .Col = 0
        If .Rows > .FixedRows Then
            .CellBorderRange 0, 1, .Rows - 1, 1, vbBlack, 0, 0, 1, 0, 0, 0
            .CellBorderRange 0, 3, .Rows - 1, 3, vbBlack, 0, 0, 1, 0, 0, 0
            .Cell(flexcpPictureAlignment, .FixedRows, 1, .Rows - 1, 1) = 4
            .Cell(flexcpPictureAlignment, .FixedRows, 3, .Rows - 1, 3) = 4
            .Cell(flexcpPictureAlignment, .FixedRows, 5, .Rows - 1, 5) = 4
            For i = 0 To mrsRole.RecordCount - 1
                lngRow = i \ 3 + .FixedRows: lngCol = (i Mod 3) * 2
                .TextMatrix(lngRow, lngCol) = mrsRole!RoleName
                .Cell(flexcpChecked, lngRow, lngCol) = IIf(mrsRole!��ѡ = 1, flexChecked, flexUnchecked)
                .Cell(flexcpChecked, lngRow, lngCol + 1) = IIf(mrsRole!Admin = 1, flexChecked, flexUnchecked)
                mrsRole.Update "չʾ", 1
                mrsRole.MoveNext
            Next
            .Row = .FixedRows: .Col = 0
        End If
        .Redraw = flexRDDirect
    End With
    Call FillModule
    FillRole = True
End Function

Private Function GetRoleFilter() As String
    Dim strSearChar As String, strGroup As String
    Dim strFilter As String

    strSearChar = Replace(UCase(Trim(txtSearch.Text)), "'", "")
    strGroup = Mid(cboRoleGroups.Text, 1, InStrRev(cboRoleGroups.Text, "(") - 1)
    If strGroup = "���н�ɫ" Then strGroup = ""
    If strGroup = "δ����" Then
        strGroup = ""
        strFilter = "����=null"
    End If
    '���������
    strFilter = IIf(strGroup = "", strFilter, "����='" & strGroup & "'")
    strFilter = strFilter & IIf(glngSysNo = -1, "", IIf(strFilter = "", "", " And ") & "ϵͳ=" & glngSysNo)
    strFilter = strFilter & IIf(chkGranted.value = 1, IIf(strFilter <> "", " And ", "") & "��ѡ=1", "")
    If strSearChar <> "" Then
        If strFilter = "" Then
            strFilter = "RoleName Like '" & strSearChar & "%' OR ���� Like '" & strSearChar & "%'"
        Else
            strFilter = "(" & strFilter & " And RoleName Like '" & strSearChar & "%' ) OR (" & strFilter & " And ���� Like '" & strSearChar & "%' )"
        End If
    End If
    GetRoleFilter = strFilter
End Function

Private Sub FillModule()
'����ѡȡ�Ľ�ɫ�г����е�ģ�鼰����
    Dim strModule As String, strFun As String
    Dim strAllFun As String, strRole As String
    Dim lngSerialNumber As Long
    
    '��ȡ������չʾ����ѡ�Ľ�ɫ
    mrsRole.Filter = "չʾ=1 And ��ѡ=1"
    vsfModule.Rows = 1
    If mrsRole.RecordCount = 0 Then Exit Sub
    Do While Not mrsRole.EOF
        strRole = strRole & " OR (��ɫ='" & mrsRole!Role & "'" & IIf(glngSysNo = -1, "", " And ���=" & glngSysNo) & ")"
        mrsRole.MoveNext
    Loop
    strRole = Mid(strRole, Len(" OR "))
    '���˽�����չʾ����ѡ�Ľ�ɫ��Ӧ��ģ�飬������Ź�������
    mrsModule.Filter = strRole
    Do While Not mrsModule.EOF
        If lngSerialNumber <> Nvl(mrsModule!���, -1) Or IsNull(mrsModule!���) Then
            vsfModule.Rows = vsfModule.Rows + 1
            If IsNull(mrsModule!����) Then
                If mrsModule!��� < 100 Then
                    vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_ϵͳ) = "��������"
                Else
                    vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_ϵͳ) = "�Զ��屨��"
                End If
            Else
                vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_ϵͳ) = mrsModule!���� & "(" & mrsModule!��� & ")"
            End If
            vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_���) = mrsModule!��� & ""
            vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_����) = mrsModule!����
            vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_����) = IIf(mrsModule!���� & "" = "����", "", mrsModule!���� & "")
            vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_˵��) = mrsModule!˵�� & ""
            lngSerialNumber = Nvl(mrsModule!���, -1)
        Else
            If mrsModule!���� & "" <> "����" Then
                vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_����) = vsfModule.TextMatrix(vsfModule.Rows - 1, VMT_����) & IIf(IsNull(mrsModule!����), "", ",") & mrsModule!����
            End If
        End If
        mrsModule.MoveNext
    Loop
    Exit Sub
errh:
    MsgBox "������Ȩģ������ʱ��������,��ϸ�Ĵ�����Ϣ����:" & vbCrLf & "  �����:" & err.Number & "  ��������:" & err.Description, vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Function LoadPerson(Optional ByVal strMenInfo As String) As Boolean
'���ܣ���������չʾ��Աѡ����
    Dim rsDept As ADODB.Recordset, strSQL As String
    Dim rsMen As ADODB.Recordset
    Dim objNode As Node, objParent As Node
    Dim strKey As String, i As Long
    
    On Error GoTo errh
    tvwPerson.Nodes.Clear: tvwPerson.Visible = False: tvwPerson.Tag = ""
     If txtNO.Tag <> "" And strMenInfo = "" Then
        strKey = "P" & txtNO.Tag
     End If
    '��ȡƥ����Ա����δ��ȡ�����˳�,�����û�ʱmstrUser = ""��ֻ��ѯδ�������û�
    strSQL = "Select a.Id, a.���, a.����, b.����id" & vbNewLine & _
                "From " & mstr������ & ".��Ա�� a, " & mstr������ & ".���ű� c, " & mstr������ & ".������Ա b" & IIf(mstrUser = "", "," & mstr������ & ".�ϻ���Ա�� D", "") & vbNewLine & _
                "Where (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And a.Id = b.��Աid And b.ȱʡ = 1 And b.����id = c.Id And" & vbNewLine & _
                "      (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) " & IIf(mstrUser = "", " And B.��Աid = d.��Աid(+) And D.��Աid is null ", "")
    If strMenInfo <> "" Then
        strMenInfo = UCase(strMenInfo)
        If IsNumeric(strMenInfo) Then
            strSQL = strSQL & "And a.��� Like '" & strMenInfo & "%'"
        ElseIf IsCharAlpha(strMenInfo) Then
            strSQL = strSQL & "And a.���� Like '" & strMenInfo & "%'"
        ElseIf IsCharChinese(strMenInfo) Then
            strSQL = strSQL & "And a.���� Like '" & strMenInfo & "%'"
        Else
            strSQL = gstrSQL & "And (a.���� Like '" & strMenInfo & "%' OR a.���� Like '" & strMenInfo & "%')"
        End If
    End If
    Set rsMen = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    If rsMen.EOF Then Exit Function 'û�в�ѯ����Ա��չʾ
    '��ȡ�����ز��������б�
    strSQL = "Select Id, ����, ����, �ϼ�id" & vbNewLine & _
                "From " & mstr������ & ".���ű�" & vbNewLine & _
                "Where ���� <> '-' And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & vbNewLine & _
                "Start With �ϼ�id Is Null" & vbNewLine & _
                "Connect By Prior Id = �ϼ�id"
    Set rsDept = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    Do While Not rsDept.EOF
        If IsNull(rsDept!�ϼ�id) Then
            Set objNode = tvwPerson.Nodes.Add(, , "K" & rsDept!id, "��" & rsDept!���� & "��" & rsDept!����, "Dept", "Dept")
        Else
            Set objNode = tvwPerson.Nodes.Add("K" & rsDept!�ϼ�id, tvwChild, "K" & rsDept!id, "��" & rsDept!���� & "��" & rsDept!����, "Dept", "Dept")
        End If
        objNode.Tag = 0
        rsDept.MoveNext
    Loop
    '���ز��ŵ�������Ա
    Do While Not rsMen.EOF
        Set objNode = tvwPerson.Nodes.Add("K" & rsMen!����id, tvwChild, "P" & rsMen!id, "��" & rsMen!��� & "��" & rsMen!����, "User", "User")
        objNode.ForeColor = RGB(0, 0, 255)
        If strKey = "" Then
            strKey = objNode.Key
        End If
        '��Ǹ����������Ѿ������Ӽ����Ӽ����Ӽ�,�ڲ�ѯģʽ�£���չ���ڵ�
        If objNode.Parent.Tag = 0 Then
            Set objParent = objNode.Parent
            Do While Not objParent Is Nothing
                If objParent.Tag = 1 Then Exit Do
                objParent.Tag = 1 '���
                If strMenInfo <> "" Then objParent.Expanded = True
                Set objParent = objParent.Parent
            Loop
        End If
        objNode.Tag = 2
        rsMen.MoveNext
    Loop
    '�Ƴ�û���¼��ĸ��ڵ�
    i = 1
    Do
        If tvwPerson.Nodes(i).Tag = 0 Then
            tvwPerson.Nodes.Remove i
        Else
             i = i + 1
        End If
    Loop While (i <= tvwPerson.Nodes.Count)
    On Error Resume Next
    If tvwPerson.Nodes.Count = 0 Then Exit Function 'û�нڵ���չʾ
    LoadPerson = True
    DoEvents
    tvwPerson.Visible = True
    If strKey <> "" Then
        tvwPerson.Nodes(strKey).Selected = True
        tvwPerson.SelectedItem.EnsureVisible
    End If
    tvwPerson.SetFocus
    With tvwPerson
        .Top = Me.ScaleTop
        .Left = cmdSelect.Left + cmdSelect.Width + fraPerson.Left
        .Height = Me.ScaleHeight
        .ZOrder
    End With
    Exit Function
errh:
    If MsgBox("װ����Աʱ�������´���" & vbCrLf & vbCrLf & _
        err.Description & vbCrLf & vbCrLf & "��Ҫ����һ����", vbQuestion Or vbYesNo, gstrSysName) = vbYes Then
        Resume
    End If
End Function

Private Function VilidateData() As Boolean
'���ܣ������ݺϷ��Խ���У��
    Dim strUser As String
    
    strUser = UCase(Trim(txtUserName.Text))
    If strUser = "SYS" Or strUser = "SYSTEM" Then
        MsgBox "����ʹ��SYS��SYSTEM�û���", vbInformation, gstrSysName
        txtUserName.Text = "": Exit Function
    End If
    
    If txtUserName.Enabled Then      '�½��û�
        If Len(strUser) = 0 Then
            MsgBox "�������û�����", vbExclamation, gstrSysName
            txtUserName.SetFocus: Exit Function
        End If
        If Len(Trim(txtPasswd)) < 2 Then
            MsgBox "���������λ�ַ����ϵ��û����롣", vbExclamation, gstrSysName
            txtPasswd.SetFocus: Exit Function
        End If
        If StrIsValid(strUser, txtUserName.MaxLength) = False Then
            txtUserName.SetFocus
            Exit Function
        End If
        If StrIsValid(Trim(txtPasswd.Text), txtPasswd.MaxLength) = False Then
            txtPasswd.SetFocus: Exit Function
        End If
        If txtPasswd <> txtVerify Then
            MsgBox "�û���������֤���벻һ��", vbExclamation, gstrSysName
            txtPasswd = "": txtVerify = ""
            txtPasswd.SetFocus: Exit Function
        End If
    End If
    
    VilidateData = True
End Function

Private Function GetUserName(ByVal str���� As String, ByVal str���� As String, ByVal strPre���� As String) As String
'���ܣ�����һ�����ݿ��û���
    If optMode(0).value = True Then
        If strPre���� <> str���� Then
            GetUserName = str����
        Else
            GetUserName = str���� & str����
        End If
    Else
        '��U+��Ŵ���
        If IsNumeric(Left(str����, 1)) Then
            GetUserName = "U" & str����
        Else
            GetUserName = str����
        End If
    End If
End Function

Private Function GetOtherPerson() As ADODB.Recordset
'���ܣ�������������ȡ������Ӧ�õ�������
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strWorkRange As String, strManPros As String
    'Ϊʲô�����ٴ������޶���
    'û���ٴ���Ϣ������в������ʹ���
    On Error GoTo errh
    If cboWorkRange.ListIndex > -1 Then
        strWorkRange = " And e.�������� = '�ٴ�'"
        If cboWorkRange.ItemData(cboWorkRange.ListIndex) < 2 Then
            strWorkRange = strWorkRange & " And e.������� in (" & cboWorkRange.ItemData(cboWorkRange.ListIndex) & ",3) "
        End If
    End If
    strManPros = cboManPros.Text
    strSQL = "Select Distinct a.Id, a.���, Decode(a.����, Null, Zlspellcode(a.����), a.����) As ����, a.����" & vbNewLine & _
                    "From " & mstr������ & ".��Ա�� a, " & mstr������ & ".������Ա b, " & mstr������ & ".�ϻ���Ա�� c" & vbNewLine
    '������Ա���ʵĶ�ȡ
    strSQL = strSQL & IIf(strManPros = "", "", ", " & mstr������ & ".��Ա����˵�� d") & vbNewLine
    '���ӹ�����Χ��ȡ
    strSQL = strSQL & IIf(strWorkRange = "", "", ", " & mstr������ & ".��������˵�� e") & vbNewLine
    '������������,���Ӳ��Ź�������
     strSQL = strSQL & _
                    "Where a.Id = b.��Աid And a.Id = c.��Աid(+) And" & vbNewLine & _
                    "      (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) And b.ȱʡ = 1 And a.Id <> " & Val(txtNO.Tag) & vbNewLine & _
                    IIf(optApplyTo(ATE_��ǰ����).value, "  And B.����id = " & Val(txtDept.Tag), "")
     '������Ա���ʵĹ���
     strSQL = strSQL & _
                   IIf(cboManPros <> "", " And A.id=D.��Աid And D.��Ա����='" & cboManPros & "'", "")
    '���ӹ�����Χ�Ĺ���
     strSQL = strSQL & _
                   IIf(strWorkRange <> "", " And b.����id = e.����id " & strWorkRange, "")
    strSQL = strSQL & "And c.�û��� is Null  Order By ����"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    If rsTmp.RecordCount = 0 Then Exit Function 'û�в�ѯ������Ȩ��Ա���˳�
    Set GetOtherPerson = rsTmp
    Exit Function
errh:
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetGrantRole(ByRef strNormalRoles As String, ByRef strAdminRoles As String) As Boolean
'���ܣ���ȡӦ����Ȩ�Ľ�ɫ
    '��ȡ��ͨ��Ȩ��ת����Ȩ
    strNormalRoles = "": strAdminRoles = ""
    mrsRole.Filter = "��ѡ=1": mrsRole.Sort = "ת��"
    Do While Not mrsRole.EOF
        If mrsRole!ת�� = 1 Then
            strAdminRoles = strAdminRoles & "," & mrsRole!Role
        Else
            strNormalRoles = strNormalRoles & "," & mrsRole!Role
        End If
        mrsRole.MoveNext
    Loop
    strNormalRoles = Mid(strNormalRoles, 2): strAdminRoles = Mid(strAdminRoles, 2)
    'ȡ�����ˣ�����Ȩʹ��
    mrsRole.Filter = "": mrsRole.Sort = "Role"
End Function

Private Function SpellCode(ByVal strAsk As String) As String
    '-------------------------------------------------------------
    '���ܣ�����ָ���ַ�����ƴ������
    '��������SSC���ƣ�
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    Dim aryStard As Variant
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    '��
    aryStard = Split("��;��;��;��;��;��;��;��;;��;��;��;��;��;ž;��;��;��;��;��;;��;��;Ѽ;��", ";")
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbProperCase)         '��ȫ��ת��Ϊ��ǣ���λת��Ϊ��д
    
    strCode = ""
    For intBit = 1 To Len(strAsk)
        If Mid(strAsk, intBit, 1) = "��" Then
            blnCan = True
            strCode = strCode & "T"
        ElseIf Asc(Mid(strAsk, intBit, 1)) < 0 Then
            blnCan = True
            For iCount = 0 To UBound(aryStard)
                If Len(aryStard(iCount)) <> 0 Then
                    If StrComp(Mid(strAsk, intBit, 1), aryStard(iCount), vbTextCompare) = -1 Then
                        strCode = strCode & Chr(65 + iCount)
                        Exit For
                    ElseIf iCount = UBound(aryStard) Then
                        strCode = strCode & "Z"
                    End If
                End If
            Next
        Else
            If Mid(strAsk, intBit, 1) >= "A" And Mid(strAsk, intBit, 1) <= "Z" Then
                strCode = strCode & Mid(strAsk, intBit, 1)
            End If
        End If
        If Len(strCode) >= 10 Then Exit For
    Next
    SpellCode = strCode
    
End Function

Public Function ApplyOnePerson(ByVal strUser As String, Optional ByVal blnMsg As Boolean, Optional ByVal blnCurPerson As Boolean) As Boolean
'���ܣ�����ɫ��ȨӦ�ø�һ���û�
'������
'         strUser=�û���
'         strPwd=����
'         blnMsg=�Ƿ���Ϣ��ʾ
'���أ��Ƿ�ɹ�
    Dim strRSQL As String, strGSQL As String
    On Error Resume Next
    '���ջ�Ȩ�ޣ�Ȼ����Ȩ
    mrsRole.MoveFirst
    Do While Not mrsRole.EOF
        If blnCurPerson Then strRSQL = "Revoke " & mrsRole!Role & " From " & strUser
        If mrsRole!��ѡ = 1 Then
            strGSQL = "Grant " & mrsRole!Role & " to " & strUser & IIf(mrsRole!ת�� = 1, " With Admin Option", "")
        Else
            strGSQL = ""
        End If
        If Not gblnDBA And mrsRole!Grantee <> gstrUserName Then
            If blnCurPerson Then Call gclsBase.ExecuteCmdText(strRSQL, Me.Caption, gcnSystem, True)
            If err.Number <> 0 Then err.Clear
            If strGSQL <> "" Then Call gclsBase.ExecuteCmdText(strGSQL, Me.Caption, gcnSystem, True)
        Else
            If blnCurPerson Then Call gclsBase.ExecuteCmdText(strRSQL, Me.Caption, , True)
            If err.Number <> 0 Then err.Clear
            If strGSQL <> "" Then Call gclsBase.ExecuteCmdText(strGSQL, Me.Caption, , True)
        End If
        If err.Number <> 0 Then
            If blnMsg Then MsgBox "��ɫ����ʧ��,������Ϣ����:" & vbCrLf & err.Description, vbExclamation, gstrSysName
            err.Clear
            If blnMsg Then Exit Function
        End If
        mrsRole.MoveNext
    Loop
    '��¼��ɫ��Ȩ��Ϣ
    If err.Number <> 0 Then err.Clear
    Call ExecuteProcedure("Zl_Zluserroles_Add('" & strUser & "')", Me.Caption)
    If err.Number <> 0 Then
        If blnMsg Then MsgBox "��ɫ����ʧ��,������Ϣ����:" & vbCrLf & err.Description, vbExclamation, gstrSysName
        err.Clear
        Exit Function
    End If
    ApplyOnePerson = True
End Function


Private Function CreateApptionUser(ByVal lngID As Long, ByVal strUser As String, ByVal strPwd As String, Optional ByVal blnNew As Boolean, Optional ByVal blnMsg As Boolean) As Boolean
 '���ܣ������û�����Ȩ�����޸��û�����Ȩ
 'lngID=��ԱID
 'strUser=�û�
 'strPwd=����
 'blnNew=�Ƿ������û�
 'blnMsg=�Ƿ���Ϣ��ʾ
    Dim strError As String
    
    If Not blnNew Then
        gcnOracle.Execute "Delete From " & mstr������ & ".�ϻ���Ա�� Where �û���='" & strUser & "'"
    Else
        Call gobjRegister.CreateUser(gcnOracle, strUser, strPwd, strError)
        If strError = "" Then
            Call gclsBase.ExecuteCmdText("Grant Connect,Alter Session,Create Session,Create Synonym,Create Table,Create View,Create Sequence,Create Database Link,Create Cluster to " & strUser, Me.Caption, , True)
            Call AlterUserTableSpaces(gcnOracle, strUser)
        Else
            If blnMsg Then MsgBox "�û��������벻����Ҫ��������û��Ƿ��Ѿ�����!" & vbCrLf & "������Ϣ����:" & vbCrLf & strError, vbExclamation, gstrSysName
            If blnMsg Then Exit Function
        End If
    End If
    gcnOracle.Execute "Insert Into " & mstr������ & ".�ϻ���Ա��(�û���,��Աid) Values ('" & strUser & "'," & lngID & ")"
     If err.Number <> 0 Then
        err.Clear
        MsgBox "���ڹ�����ԭ���û������Ϣ(���š�����)����ʧ�ܡ�" & vbNewLine & err.Description, vbExclamation, gstrSysName
    End If
    mstrCreateUserList = IIf(mstrCreateUserList = "", "", mstrCreateUserList & "��") & strUser
    CreateApptionUser = True
End Function

Private Sub ChangePerson(ByVal objNode As Node)
'���ܣ��޸���Աʱ�������ݴ���
    Dim arrTmp As Variant
    
    txtNO.Tag = Val(Mid(objNode.Key, 2))
    arrTmp = Split(objNode.Text, "��")
    txtName.Text = Mid(objNode.Text, Len(arrTmp(0)) + 2)
    txtNO.Text = Mid(arrTmp(0), 2)
    lblNO.Tag = txtNO.Tag & "," & txtNO.Text
    lblName.Tag = txtName.Text
    arrTmp = Split(objNode.Parent.Text, "��")
    txtDept.Text = Mid(objNode.Parent.Text, Len(arrTmp(0)) + 2)
    txtDept.Tag = Mid(objNode.Parent.Key, 2)
    tvwPerson.Visible = False
    If txtUserName.Enabled Then
        If optMode(0).value Then
            txtUserName.Text = SpellCode(txtName.Text)
        Else
            If UCase(Left(txtNO.Text, 1)) >= "A" And UCase(Left(txtNO.Text, 1)) <= "Z" Then
                txtUserName.Text = txtNO.Text
            Else
                txtUserName.Text = "U" & txtNO.Text
            End If
        End If
        txtUserName.SetFocus
    Else
        If vsRole.Enabled Then vsRole.SetFocus
    End If
    Call LocateManPros(Val(txtNO.Tag))
    Call LocateWorkRange(Val(txtDept.Tag))
End Sub

