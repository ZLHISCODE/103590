VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmILLSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "����ѡ����"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   Icon            =   "frmILLSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraBottom 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   45
      TabIndex        =   13
      Top             =   4890
      Width           =   8880
      Begin VB.CommandButton cmdCommon 
         Caption         =   "���˳���(&P)"
         Height          =   350
         Index           =   1
         Left            =   100
         TabIndex        =   16
         Top             =   135
         Width           =   1230
      End
      Begin VB.CommandButton cmdUnUse 
         Caption         =   "ȡ������(&U)"
         Height          =   350
         Left            =   4485
         TabIndex        =   9
         Top             =   135
         Width           =   1230
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   2610
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   165
         Width           =   1590
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   6255
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   7350
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCommon 
         Caption         =   "���ҳ���(&M)"
         Height          =   350
         Index           =   0
         Left            =   1335
         TabIndex        =   7
         Top             =   135
         Width           =   1230
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   4245
      Left            =   3315
      TabIndex        =   4
      Top             =   615
      Width           =   5745
      _cx             =   10134
      _cy             =   7488
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
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmILLSelect.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   -1  'True
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
   Begin MSComctlLib.ImageList iimg16 
      Left            =   1125
      Top             =   3405
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
            Picture         =   "frmILLSelect.frx":06A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":0C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":11D4
            Key             =   "wubi"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmILLSelect.frx":176E
            Key             =   "spell"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTop 
      Height          =   645
      Left            =   0
      TabIndex        =   11
      Top             =   -75
      Width           =   9070
      Begin VB.TextBox txtLocate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   3840
         TabIndex        =   14
         ToolTipText     =   "������һ����F3��س�����λ�����F4"
         Top             =   225
         Width           =   1665
      End
      Begin VB.ComboBox cbo��� 
         Height          =   300
         Left            =   6765
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   2160
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   1005
         TabIndex        =   1
         Top             =   225
         Width           =   2160
      End
      Begin VB.Image imgCodeType 
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   5550
         Top             =   250
         Width           =   240
      End
      Begin VB.Label lblLocate 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   3360
         TabIndex        =   15
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   5970
         TabIndex        =   12
         Top             =   285
         Width           =   720
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ӧ����"
         Height          =   180
         Left            =   210
         TabIndex        =   0
         Top             =   285
         Width           =   720
      End
   End
   Begin VB.Frame fraLR 
      BorderStyle     =   0  'None
      Height          =   4245
      Left            =   3225
      MousePointer    =   9  'Size W E
      TabIndex        =   10
      Top             =   615
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvwTree_s 
      Height          =   4245
      Left            =   15
      TabIndex        =   3
      Top             =   630
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   7488
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "iimg16"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmILLSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const S_SUB As String = ",28," '��չ��
'�������½�
Private Const S_MAIN As String = ",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,24,25,27," ' ���� ��ͨ���ѡ��Χ
Private Const S_SSZD As String = ",23," '�����ж�
Private Const S_BLZD As String = ",2,"
Private Const S_ZYZD As String = ",26,"

'��ڲ���
Private mfrmParent As Object
Private mstr��� As String
Private mlng���˿���ID As Long
Private mstr�Ա� As String
Private mblnMultiSel As Boolean
Private mblnICD10 As Boolean
Private mbln����ϵͳ As Boolean


Private mrsList As ADODB.Recordset
Private mblnOK As Boolean
Private mstrLike As String
Private mlngPreDept As Long
Private mintPreClass As Integer
Private mstrPreNode As String
Private mint���� As Integer
Private mbln�����޸� As Boolean
Private mstrSel���� As String
Private mlngUserID As Long 'ҽ��id/����Աid
Private mInt���÷�Χ As Integer    '1-���ﲡ��;2-סԺ����;0-ȫԺ

Private mlngICD11 As Long '-1-�ж�ϵͳ������0-��ICD11,1-��ICD11¼�룬
Private mbln�� As Boolean '�Ƿ���¼�������룬true ¼�������룬false ��չ��

Private mlngDiagType As Long '���¼�����ͣ�
        '˵����1-��ҽ�������;2-��ҽ��Ժ���;3-��ҽ��Ժ���;5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;8-��ǰ���;9-�������;10-����֢;11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���;21-��ԭѧ���;22-Ӱ��ѧ���
        'Ŀǰֻ�������ࣺ7 �� ��7��

Private mstr�½� As String '��ǰ���˵ı��뷶Χ
Private mbytModel As Byte  '=1 �ٴ�·�����������õ���

Private mblnParICD11 As Boolean

Private Enum DiseaseCols
    ColSel = 0
    Col���� = 1
    col���� = 2 'ֻ�м���������
    Col���� = 3
    col˵�� = 4
    col���� = 5 'ֻ����ϱ�����
    Col����ID = 6
    Col���� = 7
    Col���id = 8 '��������ʱʹ�ã�������Ӧ�����
    Col����Id = 9 '��ϱ���ʱʹ��,��϶�Ӧ�ļ���
End Enum

Public Function ShowMe(frmParent As Object, ByVal str��� As String, ByVal lng���˿���ID As Long, Optional ByVal str�Ա� As String, _
    Optional ByVal blnMultiSel As Boolean, Optional ByVal blnICD10 As Boolean = True, Optional ByVal strSel���� As String, Optional ByVal lngSys As Long = 100, Optional ByVal intPatiType As Integer, _
    Optional ByVal lngICD11 As Long, Optional ByVal bln�� As Boolean, Optional ByVal lngDiagType As Long, Optional ByVal bytModel As Byte) As ADODB.Recordset
    
    mstr��� = str���
    mlng���˿���ID = lng���˿���ID
    mstr�Ա� = str�Ա�
    mblnMultiSel = blnMultiSel
    mblnICD10 = blnICD10
    mstrSel���� = strSel����
    mbln����ϵͳ = (lngSys \ 100 = 3)
    mlngICD11 = lngICD11
    mbln�� = bln��
    mlngDiagType = lngDiagType
    Set mfrmParent = frmParent
    mInt���÷�Χ = intPatiType
    mbytModel = bytModel
    Me.Show 1, frmParent
    
    If mblnOK Then Set ShowMe = mrsList
End Function

Private Sub cbo����_Click()
    Call SetControlEnabled
End Sub

Private Sub cbo����_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim intIdx As Integer, blnDo As Boolean
    Dim vRect As Variant, blnCancel As Boolean
        
    If cbo����.ListIndex = -1 Then Exit Sub
    If cbo����.ItemData(cbo����.ListIndex) = mlngPreDept And cbo����.ItemData(cbo����.ListIndex) <> -1 Then Exit Sub
    
    blnDo = True
    If cbo����.ItemData(cbo����.ListIndex) = -1 Then
        'ѡ����������
        strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
            " From ���ű� A,��������˵�� B" & _
            " Where A.ID=B.����ID And B.������� IN(2,3)" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " Order by A.����"
        vRect = gobjComlib.zlControl.GetControlRect(cbo����.hWnd)
        Set rsTmp = gobjComlib.zlDatabase.ShowSelect(Me, strSQL, 0, IIF(mblnICD10, "ѡ�񼲲�", "ѡ�����"), , , , , , True, vRect.Left, vRect.Top, cbo����.Height, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            intIdx = gobjComlib.cbo.FindIndex(cbo����, rsTmp!ID)
            '������Click�¼�,�ڱ��¼�����ʱһ������
            If intIdx <> -1 Then
                Call gobjComlib.zlControl.CboSetIndex(cbo����.hWnd, intIdx)
            Else
                cbo����.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����.ListCount - 1
                cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
                Call gobjComlib.zlControl.CboSetIndex(cbo����.hWnd, cbo����.NewIndex)
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п������ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
            End If
            '�ָ������еĿ���(������Click)
            intIdx = gobjComlib.cbo.FindIndex(cbo����, mlngPreDept)
            Call gobjComlib.zlControl.CboSetIndex(cbo����.hWnd, intIdx)
            blnDo = False
        End If
    End If
    mlngPreDept = cbo����.ItemData(cbo����.ListIndex)
    
    '��ȡ����
    If blnDo Then
        Call SetControlEnabled
        Call FillTreeData
    End If
End Sub

Private Sub cbo����_GotFocus()
    Call gobjComlib.zlControl.TxtSelAll(cbo����)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo����.ListIndex = -1 Then
            Call cbo����_Validate(blnCancel)
        
        End If
        If Not blnCancel Then
            Call cbo����_Validate(False)
            Call gobjComlib.ZLCommFun.PressKey(vbKeyTab)
        End If
        '����ϵͳû���ҵ�����
        If cbo����.ListIndex = -1 And mbln����ϵͳ Then cbo����.ListIndex = 0
    Else
        If mbln����ϵͳ Then KeyAscii = 0
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, intIdx As Long
    Dim vRect As Variant, blnCancel As Boolean
    Dim strInput As String
    
    If cbo����.ListIndex <> -1 Then Exit Sub '��ѡ��,���ô���
    If cbo����.Text = "" Then Cancel = True: Exit Sub '������
    
    On Error GoTo errH
    
    strInput = UCase(gobjComlib.ZLCommFun.GetNeedName(cbo����.Text))
    strSQL = "Select Distinct A.ID,A.����,A.����,A.����" & _
        " From ���ű� A,��������˵�� B" & _
        " Where A.ID=B.����ID And B.������� IN(2,3)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
        " Order by A.����"
    
    vRect = gobjComlib.zlControl.GetControlRect(cbo����.hWnd)
    Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(mblnICD10, "����ѡ��", "���ѡ��"), False, "", "", False, False, _
        True, vRect.Left, vRect.Top, cbo����.Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
    If Not rsTmp Is Nothing Then
        intIdx = gobjComlib.cbo.FindIndex(cbo����, rsTmp!ID)
        If intIdx <> -1 Then
            cbo����.ListIndex = intIdx
        Else
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����, cbo����.ListCount - 1
            cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
            cbo����.ListIndex = cbo����.NewIndex
        End If
    Else
        If Not blnCancel Then
            MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, gstrSysName
        End If
        Cancel = True: Exit Sub
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cbo���_Click()
    If mintPreClass = cbo���.ListIndex Then Exit Sub
    mintPreClass = cbo���.ListIndex
    
    Call FillTreeData
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCommon_Click(Index As Integer)
    Dim arrSQL As Variant, i As Long
    Dim strPar As String
    
    If Index = 0 Then '���ҳ���
        If cbo����.ListIndex = -1 Then
            MsgBox "��ָ����ǰ" & IIF(mblnICD10, "����", "���") & "�ĳ��ÿ��ҡ�", vbInformation, gstrSysName
            cbo����.SetFocus: Exit Sub
        End If
        If cbo����.ItemData(cbo����.ListIndex) = cbo����.ItemData(cbo����.ListIndex) Then
            MsgBox "��" & IIF(mblnICD10, "����", "���") & "�Ѿ�����Ϊ""" & cbo����.Text & """�ĳ���" & IIF(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
            cbo����.SetFocus: Exit Sub
        End If
        strPar = cbo����.ItemData(cbo����.ListIndex)
    ElseIf Index = 1 Then '���˳���
        If mlngUserID = cbo����.ItemData(cbo����.ListIndex) Then
            MsgBox "��" & IIF(mblnICD10, "����", "���") & "�Ѿ�����Ϊ���˵ĳ���" & IIF(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
            cbo����.SetFocus: Exit Sub
        End If
        strPar = "Null," & mlngUserID
    End If
    
    arrSQL = Array()
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And .RowData(i) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mblnICD10 Then
                        arrSQL(UBound(arrSQL)) = "zl_�����������_Insert(" & .RowData(i) & "," & strPar & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "zl_������Ͽ���_Insert(" & .RowData(i) & "," & strPar & ")"
                    End If
                End If
            Next
        End If
        If UBound(arrSQL) = -1 Then
            If .RowData(.Row) = 0 Then
                MsgBox "û��ѡ��" & IIF(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
                Exit Sub
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mblnICD10 Then
                arrSQL(UBound(arrSQL)) = "zl_�����������_Insert(" & .RowData(.Row) & "," & strPar & ")"
            Else
                arrSQL(UBound(arrSQL)) = "zl_������Ͽ���_Insert(" & .RowData(.Row) & "," & strPar & ")"
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call gobjComlib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
        
    MsgBox "�����á�", vbInformation, gstrSysName
    vsList.SetFocus
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim strFilter As String
    Dim i As Long
    
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And Val(.RowData(i)) <> 0 Then
                    strFilter = strFilter & " Or ��ĿID=" & .RowData(i)
                End If
            Next
            strFilter = Mid(strFilter, 5)
        End If
        If strFilter = "" Then
            If Val(.RowData(.Row)) = 0 Then
                MsgBox "û��ѡ��" & IIF(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
                Exit Sub
            End If
            strFilter = "��ĿID=" & .RowData(.Row)
        End If
        
        mrsList.Filter = strFilter
        If mrsList.EOF Then
            MsgBox "û��ѡ��" & IIF(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdUnUse_Click()
    Dim arrSQL As Variant, i As Long
    Dim strPar As String
    Dim strTmp As String
    
    If cbo����.List(cbo����.ListIndex) = IIF(mblnICD10, "���м���", "�������") Then 'ɾȫ��
        strPar = cbo����.ItemData(cbo����.ListIndex) & "," & mlngUserID
        strTmp = "���˳��ú�" & gobjComlib.ZLCommFun.GetNeedName(cbo����.Text)
    ElseIf cbo����.List(cbo����.ListIndex) = "���˳���" Then 'ɾ���˳���
        strPar = "Null," & mlngUserID
        strTmp = "���˳���"
    Else 'ɾ���ҳ���
        strPar = cbo����.ItemData(cbo����.ListIndex)
        strTmp = gobjComlib.ZLCommFun.GetNeedName(cbo����.Text)
    End If
    
    If MsgBox("ȷʵҪ��ѡ���" & IIF(mblnICD10, "����", "���") & "��" & strTmp & "��ȡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    arrSQL = Array()
    With vsList
        If mblnMultiSel Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 And .RowData(i) <> 0 Then
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    If mblnICD10 Then
                        arrSQL(UBound(arrSQL)) = "Zl_�����������_Delete(" & .RowData(i) & "," & strPar & ")"
                    Else
                        arrSQL(UBound(arrSQL)) = "Zl_������Ͽ���_Delete(" & .RowData(i) & "," & strPar & ")"
                    End If
                End If
            Next
        End If
        If UBound(arrSQL) = -1 Then
            If .RowData(.Row) = 0 Then
                MsgBox "û��ѡ��" & IIF(mblnICD10, "����", "���") & "��", vbInformation, gstrSysName
                Exit Sub
            End If
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If mblnICD10 Then
                arrSQL(UBound(arrSQL)) = "Zl_�����������_Delete(" & .RowData(.Row) & "," & strPar & ")"
            Else
                arrSQL(UBound(arrSQL)) = "Zl_������Ͽ���_Delete(" & .RowData(.Row) & "," & strPar & ")"
            End If
        End If
    End With
    
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call gobjComlib.zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    
    mstrPreNode = ""
    Call tvwTree_s_NodeClick(tvwTree_s.SelectedItem)
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        txtLocate_KeyPress (vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim blnDept As Boolean, blnHave As Boolean
    Dim blnDoc As Boolean

    mstr�½� = ""
    If mlngICD11 = 1 Then
        mblnICD10 = True
        mstr��� = "E"
        'ȷ������ķ�Χ
        If mbytModel = 1 Then
            If mlng���˿���ID > 0 Then
                If gobjComlib.sys.DeptHaveProperty(mlng���˿���ID, "��ҽ��") Then
                    mstr�½� = S_MAIN & "," & S_ZYZD
                Else
                    mstr�½� = S_MAIN
                End If
            Else
                mstr�½� = S_MAIN & S_ZYZD
            End If
        Else
            If mlngDiagType = 7 Then
                mstr�½� = S_SSZD
            ElseIf mlngDiagType = 6 Then
                mstr�½� = S_BLZD
            ElseIf mlngDiagType = 11 Or mlngDiagType = 12 Or mlngDiagType = 13 Then
                mstr�½� = S_ZYZD
            Else
                mstr�½� = S_MAIN
            End If
        
            If Not mbln�� Then
                If mlngDiagType = 11 Or mlngDiagType = 12 Or mlngDiagType = 13 Then
                    mstr�½� = S_ZYZD
                Else
                    mstr�½� = S_SUB
                End If
            End If
        End If
    End If
    
    '��������
    With vsList
        If mlngICD11 = 1 Then
            .ColHidden(col����) = True
        Else
            .ColHidden(col����) = Not mblnICD10
        End If
        .ColHidden(col����) = mblnICD10
        .Rows = 1: .Rows = .FixedRows + 1
    End With
    If Not mblnICD10 Then Me.Caption = "���ѡ����"
    Call gobjComlib.RestoreWinState(Me, App.ProductName, mfrmParent.Name & IIF(mblnICD10, 1, 0))
    
    If mbln����ϵͳ Then
        '����ϵͳ����ʾ������Ŀ
        lbl����.Caption = "����ѡ��"
        cmdCommon(0).Visible = False: cmdCommon(0).Enabled = False
        cbo����.Visible = False: cbo����.Enabled = False
        cmdUnUse.Left = cmdCommon(0).Left
    End If
    cbo����.AddItem IIF(mblnICD10, "���м���", "�������")
    cbo����.AddItem "���˳���"
    
    
    mblnOK = False
    mlngPreDept = -1
    mintPreClass = -1
    mstrPreNode = ""
    Set mrsList = Nothing
    mstrLike = IIF(Val(gobjComlib.zlDatabase.GetPara("����ƥ��")) = 0, "%", "") '����ƥ�䷽ʽ
    
    On Error GoTo errH
    Call gobjComlib.zlDatabase.GetUserInfo
    mlngUserID = UserInfo.ID
    cbo����.ItemData(cbo����.NewIndex) = mlngUserID '���˳�����Ŀ
    
    '����ϵͳ��������������
    If Not mbln����ϵͳ Then
        '����Ƿ��Ӧ����Ա
        If mlngUserID <> 0 Then
            strSQL = "select * from " & IIF(mblnICD10, "�����������", "������Ͽ���") & " where ��Աid=[1] and Rownum<2"
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngUserID)
            If Not rsTmp.EOF Then blnDoc = True
        End If
        
        '����Ƿ��ж�Ӧ����
        If Not blnDoc Then
            If mblnICD10 Then
                If mstr��� = "" Then
                    strSQL = "Select A.* From ����������� A,������Ա B,�ϻ���Ա�� C" & _
                        " Where A.����ID=B.����ID And B.��ԱID=C.��ԱID And C.�û���=User And Rownum=1"
                Else
                    strSQL = "Select A.* From ��������Ŀ¼ I,����������� A,������Ա B,�ϻ���Ա�� C" & _
                        " Where I.ID=A.����ID And A.����ID=B.����ID And B.��ԱID=C.��ԱID" & _
                        " And (I.����ʱ�� is Null Or I.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                        " And C.�û���=User And Instr([1],I.���)>0 And Rownum=1"
                End If
            Else
                If mstr��� = "" Then mstr��� = "1,2"
                strSQL = "Select A.* From �������Ŀ¼ I,������Ͽ��� A,������Ա B,�ϻ���Ա�� C" & _
                    " Where I.ID=A.���ID And A.����ID=B.����ID And B.��ԱID=C.��ԱID" & _
                    " And (I.����ʱ�� is Null Or I.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And C.�û���=User And Instr([1],I.���)>0 And Rownum=1"
            End If
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr���)
            If Not rsTmp.EOF Then blnDept = True
        End If
        
        '��ʾ��ǰ��Ա����
        strSQL = "Select A.ID,A.����,A.����,A.����,Max(Nvl(C.ȱʡ,0)) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C,�ϻ���Ա�� D" & _
            " Where A.ID=B.����ID And B.�������� IN('�ٴ�','���','����','����','����','Ӫ��')" & _
            " And A.ID=C.����ID And C.��ԱID=D.��ԱID And D.�û���=User" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " Group by A.ID,A.����,A.����,A.����" & _
            " Order by A.����"
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        
        If blnDoc Then Call gobjComlib.zlControl.CboSetIndex(cbo����.hWnd, cbo����.NewIndex)
        
        Do While Not rsTmp.EOF
            blnHave = True
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
            If blnDept Then
                If rsTmp!ID = mlng���˿���ID Then
                    Call gobjComlib.zlControl.CboSetIndex(cbo����.hWnd, cbo����.NewIndex)
                ElseIf cbo����.ListIndex = -1 And rsTmp!ȱʡ = 1 Then
                    Call gobjComlib.zlControl.CboSetIndex(cbo����.hWnd, cbo����.NewIndex)
                End If
            End If
            
            cbo����.AddItem rsTmp!����
            cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
            If rsTmp!ID = mlng���˿���ID Then
                cbo����.ListIndex = cbo����.NewIndex
            ElseIf cbo����.ListIndex = -1 And rsTmp!ȱʡ = 1 Then
                cbo����.ListIndex = cbo����.NewIndex
            End If
            
            rsTmp.MoveNext
        Loop
        cbo����.AddItem "<��������...>"
        cbo����.ItemData(cbo����.NewIndex) = -1
        
        If cbo����.ListCount > 0 And cbo����.ListIndex = -1 Then
            cbo����.ListIndex = 0
        End If
    End If
    
    If cbo����.ListIndex = -1 Then
        If Not blnDept Or Not blnHave Or Not blnDoc Then
            '���κμ�����Ӧ��������ʱ,������Ա�޶�Ӧ����ʱ��ȱʡ��ʾ���м���
            Call gobjComlib.zlControl.CboSetIndex(cbo����.hWnd, 0) '����ϵͳ����Ϊ���м���
        Else
            Call gobjComlib.zlControl.CboSetIndex(cbo����.hWnd, 1)
        End If
    End If

    '��ʾ�����������
    If mblnICD10 Then
        If mstr��� = "" Then
            strSQL = "Select ����,���,�Ƿ���� From ����������� Order by ���ȼ�"
        Else
            strSQL = "Select ����,���,�Ƿ���� From ����������� Where Instr([1],����)>0 Order by ���ȼ�"
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr���)
        Do While Not rsTmp.EOF
            cbo���.AddItem rsTmp!���� & ". " & rsTmp!���
            cbo���.ItemData(cbo���.NewIndex) = Nvl(rsTmp!�Ƿ����, 0)
            If mstr��� <> "" Then
                If rsTmp!���� & "" = Mid(mstr���, 1, 1) Then
                    Call gobjComlib.zlControl.CboSetIndex(cbo���.hWnd, cbo���.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Loop
        If mstr��� = "" Then Call gobjComlib.zlControl.CboSetIndex(cbo���.hWnd, 0)
        If cbo���.ListCount = 1 Then cbo���.Locked = True
    Else
        lbl���.Visible = False
        cbo���.Visible = False
    End If
    
    mint���� = Val(gobjComlib.zlDatabase.GetPara("���뷽ʽ"))
    mbln�����޸� = Val(gobjComlib.zlDatabase.GetPara("����ƥ�䷽ʽ�л�")) = 1
    
    If mint���� = 1 Then
        imgCodeType.Picture = iimg16.ListImages("wubi").Picture
        imgCodeType.Tag = "wubi"
    Else
        imgCodeType.Picture = iimg16.ListImages("spell").Picture
        imgCodeType.Tag = "spell"
    End If
    
    'ȱʡ��ȡ����
    Call FillTreeData
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraTop.Left = 0
    fraTop.Top = -75
    fraTop.Width = Me.ScaleWidth
    
    If fraTop.Width - cbo���.Width - 200 > 4135 Then
        cbo���.Left = fraTop.Width - cbo���.Width - 200
        lbl���.Left = cbo���.Left - lbl���.Width - 75
    End If
    
    fraBottom.Left = 0
    fraBottom.Top = Me.ScaleHeight - fraBottom.Height
    fraBottom.Width = Me.ScaleWidth
    
    If fraBottom.Width - cmdCancel.Width - 550 > 7000 Then
        cmdCancel.Left = fraBottom.Width - cmdCancel.Width - 800
        cmdOK.Left = cmdCancel.Left - cmdOK.Width
    End If
    
    tvwTree_s.Left = 0
    tvwTree_s.Top = fraTop.Top + fraTop.Height + 15
    tvwTree_s.Height = Me.ScaleHeight - tvwTree_s.Top - fraBottom.Height
    
    fraLR.Top = tvwTree_s.Top
    fraLR.Left = tvwTree_s.Left + tvwTree_s.Width
    fraLR.Height = tvwTree_s.Height
    
    vsList.Top = tvwTree_s.Top
    vsList.Left = IIF(tvwTree_s.Visible, fraLR.Left + fraLR.Width, 0)
    vsList.Width = Me.ScaleWidth - vsList.Left
    vsList.Height = tvwTree_s.Height
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gobjComlib.SaveWinState(Me, App.ProductName, mfrmParent.Name & IIF(mblnICD10, 1, 0))
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvwTree_s.Width + X < 1000 Or vsList.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvwTree_s.Width = tvwTree_s.Width + X
        vsList.Left = vsList.Left + X
        vsList.Width = vsList.Width - X
    End If
End Sub

Private Sub FillTreeData()
'���ܣ���ȡ�����������ݣ������ǿ��Ҷ�Ӧ����ֻ��Ӧ�ķ���
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim objNode As Node
    Dim strICD11 As String
    Dim str��� As String
    Dim lng������� As Long, lng֤����� As Long
    
    '�������
    Set mrsList = Nothing
    tvwTree_s.Nodes.Clear
    vsList.Rows = vsList.FixedRows
    vsList.Rows = vsList.FixedRows + 1
    
    'ICD-10����Ƿ��з���
    If mblnICD10 Then
        If cbo���.ItemData(cbo���.ListIndex) = 0 Then
            tvwTree_s.Visible = False
            fraLR.Visible = False
        Else
            tvwTree_s.Visible = True
            fraLR.Visible = True
        End If
        Call Form_Resize
    End If
    
    Screen.MousePointer = 11
    Me.Refresh
    
    On Error GoTo errH
    
    If mblnICD10 Then
        If cbo���.ItemData(cbo���.ListIndex) <> 0 Then 'Ϊ0��ʾ���ּ���û�з���
            If mlngICD11 = 1 Then
                If mstr�½� <> "" Then
                    strICD11 = " And instr('" & mstr�½� & "',','||A.�½�||',')>0"
                    If InStr(mstr�½�, S_ZYZD) > 0 Then
                        strSQL = "Select ��� From ����������� Where �½� = '26' And ���� = '��ͳҽѧ������TM1��' And ���� = 'L1-SA0'"
                        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "�����������")
                        If Not rsTmp.EOF Then
                            lng������� = Val("" & rsTmp!���)
                        End If
                        
                        strSQL = "Select ��� From ����������� Where �½� = '26' And ���� = '��ͳҽѧ֤��TM1��' And ���� = 'L1-SE7'"
                        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "�����������")
                        If Not rsTmp.EOF Then
                            lng֤����� = Val("" & rsTmp!���)
                        End If
                        If mbln�� Then
                            strICD11 = strICD11 & IIF(lng֤����� <> 0, " And a.��� Not In (Select e.��� From ����������� e Where e.�½� = '26' And e.��� >=" & lng֤����� & ")", "")
                        Else
                            strICD11 = strICD11 & IIF(lng֤����� <> 0 And lng������� <> 0, " And a.��� Not In (Select e.��� From ����������� e Where e.�½� = '26' And (e.��� >=" & lng������� & " And e.��� <" & lng֤����� & "))", "")
                        End If
                    End If
                End If
            End If
            
            If cbo����.ItemData(cbo����.ListIndex) = 0 Then
                strSQL = "Select a.ID,a.�ϼ�ID,a.���,decode(a.���,'E', a.����||' ',null) || a.���� as ���� From ����������� a Where a.���=[1]" & _
                    " And (a.����ʱ�� is Null Or a.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) " & strICD11 & vbNewLine & _
                    " Start With a.�ϼ�ID is Null Connect by Prior a.ID=a.�ϼ�ID Order by Level,a.���"
            Else
                strSQL = _
                    " Select Distinct B.����id From ����������� A, ��������Ŀ¼ B Where A.����id = B.ID" & _
                    IIF(cbo����.List(cbo����.ListIndex) = "���˳���", " And A.��Աid = [3]", " And A.����id = [2]") & _
                    IIF(mInt���÷�Χ = 0, "", " And (Nvl(B.���÷�Χ,0) = 0 Or B.���÷�Χ = " & CStr(mInt���÷�Χ) & ") ") & _
                    " And (B.����ʱ�� is Null Or B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
                strSQL = _
                    "Select Max(Level) as ��ID, a.ID, a.�ϼ�id, a.���, decode(a.���,'E', a.����||' ',null) || a.���� as ���� " & vbNewLine & _
                    "From ����������� a Where a.���=[1] " & strICD11 & "  And (a.����ʱ�� is Null Or a.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) " & vbNewLine & _
                    "Start With a.ID In (" & strSQL & ")" & vbNewLine & _
                    "Connect By Prior a.�ϼ�id = a.ID" & vbNewLine & _
                    "Group By a.ID, a.�ϼ�ID, a.���, a.����,a.���,a.����" & vbNewLine & _
                    "Order By ��ID Desc"
                strSQL = "Select a.ID, a.�ϼ�id, a.���, a.���� From (" & strSQL & ") a"
            End If
            Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Left(cbo���.Text, 1), cbo����.ItemData(cbo����.ListIndex), mlngUserID)
            Do Until rsTmp.EOF
                If "E" = Left(cbo���.Text, 1) Then
                    str��� = ""
                Else
                    str��� = "��" & rsTmp!��� & "��"
                End If
                If IsNull(rsTmp!�ϼ�ID) Then
                    Set objNode = tvwTree_s.Nodes.Add(, , "_" & rsTmp!ID, str��� & Trim(rsTmp!����), 1, 2)
                Else
                    Set objNode = tvwTree_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, str��� & Trim(rsTmp!����), 1, 2)
                End If
                rsTmp.MoveNext
            Loop
        End If
    Else
        If cbo����.ItemData(cbo����.ListIndex) = 0 Then
            strSQL = "Select ID,�ϼ�ID,����,���� From ������Ϸ��� Where Instr([1],���)>0" & _
                " Start With �ϼ�ID is Null Connect by Prior ID=�ϼ�ID Order by Level,����"
        Else
            strSQL = _
                " Select Distinct C.����ID From ������Ͽ��� A, �������Ŀ¼ B,����������� C" & _
                " Where A.���ID = B.ID And B.ID=C.���ID" & _
                " And (B.����ʱ�� is Null Or B.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) " & _
                IIF(mInt���÷�Χ = 0, "", " And (Nvl(B.���÷�Χ,0) = 0 Or B.���÷�Χ = " & CStr(mInt���÷�Χ) & ") ") & _
                IIF(cbo����.List(cbo����.ListIndex) = "���˳���", " And A.��Աid = [3]", " And A.����id = [2]")
            strSQL = _
                "Select Max(Level) as ��ID, ID, �ϼ�id, ����, ����" & vbNewLine & _
                "From ������Ϸ��� Where Instr([1],���)>0" & vbNewLine & _
                "Start With ID In (" & strSQL & ")" & vbNewLine & _
                "Connect By Prior �ϼ�id = ID" & vbNewLine & _
                "Group By ID, �ϼ�ID, ����, ����" & vbNewLine & _
                "Order By ��ID Desc"
            strSQL = "Select ID, �ϼ�id, ����, ���� From (" & strSQL & ")"
        End If
        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr���, cbo����.ItemData(cbo����.ListIndex), mlngUserID)
        Do Until rsTmp.EOF
            If IsNull(rsTmp!�ϼ�ID) Then
                Set objNode = tvwTree_s.Nodes.Add(, , "_" & rsTmp!ID, "[" & rsTmp!���� & "]" & Trim(rsTmp!����), 1, 2)
            Else
                Set objNode = tvwTree_s.Nodes.Add("_" & rsTmp!�ϼ�ID, 4, "_" & rsTmp!ID, "[" & rsTmp!���� & "]" & Trim(rsTmp!����), 1, 2)
            End If
            rsTmp.MoveNext
        Loop
    End If
    
    If tvwTree_s.Nodes.Count > 0 Then
        tvwTree_s.Nodes(1).Selected = True
        tvwTree_s.Nodes(1).Expanded = True
        tvwTree_s.Nodes(1).EnsureVisible
    End If
    
    Screen.MousePointer = 0
    Call FillListData
    Exit Sub
errH:
    Screen.MousePointer = 0
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub FillListData()
    Dim strSQL As String, strSQLTmp As String
    Dim str�Ա� As String
    Dim lng����ID As Long, str��� As String
    Dim i As Long
    Dim str���� As String
    Dim strICD11 As String
    Dim lng������� As Long, lng֤����� As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    vsList.Rows = vsList.FixedRows
    vsList.Rows = vsList.FixedRows + 1
    vsList.ColHidden(0) = Not mblnMultiSel
    
    If mstr�Ա� Like "*��*" Then
        str�Ա� = "��"
    ElseIf mstr�Ա� Like "*Ů*" Then
        str�Ա� = "Ů"
    End If
    
    If mlngICD11 = 1 Then
        If mstr�½� <> "" Then
            strICD11 = " And instr('" & mstr�½� & "',','||A.�½�||',')>0"
            If InStr(mstr�½�, S_ZYZD) > 0 Then
                strSQL = "Select ��� From ����������� Where �½� = '26' And ���� = '��ͳҽѧ������TM1��' And ���� = 'L1-SA0'"
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "�����������")
                If Not rsTmp.EOF Then
                    lng������� = Val("" & rsTmp!���)
                End If
                
                strSQL = "Select ��� From ����������� Where �½� = '26' And ���� = '��ͳҽѧ֤��TM1��' And ���� = 'L1-SE7'"
                Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "�����������")
                If Not rsTmp.EOF Then
                    lng֤����� = Val("" & rsTmp!���)
                End If
                If mbln�� Then
                    strICD11 = strICD11 & IIF(lng֤����� <> 0, " And a.��� Not In (Select e.��� From ����������� e Where e.�½� = '26' And e.��� >=" & lng֤����� & ")", "")
                Else
                    strICD11 = strICD11 & IIF(lng֤����� <> 0 And lng������� <> 0, " And a.��� Not In (Select e.��� From ����������� e Where e.�½� = '26' And (e.��� >=" & lng������� & " And e.��� <" & lng֤����� & "))", "")
                End If
            End If
        End If
    End If
        
    If mblnICD10 Then
        If cbo���.ItemData(cbo���.ListIndex) <> 0 Then 'Ϊ0��ʾ���ּ���û�з���
            If tvwTree_s.SelectedItem Is Nothing Then
                vsList.Row = 1: Screen.MousePointer = 0: Exit Sub
            End If
            lng����ID = Val(Mid(tvwTree_s.SelectedItem.Key, 2))
            strSQLTmp = " And (A.����id = [3] " & strICD11 & " Or" & vbNewLine & _
                        "      A.����id In (Select A.Id" & vbNewLine & _
                        "                  From ����������� A, ����������� B" & vbNewLine & _
                        "                  Where A.��� = [2] " & strICD11 & " And (A.�ϼ�id = B.Id Or B.�ϼ�id Is Null) And A.��� = B.��� And B.Id = [3]))"
        End If
    Else
        If tvwTree_s.SelectedItem Is Nothing Then
            vsList.Row = 1: Screen.MousePointer = 0: Exit Sub
        End If
        lng����ID = Val(Mid(tvwTree_s.SelectedItem.Key, 2))
    End If
    
    If cbo����.ItemData(cbo����.ListIndex) <> 0 Then
        If mblnICD10 Then '��������������
            If mbln����ϵͳ Then
                strSQL = "Select A.Id As ��Ŀid, A.����, A.���, A.����, Null ����ID, Null ��������, A.����, A.˵��, Null ����, A.����id, A.����, A.��Ч����, A.����, C.�Ƿ���,A.���� ��������, A.Id ����id,A.��� �������, Null ���id" & vbNewLine & _
                        "From ��������Ŀ¼ A, ����������� B, ����������� C " & vbNewLine & _
                        "Where A.Id = B.����id And A.��� = [2] And A.����id = C.Id(+)" & IIF(cbo����.List(cbo����.ListIndex) = "���˳���", " And b.��Աid = [5]", " ") & vbNewLine & _
                        "  And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIF(str�Ա� <> "", " And (A.�Ա�����=[4] Or A.�Ա����� is Null) ", " ") & IIF(mInt���÷�Χ = 0, "", " And (Nvl(A.���÷�Χ,0) = 0 Or A.���÷�Χ = " & CStr(mInt���÷�Χ) & ") ") & strSQLTmp
            Else
                strSQL = "Select A.Id As ��Ŀid, A.����, A.���, A.����, Null ����ID, Null ��������, A.����, A.˵��, Null ����, A.����id, A.����, A.��Ч����, A.����, C.�Ƿ���,A.���� ��������, A.Id ����id,A.��� �������,Max(D.���id) ���id" & vbNewLine & _
                        "From ��������Ŀ¼ A, ����������� B, ����������� C, ������϶��� D" & vbNewLine & _
                        "Where A.Id = B.����id And A.��� = [2] And A.����id = C.Id(+) And A.Id = D.����id(+) And" & vbNewLine & _
                        IIF(cbo����.List(cbo����.ListIndex) = "���˳���", " b.��Աid = [5] And ", "  b.����id = [1] And ") & _
                        " (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIF(str�Ա� <> "", " And (A.�Ա�����=[4] Or A.�Ա����� is Null) ", " ") & strSQLTmp & vbNewLine & _
                        IIF(mInt���÷�Χ = 0, "", " And (Nvl(A.���÷�Χ,0) = 0 Or A.���÷�Χ = " & CStr(mInt���÷�Χ) & ") ") & _
                        "Group By A.Id, A.����, A.���, A.����, A.����, A.˵��, A.����id, A.����, A.��Ч����, A.����, C.�Ƿ���,A.���"
            End If
        Else '���������
            strSQL = "Select A.Id As ��Ŀid, A.����, Null ���, Null ����,Null ����ID, Null ��������, A.����, A.˵��, A.����, C.����id, '' As ����, 0 ��Ч����, 0 ����, 0 �Ƿ���, Max(D.����id) ����id," & vbNewLine & _
                    "       A.Id ���id" & vbNewLine & _
                    "From �������Ŀ¼ A, ������Ͽ��� B, ����������� C, ������϶��� D" & vbNewLine & _
                    "Where A.Id = B.���id And A.Id = D.����id(+) And A.Id = C.���id And Instr([2], A.���) > 0 " & IIF(cbo����.List(cbo����.ListIndex) = "���˳���", " And b.��Աid = [5]", " And b.����id = [1]") & vbNewLine & _
                    " And (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) " & _
                    "     And C.����id In ((Select ID From ������Ϸ��� Where Instr([2], ���) > 0 And ID = [3] Or �ϼ�id = [3]))" & IIF(mInt���÷�Χ = 0, "", " And (Nvl(A.���÷�Χ,0) = 0 Or A.���÷�Χ = " & CStr(mInt���÷�Χ) & ") ") & vbNewLine & _
                    "Group By A.Id, A.����,A.����, A.˵��, A.����, C.����id"
            '��ȡ��Ӧ�ļ�������,����
            strSQL = "Select A.��Ŀid, A.����, B.���, B.����, Null  ����ID, Null ��������, A.����, A.˵��, Null ����, A.����id, A.����, A.��Ч����, A.����, A.�Ƿ���,B.���� ��������, B.Id ����id,B.��� �������,A.���id" & vbNewLine & _
                            "From (" & strSQL & ") A,��������Ŀ¼ B " & vbNewLine & _
                            "Where A.����id=B.ID(+) "
        End If
    Else
        If mblnICD10 Then '��������������
            If mbln����ϵͳ Then
                strSQL = "Select A.Id As ��Ŀid, A.����, A.���, A.����,Null ����ID, Null ��������, A.����, A.˵��, Null ����, A.����id, A.����,  A.��Ч����, A.����, C.�Ƿ���,A.���� ��������, A.Id ����id,A.��� �������, Null ���id" & vbNewLine & _
                    "From ��������Ŀ¼ A, ����������� C" & vbNewLine & _
                    "Where A.��� = [2] And A.����id = C.Id(+)  And" & vbNewLine & _
                    "      (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIF(str�Ա� <> "", " And (A.�Ա�����=[4] Or A.�Ա����� is Null) ", " ") & IIF(mInt���÷�Χ = 0, "", " And (Nvl(A.���÷�Χ,0) = 0 Or A.���÷�Χ = " & CStr(mInt���÷�Χ) & ") ") & strSQLTmp
            Else
                strSQL = "Select A.Id As ��Ŀid, A.����, A.���, A.����,Null ����ID, Null ��������, A.����, A.˵��, Null ����, A.����id, A.����,  A.��Ч����, A.����, C.�Ƿ���,A.���� ��������, A.Id ����id,A.��� �������," & vbNewLine & _
                        "       Max(B.���id) ���id" & vbNewLine & _
                        "From ��������Ŀ¼ A, ������϶��� B, ����������� C" & vbNewLine & _
                        "Where A.��� = [2] And A.Id = B.����id(+) And A.����id = C.Id(+)  And" & vbNewLine & _
                        "      (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & IIF(str�Ա� <> "", " And (A.�Ա�����=[4] Or A.�Ա����� is Null) ", " ") & strSQLTmp & vbNewLine & _
                        IIF(mInt���÷�Χ = 0, "", " And (Nvl(A.���÷�Χ,0) = 0 Or A.���÷�Χ = " & CStr(mInt���÷�Χ) & ") ") & _
                        "Group By A.Id, A.����, A.���, A.����, A.����, A.˵��, A.����id, A.����, A.��Ч����, A.����,A.���, C.�Ƿ���"
            End If
        Else '���������
            strSQL = "Select A.Id As ��Ŀid, A.����, Null ���, Null ����,Null ����ID, Null ��������, A.����, A.˵��, A.����, B.����ID, '' As ����, 0 ��Ч����, 0 ����, 0 �Ƿ���," & vbNewLine & _
                    "       Max(D.����id) ����id, A.Id ���id" & vbNewLine & _
                    "From �������Ŀ¼ A, ����������� B, ������϶��� D" & vbNewLine & _
                    "Where Instr([2], A.���) > 0 And A.Id = B.���id And A.Id = D.����id(+) And" & vbNewLine & _
                    "  (A.����ʱ�� Is Null Or A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And " & _
                    "      B.����id In (Select ID From ������Ϸ��� Where Instr([2], ���) > 0 And ID = [3] Or �ϼ�id = [3])" & vbNewLine & _
                    IIF(mInt���÷�Χ = 0, "", " And (Nvl(A.���÷�Χ,0) = 0 Or A.���÷�Χ = " & CStr(mInt���÷�Χ) & ") ") & _
                    "Group By A.Id, A.����, A.����, A.˵��, A.����, B.����ID"
            '��ȡ��Ӧ�ļ�������,����
            strSQL = "Select A.��Ŀid, A.����, B.���, B.����, Null  ����ID, Null ��������, A.����, A.˵��, Null ����, A.����id, A.����, A.��Ч����, A.����, A.�Ƿ���,B.���� ��������, B.Id ����id,B.��� �������,A.���id" & vbNewLine & _
                            "From (" & strSQL & ") A,��������Ŀ¼ B " & vbNewLine & _
                            "Where A.����id=B.ID(+) "
        End If
    End If
    If mblnICD10 Then
        str��� = Left(cbo���.Text, 1)
    Else
        str��� = mstr���
    End If
    strSQL = strSQL & " Order by A.����" & IIF(mblnICD10, ",A.���", "")
    Set mrsList = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cbo����.ItemData(cbo����.ListIndex), str���, lng����ID, str�Ա�, mlngUserID)
    
    If Not mrsList.EOF Then
        With vsList
            .Redraw = flexRDNone
            .Rows = .FixedRows + mrsList.RecordCount
            For i = 1 To mrsList.RecordCount
                .RowData(i) = Val(mrsList!��ĿID & "")
                str���� = mrsList!���� & ""
                .TextMatrix(i, Col����) = str����
                .TextMatrix(i, Col����) = mrsList!���� & ""
                .TextMatrix(i, Col����ID) = mrsList!����ID & ""
                .TextMatrix(i, Col����) = mrsList!���� & ""
                .TextMatrix(i, col˵��) = mrsList!˵�� & ""
                .TextMatrix(i, col����) = mrsList!���� & ""
                .TextMatrix(i, col����) = mrsList!���� & ""
                .TextMatrix(i, Col���id) = mrsList!���id & ""
                .TextMatrix(i, Col����Id) = mrsList!����Id & ""
                .Cell(flexcpData, i, Col����) = CStr(str����)
                If mstrSel���� <> "" Then
                    If InStr(mstrSel����, "," & str���� & ",") > 0 Then
                        .TextMatrix(i, ColSel) = 1
                    Else
                        .TextMatrix(i, ColSel) = 0
                    End If
                Else
                    .TextMatrix(i, ColSel) = 0
                End If
                
                If mblnICD10 Then
                    If str���� = .Cell(flexcpData, i - 1, Col����) Then
                        If Not IsNull(mrsList!���) Then
                            .TextMatrix(i, Col����) = .TextMatrix(i, Col����) & "." & mrsList!���
                            If .TextMatrix(i - 1, Col����) = .Cell(flexcpData, i - 1, Col����) And mrsList!��� = 2 Then
                                .TextMatrix(i - 1, Col����) = .TextMatrix(i - 1, Col����) & ".1"
                            End If
                        End If
                    End If
                End If
                
                mrsList.MoveNext
            Next
            .Redraw = flexRDDirect
        End With
    End If

    vsList.Row = 1: vsList.Col = 1
    Screen.MousePointer = 0
    Call vsList_AfterRowColChange(-1, -1, vsList.Row, vsList.Col)
    Exit Sub
errH:
    Screen.MousePointer = 0
    If gobjComlib.ErrCenter() = 1 Then
        Resume
    End If
    Call gobjComlib.SaveErrLog
End Sub

Private Sub imgCodeType_Click()
    If Not mbln�����޸� Then Exit Sub
    If imgCodeType.Tag = "spell" Then
        Call gobjComlib.zlDatabase.SetPara("���뷽ʽ", 1)
        mint���� = 1
        imgCodeType.Picture = iimg16.ListImages("wubi").Picture
        imgCodeType.Tag = "wubi"
    Else
        Call gobjComlib.zlDatabase.SetPara("���뷽ʽ", 0)
        mint���� = 0
        imgCodeType.Picture = iimg16.ListImages("spell").Picture
        imgCodeType.Tag = "spell"
    End If
End Sub

Private Sub tvwTree_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrPreNode = Node.Key Then Exit Sub
    mstrPreNode = Node.Key
    Call FillListData
End Sub

Private Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIF(IsNull(varValue), DefaultValue, varValue)
End Function

Private Sub txtLocate_GotFocus()
    gobjComlib.zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long
    Dim strSQL As String, str�Ա� As String
    Dim strInput As String
    Dim rsTmp As ADODB.Recordset
    Dim vRect As Variant
    Dim blnCancle As Boolean
    Dim strICD11 As String
    Dim lng������� As Long, lng֤����� As Long
    
    If KeyAscii = vbKeyReturn Then
        On Error GoTo errH
        strInput = UCase(Trim(txtLocate.Text))
        
        If Not mblnICD10 Then
            '���Ŀ¼
            If gobjComlib.ZLCommFun.IsCharChinese(strInput) Then
                strSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
            ElseIf gobjComlib.ZLCommFun.IsCharAlpha(strInput) Then
                strSQL = "B.���� Like [2] Or B.���� Like [2]"
            Else
                strSQL = "A.���� Like [1] Or B.���� Like [2]"
            End If
            strSQL = _
                " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����,D.����ID" & _
                " From �������Ŀ¼ A,������ϱ��� B,������Ͽ��� C,����������� D" & _
                " Where  A.ID=C.���ID(+) And A.ID=B.���ID AND a.Id = D.���id " & _
                " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) " & _
                IIF(Val(cbo����.ItemData(cbo����.ListIndex)) <> 0, " And C.����ID=[3]", "") & _
                " And B.����=[5] and instr([6],A.���)>0 And (" & strSQL & ")" & _
                " Order by A.����"
        Else
            
            If mlngICD11 = 1 Then
                If mstr�½� <> "" Then
                    strICD11 = " And instr('" & mstr�½� & "',','||A.�½�||',')>0"
                    If InStr(mstr�½�, S_ZYZD) > 0 Then
                        strSQL = "Select ��� From ����������� Where �½� = '26' And ���� = '��ͳҽѧ������TM1��' And ���� = 'L1-SA0'"
                        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "�����������")
                        If Not rsTmp.EOF Then
                            lng������� = Val("" & rsTmp!���)
                        End If
                        
                        strSQL = "Select ��� From ����������� Where �½� = '26' And ���� = '��ͳҽѧ֤��TM1��' And ���� = 'L1-SE7'"
                        Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, "�����������")
                        If Not rsTmp.EOF Then
                            lng֤����� = Val("" & rsTmp!���)
                        End If
                        If mbln�� Then
                            strICD11 = strICD11 & IIF(lng֤����� <> 0, " And a.����ID Not In (Select e.ID From ����������� e where e.�½�='26' And e.���>=" & lng֤����� & ")", "")
                        Else
                            strICD11 = strICD11 & IIF(lng֤����� <> 0 And lng������� <> 0, " And a.����ID Not In (Select e.ID From ����������� e Where e.�½� = '26' And (e.��� >=" & lng������� & " And e.��� <" & lng֤����� & "))", "")
                        End If
                    End If
                End If
            End If
    
            If mstr�Ա� Like "*��*" Then
                str�Ա� = "��"
            ElseIf mstr�Ա� Like "*Ů*" Then
                str�Ա� = "Ů"
            End If
            If gobjComlib.ZLCommFun.IsCharChinese(strInput) Then
                strSQL = "A.���� Like [2]" '���뺺��ʱֻƥ������
            ElseIf gobjComlib.ZLCommFun.IsCharAlpha(strInput) Then
                strSQL = "A.���� Like [2] Or " & IIF(mint���� = 0, "a.����", "a.�����") & " Like [2]"
            Else
                strSQL = "A.���� Like [1] Or A.���� Like [2]"
            End If
            strSQL = _
                " Select A.ID,A.ID as ��ĿID,A.����,A.����,A.����," & IIF(mint���� = 0, "a.����", "a.����� as ����") & ",A.˵��,A.����ID" & _
                " From ��������Ŀ¼ A,����������� B Where A.ID=B.����ID(+) " & _
                IIF(Val(cbo����.ItemData(cbo����.ListIndex)) <> 0, " And B.����ID=[3]", "") & _
                " And (" & strSQL & ") And a.���=[6]" & strICD11 & _
                IIF(str�Ա� <> "", " And (A.�Ա�����=[4] Or A.�Ա����� is NULL)", "") & _
                " And (A.����ʱ�� is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Order by A.����"
        End If
        vRect = gobjComlib.zlControl.GetControlRect(txtLocate.hWnd)
        
        Set rsTmp = gobjComlib.zlDatabase.ShowSQLSelect(Me, strSQL, 0, IIF(Not mblnICD10, "�������", "��������"), False, "", "", False, False, True, _
            vRect.Left, vRect.Bottom, 0, blnCancle, False, True, strInput & "%", mstrLike & strInput & "%", Val(cbo����.ItemData(cbo����.ListIndex)), str�Ա�, mint���� + 1, IIF(mblnICD10, Left(cbo���.Text, 1), mstr���))

        '���������뷽ʽ
        If blnCancle Then Exit Sub
        If rsTmp Is Nothing Then
            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
        Else
            '��λ
            If txtLocate.Tag <> txtLocate.Text Then
                lblLocate.Tag = ""
                txtLocate.Tag = txtLocate.Text
            End If
            
            lngStart = Val("" & lblLocate.Tag) + 1
            If lngStart >= vsList.Rows Then lngStart = 1
            'ȷ��������ڵ�
            If tvwTree_s.Visible Then
                tvwTree_s.Nodes("_" & rsTmp!����ID).Selected = True
                tvwTree_s_NodeClick tvwTree_s.Nodes("_" & rsTmp!����ID)
            End If
            'ȷ�� VSLIST ��Ŀ
            For i = lngStart To vsList.Rows - 1
                If Val(vsList.RowData(i) & "") = Val(rsTmp!ID & "") Then
                    vsList.Row = i
                    vsList.TopRow = i
                    lblLocate.Tag = i
                    vsList.SetFocus
                    Exit For
                End If
            Next
        End If
    End If
    Exit Sub
errH:
    If gobjComlib.ErrCenter() = 1 Then Resume
    Call gobjComlib.SaveErrLog
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnEnabled As Boolean
    
    Call SetControlEnabled
    
    '�������ݵ�����£�ֻ��ȡ�������������ҵĳ��ü���
    If vsList.RowData(vsList.Row) <> 0 Then
        blnEnabled = True
    End If
    cmdUnUse.Enabled = blnEnabled
End Sub

Private Sub vsList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 Then
        If Val(vsList.TextMatrix(Row, 0)) <> 0 Then
            vsList.Cell(flexcpBackColor, Row, 0, Row, vsList.Cols - 1) = &HC0FFFF
        Else
            vsList.Cell(flexcpBackColor, Row, 0, Row, vsList.Cols - 1) = vsList.BackColor
        End If
    End If
End Sub

Private Sub vsList_DblClick()
    If vsList.MouseRow >= vsList.FixedRows Then
        vsList.TextMatrix(vsList.RowSel, ColSel) = 1
        Call cmdOK_Click
    End If
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    ElseIf KeyAscii = 32 Then
        If mblnMultiSel And vsList.Col > 0 And vsList.RowData(vsList.Row) <> 0 Then
            vsList.TextMatrix(vsList.Row, 0) = IIF(Val(vsList.TextMatrix(vsList.Row, 0)) = 0, 1, 0)
        End If
    End If
End Sub

Private Sub SetControlEnabled()
    Dim blnEnabled As Boolean
    Dim bln���˳��� As Boolean
    
    '��Ϊ���õĿ�����
    blnEnabled = True: bln���˳��� = True
    If cbo����.ListIndex = -1 Then
        blnEnabled = False
    ElseIf cbo����.ListIndex <> -1 And cbo����.ListIndex <> -1 Then
        If cbo����.ItemData(cbo����.ListIndex) = cbo����.ItemData(cbo����.ListIndex) Then
            blnEnabled = False
        End If
        If cbo����.List(cbo����.ListIndex) = "���˳���" Then
            bln���˳��� = False
        End If
    End If
    If blnEnabled Or bln���˳��� Then
        If vsList.Row >= vsList.FixedRows Then
            blnEnabled = IIF(blnEnabled, vsList.RowData(vsList.Row) <> 0, blnEnabled)
            bln���˳��� = IIF(bln���˳���, vsList.RowData(vsList.Row) <> 0, bln���˳���)
        End If
    End If
    
    cmdCommon(0).Enabled = blnEnabled ' ���ҳ���
    cmdCommon(1).Enabled = bln���˳��� ' ���˳���
    
    'ȷ����ť�Ŀ�����
    blnEnabled = True
    If vsList.Row >= vsList.FixedRows Then
        blnEnabled = vsList.RowData(vsList.Row) <> 0
    Else
        blnEnabled = False
    End If
    cmdOK.Enabled = blnEnabled
End Sub

Private Sub vsList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    With vsList
        lngRow = .MouseRow
        If lngRow >= .FixedRows Then
            Call gobjComlib.ZLCommFun.ShowTipInfo(.hWnd, .TextMatrix(lngRow, col˵��), True)     '·������Ŀ�����ԭ��
        Else
            Call gobjComlib.ZLCommFun.ShowTipInfo(.hWnd, "")
        End If
    End With
End Sub

Private Sub vsList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsList.RowData(Row) = 0 Then
        Cancel = True
    ElseIf Col <> 0 Then
        Cancel = True
    End If
End Sub

