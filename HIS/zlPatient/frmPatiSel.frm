VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiSel 
   Caption         =   "����ѡ��"
   ClientHeight    =   4650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   Icon            =   "frmPatiSel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6945
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   4350
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   529
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   1875
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSel.frx":058A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3540
      Left            =   2100
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3540
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   45
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   3600
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   6350
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   476
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   6945
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3750
      Width           =   6945
      Begin VB.CommandButton cmdFilter 
         Caption         =   "����(&F)"
         Height          =   350
         Left            =   195
         TabIndex        =   4
         ToolTipText     =   "ɸѡ���������Ĳ���(Ctrl+F)"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "��λ(&G)"
         Height          =   350
         Left            =   1410
         TabIndex        =   5
         ToolTipText     =   "��λ�����������Ĳ�����(Ctrl+G)"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5445
         TabIndex        =   3
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4215
         TabIndex        =   2
         Top             =   120
         Width           =   1100
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPati 
      Height          =   3585
      Left            =   2265
      TabIndex        =   1
      Top             =   75
      Width           =   4440
      _ExtentX        =   7832
      _ExtentY        =   6324
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmPatiSel.frx":06E4
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmPatiSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlng����ID As Long
Public mstrPrivs As String

Private mrsPati As ADODB.Recordset
Private mlngCurRow As Long, mlngTopRow As Long
Private mlngGo As Long, mblnDown As Boolean, mblnGo As Boolean
Private mstrFilter As String
Private mfrmFilter As frmPatiFilter
Private mfrmFind As frmPatiFind
Private mstrUnitIDs As String '����Ա���ڲ����������������

Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    �Ǽ�ʱ��B As Date
    �Ǽ�ʱ��E As Date
    ����ʱ��B As Date
    ����ʱ��E As Date
    ��Ժʱ��B As Date
    ��Ժʱ��E As Date
    ��Ժʱ��B As Date
    ��Ժʱ��E As Date
    סԺ�� As String
    �Ա� As String
    �ѱ� As String
    ���� As String
    Patient As String
End Type
Private SQLCondition As Type_SQLCondition

Private Sub cmdCancel_Click()
    mlng����ID = 0
    
    SaveWinState Me, App.ProductName
    
    Hide
End Sub

Private Sub cmdFilter_Click()
    Dim blnOK As Boolean
    
    blnOK = gblnOK
    mfrmFilter.mbytType = Val(mshPati.Tag)
    mfrmFilter.Show 1, Me
    If gblnOK Then
        With mfrmFilter
            mstrFilter = .mstrFilter
            SQLCondition.�Ǽ�ʱ��B = .dtp�Ǽ�B
            SQLCondition.�Ǽ�ʱ��E = .dtp�Ǽ�E
            SQLCondition.����ʱ��B = .dtp����B
            SQLCondition.����ʱ��E = .dtp����E
            
            SQLCondition.��Ժʱ��B = .dtp��ԺB
            SQLCondition.��Ժʱ��E = .dtp��ԺE
            SQLCondition.��Ժʱ��B = .dtp��ԺB
            SQLCondition.��Ժʱ��E = .dtp��ԺE
            
            SQLCondition.סԺ�� = Trim(.txtסԺ��.Text)
            SQLCondition.�Ա� = zlCommFun.GetNeedName(.cbo�Ա�.Text)
            SQLCondition.�ѱ� = zlCommFun.GetNeedName(.cbo�ѱ�.Text)
            SQLCondition.���� = zlCommFun.GetNeedName(.txt����.Text)
            
            If .PatiIdentify.GetCurCard.���� = "����" And .mlngPatiId = 0 And (.chk�Ǽ�.Value = 1 Or .chk��Ժ.Value = 1 Or .chk��Ժ.Value = 1) Then       '����
                SQLCondition.Patient = Trim(.PatiIdentify.Text) & "%"
            Else
                SQLCondition.Patient = IIf(.mlngPatiId <> 0, .mlngPatiId, .PatiIdentify.Text)
            End If
        End With
        
        Call ShowPatis(mstrFilter)
    End If
    gblnOK = blnOK
End Sub

Private Sub cmdFind_Click()
    Dim blnOK As Boolean
    blnOK = gblnOK
    mfrmFind.mbytType = Val(mshPati.Tag)
    mfrmFind.Show 1, Me
    If gblnOK Then Call SeekPati(mfrmFind.optHead)
    gblnOK = blnOK
End Sub

Private Sub cmdOK_Click()
    If Val(mshPati.TextMatrix(mshPati.Row, 0)) = 0 Then
        If glngSys Like "8??" Then
            MsgBox "û�пͻ�����ѡ��", vbInformation, gstrSysName
        Else
            MsgBox "û�в��˿���ѡ��", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    mlng����ID = Val(mshPati.TextMatrix(mshPati.Row, 0))
    
    SaveWinState Me, App.ProductName
    
    Hide
End Sub

Private Sub Form_Activate()
    mshPati.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            Call SeekPati(False)
        Case vbKeyReturn
            cmdOK_Click
        Case vbKeyEscape
            mblnGo = False
        Case vbKeyF
            If Shift = 2 Then cmdFilter_Click
        Case vbKeyG
            If Shift = 2 Then cmdFind_Click
    End Select
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    mlng����ID = 0
    If glngSys Like "8??" Then
        Caption = "�ͻ�ѡ��"
        tvw_s.Visible = False
        pic.Visible = False
    End If
    
    Set mfrmFilter = New frmPatiFilter
    Set mfrmFind = New frmPatiFind
        
    mstrUnitIDs = GetUserUnits
    Call InitUnits
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    tvw_s.Left = 0
    tvw_s.Top = 0
    tvw_s.Height = Me.ScaleHeight - picCmd.Height - sta.Height
    
    pic.Top = 0
    pic.Left = tvw_s.Width
    pic.Height = tvw_s.Height
    
    mshPati.Top = 0
    mshPati.Left = IIf(pic.Visible, pic.Width, 0) + IIf(tvw_s.Visible, tvw_s.Width, 0)
    mshPati.Width = Me.ScaleWidth - IIf(pic.Visible, pic.Width, 0) - IIf(tvw_s.Visible, tvw_s.Width, 0)
    mshPati.Height = tvw_s.Height
    
    If ScaleWidth - cmdCancel.Width - 300 > 4000 Then
        cmdCancel.Left = ScaleWidth - cmdCancel.Width - 300
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmFind
    Unload mfrmFilter
    mstrFilter = ""
    mlng����ID = 0
End Sub

Private Sub mshPati_DblClick()
    cmdOK_Click
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or mshPati.Width - X < 1000 Then Exit Sub
        pic.Left = pic.Left + X
        tvw_s.Width = tvw_s.Width + X
        mshPati.Left = mshPati.Left + X
        mshPati.Width = mshPati.Width - X
    End If
End Sub

Private Function InitUnits() As Boolean
'���ܣ���ʼ�����˲����ֲ��б�
'˵�����Բ����ֲ�
    Dim rsTmp As New ADODB.Recordset
    Dim objNode As Node, i As Integer
    Dim strPreKey  As String
    
    On Error GoTo errH
    
    strPreKey = ""
    If Not tvw_s.SelectedItem Is Nothing Then strPreKey = tvw_s.SelectedItem.Key
    
    If glngSys Like "8??" Then
        tvw_s.Nodes.Clear
        Set objNode = tvw_s.Nodes.Add(, , "Clinic", "���пͻ�", 1)
        objNode.Expanded = True
        objNode.Selected = True
    Else
        tvw_s.Nodes.Clear
        Set objNode = tvw_s.Nodes.Add(, , "All", "���в���", 1)
        objNode.Expanded = True
        
        Set objNode = tvw_s.Nodes.Add("All", 4, "In", "��Ժ����", 1)
        Set objNode = tvw_s.Nodes.Add("All", 4, "Out", "��Ժ����", 1)
        Set objNode = tvw_s.Nodes.Add("All", 4, "Clinic", "���ﲡ��", 1)
        Set objNode = tvw_s.Nodes.Add("All", 4, "Temp", "���۲���", 1)
        objNode.Expanded = True
        If objNode.Key = strPreKey Then objNode.Selected = True
                
        Set rsTmp = GetUnit(InStr(mstrPrivs, "���в���") = 0, "1,2,3", "����")
        If Not rsTmp.EOF Then
            For i = 1 To rsTmp.RecordCount
                Set objNode = tvw_s.Nodes.Add("In", 4, "D" & rsTmp!ID, "[" & rsTmp!���� & "]" & rsTmp!����, 1)
                
                If rsTmp!ID = UserInfo.����ID Then objNode.Selected = True
                If objNode.Key = strPreKey Then objNode.Selected = True
                objNode.Expanded = True
                
                rsTmp.MoveNext
            Next
        End If
        If tvw_s.SelectedItem Is Nothing Then tvw_s.Nodes("In").Selected = True
    End If
    
    InitUnits = True
    
    Call tvw_s_NodeClick(tvw_s.SelectedItem)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvw_s.Tag = Node.Key Then Exit Sub
    tvw_s.Tag = Node.Key
    
    SaveFlexState mshPati, App.ProductName & "\" & Me.Name
    Call ShowPatis("", , True)  '�л���������ʱ,�������,ʹ��ȱʡ����
End Sub

Private Sub ShowPatis(Optional ByVal strIF As String, Optional blnSort As Boolean, Optional blnSet As Boolean)
'���ܣ����ݵ�ǰ�˵����Ҫ��(�Զ���������),��ȡ������Ϣ
'������strIF=" And ...."��ʽ�Ĺ�������
    Dim i As Integer, strSQL As String
    Dim strInfo As String, Curdate As Date
    Dim strCard As String, lngUnitID As String
    Dim blnLimitUnit As Boolean, blnFirst As Boolean
    
    On Error GoTo errH
    
    If Not blnSort Then
        blnLimitUnit = InStr(mstrPrivs, "���в���") = 0
        
        If strIF = "" Then
            blnFirst = True
            If InStr(1, ",All,Clinic,Temp,", "," & tvw_s.SelectedItem.Key & ",") > 0 Then
                strIF = " And A.�Ǽ�ʱ�� Between trunc(Sysdate) And Sysdate"
            ElseIf tvw_s.SelectedItem.Key = "Out" Then
                strIF = " And P.��Ժ���� Between trunc(Sysdate) And Sysdate"
            ElseIf tvw_s.SelectedItem.Key = "In" Then
                strIF = " And P.��Ժ���� Between trunc(Sysdate) And Sysdate"
            End If
        End If
        strIF = strIF & " And A.ͣ��ʱ�� is NULL"
        '���￨����ʾ
        strCard = "Decode(" & IIf(gblnShowCard, 1, 0) & ",1,A.���￨��,LPAD('*',Length(A.���￨��),'*')) as ���￨,"
        
        
        If tvw_s.SelectedItem.Key = "All" Then '���в���
            strIF = strIF & IIf(blnLimitUnit, " And (A.��ǰ����ID Is NULL Or Instr(','||[2]||',',','||A.��ǰ����ID||',')>0)", "")
            
            '����25886 by lesfeng 2009-10-28 �������ͷ����������SQL�������������¿���֮����λ b
'            strSQL = "Select A.����ID,A.�����,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,A.�ѱ� as ����ѱ�," & _
'            " C.���� as ����,A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
'            " To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
'            " A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��," & _
'            " Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
'            " From ������ҳ P,������Ϣ A,���ű� C" & _
'            " Where A.��ǰ����ID=C.ID(+) And A.����ID=P.����ID(+) And A.��ҳID=P.��ҳID(+) " & strIF & _
'            " Order by A.�Ǽ�ʱ�� Desc"
            strSQL = "Select A.����ID,A.�����,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,A.�ѱ� as ����ѱ�," & _
            " B.���� as ����,C.���� as ����,A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
            " To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
            " A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��," & _
            " Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
            " From ������ҳ P,������Ϣ A,���ű� B,���ű� C" & _
            " Where A.��ǰ����ID=B.ID(+) And A.��ǰ����ID=C.ID(+) And A.����ID=P.����ID(+) And A.��ҳID=P.��ҳID(+) " & strIF & _
            " Order by A.�Ǽ�ʱ�� Desc"
            '����25886 by lesfeng 2009-10-28 �������ͷ����������SQL�������������¿���֮����λ b
            strInfo = "���ڶ�ȡ���в����嵥,���Ժ� ..."
            If Val(mshPati.Tag) <> 0 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 0
        ElseIf tvw_s.SelectedItem.Key = "In" Or Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then  '��Ժ����
            '58842,������,2013-02-25,��Ժ���˶�ȡ(����Ժ�����ж�ȡ)
            If Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then
                lngUnitID = Mid(tvw_s.SelectedItem.Key, 2)
                strIF = strIF & " And E.����ID= [1] "
            Else
                If blnLimitUnit Then
                    strIF = strIF & " And Instr(','||[2]||',',','||E.����ID||',')>0"
                End If
            End If
            
            strSQL = "Select A.����ID,A.סԺ��," & strCard & "NVL(P.����,A.����) ����,NVL(P.�Ա�,A.�Ա�) �Ա�,NVL(P.����,A.����) ����,P.�ѱ� as סԺ�ѱ�," & _
                " B.���� as ����,C.���� as ����,A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
                " A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                " A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
                " From ������ҳ P,������Ϣ A,���ű� B,���ű� C,��Ժ���� E" & _
                " Where A.��ǰ����ID=B.ID And A.��ǰ����ID=C.ID" & strIF & _
                " And A.����ID=P.����ID And A.��ҳID=P.��ҳID And A.����ID=E.����ID And Nvl(P.��ҳID,0)<>0 " & _
                " Order by A.��Ժʱ�� Desc,A.סԺ�� Desc"
            
            strInfo = "���ڶ�ȡ��Ժ�����嵥,���Ժ� ..."
            If Val(mshPati.Tag) <> 1 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 1
        ElseIf tvw_s.SelectedItem.Key = "Out" Then '��Ժ����
            strIF = strIF & IIf(blnLimitUnit, " And Instr(','||[2]||',',','||P.��ǰ����ID||',')>0", "")
                    
            strSQL = "Select A.����ID,A.סԺ��," & strCard & "NVL(P.����,A.����) ����,NVL(P.�Ա�,A.�Ա�) �Ա�,NVL(P.����,A.����) ����,P.�ѱ� as סԺ�ѱ�," & _
                " To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
                " A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                " A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
                " From ������ҳ P,������Ϣ A" & _
                " Where A.����ID=P.����ID And A.��ҳID=P.��ҳID" & _
                " And Nvl(P.��ҳID,0)<>0 And P.��Ժ���� Is Not NULL " & strIF & _
                " Order by A.��Ժʱ�� Desc,A.סԺ��"
            
            strInfo = "���ڶ�ȡ��Ժ�����嵥,���Ժ� ..."
            If Val(mshPati.Tag) <> 2 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 2
        ElseIf tvw_s.SelectedItem.Key = "Clinic" Then '���ﲡ��
            strSQL = "Select A.����ID,A.�����," & strCard & "A.����,A.�Ա�,A.����," & _
                " A.�ѱ� as " & IIf(glngSys Like "8??", "��Ա�ȼ�", "����ѱ�") & "," & _
                " To_Char(A.��������,'YYYY-MM-DD') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                " A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��,Decode(A.����,Null,'��ͨ����','ҽ������') ��������" & _
                " From ������Ϣ A " & _
                " Where A.��ǰ����ID is NULL And A.��ǰ����ID is NULL And A.��ҳID is NULL" & strIF & _
                " Order by A.�Ǽ�ʱ�� Desc,A.����� Desc"
            
            If glngSys Like "8??" Then
                strInfo = "���ڶ�ȡ�ͻ��嵥,���Ժ� ..."
            Else
                strInfo = "���ڶ�ȡ���ﲡ���嵥,���Ժ� ..."
            End If
            
            If Val(mshPati.Tag) <> 3 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 3
        ElseIf tvw_s.SelectedItem.Key = "Temp" Then
            '�������ۺ�סԺ���۲���
            strSQL = "Select Distinct A.����ID,Decode(P.��������,1,'��������','סԺ����') as ����, A.�����," & strCard & "NVL(P.����,A.����) ����,NVL(P.�Ա�,A.�Ա�) �Ա�,NVL(P.����,A.����) ����," & _
                " A.�ѱ� as " & IIf(glngSys Like "8??", "��Ա�ȼ�", "����ѱ�") & "," & _
                " To_Char(A.��������,'YYYY-MM-DD') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
                " A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
                " From ������ҳ P,������Ϣ A " & _
                " Where A.����ID=P.����ID And P.��������<>0 And A.סԺ�� is Null " & strIF & _
                " Order by ����,�Ǽ�ʱ�� Desc"
            
            strInfo = "���ڶ�ȡ���۲����嵥,���Ժ� ..."
            
            If Val(mshPati.Tag) <> 4 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 4
        End If
        
        tvw_s.Tag = tvw_s.SelectedItem.Key
        sta.SimpleText = strInfo
        Screen.MousePointer = 11
        DoEvents
        Me.Refresh
        
        With SQLCondition
            Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngUnitID, mstrUnitIDs, .�Ǽ�ʱ��B, .�Ǽ�ʱ��E, .����ʱ��B, .����ʱ��E, _
            .��Ժʱ��B, .��Ժʱ��E, .��Ժʱ��B, .��Ժʱ��E, .סԺ��, .�Ա�, .����, .�ѱ�, .Patient)
        End With
    End If
    
    mshPati.Clear
    mshPati.Rows = 2
    
    If mrsPati.EOF Then
        Call SetHeader(blnSet)
        If glngSys Like "8??" Then
            sta.SimpleText = IIf(blnFirst, "����", "") & "û���ҵ����������Ŀͻ�,����[ɸѡ],ѡ���ѯ����."
        Else
            sta.SimpleText = IIf(blnFirst, "����", "") & "û���ҵ����������Ĳ���,����[ɸѡ],ѡ���ѯ����."
        End If
    Else
        Set mshPati.DataSource = mrsPati
        Call SetHeader(blnSet)
        If glngSys Like "8??" Then
            sta.SimpleText = IIf(blnFirst, "����", "") & "���ҵ� " & mrsPati.RecordCount & " λ���������Ŀͻ�"
        Else
            sta.SimpleText = IIf(blnFirst, "����", "") & "���ҵ� " & mrsPati.RecordCount & " λ���������Ĳ���."
        End If
    End If
    
    Screen.MousePointer = 0
    
    Me.Refresh
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetHeader(Optional blnSet As Boolean)
    Dim strHead As String
    Dim i As Integer
    
    If tvw_s.SelectedItem.Key = "All" Then '���в���
        strHead = "����ID,1,750|�����,1,750|סԺ��,1,750|���￨,4,850|����,1,800|�Ա�,4,500|����,4,800|����ѱ�,4,850|" & _
            "����,1,850|����,1,850|����,4,500|��Ժʱ��,4,1000|��Ժʱ��,4,1000|סԺ����,4,850|��������,4,1000|" & _
            "����,4,500|����,4,800|����,1,600|ѧ��,4,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,4,1000|��������,1,800"
    ElseIf tvw_s.SelectedItem.Key = "Clinic" Then '���ﲡ��
        If glngSys Like "8??" Then
            strHead = "�ͻ�ID,1,750|�ͻ���,1,750|��Ա��,4,850|����,1,800|�Ա�,4,500|����,4,800|��Ա�ȼ�,4,850|" & _
                "��������,4,1000|����,4,500|����,4,800|����,1,600|ѧ��,4,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|" & _
                "��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,4,1000|��������,1,800"
        Else
            strHead = "����ID,1,750|�����,1,750|���￨,4,850|����,1,800|�Ա�,4,500|����,4,800|����ѱ�,4,850|" & _
                "��������,4,1000|����,4,500|����,4,800|����,1,600|ѧ��,4,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|" & _
                "��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,4,1000|��������,1,800"
        End If
    ElseIf tvw_s.SelectedItem.Key = "Temp" Then  '���۲���
         strHead = "����ID,1,750|����,1,1000|�����,1,750|���￨,4,850|����,1,800|�Ա�,4,500|����,4,800|����ѱ�,4,850|" & _
                "��������,4,1000|����,4,500|����,4,800|����,1,600|ѧ��,4,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|" & _
                "��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,4,1000|��������,1,800"
    ElseIf tvw_s.SelectedItem.Key = "Out" Then '��Ժ����
        strHead = "����ID,1,750|סԺ��,1,750|���￨,4,850|����,1,800|�Ա�,4,500|����,4,800|סԺ�ѱ�,4,850|" & _
            "��Ժʱ��,4,1000|��Ժʱ��,4,1000|סԺ����,4,850|��������,4,1000|����,4,500|����,4,800|����,1,600|" & _
            "ѧ��,4,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,4,1000|��������,1,800"
    ElseIf tvw_s.SelectedItem.Key = "In" Or InStr("D", Left(tvw_s.SelectedItem.Key, 1)) > 0 Then '��Ժ����
        strHead = "����ID,1,750|סԺ��,1,750|���￨,4,850|����,1,800|�Ա�,4,500|����,4,800|סԺ�ѱ�,4,850|" & _
            "����,1,850|����,1,850|����,4,500|��Ժʱ��,4,1000|סԺ����,4,850|��������,4,1000|" & _
            "����,4,500|����,4,800|����,1,600|ѧ��,4,500|ְҵ,1,1000|���,1,750|���֤��,1,2000|��ͥ��ַ,1,2000|������λ,1,2000|�Ǽ�ʱ��,4,1000|��������,1,800"
    End If
    
    With mshPati
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Or blnSet Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Or blnSet Then Call RestoreFlexState(mshPati, App.ProductName & "\" & Me.Name)
        
        If glngSys Like "8??" Then .ColWidth(1) = 0
        .RowHeight(0) = 320
        
        '�ָ��ϴ���
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .Cols - 1
        Call mshPati_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub mshPati_EnterCell()
    If glngSys Like "8??" Then
        If mshPati.Row = 0 Or mshPati.TextMatrix(mshPati.Row, GetColNum("�ͻ�ID")) = "" Then Exit Sub
    Else
        If mshPati.Row = 0 Or mshPati.TextMatrix(mshPati.Row, GetColNum("����ID")) = "" Then Exit Sub
    End If
    mlngGo = mshPati.Row
    mlngCurRow = mshPati.Row: mlngTopRow = mshPati.TopRow
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Integer
    For i = 0 To mshPati.Cols - 1
        If mshPati.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mshPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshPati.MouseRow = 0 Then
        mshPati.MousePointer = 99
    Else
        mshPati.MousePointer = Screen.MousePointer
    End If
End Sub

Private Sub mshPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshPati.MouseCol
    
    If Button = 1 And mshPati.MousePointer = 99 And mblnDown Then '˫�����ʱ��ִ��
        mblnDown = False
        
        If mshPati.TextMatrix(0, lngCol) = "" Then Exit Sub
        If glngSys Like "8??" Then
            If mshPati.TextMatrix(1, GetColNum("�ͻ�ID")) = "" Then Exit Sub
        Else
            If mshPati.TextMatrix(1, GetColNum("����ID")) = "" Then Exit Sub
        End If
        
        Set mshPati.DataSource = Nothing
        
        Select Case mshPati.TextMatrix(0, lngCol)
            Case "�ͻ�ID"
                mrsPati.Sort = "����ID" & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
            Case "��Ա��"
                mrsPati.Sort = "���￨" & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
            Case Else
                mrsPati.Sort = mshPati.TextMatrix(0, lngCol) & IIf(mshPati.ColData(lngCol) = 0, "", " DESC")
        End Select
        
        mshPati.ColData(lngCol) = (mshPati.ColData(lngCol) + 1) Mod 2
        
        Call ShowPatis(, True)
    End If
End Sub

Private Sub mshPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mblnDown = True
End Sub

Private Sub SeekPati(blnHead As Boolean)
    Dim i As Long
    Dim blnFill As Boolean
    
    Screen.MousePointer = 11
    mblnGo = True
    If glngSys Like "8??" Then
        sta.SimpleText = "���ڶ�λ���������Ŀͻ�,��ESC��ֹ ..."
    Else
        sta.SimpleText = "���ڶ�λ���������Ĳ���,��ESC��ֹ ..."
    End If
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With mfrmFind
            If .txt����ID.Text <> "" Then
                If glngSys Like "8??" Then
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("�ͻ�ID")) = .txt����ID.Text
                Else
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����ID")) = .txt����ID.Text
                End If
            End If
            If .txt���￨.Text <> "" Then
                If glngSys Like "8??" Then
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("��Ա��")) = .txt���￨.Text
                Else
                    blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("���￨")) = .txt���￨.Text
                End If
            End If
            If .txt�����.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("�����")) = .txt�����.Text
            End If
            If .txtסԺ��.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("סԺ��")) = .txtסԺ��.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����")) = .txt����.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����")) Like "*" & .txt����.Text & "*"
            End If
            If .txt���֤.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("���֤��")) = .txt���֤.Text
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mlngGo = i + 1
            If i <= mshPati.Rows - 1 Then mshPati.Row = i: mshPati.TopRow = i
            mshPati.Col = 0: mshPati.ColSel = mshPati.Cols - 1
            sta.SimpleText = "�ҵ�һ����¼"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '��ESCȡ��
        If mblnGo = False Then
            sta.SimpleText = "�û�ȡ����λ����"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    sta.SimpleText = "�Ѷ�λ���嵥β��"
    Screen.MousePointer = 0
End Sub
