VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
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
Private mlng����ID As Long
Private mstrPrivs As String
Private mcllFilter As Collection
Private mrsPati As ADODB.Recordset
Private mlngCurRow As Long, mlngTopRow As Long
Private mlngGo As Long, mblnDown As Boolean, mblnGo As Boolean
Private mfrmFilter As frmPatiFilter
Private mfrmFind As frmPatiFind
Private mstrUnitIDs As String '����Ա���ڲ����������������
  
Private mblnShowCard As Boolean '�����Ƿ������ʾ

Private mcnOracle As ADODB.Connection
Private mobjDataBase As clsDataBase
Private mobjOneDataObject As clsOneCardDataObject
Private mblnOk As Boolean

Public Function zlShowCard(ByVal cnOracle As ADODB.Connection, frmMain As Object, ByVal strPrivs As String, Optional lng����ID_Out As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ����ѡ����
    '���:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-12-05 16:56:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCard As Card
    mlng����ID = 0: mstrPrivs = strPrivs
    If zlGetOneDataBase(cnOracle, mobjDataBase) = False Then Exit Function
    If zlGetOneCardDataObject(cnOracle, mobjOneDataObject) = False Then Exit Function
    
    Set mcllFilter = New Collection
    If mobjOneDataObject.zlGetCardFromCardTypeID("���￨", False, objCard) Then
        mblnShowCard = objCard.�������Ĺ��� = ""
    End If
    
    If frmMain Is Nothing Then
         Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    lng����ID_Out = mlng����ID
    zlShowCard = mblnOk
End Function


Private Sub cmdCancel_Click()
    mlng����ID = 0
    mblnOk = False
    If gobjComLib Is Nothing Then zlInitCommLib
    If gobjComLib Is Nothing Then Exit Sub
    gobjComLib.saveWinState Me, App.ProductName
    Hide
End Sub

Private Sub cmdFilter_Click()
    If mfrmFilter.zlShowCard(Me, Val(mshPati.Tag), mcllFilter, mcnOracle) = False Then Exit Sub
    Call ShowPatis(mcllFilter)
End Sub

Private Sub cmdFind_Click()
    Dim blnOK As Boolean
    blnOK = gblnOk
    mfrmFind.mbytType = Val(mshPati.Tag)
    mfrmFind.Show 1, Me
    If gblnOk Then Call SeekPati(mfrmFind.optHead)
    gblnOk = blnOK
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
    mblnOk = True
    
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing Then
        gobjComLib.saveWinState Me, App.ProductName
    End If
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
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing Then
        gobjComLib.RestoreWinState Me, App.ProductName
     End If
    
    If glngSys Like "8??" Then
        Caption = "�ͻ�ѡ��"
        tvw_s.Visible = False
        pic.Visible = False
    End If
    Set mfrmFilter = New frmPatiFilter
    Set mfrmFind = New frmPatiFind
        
    mstrUnitIDs = mobjOneDataObject.zlGetUserUnits
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
    Err = 0: On Error Resume Next
    If Not mfrmFind Is Nothing Then Unload mfrmFind: Set mfrmFind = Nothing
    If Not mfrmFilter Is Nothing Then Unload mfrmFilter: Set mfrmFilter = Nothing
    If Not mrsPati Is Nothing Then Set mrsPati = Nothing
    If Not mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    If Not mobjOneDataObject Is Nothing Then Set mobjOneDataObject = Nothing
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing Then
        gobjComLib.saveWinState Me, App.ProductName
    End If
    
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
    
    On Error GoTo ErrH
    
    strPreKey = ""
    If Not tvw_s.SelectedItem Is Nothing Then strPreKey = tvw_s.SelectedItem.Key
    
    tvw_s.Nodes.Clear
    Set objNode = tvw_s.Nodes.Add(, , "All", "���в���", 1)
    objNode.Expanded = True
    Set objNode = tvw_s.Nodes.Add("All", 4, "In", "��Ժ����", 1)
    Set objNode = tvw_s.Nodes.Add("All", 4, "Out", "��Ժ����", 1)
    Set objNode = tvw_s.Nodes.Add("All", 4, "Clinic", "���ﲡ��", 1)
    Set objNode = tvw_s.Nodes.Add("All", 4, "Temp", "���۲���", 1)
    
    objNode.Expanded = True
    If objNode.Key = strPreKey Then objNode.Selected = True
    Set rsTmp = mobjOneDataObject.zlGetUnitRecordFromDepdIDs(InStr(mstrPrivs, "���в���") = 0, "1,2,3", "����")
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            Set objNode = tvw_s.Nodes.Add("In", 4, "D" & rsTmp!id, "[" & rsTmp!���� & "]" & rsTmp!����, 1)
            
            If rsTmp!id = UserInfo.����ID Then objNode.Selected = True
            If objNode.Key = strPreKey Then objNode.Selected = True
            objNode.Expanded = True
            
            rsTmp.MoveNext
        Next
    End If
    If tvw_s.SelectedItem Is Nothing Then tvw_s.Nodes("In").Selected = True

    
    InitUnits = True
    
    Call tvw_s_NodeClick(tvw_s.SelectedItem)
    Exit Function
ErrH:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
    Call mobjDataBase.SaveErrLog
End Function

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If tvw_s.Tag = Node.Key Then Exit Sub
    tvw_s.Tag = Node.Key
    If gobjComLib Is Nothing Then zlInitCommLib
    If Not gobjComLib Is Nothing Then gobjComLib.SaveFlexState mshPati, App.ProductName & "\" & Me.Name
    Call ShowPatis(Nothing, , True)      '�л���������ʱ,�������,ʹ��ȱʡ����
End Sub

  



Private Sub ShowPatis(ByVal cllFilter As Collection, Optional blnSort As Boolean, Optional blnSet As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݵ�ǰ�˵����Ҫ��(�Զ���������),��ȡ������Ϣ
    '���:cllfilter-��������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-30 16:49:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str����IDs As String, strInfo As String, curDate As Date
    Dim blnLimitUnit As Boolean, blnFirst As Boolean
    Dim blnPatiQuery As Boolean    '�Ƿ񰴲�����Ϣ��ѯ
    
    On Error GoTo ErrH
    If Not blnSort Then
        blnLimitUnit = InStr(mstrPrivs, ";���в���;") = 0
        If cllFilter.count = 0 Then
            blnFirst = True
            If InStr(1, ",All,Clinic,Temp,", "," & tvw_s.SelectedItem.Key & ",") > 0 Then
                 curDate = mobjDataBase.Currentdate
                 cllFilter.Add Array("�Ǽ�ʱ��", Format(curDate, "yyyy-mm-dd 00:00:00"), Format(curDate, "yyyy-mm-dd 23:59:59"), "�Ǽ�ʱ��")
            ElseIf tvw_s.SelectedItem.Key = "Out" Then
                 curDate = mobjDataBase.Currentdate
                 cllFilter.Add Array("��Ժ����", Format(curDate, "yyyy-mm-dd 00:00:00"), Format(curDate, "yyyy-mm-dd 23:59:59"), "��Ժ����")
            ElseIf tvw_s.SelectedItem.Key = "In" Then
                 curDate = mobjDataBase.Currentdate
                 cllFilter.Add Array("��Ժ����", Format(curDate, "yyyy-mm-dd 00:00:00"), Format(curDate, "yyyy-mm-dd 23:59:59"), "��Ժ����")
            End If
        End If
        
         
        If tvw_s.SelectedItem.Key = "All" Then '���в���
            
            strInfo = "���ڶ�ȡ���в����嵥,���Ժ� ..."
            gobjCommFun.ShowFlash strInfo
            If Val(mshPati.Tag) <> 0 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 0
            Call GetHospitalizationPatientData(IIf(blnLimitUnit, mstrUnitIDs, ""), cllFilter, mrsPati)
             
        ElseIf tvw_s.SelectedItem.Key = "In" Or Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then  '��Ժ����
            strInfo = "���ڶ�ȡ��Ժ�����嵥,���Ժ� ..."
            gobjCommFun.ShowFlash strInfo
            If Val(mshPati.Tag) <> 1 Then
               Unload mfrmFind
               Unload mfrmFilter
            End If
            mshPati.Tag = 1
                    
            str����IDs = IIf(blnLimitUnit, mstrUnitIDs, "")
            If Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then
                str����IDs = Mid(tvw_s.SelectedItem.Key, 2)
            End If
            Call GetHospitalizationPatientData(str����IDs, cllFilter, mrsPati)
        ElseIf tvw_s.SelectedItem.Key = "Out" Then '��Ժ����
            strInfo = "���ڶ�ȡ��Ժ�����嵥,���Ժ� ..."
            gobjCommFun.ShowFlash strInfo
            If Val(mshPati.Tag) <> 2 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 2
            Call GetLeavePatientData(IIf(blnLimitUnit, mstrUnitIDs, ""), cllFilter, mrsPati)
        ElseIf tvw_s.SelectedItem.Key = "Clinic" Then '���ﲡ��
            strInfo = "���ڶ�ȡ���ﲡ���嵥,���Ժ� ..."
            gobjCommFun.ShowFlash strInfo
            tvw_s.Tag = tvw_s.SelectedItem.Key
            sta.SimpleText = strInfo
            Screen.MousePointer = 11
            DoEvents
            Me.Refresh
            If Val(mshPati.Tag) <> 3 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 3
            Call GetOutPatientData(cllFilter, mrsPati) '��ȡ���ﲡ����Ϣ'
        ElseIf tvw_s.SelectedItem.Key = "Temp" Then
            '�������ۺ�סԺ���۲���
            strInfo = "���ڶ�ȡ���۲����嵥,���Ժ� ..."
            gobjCommFun.ShowFlash strInfo
            If Val(mshPati.Tag) <> 4 Then
                Unload mfrmFind
                Unload mfrmFilter
            End If
            mshPati.Tag = 4
            Call GetObservationPatiData(cllFilter, mrsPati) '��ȡ���۲�������

        End If
        
        tvw_s.Tag = tvw_s.SelectedItem.Key
        sta.SimpleText = strInfo
        Screen.MousePointer = 11
        DoEvents
        Me.Refresh
        gobjCommFun.StopFlash
    End If
    
    mshPati.Clear
    mshPati.Rows = 2
    If mrsPati.EOF Then
        Call SetHeader(blnSet)
        sta.SimpleText = IIf(blnFirst, "����", "") & "û���ҵ����������Ĳ���,����[ɸѡ],ѡ���ѯ����."
    Else
        Set mshPati.DataSource = mrsPati
        Call SetHeader(blnSet)
        sta.SimpleText = IIf(blnFirst, "����", "") & "���ҵ� " & mrsPati.RecordCount & " λ���������Ĳ���."
    End If
    
    Screen.MousePointer = 0
    Me.Refresh
    Exit Sub
ErrH:
    Screen.MousePointer = 0
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
   Call gobjComLib.SaveErrLog
     gobjCommFun.StopFlash
End Sub


Private Function GetAllPatientData(ByVal str����IDs As String, ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���в�����Ϣ
    '���:cllFilter-��������
    '     str����IDs-ָ������
    '����:rsPatiInfo_Out-���ز������ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiCons As Collection, cllPageCons As Collection
    Dim cllPatiData As Collection, cllPageData As Collection
    Dim cllTemp As Collection, cllTemp1 As Collection, cllPageTemp As Collection
    Dim blnPatiQuery As Boolean, str����Ids As String
    Dim i As Long, lng����ID As Long, J As Long
    
    
    
    On Error GoTo errHandle


'    strSQL = "" & _
'    "   Select A.����ID,A.�����,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,A.�ѱ� as ����ѱ�," & _
'    "           B.���� as ����,C.���� as ����,A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
'    "           To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
'    "           A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���,A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��," & _
'    "           Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
'    " From ������ҳ P,������Ϣ A,���ű� B,���ű� C" & _
'    " Where A.��ǰ����ID=B.ID(+) And A.��ǰ����ID=C.ID(+) And A.����ID=P.����ID(+) And A.��ҳID=P.��ҳID(+) " & strIF & _
'    " Order by A.�Ǽ�ʱ�� Desc"
    
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "����ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "�����", adLongVarChar, 20, adFldIsNullable
        .fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        .fields.Append "���￨��", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .fields.Append "����ѱ�", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "��Ժʱ��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��Ժʱ��", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "סԺ����", adVarNumeric, 18, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ѧ��", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ְҵ", adLongVarChar, 100, adFldIsNullable
        .fields.Append "���", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "���֤��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��ͥ��ַ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "������λ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "�Ǽ�ʱ��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
     '�Ƿ񰴲�����ϢΪ����ѯ
    Set cllPatiCons = New Collection
    Set cllPageCons = New Collection
    
    blnPatiQuery = False
    For i = 1 To cllFilter.count
        If InStr(",�Ǽ�ʱ��,����ID,����,���￨��,�����,ҽ����,���֤��,IC����,", "," & cllFilter(i)(0) & ",") > 0 Then
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
            blnPatiQuery = True
        ElseIf InStr(",��Ժ����,��Ժ����,", "," & cllFilter(i)(0) & ",") Then
            'סԺ����
             cllPageCons.Add cllFilter(i), cllFilter(i)(0)
        Else
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
        End If
    Next
    
    If blnPatiQuery = False Then
    
        '1.�Ȳ�ѯ������ҳ������
        '  (0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ )
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", "", cllPageData) = False Then Exit Function
        If cllPageData Is Nothing Then Set cllPageData = New Collection
        
        str����Ids = ""
        For i = 1 To cllPageData.count
            '��������
            Set cllTemp = cllPageData(i)
            If InStr(str����Ids & ",", "," & cllTemp("_pati_id") & ",") = 0 Then
                str����Ids = str����Ids & "," & cllTemp("_pati_id")
            End If
        Next
        If str����Ids = "" Then Exit Function
        str����Ids = Mid(str����Ids, 2)
       '2.���Բ���Ϊ��ѯ����
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, str����Ids, str����IDs) = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
    Else
        '1.��������Ϣ��ѯ
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, "", "") = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
        
        str����Ids = ""
        For i = 1 To cllPatiData.count
            Set cllTemp = cllPatiData(i)
            If InStr(str����Ids & ",", "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid") & ",") = 0 Then
                str����Ids = str����Ids & "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid")
            End If
        Next
        If str����Ids = "" Then Exit Function
        
        '2.��ѯ������ҳ��Ϣ
        str����Ids = Mid(str����Ids, 2)
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", "", cllPageData) = False Then Exit Function
        
    End If
    
    '��װ����
    For i = 1 To cllPatiData.count
       Set cllTemp = cllPatiData(i)
       lng����ID = Val(cllTemp("_pati_id"))
    
       Set cllPageTemp = Nothing
       For J = 1 To cllPageData.count
          Set cllTemp1 = cllPatiData(J)
          If cllTemp1("_pati_id") = lng����ID And cllTemp1("pati_pageid") = cllTemp("pati_pageid") Then Set cllPageTemp = cllTemp1: Exit For
       Next
       If Not cllPageTemp Is Nothing Then
           With rsPatiInfo_Out
                  !����ID = lng����ID
                  !����� = cllTemp("_outpatient_num")
                  !סԺ�� = cllTemp("_inpatient_num")

                  If mblnShowCard Then
                     !���￨�� = LPAD("*", Len(cllTemp("_vcard_no")))
                  Else
                     !���￨�� = cllTemp("_vcard_no")
                  End If
                  
                  !���� = cllTemp("_pati_name")
                  !�Ա� = cllTemp("_pati_sex")
                  !���� = cllTemp("_pati_age")
                  !����ѱ� = cllTemp("_fee_category")
                  !���� = cllTemp("_pati_wardarea_name")
                  !���� = cllTemp("_pati_dept_name")
                  !���� = cllTemp("_pati_bed")
                  
                  !��Ժʱ�� = cllPageTemp("_adta_time")
                  !��Ժʱ�� = cllPageTemp("_adtd_time")
                  !סԺ���� = cllTemp("_inp_times")
                  
                  !�������� = cllTemp("_pati_birthdate")
                  !���� = cllTemp("_country_name")
                  !���� = cllTemp("_pati_nation")
                  !���� = cllTemp("_pati_area")
                  !ѧ�� = cllTemp("_pati_education")
                  !ְҵ = cllTemp("_ocpt_name")
                  !��� = cllTemp("_pati_identity")
                  !���֤�� = cllTemp("_pati_idcard")
                  !��ͥ��ַ = cllTemp("_pat_home_addr")
                  !������λ = cllTemp("_emp_name")
                  !�Ǽ�ʱ�� = cllTemp("_create_time")
                  If cllPageTemp("_pati_type") <> "" Then
                       !�������� = cllPageTemp("_pati_type")
                  Else
                       !�������� = IIf(Val(cllPageTemp("_insurance_type")) = 0, "��ͨ����", "ҽ������")
                  End If
               rsPatiInfo_Out.Update
           End With
       End If

      Next

    GetAllPatientData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function




Private Function GetHospitalizationPatientData(ByVal str����IDs As String, ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ժ������Ϣ
    '���:cllFilter-��������
    '     str����IDs-ָ������
    '����:rsPatiInfo_Out-���ز������ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiCons As Collection, cllPageCons As Collection
    Dim cllPatiData As Collection, cllPageData As Collection
    Dim cllTemp As Collection, cllTemp1 As Collection, cllPageTemp As Collection
    Dim blnPatiQuery As Boolean, str����Ids As String, lng����ID As Long
    Dim i As Long, J As Long
    
    On Error GoTo errHandle

'    If Mid(tvw_s.SelectedItem.Key, 1, 1) = "D" Then
'        lngUnitID = Mid(tvw_s.SelectedItem.Key, 2)
'        strIF = strIF & " And P.��ǰ����ID+0= [1] "
'    Else
'        If blnLimitUnit Then
'            strIF = strIF & " And Instr(','||[2]||',',','||P.��ǰ����ID||',')>0"
'        End If
'    End If
'
'    strSQL = "Select A.����ID,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,P.�ѱ� as סԺ�ѱ�," & _
'        " B.���� as ����,C.���� as ����,A.��ǰ���� as ����,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
'        " A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
'        " A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
'        " From ������ҳ P,������Ϣ A,���ű� B,���ű� C" & _
'        " Where A.��Ժ=1 And A.��ǰ����ID=B.ID And A.��ǰ����ID=C.ID" & strIF & _
'        " And A.����ID=P.����ID And A.��ҳID=P.��ҳID And Nvl(P.��ҳID,0)<>0 And P.��Ժ���� is NULL " & _
'        " Order by A.��Ժʱ�� Desc,A.סԺ�� Desc"
    '
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "����ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "סԺ��", adLongVarChar, 18, adFldIsNullable
        .fields.Append "���￨��", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .fields.Append "סԺ�ѱ�", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "��Ժʱ��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��Ժʱ��", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "סԺ����", adVarNumeric, 18, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ѧ��", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ְҵ", adLongVarChar, 100, adFldIsNullable
        .fields.Append "���", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "���֤��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��ͥ��ַ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "������λ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "�Ǽ�ʱ��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
     '�Ƿ񰴲�����ϢΪ����ѯ
    Set cllPatiCons = New Collection
    Set cllPageCons = New Collection
    
    blnPatiQuery = False
    For i = 1 To cllFilter.count
        If InStr(",�Ǽ�ʱ��,����ID,����,���￨��,�����,ҽ����,���֤��,IC����,", "," & cllFilter(i)(0) & ",") > 0 Then
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
            blnPatiQuery = True
        ElseIf InStr(",��Ժ����,��Ժ����,", "," & cllFilter(i)(0) & ",") Then
            'סԺ����
             cllPageCons.Add cllFilter(i), cllFilter(i)(0)
        Else
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
        End If
    Next
    
    If blnPatiQuery = False Then
    
        '1.�Ȳ�ѯ������ҳ������
        '  (0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ )
        If zl_CisSvr_GetPatPageInfByRange(0, cllPageCons, "", str����IDs, cllPageData) = False Then Exit Function
        If cllPageData Is Nothing Then Set cllPageData = New Collection
        
        str����Ids = ""
        For i = 1 To cllPageData.count
            '��������
            Set cllTemp = cllPageData(i)
            If InStr(str����Ids & ",", "," & cllTemp("_pati_id") & ",") = 0 Then
                str����Ids = str����Ids & "," & cllTemp("_pati_id")
            End If
        Next
        If str����Ids = "" Then Exit Function
        str����Ids = Mid(str����Ids, 2)
       '2.���Բ���Ϊ��ѯ����
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, str����Ids, "") = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
    Else
        '1.��������Ϣ��ѯ
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, "", "") = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
        
        str����Ids = ""
        For i = 1 To cllPatiData.count
            Set cllTemp = cllPatiData(i)
            If InStr(str����Ids & ",", "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid") & ",") = 0 Then
                str����Ids = str����Ids & "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid")
            End If
        Next
        If str����Ids = "" Then Exit Function
        
        '2.��ѯ������ҳ��Ϣ
        str����Ids = Mid(str����Ids, 2)
        If zl_CisSvr_GetPatPageInfByRange(0, cllPageCons, "", str����IDs, cllPageData) = False Then Exit Function
        
    End If
    
    '��װ����
    For i = 1 To cllPatiData.count
       Set cllTemp = cllPatiData(i)
       lng����ID = Val(cllTemp("_pati_id"))
    
       Set cllPageTemp = Nothing
       For J = 1 To cllPageData.count
          Set cllTemp1 = cllPatiData(J)
          If cllTemp1("_pati_id") = lng����ID And cllTemp1("pati_pageid") = cllTemp("pati_pageid") Then Set cllPageTemp = cllTemp1: Exit For
       Next
       If Not cllPageTemp Is Nothing Then
           With rsPatiInfo_Out
               rsPatiInfo_Out.AddNew
                  !����ID = lng����ID
                    !סԺ�� = cllTemp("_inpatient_num")

                  If mblnShowCard Then
                     !���￨�� = LPAD("*", Len(cllTemp("_vcard_no")))
                  Else
                     !���￨�� = cllTemp("_vcard_no")
                  End If
                  
                  !���� = cllTemp("_pati_name")
                  !�Ա� = cllTemp("_pati_sex")
                  !���� = cllTemp("_pati_age")
                  !���� = cllTemp("_pati_wardarea_name")
                  !���� = cllTemp("_pati_dept_name")
                  !���� = cllTemp("_pati_bed")
                  
                  !��Ժʱ�� = cllPageTemp("_adta_time")
                  !��Ժʱ�� = cllPageTemp("_adtd_time")
                  !סԺ���� = cllTemp("_inp_times")
                  
                  !�������� = cllTemp("_pati_birthdate")
                  !���� = cllTemp("_country_name")
                  !���� = cllTemp("_pati_nation")
                  !���� = cllTemp("_pati_area")
                  !ѧ�� = cllTemp("_pati_education")
                  !ְҵ = cllTemp("_ocpt_name")
                  !��� = cllTemp("_pati_identity")
                  !���֤�� = cllTemp("_pati_idcard")
                  !��ͥ��ַ = cllTemp("_pat_home_addr")
                  !������λ = cllTemp("_emp_name")
                  !�Ǽ�ʱ�� = cllTemp("_create_time")
                  If cllPageTemp("_pati_type") <> "" Then
                       !�������� = cllPageTemp("_pati_type")
                  Else
                       !�������� = IIf(Val(cllPageTemp("_insurance_type")) = 0, "��ͨ����", "ҽ������")
                  End If
               rsPatiInfo_Out.Update
           End With
        End If
    Next

    
    GetHospitalizationPatientData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function GetLeavePatientData(ByVal str����IDs As String, ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��Ժ������Ϣ
    '���:cllFilter-��������
    '     str����IDs-ָ������
    '����:rsPatiInfo_Out-���ز������ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiCons As Collection, cllPageCons As Collection
    Dim cllPatiData As Collection, cllPageData As Collection
    Dim cllTemp As Collection, cllTemp1 As Collection, cllPageTemp As Collection
    Dim blnPatiQuery As Boolean, str����Ids As String, lng����ID As Long
    Dim i As Long, J As Long
    On Error GoTo errHandle
    
'
'    strIF = strIF & IIf(blnLimitUnit, " And Instr(','||[2]||',',','||P.��ǰ����ID||',')>0", "")
'
'    strSQL = "Select A.����ID,A.סԺ��," & strCard & "A.����,A.�Ա�,A.����,P.�ѱ� as סԺ�ѱ�," & _
'    " To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��,To_Char(A.��Ժʱ��,'YYYY-MM-DD') as ��Ժʱ��," & _
'    " A.סԺ����,To_Char(A.��������,'YYYY-MM-DD') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
'    " A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
'    " From ������ҳ P,������Ϣ A" & _
'    " Where A.����ID=P.����ID And A.��ҳID=P.��ҳID" & _
'    " And Nvl(P.��ҳID,0)<>0 And P.��Ժ���� Is Not NULL " & strIF & _
'    " Order by A.��Ժʱ�� Desc,A.סԺ��"
    '
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "����ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "סԺ��", adLongVarChar, 18, adFldIsNullable
        .fields.Append "���￨��", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .fields.Append "סԺ�ѱ�", adLongVarChar, 50, adFldIsNullable
        
        .fields.Append "��Ժʱ��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��Ժʱ��", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "סԺ����", adVarNumeric, 18, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ѧ��", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ְҵ", adLongVarChar, 100, adFldIsNullable
        .fields.Append "���", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "���֤��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��ͥ��ַ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "������λ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "�Ǽ�ʱ��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
     '�Ƿ񰴲�����ϢΪ����ѯ
    Set cllPatiCons = New Collection
    Set cllPageCons = New Collection
    
    blnPatiQuery = False
    For i = 1 To cllFilter.count
        If InStr(",�Ǽ�ʱ��,����ID,����,���￨��,�����,ҽ����,���֤��,IC����,", "," & cllFilter(i)(0) & ",") > 0 Then
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
            blnPatiQuery = True
        ElseIf InStr(",��Ժ����,��Ժ����,", "," & cllFilter(i)(0) & ",") Then
            'סԺ����
             cllPageCons.Add cllFilter(i), cllFilter(i)(0)
        Else
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
        End If
    Next
    
    If blnPatiQuery = False Then
    
        '1.�Ȳ�ѯ������ҳ������
        '  (0-��Ժ����;1-��Ժ����;2-��Ժ���Ժ )
        If zl_CisSvr_GetPatPageInfByRange(1, cllPageCons, "", str����IDs, cllPageData) = False Then Exit Function
        If cllPageData Is Nothing Then Set cllPageData = New Collection
        
        str����Ids = ""
        For i = 1 To cllPageData.count
            '��������
            Set cllTemp = cllPageData(i)
            If InStr(str����Ids & ",", "," & cllTemp("_pati_id") & ",") = 0 Then
                str����Ids = str����Ids & "," & cllTemp("_pati_id")
            End If
        Next
        If str����Ids = "" Then Exit Function
        str����Ids = Mid(str����Ids, 2)
       '2.���Բ���Ϊ��ѯ����
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, str����Ids, "") = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
    Else
        '1.��������Ϣ��ѯ
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData) = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
        
        str����Ids = ""
        For i = 1 To cllPatiData.count
            Set cllTemp = cllPatiData(i)
            If InStr(str����Ids & ",", "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid") & ",") = 0 Then
                str����Ids = str����Ids & "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid")
            End If
        Next
        If str����Ids = "" Then Exit Function
        
        '2.��ѯ������ҳ��Ϣ
        str����Ids = Mid(str����Ids, 2)
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", str����IDs, cllPageData) = False Then Exit Function
        
    End If
    
    '��װ����
    For i = 1 To cllPatiData.count
       Set cllTemp = cllPatiData(i)
       lng����ID = Val(cllTemp("_pati_id"))
    
       Set cllPageTemp = Nothing
       For J = 1 To cllPageData.count
          Set cllTemp1 = cllPatiData(J)
          If cllTemp1("_pati_id") = lng����ID And cllTemp1("pati_pageid") = cllTemp("pati_pageid") Then Set cllPageTemp = cllTemp1: Exit For
       Next
       If Not cllPageTemp Is Nothing Then
           With rsPatiInfo_Out
               rsPatiInfo_Out.AddNew
                  !����ID = lng����ID
                    !סԺ�� = cllTemp("_inpatient_num")

                  If mblnShowCard Then
                     !���￨�� = LPAD("*", Len(cllTemp("_vcard_no")))
                  Else
                     !���￨�� = cllTemp("_vcard_no")
                  End If
                  
                  !���� = cllTemp("_pati_name")
                  !�Ա� = cllTemp("_pati_sex")
                  !���� = cllTemp("_pati_age")
                  !סԺ�ѱ� = cllPageTemp("_fee_category")
                  !��Ժʱ�� = cllPageTemp("_adta_time")
                  !��Ժʱ�� = cllPageTemp("_adtd_time")
                  !סԺ���� = cllTemp("_inp_times")
                  
                  !�������� = cllTemp("_pati_birthdate")
                  !���� = cllTemp("_country_name")
                  !���� = cllTemp("_pati_nation")
                  !���� = cllTemp("_pati_area")
                  !ѧ�� = cllTemp("_pati_education")
                  !ְҵ = cllTemp("_ocpt_name")
                  !��� = cllTemp("_pati_identity")
                  !���֤�� = cllTemp("_pati_idcard")
                  !��ͥ��ַ = cllTemp("_pat_home_addr")
                  !������λ = cllTemp("_emp_name")
                  !�Ǽ�ʱ�� = cllTemp("_create_time")
                  If cllPageTemp("_pati_type") <> "" Then
                       !�������� = cllPageTemp("_pati_type")
                  Else
                       !�������� = IIf(Val(cllPageTemp("_insurance_type")) = 0, "��ͨ����", "ҽ������")
                  End If
               rsPatiInfo_Out.Update
           End With
        End If
    Next
    
    GetLeavePatientData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function GetOutPatientData(ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���ﲡ����Ϣ
    '���:cllFilter-��������
    '����:rsPatiInfo_Out-���ز������ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiData As Collection, cllTemp As Collection
    Dim i As Long, lng����ID As Long
    
    
     On Error GoTo errHandle
    
       
    'strSQL = "Select A.����ID,A.�����," & strCard & "A.����,A.�Ա�,A.����," & _
    " A.�ѱ� as " & IIf(glngSys Like "8??", "��Ա�ȼ�", "����ѱ�") & "," & _
    " To_Char(A.��������,'YYYY-MM-DD') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
    " A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��,Decode(A.����,Null,'��ͨ����','ҽ������') ��������" & _
    " From ������Ϣ A " & _
    " Where A.��ǰ����ID is NULL And A.��ǰ����ID is NULL And A.��ҳID is NULL" & strIF & _
    " Order by A.�Ǽ�ʱ�� Desc,A.����� Desc"
    '
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "����ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "�����", adVarNumeric, 18, adFldIsNullable
        .fields.Append "���￨��", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .fields.Append "����ѱ�", adLongVarChar, 50, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ѧ��", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ְҵ", adLongVarChar, 100, adFldIsNullable
        .fields.Append "���", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "���֤��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��ͥ��ַ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "������λ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "�Ǽ�ʱ��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    If zl_PatiSvr_GetPatiInfsByRange(0, cllFilter, cllPatiData) = False Then Exit Function
    If cllPatiData Is Nothing Then Set cllPatiData = New Collection
    If cllPatiData.count = 0 Then Exit Function
    '��װ����
    For i = 1 To cllPatiData.count
        Set cllTemp = cllPatiData(i)
        lng����ID = Val(cllTemp("_pati_id"))
        With rsPatiInfo_Out
            rsPatiInfo_Out.AddNew
                !����ID = lng����ID
                !����� = cllTemp("_outpatient_num")
                
                If mblnShowCard Then
                    !���￨�� = LPAD("*", Len(cllTemp("_vcard_no")))
                Else
                    !���￨�� = cllTemp("_vcard_no")
                End If
                
                !���� = cllTemp("_pati_name")
                !�Ա� = cllTemp("_pati_sex")
                !���� = cllTemp("_pati_age")
                !����ѱ� = cllTemp("_fee_category")
                !�������� = cllTemp("_pati_birthdate")
                !���� = cllTemp("_country_name")
                !���� = cllTemp("_pati_nation")
                !���� = cllTemp("_pati_area")
                !ѧ�� = cllTemp("_pati_education")
                !ְҵ = cllTemp("_ocpt_name")
                !��� = cllTemp("_pati_identity")
                !���֤�� = cllTemp("_pati_idcard")
                !��ͥ��ַ = cllTemp("_pat_home_addr")
                !������λ = cllTemp("_emp_name")
                !�Ǽ�ʱ�� = cllTemp("_create_time")
                If cllTemp("_pati_type") <> "" Then
                    !�������� = cllTemp("_pati_type")
                Else
                    !�������� = IIf(Val(cllTemp("_insurance_type")) = 0, "��ͨ����", "ҽ������")
                End If
            rsPatiInfo_Out.Update
        End With
    Next
     
    
    GetOutPatientData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function GetObservationPatiData(ByVal cllFilter As Collection, ByRef rsPatiInfo_Out As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���۲�����Ϣ
    '���:cllFilter-��������
    '����:rsPatiInfo_Out-���ز������ݼ�
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2019-10-31 15:06:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPatiCons As Collection, cllPageCons As Collection
    Dim cllPatiData As Collection, cllPageData As Collection
    Dim cllTemp As Collection, cllTemp1 As Collection, cllPageTemp As Collection
    Dim blnPatiQuery As Boolean, str����Ids As String
    Dim i As Long, lng����ID As Long, J As Long
    
    
    On Error GoTo errHandle

    '    '�������ۺ�סԺ���۲���
    '    strSQL = "Select Distinct A.����ID,Decode(P.��������,1,'��������','סԺ����') as ����, A.�����," & strCard & "A.����,A.�Ա�,A.����," & _
    '    " A.�ѱ� as " & IIf(glngSys Like "8??", "��Ա�ȼ�", "����ѱ�") & "," & _
    '    " To_Char(A.��������,'YYYY-MM-DD') as ��������,A.����,A.����,A.����,A.ѧ��,A.ְҵ,A.���," & _
    '    " A.���֤��,A.��ͥ��ַ,A.������λ,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD') as �Ǽ�ʱ��,Nvl(P.��������,Decode(P.����,Null,'��ͨ����','ҽ������')) ��������" & _
    '    " From ������ҳ P,������Ϣ A " & _
    '    " Where A.����ID=P.����ID And P.��������<>0 And A.סԺ�� is Null " & strIF & _
    '    " Order by ����,�Ǽ�ʱ�� Desc"
    '
    Set rsPatiInfo_Out = New ADODB.Recordset
    With rsPatiInfo_Out
        If .State = 1 Then .Close
        .fields.Append "����ID", adVarNumeric, 18, adFldIsNullable
        .fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .fields.Append "�����", adLongVarChar, 18, adFldIsNullable
        .fields.Append "���￨��", adLongVarChar, 100, adFldIsNullable
    
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "�Ա�", adLongVarChar, 20, adFldIsNullable
        .fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .fields.Append "����ѱ�", adLongVarChar, 50, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 30, adFldIsNullable
        
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ѧ��", adLongVarChar, 100, adFldIsNullable
        .fields.Append "ְҵ", adLongVarChar, 100, adFldIsNullable
        .fields.Append "���", adLongVarChar, 100, adFldIsNullable
        
        .fields.Append "���֤��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��ͥ��ַ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "������λ", adLongVarChar, 200, adFldIsNullable
        .fields.Append "�Ǽ�ʱ��", adLongVarChar, 30, adFldIsNullable
        .fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
     '�Ƿ񰴲�����ϢΪ����ѯ
    Set cllPatiCons = New Collection
    Set cllPageCons = New Collection
    
    blnPatiQuery = False
    For i = 1 To cllFilter.count
        If InStr(",�Ǽ�ʱ��,����ID,����,���￨��,�����,ҽ����,���֤��,IC����,", "," & cllFilter(i)(0) & ",") > 0 Then
            cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
            blnPatiQuery = True
        ElseIf InStr(",��Ժ����,��Ժ����,", "," & cllFilter(i)(0) & ",") Then
            'סԺ����
             cllPageCons.Add cllFilter(i), cllFilter(i)(0)
        Else
           cllPatiCons.Add cllFilter(i), cllFilter(i)(0)
        End If
    Next
    
    cllPageCons.Add Array("��������", "1,2"), "��������"
    
    'ֻ�����۲���
    If blnPatiQuery = False Then
    
        '1.�Ȳ�ѯ������ҳ������
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", "", cllPageData) = False Then Exit Function
        If cllPageData Is Nothing Then Set cllPageData = New Collection
        str����Ids = ""
        For i = 1 To cllPageData.count
            '��������
            Set cllTemp = cllPageData(i)
            If InStr(str����Ids & ",", "," & cllTemp("_pati_id") & ",") = 0 Then
                str����Ids = str����Ids & "," & cllTemp("_pati_id")
            End If
        Next
        If str����Ids = "" Then Exit Function
        str����Ids = Mid(str����Ids, 2)
       '2.���Բ���Ϊ��ѯ����
        
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData, str����Ids) = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
    Else
        '1.��������Ϣ��ѯ
        If zl_PatiSvr_GetPatiInfsByRange(2, cllPatiCons, cllPatiData) = False Then Exit Function
        If cllPatiData Is Nothing Then Set cllPatiData = New Collection
        If cllPatiData.count = 0 Then Exit Function
        
        
        str����Ids = ""
        For i = 1 To cllPatiData.count
            Set cllTemp = cllPatiData(i)
            If InStr(str����Ids & ",", "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid") & ",") = 0 Then
                str����Ids = str����Ids & "," & cllTemp("_pati_id") & ";" & cllTemp("_pati_pageid")
            End If
        Next
        If str����Ids = "" Then Exit Function
        
        '2.��ѯ������ҳ��Ϣ
        str����Ids = Mid(str����Ids, 2)
        If zl_CisSvr_GetPatPageInfByRange(2, cllPageCons, "", "", cllPageData) = False Then Exit Function
    End If
    
    '��װ����
    For i = 1 To cllPatiData.count
       Set cllTemp = cllPatiData(i)
       lng����ID = Val(cllTemp("_pati_id"))
    
       Set cllPageTemp = Nothing
       For J = 1 To cllPageData.count
          Set cllTemp1 = cllPatiData(J)
          If cllTemp1("_pati_id") = lng����ID And cllTemp1("pati_pageid") = cllTemp("pati_pageid") Then Set cllPageTemp = cllTemp1: Exit For
       Next
       If Not cllPageTemp Is Nothing Then
           With rsPatiInfo_Out
               rsPatiInfo_Out.AddNew
                  !����ID = lng����ID
                  !���� = Decode(Val(cllPageTemp("_pati_nature")), 1, "��������", 2, "סԺ����", "������")
                  !����� = cllTemp("_outpatient_num")

                  If mblnShowCard Then
                     !���￨�� = LPAD("*", Len(cllTemp("_vcard_no")))
                  Else
                     !���￨�� = cllTemp("_vcard_no")
                  End If
                  
                  !���� = cllTemp("_pati_name")
                  !�Ա� = cllTemp("_pati_sex")
                  !���� = cllTemp("_pati_age")
                  !����ѱ� = cllTemp("_fee_category")
                  !�������� = cllTemp("_pati_birthdate")
                  !���� = cllTemp("_country_name")
                  !���� = cllTemp("_pati_nation")
                  !���� = cllTemp("_pati_area")
                  !ѧ�� = cllTemp("_pati_education")
                  !ְҵ = cllTemp("_ocpt_name")
                  !��� = cllTemp("_pati_identity")
                  !���֤�� = cllTemp("_pati_idcard")
                  !��ͥ��ַ = cllTemp("_pat_home_addr")
                  !������λ = cllTemp("_emp_name")
                  !�Ǽ�ʱ�� = cllTemp("_create_time")
                  If cllPageTemp("_pati_type") <> "" Then
                       !�������� = cllPageTemp("_pati_type")
                  Else
                       !�������� = IIf(Val(cllPageTemp("_insurance_type")) = 0, "��ͨ����", "ҽ������")
                  End If
               rsPatiInfo_Out.Update
           End With
        End If
    Next
    GetObservationPatiData = True
    Exit Function
errHandle:
    If gobjComLib.ErrCenter() = 1 Then
        Resume
    End If
End Function


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
        
        If Not Visible Or blnSet Then
            If gobjComLib Is Nothing Then zlInitCommLib
            If Not gobjComLib Is Nothing Then Call gobjComLib.RestoreFlexState(mshPati, App.ProductName & "\" & Me.Name)
        End If
        
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
        
        .col = 0: .ColSel = .Cols - 1
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
        
        Call ShowPatis(Nothing, True)
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
 
    sta.SimpleText = "���ڶ�λ���������Ĳ���,��ESC��ֹ ..."
 
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshPati.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With mfrmFind
            If .txt����ID.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("����ID")) = .txt����ID.Text
            End If
            If .txt���￨.Text <> "" Then
                blnFill = blnFill And mshPati.TextMatrix(i, GetColNum("���￨")) = .txt���￨.Text
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
            mshPati.col = 0: mshPati.ColSel = mshPati.Cols - 1
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

