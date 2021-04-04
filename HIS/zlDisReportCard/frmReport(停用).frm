VERSION 5.00
Object = "{09B13292-AC31-4C5D-B44A-C83E7AAD70E6}#1.1#0"; "zlSubclass.ocx"
Begin VB.Form frmReport 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   16755
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   14670
   LinkTopic       =   "Form1"
   ScaleHeight     =   16755
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin zlSubclass.Subclass Subclass1 
      Left            =   900
      Top             =   3495
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.HScrollBar hsbReport 
      Height          =   255
      LargeChange     =   500
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   4
      Top             =   0
      Width           =   8535
   End
   Begin VB.VScrollBar vsbReport 
      Height          =   7335
      LargeChange     =   50
      Left            =   0
      Max             =   100
      SmallChange     =   10
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   15640
      Left            =   1800
      ScaleHeight     =   15615
      ScaleWidth      =   11865
      TabIndex        =   6
      Top             =   -480
      Width           =   11895
      Begin zlDisReportCard.PaneFour PaneFour 
         Height          =   2340
         Left            =   1050
         TabIndex        =   3
         Top             =   12495
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   5001
      End
      Begin zlDisReportCard.PaneThree PaneThree 
         Height          =   4500
         Left            =   1050
         TabIndex        =   2
         Top             =   7860
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   7938
      End
      Begin zlDisReportCard.PaneOne PaneOne 
         Height          =   1065
         Left            =   1020
         TabIndex        =   0
         Top             =   1005
         Width           =   9825
         _ExtentX        =   17727
         _ExtentY        =   1879
      End
      Begin zlDisReportCard.PaneTwo PaneTwo 
         Height          =   5550
         Left            =   1050
         TabIndex        =   1
         Top             =   2145
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   9790
      End
      Begin VB.Line Line2 
         X1              =   1050
         X2              =   10875
         Y1              =   12435
         Y2              =   12435
      End
      Begin VB.Line Line1 
         X1              =   1050
         X2              =   10875
         Y1              =   7815
         Y2              =   7815
      End
      Begin VB.Shape Shape1 
         Height          =   12735
         Left            =   1035
         Top             =   2115
         Width           =   9855
      End
   End
   Begin VB.PictureBox picShadow 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   1770
      Left            =   750
      ScaleHeight     =   1770
      ScaleWidth      =   1140
      TabIndex        =   7
      Top             =   660
      Width           =   1140
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private marrSql() As Variant        '��������ʱ��ִ�е�SQL
Private mColCls As New Collection   '��Ҫ���浽���ݿ������
Private mColData As New Collection  '��������ݿ��ȡ��������
Public Event HaveSavedSQL()     'ִ�б���SQLʱ����,ÿִ��һ������һ��
Public blnHaveStatus As Boolean  '�Ƿ����״̬��
Private blnFirstGot As Boolean  '��һ�λ�ý���

Private mlngPatiID As Long '����id
Private mlngPageID As Long '��ҳID�����ﴫ�Һ�ID��
Private mbytType As Byte   '�༭��ʽ0-������1-�޸ģ�����������ȡ����
Private mbytFrom As Byte   '������Դ1-���� 2-סԺ
Private mlngDeptID As Long '��ǰ����ID
Private mlngFileID As Long   '�ļ�ID,��Դ�ڵ��Ӳ�����¼.ID
Private mbytBabyNo As Long 'Ӥ��ID
Private mbln���֤���� As Boolean '���֤��Ϣ���� ��������Ⱦ���������֤�������

Private mstrChkType_2016 As String '���ݸ�ʽ�ǣ�"[��][���̲�][AIDS][...]......"
Private mstrChkType_2014 As String '���ݸ�ʽ�ǣ�"[��][���̲�][AIDS][...]......"


Private Type POINTAPI
        x As Long
        y As Long
End Type

Public Sub SetMyFocus()
    If picReport.Enabled = True Then
        Call picReport.SetFocus
    End If
End Sub

Public Function HaveChanged() As Boolean
'���ܣ��ж��ĸ��Զ���ؼ��������ʾ��Ϣ�Ƿ����ı�
    HaveChanged = False
    If PaneOne.HaveChanged = True Then
        HaveChanged = True
    ElseIf PaneTwo.HaveChanged = True Then
        HaveChanged = True
    ElseIf PaneThree.HaveChanged = True Then
        HaveChanged = True
    ElseIf PaneFour.HaveChanged = True Then
        HaveChanged = True
    End If
End Function

Public Sub CanWrite()
'���ܣ��ǽ�����Ա༭
    picReport.Enabled = True
    If gblnLock = True Then
        PaneThree.lockCtl (Not gblnLock)
    End If
End Sub

Public Sub PrintReport(ByVal frmParent As Object, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal strPrintDeviceName As String)
'���ܣ���ӡ����
    Dim strSQL As String
    Dim strPos As String
    Dim strPosInfo() As String
    Dim strPosTmp() As String
    Dim i As Integer
    
    On Error GoTo errHand
    
    Call zlRefresh(lngPatiID, lngPageID, lngFileId, False)

    If Trim(strPrintDeviceName) <> "" Then
        For i = 0 To Printers.Count - 1
            If Trim(Printers(i).DeviceName) = Trim(strPrintDeviceName) Then
                Set Printer = Printers(i)
                Exit For
            End If
            If i = Printers.Count - 1 Then
                MsgBox "û���ҵ���Ӧ�Ĵ�ӡ������˶Դ�ӡ�����ƣ�", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            End If
        Next
    End If
    Printer.PaperSize = vbPRPSA4 'A4ֽ
    Printer.ScaleMode = vbPixels

    glngOffsetX = -GetDeviceCaps(Printer.hdc, PHYSICALOFFSETX) '�ɴ�ӡ���Ե
    glngOffsetY = -GetDeviceCaps(Printer.hdc, PHYSICALOFFSETY) '�ɴ�ӡ�ϱ�Ե

    Call PaneOne.PrintOne
    Call PaneTwo.PrintTwo
    Call PaneThree.PrintThree
    Call PaneFour.PrintFour

'    strPos = "69,142,725,142|69,142,69,1069|69,1069,725,1069|725,142,725,1069|" & _
'             "69,466,725,466|69,514,725,514|69,678,725,678|69,749,725,749|" & _
'             "69,793,725,793|69,934,725,934|69,1025,725,1025"
    
             
    strPos = "69,142,725,142|69,142,69,928|69,928,725,928|725,142,725,928|" & _
             "69,466,725,466|69,514,725,514|69,678,725,678|69,749,725,749|" & _
             "69,793,725,793|69,884,725,884"
    
    
    strPosInfo = Split(strPos, "|")
    For i = 0 To UBound(strPosInfo)
        strPosTmp = (Split(strPosInfo(i), ","))
        Printer.Line (glngOffsetX + PScaleX(val(strPosTmp(0))), glngOffsetY + PScaleY(val(strPosTmp(1))))-(glngOffsetX + PScaleX(val(strPosTmp(2))), glngOffsetY + PScaleY(val(strPosTmp(3)))), &H0&, B
    Next
    
    Printer.EndDoc
    
    strSQL = "Zl_���Ӳ�����ӡ_Insert(" & mlngFileID & ",20," & mlngPatiID & "," & mlngPageID & ",'" & UserInfo.���� & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, "")
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub zlRefresh(ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal lngFileId As Long, ByVal blnMoved As Boolean)
    mlngPatiID = lngPatiID
    mlngPageID = lngPageID
    mlngFileID = lngFileId
 
    Call PaneOne.ClearMe
    Call PaneTwo.ClearMe
    Call PaneThree.ClearMe
    Call PaneFour.ClearMe
    Call InitReport(mbytType, mlngPatiID, mlngPageID, mbytFrom, 0, mlngDeptID, mlngFileID)
    If lngPatiID <> 0 Then
        Call LoadData(1, blnMoved)
    End If
End Sub

Public Sub InitReport(ByVal bytType As Byte, ByVal lngPatiID As Long, ByVal lngPageID As Long, ByVal bytFrom As Byte, ByVal bytBabyNo As Byte, ByVal lngDeptID As Long, ByVal lngFileId As Long)
    mbytType = bytType
    mlngPatiID = lngPatiID
    mlngPageID = lngPageID
    mbytFrom = bytFrom
    mlngDeptID = lngDeptID
    mlngFileID = lngFileId
    mbytBabyNo = bytBabyNo
End Sub

Public Function SaveData(ByVal blnFinish As Boolean) As Boolean
    Dim i As Integer
    Dim strSQL As String
    Dim blnBegin As Boolean
    Dim SLevel As SignLevel
    Dim lngFileId As Long       '�ļ�ID ��Դ�ڲ����ļ��б�
    Dim strFileName As String   '�ļ����� ��Դ�ڲ����ļ��б�
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHand
    
    SaveData = False
    
    '������Ҫ��ȡ�µ��ļ�ID
    If mbytType = 0 Then
        mlngFileID = zlDatabase.GetNextId("���Ӳ�����¼")
        mbytType = 1
    End If
    
    SLevel = GetUserSignLevel(UserInfo.ID, mlngPatiID, mlngPageID)
    
    strSQL = "select t.id,t.���� from �����ļ��б� t where t.����=5 and t.���='000'"
    Set rsTemp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "")
    lngFileId = Nvl(rsTemp!ID, 0)
    strFileName = Nvl(rsTemp!����, "")
    strSQL = "Zl_��Ⱦ�����濨��¼_Update(" & mlngFileID & "," & mbytFrom & "," & mlngPatiID & "," & _
              mlngPageID & "," & mlngDeptID & ",'" & UserInfo.���� & "'," & lngFileId & ",'" & strFileName & _
               "','" & UserInfo.���� & "'," & IIf(blnFinish, 1, 0) & "," & IIf(blnFinish, SLevel, "Null") & "," & mbytBabyNo & ")"
    
    Call MakeSaveSql(marrSql, mColCls, mlngFileID)

    gcnOracle.BeginTrans
    blnBegin = True
    Call zlDatabase.ExecuteProcedure(strSQL, "")
    For i = LBound(marrSql) To UBound(marrSql)
        Call zlDatabase.ExecuteProcedure(CStr(marrSql(i)), "")
        RaiseEvent HaveSavedSQL
    Next
    gcnOracle.CommitTrans
    blnBegin = False
    SaveData = True
    If blnFinish = True Then
        picReport.Enabled = False
    End If
    Exit Function
errHand:
    If blnBegin Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Sub LoadData(ByVal bytType As Byte, Optional blnMoved As Boolean)
    Dim strSQL As String
    Dim strKey As String
    Dim strNo As String
    Dim strID As String
    Dim strTmp As String
    Dim strInfo() As String
    Dim objCls As clsReport
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errHand
    Set mColCls = New Collection
    mstrChkType_2014 = ""
    mstrChkType_2016 = ""
'   1-�޸�
    If bytType = 1 Then
        Set mColData = New Collection
        strSQL = "select t.id,t.�������,t.�����ı�,t.Ҫ������ from ���Ӳ������� t where t.�ļ�id=[1]"
        If blnMoved = True Then
            strSQL = Replace(strSQL, "���Ӳ�������", "H���Ӳ�������")
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���Ӳ�������", mlngFileID)
        
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp.EOF = False Then
                strID = Nvl(rsTemp!ID)
                strNo = Nvl(rsTemp!�������)
                strTmp = Nvl(rsTemp!�����ı�)
                strKey = "K" & Trim(strNo)
                mColData.Add strTmp, strKey

                If InStr(GSTR_OBJNO_2014, "," & strNo & ",") > 0 Then
                    mstrChkType_2014 = mstrChkType_2014 & "[" & strNo & "," & Trim(strTmp) & "]"
                End If
                
                If InStr(GSTR_OBJNO_2016, "," & strNo & ",") > 0 Then
                    mstrChkType_2016 = mstrChkType_2016 & "[" & strNo & "," & Trim(strTmp) & "]"
                End If
                
                Set objCls = New clsReport
                objCls.ID = strID
                mColCls.Add objCls, strKey
                rsTemp.MoveNext
            End If
        Next
'   0-����
    ElseIf bytType = 0 Then
        For i = 1 To 44
            Set objCls = New clsReport
            strKey = "K" & i
            objCls.ID = 0
            mColCls.Add objCls, strKey
        Next
        Set mColData = New Collection
        strTmp = "����|���֤��|�Ա�|��������|����|������λ|��ϵ�˵绰|��ͥ�绰|��λ�绰|����״��|ѧ��|��λ����|��ǰ����|��ͥ��ַ"
        strInfo = Split(strTmp, "|")
        
        For i = 0 To UBound(strInfo)
            If mbytBabyNo <> 0 And Trim(strInfo(i)) = "����" Then
                strSQL = "select Zl_Replace_Element_Value([1],[2],[3],[4],null,[5]) as ��Ϣ from dual"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", strInfo(i), mlngPatiID, mlngPageID, mbytFrom, mbytBabyNo)
                strTmp = Nvl(rsTemp!��Ϣ)
            ElseIf Trim(strInfo(i)) = "��λ����" Then
                strTmp = Nvl(zlRegInfo("��λ����"))
            Else
                strSQL = "select Zl_Replace_Element_Value([1],[2],[3],[4]) as ��Ϣ from dual"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", strInfo(i), mlngPatiID, mlngPageID, mbytFrom)
                strTmp = Nvl(rsTemp!��Ϣ)
            End If
            strNo = i
            mColData.Add strTmp, "K" & Trim(strNo)
        Next
        '�ҳ�����
        If mbytBabyNo <> 0 Then
            strSQL = "select Zl_Replace_Element_Value([1],[2],[3],[4],null,[5]) as ��Ϣ from dual"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", "�ҳ�����", mlngPatiID, mlngPageID, mbytFrom, mbytBabyNo)
            strTmp = Nvl(rsTemp!��Ϣ)
            mColData.Add strTmp, "KParent"
        Else
            mColData.Add "", "KParent"
        End If
        '��������
        If mbytFrom = 1 Then
            strSQL = "select t.�Ǽ�ʱ�� as �������� from ���˹Һż�¼ t where t.id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", mlngPageID)
        Else
            strSQL = "select t.��Ժ���� as �������� from ������ҳ t where t.����id=[1] and t.��ҳid=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", mlngPatiID, mlngPageID)
        End If
        
        If rsTemp.RecordCount <> 0 Then
            mColData.Add Format(Nvl(rsTemp!��������), "yyyy-mm-dd"), "K14"
        Else
            mColData.Add "--", "K14"
        End If
        '�������
        strSQL = "select decode(t.����ʱ��,null,t.��¼����,t.����ʱ��) as ������� from ������ϼ�¼ t where t.����id=[1] and t.��ҳid=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", mlngPatiID, mlngPageID)
    
        If rsTemp.RecordCount <> 0 Then
            mColData.Add Format(Nvl(rsTemp!�������), "yyyy-mm-dd-hh"), "K15"
        Else
            mColData.Add "---", "K15"
        End If
        '��������
        strSQL = " Select a.��ʼִ��ʱ�� as �������� " & _
                 " From ����ҽ����¼ A, ������ĿĿ¼ B " & _
                 " Where a.������Ŀid = b.Id And b.��� = 'Z' And " & _
                 " b.�������� = '11'  And a.������Դ = [1] And a.����id=[2] and a.��ҳid=[3] "
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", mbytFrom, mlngPatiID, mlngPageID)
        If rsTemp.RecordCount <> 0 Then
            mColData.Add Format(Nvl(rsTemp!��������), "yyyy-mm-dd"), "K17"
        Else
            mColData.Add "--", "K17"
        End If
        '����
        strSQL = "Select a.Id, b.�ļ�id, b.���没��, a.����id, a.��ҳid, a.ҽ��id, a.�������, a.����id, a.���id" & _
                 " From ������ϼ�¼ A, ��������ǰ�� B " & _
                 " Where (a.����id = b.����id Or " & _
                 " a.���id = b.���id Or " & _
                 " b.���id = (Select c.���id From ������϶��� c Where c.����id =a.����id) or " & _
                 " b.����id = (select d.����id from ������϶��� d where d.���id=a.���id)) And " & _
                 " b.�ļ�id =(select e.id from �����ļ��б� e where e.����=5  and e.����='�л����񹲺͹���Ⱦ�����濨' and e.����=4 ) and " & _
                 " a.��¼��Դ=3 and a.����id=[1] and a.��ҳid=[2]"
        
        strTmp = ""
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݶ�ȡ", mlngPatiID, mlngPageID)
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp.EOF = False Then
                strTmp = strTmp & Nvl(rsTemp!���没��) & "|"
                rsTemp.MoveNext
            End If
        Next
        mColData.Add strTmp, "K16"
    End If
    '�޸�ʱ�����������44��(2016�汾38��)�������44��(2016�汾38��)˵�������ļ�������
    '����ʱ�����������19�����������19��˵����Ϣ��Դ�ƻ�
    
    If bytType = 1 And mColData.Count = 44 Then
        glngVersion = VL_2014                                  '�޸�ʱ�����������44��,2014�汾
    ElseIf bytType = 1 And mColData.Count = 38 Then
        glngVersion = VL_2016                                  '�޸�ʱ�����������38��,2016�汾
    ElseIf mColData.Count = 19 And bytType = 0 Then
        glngVersion = VL_2016
    End If
    
    If glngVersion = VL_2014 Then
        Call PaneOne.LoadData(mColData, bytType, mstrChkType_2014)
        Call PaneTwo.LoadData(mColData, bytType, mstrChkType_2014)
        Call PaneThree.LoadData(mColData, bytType, mstrChkType_2014)
        Call PaneFour.LoadData(mColData, bytType, mstrChkType_2014)
    ElseIf glngVersion = VL_2016 Then
        Call PaneOne.LoadData(mColData, bytType, mstrChkType_2016)
        Call PaneTwo.LoadData(mColData, bytType, mstrChkType_2016)
        Call PaneThree.LoadData(mColData, bytType, mstrChkType_2016)
        Call PaneFour.LoadData(mColData, bytType, mstrChkType_2016)
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Sub

Public Sub SetCaption���֤()
    mbln���֤���� = val(zlDatabase.GetPara("��Ⱦ���������֤�������", glngSys, 1277, 0)) = 1
    Call PaneTwo.SetCaption���֤(mbln���֤����)
End Sub

Private Sub Form_Load()
        
    blnFirstGot = True
    gbytDiseaseType = 5
    gbytAcute = 3
    
    picReport.ScaleHeight = Me.ScaleY(297, 6, 3)
    picReport.ScaleWidth = Me.ScaleX(210, 6, 3)
    picReport.Top = Me.ScaleTop + 200
    marrSql = Array()
    Subclass1.hWnd = Me.hWnd
    Subclass1.Messages(WM_MOUSEWHEEL) = True
    mbln���֤���� = val(zlDatabase.GetPara("��Ⱦ���������֤�������", glngSys, 1277, 0)) = 1
    Call PaneTwo.SetCaption���֤(mbln���֤����)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picReport.Left = Me.ScaleLeft + (Me.ScaleWidth / 2) - (picReport.Width / 2)
    
    If Me.ScaleWidth < picReport.Width Then
        hsbReport.Visible = True
    Else
        hsbReport.Visible = False
    End If
    
    vsbReport.Top = Me.ScaleTop
    vsbReport.Left = Me.ScaleLeft + Me.ScaleWidth - vsbReport.Width
    vsbReport.Height = Me.ScaleHeight - IIf(hsbReport.Visible = True, hsbReport.Height, 0) - IIf(blnHaveStatus = True, 375, 0)
    vsbReport.LargeChange = 100 / ((picReport.Height + 800) / Me.ScaleHeight)
    vsbReport.SmallChange = vsbReport.LargeChange
    
    hsbReport.Top = vsbReport.Top + vsbReport.Height
    hsbReport.Left = Me.ScaleLeft
    hsbReport.Width = Me.ScaleLeft + Me.ScaleWidth
    hsbReport.LargeChange = 100 / (picReport.Width / Me.ScaleWidth)
    hsbReport.SmallChange = hsbReport.LargeChange
    
    picShadow.Move picReport.Left + 50, picReport.Top + 50, picReport.Width, picReport.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mColCls = Nothing
    Set mColData = Nothing
    Erase marrSql
    gstrKey = ""
    Subclass1.Messages(WM_MOUSEWHEEL) = False
End Sub

Private Sub hsbReport_Change()
    picReport.Left = -((picReport.Width - Me.Width) * (hsbReport.Value / 100))
    picShadow.Left = picReport.Left + 50
End Sub

Private Sub picReport_GotFocus()
    If blnFirstGot = True And picReport.Enabled = True Then
        Call PaneOne.SetMyFoucs
    End If
    blnFirstGot = False
End Sub

Private Sub Subclass1_WndProc(Msg As Long, wParam As Long, lParam As Long, result As Long)
    '�Զ������Ϣ������
    Dim tP As POINTAPI
    Dim sngX As Single, sngY As Single   '�������
    Dim intShift As Integer              '��갴��
    Dim bWay As Boolean                  '��귽��
    Dim bMouseFlag As Boolean            '����¼������־
    Dim wzDelta, wKeys As Integer
    Select Case Msg
        Case WM_MOUSEWHEEL   '����
            wzDelta = HIWORD(wParam)
            If wzDelta > 0 Then
                vsbReport.Value = IIf(vsbReport.Value > 10, vsbReport.Value - 10, 0)
            Else
                vsbReport.Value = IIf(vsbReport.Value < 90, vsbReport.Value + 10, 100)
            End If
    End Select
End Sub

Private Sub vsbReport_Change()
    picReport.Top = 200 - ((picReport.Height + 800 - Me.Height) * (vsbReport.Value / 100))
    picShadow.Top = picReport.Top + 50
End Sub

Public Function MakeSaveSql(arrSql() As Variant, colCls As Collection, ByVal strFileId As String) As Boolean
    Call PaneOne.MakeSaveSql(arrSql, colCls, strFileId)
    Call PaneTwo.MakeSaveSql(arrSql, colCls, strFileId)
    Call PaneThree.MakeSaveSql(arrSql, colCls, strFileId)
    Call PaneFour.MakeSaveSql(arrSql, colCls, strFileId)
End Function

Public Sub ClearEnterInfo()
    Call PaneFour.ClearEnterInfo
End Sub

Public Sub SetEnterInfo()
    Dim strDate As String
    If mColData.Count < 44 Then
        strDate = Trim(CStr(mColData("K12")))
    Else
        strDate = Trim(CStr(mColData("K43")))
    End If
    If strDate = "" Or strDate = "--" Then
        strDate = zlDatabase.Currentdate
    End If
    Call PaneFour.SetEnterInfo(UserInfo.����, strDate)
End Sub

Public Function CheckValidity() As Boolean
    Dim strMsg As String
    Dim strTmp As String
    Dim strMsgInfo() As String
    Dim i As Integer
    On Error GoTo errHand
    
    strMsg = ""
    strTmp = ""
    Call PaneTwo.CheckValidity(strMsg)
    Call PaneThree.CheckValidity(strMsg)
    If Trim(strMsg) = "" Then
        CheckValidity = True
    Else
        strMsgInfo = Split(strMsg, "$")
        For i = 0 To UBound(strMsgInfo) - 1
            strTmp = strTmp & i + 1 & ". " & strMsgInfo(i) & vbCrLf
        Next
        Call ShowMsg(strTmp)
        CheckValidity = False
    End If

    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Err = 0
End Function

Public Function RelateFeedback(ByVal isRelated As Boolean) As Boolean
'���ܣ���Ⱦ�����濨���������Խ��������������ȡ������
'������isRelated  true-������false-ȡ������
    Dim objDisease As Object
  
On Error GoTo errHand
    Set objDisease = CreateObject("zl9Disease.cDockDisease")
    If objDisease Is Nothing Then Exit Function
    Call objDisease.InitDockDisease(glngSys, gcnOracle)
    Call objDisease.RelateFeedback(Me, mlngFileID, mlngPatiID, mlngPageID, mbytFrom, isRelated)
    Set objDisease = Nothing
    RelateFeedback = True
    Exit Function
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
