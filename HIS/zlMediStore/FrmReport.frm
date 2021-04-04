VERSION 5.00
Begin VB.Form FrmReport 
   Caption         =   "�����������"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   8415
   StartUpPosition =   2  '��Ļ����
   Begin VB.Menu Popup 
      Caption         =   "�����˵�"
      Begin VB.Menu mnuBill 
         Caption         =   "����(&D)"
      End
   End
End
Attribute VB_Name = "FrmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents ObjReport As zl9Report.clsReport
Attribute ObjReport.VB_VarHelpID = -1
Dim gstrSQL As String
Private lngCurReport As Long
Private CurSheet As Object
Dim strNoS As String

Private Sub Form_Load()
    Set ObjReport = New zl9Report.clsReport
End Sub

Private Sub mnuBill_Click()
    Dim StrNo As String
    Dim byt���� As Integer
    Dim byt��¼״̬ As Integer
      
    
    Select Case strNoS
        Case "ZL1_INSIDE_1309_1"   '����
            StrNo = Mid(Trim(CurSheet.TextMatrix(CurSheet.Row, 3)), 3)
            byt���� = Val(CurSheet.TextMatrix(CurSheet.Row, 1))
            byt��¼״̬ = 1
        Case "ZL1_INSIDE_1309_2"   '��ϸ��
            StrNo = Trim(CurSheet.TextMatrix(CurSheet.Row, 3))
            byt���� = Val(CurSheet.TextMatrix(CurSheet.Row, 2))
            byt��¼״̬ = Val(CurSheet.TextMatrix(CurSheet.Row, 1))
        Case "ZL1_INSIDE_1309_3"   '��ϸ��
        
    End Select
    
    If StrNo = "" Or byt���� = 0 Or byt��¼״̬ = 99 Then Exit Sub
    If byt���� = 0 Then Exit Sub
    ShowBill frmWin, StrNo, byt��¼״̬, byt����
    
End Sub

Private Sub ObjReport_ReportActive(ByVal StrNo As String, Form As Object)
    lngCurReport = Form.hwnd
    strNoS = StrNo
    If UCase(StrNo) = "ZL1_INSIDE_1309_3" Then
       SetMenu 0
    End If
End Sub


Private Sub ObjReport_SheetDblClick(ByVal StrNo As String, Sheet As Object, frmParent As Object)
    lngCurReport = frmParent.hwnd
    strNoS = StrNo
    Set CurSheet = Sheet
    If UCase(StrNo) = "ZL1_INSIDE_1309_3" Then Exit Sub
    mnuBill_Click
End Sub

Private Sub ObjReport_SheetMouseDown(ByVal StrNo As String, Button As Integer, Shift As Integer, x As Single, y As Single, Sheet As Object, frmParent As Object)
    lngCurReport = frmParent.hwnd
    strNoS = StrNo
    Set CurSheet = Sheet
    If UCase(StrNo) <> "ZL1_INSIDE_1309_3" Then
        If Button = 2 Then PopupMenu Popup, 2
    End If
End Sub

Private Sub SetMenu(ByVal IntState As Integer)
    If IntState = 0 Then Popup.Visible = False: Exit Sub
    
End Sub


Public Sub ShowBill(frmObject As Object, StrNo As String, int��¼״̬ As Integer, int���� As Integer, Optional bln���� As Boolean = False)
    '--------------------------------------------------------------------------------------
    '����:��ʾָ������
    '����:
    '       frmObject:����
    '           strNo:���ݺ�
    '     int��¼״̬:����״̬(mod(��¼״̬,3)=1-������¼;mod(��¼״̬,3)=2-������¼;mod(��¼״̬,3)=0-�Ѿ������ļ�¼)
    '         int����:�������( �ⷿ:1-�⹺��ⵥ;2-�������;3-�ƿⵥ;4-����;5-��������;6-�̴�;7-������;
    '                           ����:1-����;2-����;3-���ϵ�;4-Ȩ�����)
    '--------------------------------------------------------------------------------------
'    frmPurchaseCard.ShowCard frmObject, StrNo, 4, int��¼״̬
    Select Case int����
        Case 1
            frmPurchaseCard.ShowCard frmObject, StrNo, 4, int��¼״̬
        Case 2
            frmSelfMakeCard.ShowCard frmObject, StrNo, 4, int��¼״̬
        Case 3
            frmAccordDrugCard.ShowCard frmObject, StrNo, 4, int��¼״̬
        Case 4
            frmOtherInputCard.ShowCard frmObject, StrNo, 4, int��¼״̬
        Case 5
            frmDiffPriceAdjustCard.ShowCard frmObject, StrNo, 4, int��¼״̬
        Case 6
            frmTransferCard.ShowCard frmObject, StrNo, 4, int��¼״̬
        Case 7
            frmDrawCard.ShowCard frmObject, StrNo, 4, int��¼״̬
        Case 11
            frmOtherOutputCard.ShowCard frmObject, StrNo, 4, int��¼״̬
        Case 12
            frmCheckCard.ShowCard frmObject, StrNo, 4, int��¼״̬
        Case 13
            Dim rsTemp As New ADODB.Recordset
            Dim StrSql As String
            With rsTemp
                StrSql = "Select id,����,NO,nvl(�۸�id,0) as �۸�id" & _
                    " From ҩƷ�շ���¼" & _
                    " Where No='" & StrNo & "'" & _
                    "       And ����=" & int����
                If .State = adStateOpen Then .Close
                .Open StrSql, gcnOracle, adOpenKeyset
                If .EOF Or .BOF Then Exit Sub
            End With
            gstrUserName = UserInfo.�û�����
            With frmAdjust
                .lngBillId = rsTemp!�۸�id
                .lngMediId = 1
                .Show 1, frmObject
            End With
        Case Else
            
            Frm����See.byt���� = int����
            Frm����See.StrNo = StrNo
            Frm����See.Show 1, frmObject
        End Select
'    End With
'    Select Case int����
'           Case 1   '�⹺���
'                frmPurchaseCard.ShowCard frmObject, StrNo, 4, int��¼״̬
'           Case 2   '�������
'                frmOtherInputCard.ShowCard frmObject, StrNo, 4, int��¼״̬
'           Case 3   '�ƿⵥ
'                frmTransferCard.ShowCard frmObject, StrNo, 4, int��¼״̬
'           Case 4   '����
'                frmDrawCard.ShowCard frmObject, StrNo, 4, int��¼״̬
'           Case 5   '��������
'                frmOtherOutputCard.ShowCard frmObject, StrNo, 4, int��¼״̬
'           Case 6   '�̵�
'                frmCheckCard.ShowCard frmObject, StrNo, 4, int��¼״̬
'           Case 7   '������
'                With Frm���ʸ������༭
'                    .EditState = 5
'                    .UnitStyle = GetMaterialUnit("���ʸ�������")
'                    .StrShowNo = StrNo
'                    .int��¼״̬ = int��¼״̬
'                    .Show 1, frmObject
'                End With
'     End Select
End Sub

