VERSION 5.00
Begin VB.Form FrmReport 
   Caption         =   "报表控制中心"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4215
   ScaleWidth      =   8415
   StartUpPosition =   2  '屏幕中心
   Begin VB.Menu Popup 
      Caption         =   "弹出菜单"
      Begin VB.Menu mnuBill 
         Caption         =   "单据(&D)"
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
    Dim byt单据 As Integer
    Dim byt记录状态 As Integer
      
    
    Select Case strNoS
        Case "ZL1_INSIDE_1309_1"   '总帐
            StrNo = Mid(Trim(CurSheet.TextMatrix(CurSheet.Row, 3)), 3)
            byt单据 = Val(CurSheet.TextMatrix(CurSheet.Row, 1))
            byt记录状态 = 1
        Case "ZL1_INSIDE_1309_2"   '明细帐
            StrNo = Trim(CurSheet.TextMatrix(CurSheet.Row, 3))
            byt单据 = Val(CurSheet.TextMatrix(CurSheet.Row, 2))
            byt记录状态 = Val(CurSheet.TextMatrix(CurSheet.Row, 1))
        Case "ZL1_INSIDE_1309_3"   '明细表
        
    End Select
    
    If StrNo = "" Or byt单据 = 0 Or byt记录状态 = 99 Then Exit Sub
    If byt单据 = 0 Then Exit Sub
    ShowBill frmWin, StrNo, byt记录状态, byt单据
    
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


Public Sub ShowBill(frmObject As Object, StrNo As String, int记录状态 As Integer, int单据 As Integer, Optional bln在用 As Boolean = False)
    '--------------------------------------------------------------------------------------
    '功能:显示指定单据
    '参数:
    '       frmObject:窗体
    '           strNo:单据号
    '     int记录状态:单据状态(mod(记录状态,3)=1-正常记录;mod(记录状态,3)=2-冲销记录;mod(记录状态,3)=0-已经冲销的记录)
    '         int单据:单据类别( 库房:1-外购入库单;2-其它入库;3-移库单;4-领用;5-其它出库;6-盘存;7-更换单;
    '                           在用:1-领用;2-销售;3-报废单;4-权属变更)
    '--------------------------------------------------------------------------------------
'    frmPurchaseCard.ShowCard frmObject, StrNo, 4, int记录状态
    Select Case int单据
        Case 1
            frmPurchaseCard.ShowCard frmObject, StrNo, 4, int记录状态
        Case 2
            frmSelfMakeCard.ShowCard frmObject, StrNo, 4, int记录状态
        Case 3
            frmAccordDrugCard.ShowCard frmObject, StrNo, 4, int记录状态
        Case 4
            frmOtherInputCard.ShowCard frmObject, StrNo, 4, int记录状态
        Case 5
            frmDiffPriceAdjustCard.ShowCard frmObject, StrNo, 4, int记录状态
        Case 6
            frmTransferCard.ShowCard frmObject, StrNo, 4, int记录状态
        Case 7
            frmDrawCard.ShowCard frmObject, StrNo, 4, int记录状态
        Case 11
            frmOtherOutputCard.ShowCard frmObject, StrNo, 4, int记录状态
        Case 12
            frmCheckCard.ShowCard frmObject, StrNo, 4, int记录状态
        Case 13
            Dim rsTemp As New ADODB.Recordset
            Dim StrSql As String
            With rsTemp
                StrSql = "Select id,单据,NO,nvl(价格id,0) as 价格id" & _
                    " From 药品收发记录" & _
                    " Where No='" & StrNo & "'" & _
                    "       And 单据=" & int单据
                If .State = adStateOpen Then .Close
                .Open StrSql, gcnOracle, adOpenKeyset
                If .EOF Or .BOF Then Exit Sub
            End With
            gstrUserName = UserInfo.用户姓名
            With frmAdjust
                .lngBillId = rsTemp!价格id
                .lngMediId = 1
                .Show 1, frmObject
            End With
        Case Else
            
            Frm单据See.byt单据 = int单据
            Frm单据See.StrNo = StrNo
            Frm单据See.Show 1, frmObject
        End Select
'    End With
'    Select Case int单据
'           Case 1   '外购入库
'                frmPurchaseCard.ShowCard frmObject, StrNo, 4, int记录状态
'           Case 2   '其它入库
'                frmOtherInputCard.ShowCard frmObject, StrNo, 4, int记录状态
'           Case 3   '移库单
'                frmTransferCard.ShowCard frmObject, StrNo, 4, int记录状态
'           Case 4   '领用
'                frmDrawCard.ShowCard frmObject, StrNo, 4, int记录状态
'           Case 5   '其它出库
'                frmOtherOutputCard.ShowCard frmObject, StrNo, 4, int记录状态
'           Case 6   '盘点
'                frmCheckCard.ShowCard frmObject, StrNo, 4, int记录状态
'           Case 7   '更换单
'                With Frm物资更换单编辑
'                    .EditState = 5
'                    .UnitStyle = GetMaterialUnit("物资更换管理")
'                    .StrShowNo = StrNo
'                    .int记录状态 = int记录状态
'                    .Show 1, frmObject
'                End With
'     End Select
End Sub

