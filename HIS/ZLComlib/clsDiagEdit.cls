VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDiagEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ShowDiagEdit(frmParent As Object, ByVal lng标识ID As Long, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, _
                        ByVal int病人来源 As Integer, ByVal lng开单科室ID As Long, ByVal str开单人 As String, ByRef str诊断IDs As String, _
                        ByRef str诊断S As String, ByVal bytSize As Byte, Optional ByVal lng医嘱组ID As Long) As Boolean
'功能：调用诊断编辑界面
'参数：lng病人ID=病人ID
'      lng就诊ID=住院:主页ID,门诊：挂号单ID
'      int病人来源=1-门诊，2-住院
'      lng开单科室ID=病人所在科室，诊断使用
'      lng标识ID =用于区分各个申请单的标识，用于保存相应的诊断
'      str开单人=操作员姓名，诊断登记人
'      str诊断IDs=该申请单相关的诊断ID,多个诊断时诊断ID以逗号分割
'      str诊断S =选择的诊断的字符串，多个诊断描述以逗号分割
'      bytSize=0-9号字体，1-12号字体
'      lng医嘱组ID=申请单对应的组医嘱ID，已经生成医嘱的申请单修改时传入
'返回： ShowDiagEdit= 是确定还是取消
'        str诊断S=返回诊断描述字符串，供申请单使用
    Dim blnOK As Boolean
    Dim i As Long
    Dim arrTmp As Variant
    
    blnOK = frmDiagEdit.ShowMe(frmParent, lng标识ID, lng病人ID, lng就诊ID, int病人来源, lng开单科室ID, str开单人, str诊断IDs, str诊断S, bytSize, lng医嘱组ID)
    '清空原有该申请单关联诊断的缓存，缓存新关联的诊断
    If blnOK Then
        Call DeleteApplyDiag(lng标识ID)
        If str诊断IDs <> "" Then
            arrTmp = Split(str诊断IDs, ",")
            With grsDiagConn
                For i = LBound(arrTmp) To UBound(arrTmp)
                    .AddNew
                    !诊断ID = Val(arrTmp(i))
                    !标识ID = lng标识ID
                    .Update
                Next
            End With
        End If
    End If
    ShowDiagEdit = blnOK
End Function

Private Sub Class_Initialize()
    Set grsDiagConn = New ADODB.Recordset
    '记录申请单关联诊断
    With grsDiagConn
        .Fields.Append "诊断ID", adInteger
        .Fields.Append "标识ID", adInteger
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub

Public Function DeleteApplyDiag(ByVal lng标识ID As Long) As Boolean
'功能：删除申请单时，清除该申请单关联的诊断
    With grsDiagConn
        .Filter = "标识ID=" & lng标识ID
        .Sort = "诊断ID"
        Do While Not .EOF
            .Delete
            .Update
            .MoveNext
        Loop
    End With
End Function

Private Sub Class_Terminate()
    Set grsDiagConn = Nothing
End Sub
