VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mfrmApplicationPage As New frmPacsApplication

Public Sub InitComponents(ByVal lngDeptID As Long, objOwner As Object)
'初始化申请单
'lngDeptId:当前科室Id
'objOwner:调用所在窗口对象

    Call mfrmApplicationPage.InitComponents(lngDeptID, objOwner)
End Sub



Public Function ShowApplicationForm(ByVal lngPatientID As Long, _
                                    ByVal lngCallFrom As Long, _
                                    ByVal lngPatientRegId As Long, _
                                    ByVal lngPatientPageId As Long, _
                                    ByVal lngUpdateAppNoOrAdvId As Long, _
                                    ByRef objAppPages() As clsApplicationData, _
                                    Optional ByVal intBabyID As Integer = 0, _
                                    Optional ByVal blnEdit As Boolean = True, _
                                    Optional ByVal lngProjectId As Long = 0) As Boolean
'显示申请单界面
'lngPatientId:      病人Id
'lngCallFrom:    接口调用来源
'lngPatientRegId:   挂号Id
'lngPatientPageId:  主页Id
'lngUpdateAdviceId: 医嘱Id，如果为空，说明是新增检查申请，否则表示修改检查申请，当更新医嘱时，如果申请序号为空，则传递对应医嘱的医嘱id作为此参数
'objAppPages():     返回医嘱的录入内容，每个数组元素表示一条检查申请
'blnEdit: 为true时可以进行编辑，反之不能进行编辑
'lngProjectId: 诊疗项目ID     新开检查申请时，不为0则根据此项目自动定位

    ShowApplicationForm = mfrmApplicationPage.ShowApplicationForm(lngPatientID, lngCallFrom, lngPatientRegId, _
                                                lngPatientPageId, lngUpdateAppNoOrAdvId, objAppPages(), intBabyID, blnEdit, lngProjectId)

End Function

Private Sub Class_Terminate()
    Set mfrmApplicationPage = Nothing
End Sub
