VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDockEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event AfterSaved(lngRecordId As Long)                        'EMR文档保存后事件

Private WithEvents mobjDoc As cEPRDocument
Attribute mobjDoc.VB_VarHelpID = -1
Public Function GetOutEPREditForm(ByVal frmParent As Object, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngBillId As Long, ByVal lngDeptId As Long) As Long
'根据传入的文件ID,病人ID，挂号ID，科室ID，查询当前病人是否创建此病历文件，如果未创建，则返回新开此病历的窗体对象；
'如果已创建，则返回已创建的病历文件编辑窗体的对象。返回的窗体为无边框模式（类似虚拟模块）
Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    If mobjDoc Is Nothing Then
        Set mobjDoc = New cEPRDocument
    End If
    
    gstrSQL = "Select ID From 电子病历记录 Where 病人id = [1] And 主页id = [2] And 文件id + 0 = [3] And 科室id = [4]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查指定病历书写与否", lngPatiID, lngBillId, lngFileID, lngDeptId)
    If rsTemp.EOF Then
        mobjDoc.InitEPRDoc cprEM_新增, cprET_单病历编辑, lngFileID, cprPF_门诊, lngPatiID, lngBillId, , lngDeptId
    Else
        mobjDoc.InitEPRDoc cprEM_修改, cprET_单病历编辑, CLng(rsTemp!ID), cprPF_门诊, lngPatiID, lngBillId, , lngDeptId
    End If
    mobjDoc.ShowEPREditor frmParent, True, -1
    
    GetOutEPREditForm = mobjDoc.frmEditor.hWnd
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function DocModified() As Boolean
    If mobjDoc Is Nothing Then
        DocModified = False
    Else
        DocModified = mobjDoc.frmEditor.Editor1.Modified
    End If
End Function
Public Sub SaveDoc()
    If mobjDoc Is Nothing Then
        Exit Sub
    Else
        Call mobjDoc.frmEditor.SaveEMRDoc
    End If
End Sub
Private Sub DocObjectClear()
    On Error Resume Next
    Set mobjDoc.Tables = Nothing
    Set mobjDoc.Pictures = Nothing
    Set mobjDoc.Compends = Nothing
    Set mobjDoc.Elements = Nothing
    Set mobjDoc.Diagnosises = Nothing
    Set mobjDoc.Signs = Nothing

    Set mobjDoc.EPRFileInfo = Nothing
    Set mobjDoc.EPRDemoInfo = Nothing
    Set mobjDoc.EPRPatiRecInfo = Nothing

    Unload mobjDoc.frmEditor
    Set mobjDoc.frmEditor = Nothing
    
    Set mobjDoc = Nothing
    Err.Clear
End Sub
Private Sub Class_Terminate()
    Call DocObjectClear
End Sub

Private Sub mobjDoc_AfterSaved(lngRecordId As Long)
    RaiseEvent AfterSaved(lngRecordId)
End Sub
