VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabDemoInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'病历文件示范
Public ID As Long                           '示范ID
Public 文件ID As Long                       '原始文件定义ID
Public 编号 As String                       '示范编号
Public 名称 As String                       '示范名称
Public 性质 As Integer                      '示范性质:0-范文;1-片段
Public 通用级 As Long                       '0-全院通用;1-部门通用;2-个人使用
Public 科室ID As Long                       '记录科室ID
Public 人员ID As Long                       '记录人员ID
Public Sub GetDemoInfo(ByVal lngFileID As Long)
Dim rsTemp As New ADODB.Recordset

    On Error GoTo errHand
    gstrSQL = "Select ID, 文件ID, 编号, 名称, 性质, 说明, 通用级, 科室ID, 人员ID From 病历范文目录 where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cTableEPR", lngFileID)  '此时的lngFileId表示文件示范的ID
    With Me
        .ID = lngFileID
        If Not rsTemp.EOF Then
            .文件ID = Nvl(rsTemp!文件ID, 0)
            .编号 = Nvl(rsTemp!编号, 0)
            .名称 = Nvl(rsTemp!名称)
            .性质 = Nvl(rsTemp!性质, 0)
            .通用级 = Nvl(rsTemp!通用级, 0)
            .科室ID = Nvl(rsTemp!科室ID, 0)
            .人员ID = Nvl(rsTemp!人员ID, 0)
        End If
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function GetSaveSQL(arrSQL As Variant) As Boolean
    On Error GoTo errHand
    
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病历范文内容_Commit(" & Me.ID & ")"
    
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

