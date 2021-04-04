VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTabPatiRecordInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'记录信息
Public ID As Long                          '记录ID
Public 病人来源 As PatiFrom            '1-门诊;2-住院;3-外来(今后用于辅诊部门接收外来病人);4-体检病人
Public 病人ID As Long                      '
Public 主页ID As Long                      '住院病人必须填写
Public 婴儿 As Long                        '如果是婴儿，显示时将姓名处理为**之婴
Public 科室ID As Long                      '就诊科室,记录书写本次病历的医生或护士所在科室
Public 病历种类 As EPRDocType          '对应于病历文件的种类
Public 文件ID As Long                      '如果对应的病历文件已经删除,则格式按缺省处理。
Public 病历名称 As String                  '
Public 创建人 As String
Public 创建时间 As Date
Public 保存人 As String
Public 保存时间 As Date
Public 归档人 As String
Public 归档日期 As Date
Public 完成时间 As Date                    '在签名时填写,平常产生新的版本保存时不填写,表示正在进行中
Public 最后版本 As Integer                 '即目标签名版本
Public 签名级别 As EPRSignLevel        '1-书写;2-主治医师审阅;3-主任医师审阅。住院病历以外的病历只有书写和审阅状态
Public 医嘱id As Long
Public Sub GetPatiRecordInfo(ByVal lngFileID As Long, ByVal blnMoved As Boolean)
Dim rsTemp As New ADODB.Recordset, lngTmp As Long

    On Error GoTo errHand
    gstrSQL = "Select ID, 病人来源, 病人ID, 主页ID, 婴儿, 科室ID, 病历种类, 文件ID, " & _
            " 病历名称, 归档人, 归档日期, 保存人, 保存时间, 创建人, 创建时间, " & _
            " 最后版本 , 签名级别 , 完成时间 From 电子病历记录 Where ID=[1] "
    If blnMoved Then gstrSQL = Replace(gstrSQL, "电子病历记录", "H电子病历记录")
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "cTableEPR", lngFileID)  '此时的lngFileId表示病人电子病历的ID
    With Me
        .ID = lngFileID
        If Not rsTemp.EOF Then
            .病人来源 = Nvl(rsTemp("病人来源"), 1)
            .病人ID = Nvl(rsTemp("病人ID"), 0)
            .主页ID = Nvl(rsTemp("主页ID"), 0)
            .婴儿 = Nvl(rsTemp("婴儿"), 0)
            .科室ID = Nvl(rsTemp("科室ID"), 0)
            .病历种类 = Nvl(rsTemp("病历种类"), 0)
            .文件ID = Nvl(rsTemp("文件ID"), 0)
            .病历名称 = Nvl(rsTemp("病历名称"))
            .创建人 = Nvl(rsTemp("创建人"))
            .创建时间 = Nvl(rsTemp("创建时间"), 0)
            .保存人 = Nvl(rsTemp("保存人"))
            .保存时间 = Nvl(rsTemp("保存时间"), 0)
            .归档人 = Nvl(rsTemp("归档人"))
            .归档日期 = Nvl(rsTemp("归档日期"), 0)
            .完成时间 = Nvl(rsTemp("完成时间"), 0)
            .最后版本 = Nvl(rsTemp("最后版本"), 1)
            lngTmp = Nvl(rsTemp("签名级别"), 0)     '记录本次编辑的初始版本状态
            If lngTmp >= 4 Then
                .签名级别 = TabSL_主任
            ElseIf lngTmp >= 2 Then
                .签名级别 = TabSL_主治
            ElseIf lngTmp >= 1 Then
                .签名级别 = TabSL_经治
            Else
                .签名级别 = TabSL_空白
            End If
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
    arrSQL(UBound(arrSQL)) = "Zl_电子病历记录_Apply(" & Me.ID & "," & Me.病人来源 & "," & Me.病人ID & "," & Me.主页ID & _
                            "," & Me.婴儿 & "," & IIf(Me.科室ID = 0, UserInfo.部门ID, Me.科室ID) & "," & Me.病历种类 & "," & Me.文件ID & ",'" & Me.病历名称 & _
                            "','" & Me.创建人 & "','" & UserInfo.姓名 & "'," & Me.医嘱id & ")"
    GetSaveSQL = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

