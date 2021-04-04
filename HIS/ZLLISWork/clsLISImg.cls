VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Get_Chart2d_File(ByVal strPath As String, ByVal lngID As Long) As Boolean
    ' 读取数据库里的图形数据，生成Chart2d控件能够直接调用的文件。
    ' 入参：
    '   strPath : 生成的图形文件保存的路径
    '   lngID   : 检验图像结果的ID，保存的图形文件名为，strpath &"\" & lngID & ".cht"
    Get_Chart2d_File = LoadImageData(strPath, lngID)
End Function

Public Function Get_ReportCode(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByRef strCode As String, ByRef strNO As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    '取LIS报表的 编号
    Get_ReportCode = GetReportCode(lng医嘱ID, lng发送号, strCode, strNO, bytMode, DataMoved)
End Function

Public Sub ShowPatientRptPrint(cnOracle As ADODB.Connection, ByVal lngSys As Long, ByVal lngPatiDeptID As Long, ByVal MainPrivs As String, ByVal frmParent As Object)
    '病区批量打印检验报告
    'lngPatiDeptID  :病区ID
    'frmMain        :调用主窗体
    Dim frmRptPrint As New frmPatientRptPrint
    Call frmRptPrint.ShowME(cnOracle, lngSys, frmParent, lngPatiDeptID, MainPrivs)
End Sub
