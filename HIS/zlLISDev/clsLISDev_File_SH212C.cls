VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_SH212C"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'江津西城 SH212C
Private maccess As New ADODB.Connection
Private mstrsql As String
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, ByVal dtStart As Date, _
                    Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    
    Dim access As New ADODB.Recordset, Rsym As New ADODB.Recordset
    Dim i As Integer, strPass As String
    Dim str结果 As String
    Dim str值 As String
    Dim aResults() As String, strFilePath As String, objFile As New Scripting.FileSystemObject, strFileErs As String
    i = UBound(Split(strFile, "/"))
    
    If i >= 1 Then
       strPass = Split(strFile, "/")(1)
       strFilePath = objFile.GetParentFolderName(Split(strFile, "/")(0)) '取路径
       strFile = Split(strFile, "/")(0)
    Else
       strFilePath = objFile.GetParentFolderName(strFile) '取路径
    End If
    
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    strFileErs = strFile

    If i >= 1 Then
       If Trim(Dir(strFileErs)) <> "" Then
          maccess.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & strFile & _
                                     ";Jet OLEDB:Database Password=" & strPass
          maccess.Open
          If maccess.State = adStateClosed Then strFileErs = ""
       Else
          strFileErs = ""
       End If
    Else
       If Trim(Dir(strFileErs)) <> "" Then
          maccess.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & strFile
          maccess.Open
          If maccess.State = adStateClosed Then strFileErs = ""
       Else
          strFileErs = ""
       End If
    End If
    ReDim aResults(0)
        
    If strFileErs <> "" Then
    
    
        mstrsql = "SELECT 送检日期 AS 检验时间,编号ID AS 标本号,[全血高切粘度(200/S)] as a11,[全血中切粘度(40/S)] as a12,[全血中切粘度(30/S)] as a13,[全血低切粘度(10/S)] as a14,[全血低切粘度(1/S)] as a15,[高切流阻(e9SI)] as a16,[中切流阻(e9SI)] as a17,[低切流阻(e9SI)] as a18,[全血卡森粘度(mPas)] as a19,[全血卡森应力(dyn/cm^2)] as a20," & _
        "[血浆粘度(mPas)] as a21,[红细胞压积(Hct)] as a22,[全血还原粘度(200/S)] as a23,[全血还原粘度(40/S)] as a24,[全血还原粘度(30/S)] as a25,[全血还原粘度(10/S)] as a26,[全血还原粘度(1/S)] as a27,[全血高切相对粘度(BR)] as a28,[全血低切相对粘度(Br)] as a29,[红细胞聚集指数(AI)] as a30,[红细胞聚集系数(VAI)] as a31,[红细胞变形性(TK)] as a32,[红细胞刚性指数(IR)] as a33,[红细胞计数(e12个/L)] as a34,[红细胞内粘度(mPas)] as a35,[红细胞电泳时间(s)] as a36,[血红蛋白浓度(g/L)] as a37," & _
        "[红细胞沉降率(mm/h)] as a38,[血小板数粘附前] as a39,[血小板数粘附后] as a40,[血小板粘附率(PAd)] as a41,[血沉方程K值(ESRK)] as a42,[体外血栓长度(mm)] as a43,[体外血栓干重(mg)] as a44,[体外血栓湿重(mg)] as a45,[纤维蛋白原(g/L)] as a46,[胆固醇(mmol/L)] as a47,[甘油三脂(mmol/L)] as a48,[血糖(mmol/L)] as a49,[高密度脂蛋白(mmol/L)] as a50,[低密度脂蛋白(mmol/L)] as a51 " & _
        " FROM 病人数据表  where "
        
        mstrsql = mstrsql & "Format(送检日期,'yyyy-MM-dd')>='" & _
        Format(dtStart, "yyyy-MM-dd") & "'"
        mstrsql = mstrsql & " And Format(送检日期,'yyyy-MM-dd')<='" & _
        Format(dtEnd, "yyyy-MM-dd") & "'"
        
        Call openaccess(Rsym, "取明细")
              
        Do While Not Rsym.EOF
            str结果 = Rsym("检验时间") & "|" & Rsym("标本号") & "|" & " |血液|0"
            
            For i = 11 To 51
               
                str值 = "a" & i
                str结果 = str结果 & "|" & i & "|" & Rsym(str值)
              
            Next
        
            If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                     aResults(UBound(aResults)) = str结果
             Rsym.MoveNext
        Loop

        ResultFromFile = aResults
        str结果 = ""
        maccess.Close
    End If
End Function

Private Sub openaccess(rsTemp As ADODB.Recordset, ByVal strCaption As String)
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open mstrsql, maccess, adOpenStatic, adLockReadOnly
End Sub



Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function




