VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public gobjRegister As Object               '注册授权部件zlRegister
Private WithEvents gclsDockAduits As zlRichEPR.clsDockAduits, mstrPrintDocIDs As String
Attribute gclsDockAduits.VB_VarHelpID = -1

Private Function blnDebug() As Boolean
Dim l As Long
    Err.Clear
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then
        Err.Clear
        blnDebug = True: Exit Function
    End If
    
    blnDebug = False: Exit Function
    
End Function
Public Function PrintDocument(ByVal objFrm As Object, ByVal strServer As String, ByVal strUser As String, ByVal strPass As String, ByVal strPrintContent As String, ByVal strDeviceName As String) As Boolean
'strPrintContent以单个打印对象用竖线|分隔,单个打印对象参数间以逗号分隔 姓名,类别,......
Dim clsPath As Object, clsTendsNew As Object, clsReport As Object, clsComlib As Object, clsDatabase As Object, clsRichEPR As Object, clsCommFun As Object, strErr As String, clsKernel As Object
Dim l As Long, arrPar As Variant, arrParOne As Variant, blnTrans As Boolean, strSQL As String, strError As String
Dim cnMain As New ADODB.Connection

    On Error Resume Next
    If gobjRegister Is Nothing Then
        Set gobjRegister = GetObject("", "zlRegister.clsRegister")
        Err.Clear
        If gobjRegister Is Nothing Then
            Set gobjRegister = CreateObject("zlRegister.clsRegister")
            Err.Clear
            If gobjRegister Is Nothing Then
                MsgBox "创建zlRegister部件对象失败,请检查文件是否存在并且正确注册。", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    On Error GoTo errHand

    If cnMain.State = adStateOpen Then cnMain.Close
    Set cnMain = gobjRegister.GetConnection(strServer, strUser, strPass, True, 0, strError)

    If strError <> "" Then MsgBox strError, vbInformation
    If cnMain.State = adStateClosed Then Exit Function
    
    Set clsComlib = CreateObject("zl9ComLib.clsComLib")
    clsComlib.InitCommon cnMain
    Set clsDatabase = CreateObject("zl9ComLib.clsDatabase")
    Set clsCommFun = CreateObject("zl9ComLib.clsCommFun")
    Set clsReport = CreateObject("zl9Report.clsReport")
    clsReport.InitOracle cnMain
    
    Err.Clear
    Call clsCommFun.PDFInitialize(strErr)
    If Err.Number <> 0 Or strErr <> "" Then
        Err.Raise vbObjectError, , "PDF设备初始化失败:" & strErr
    End If
    
    arrPar = Split(strPrintContent, "|")
    For l = 0 To UBound(arrPar)
        arrParOne = Split(arrPar(l), ",")
        Select Case arrParOne(1)
            Case "R5"               '首页
                '系统号,报表编号,病人id,主页id,正/反/附一/附二,PDFFileName
                objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "首页"
                
                Call SaveSetting("ZLSOFT", "私有模块\" & strUser & "\zl9Report\LocalSet\" & arrParOne(3), "Printer", strDeviceName)
                Call clsReport.ReportOpen(cnMain, CLng(arrParOne(2)), arrParOne(3), objFrm, "病人id=" & arrParOne(4), "主页id=" & arrParOne(5), "ReportFormat=" & arrParOne(6), "PDF=" & arrParOne(7), 4)
            Case "R1"               '医嘱
                '系统号,报表编号,病人id,主页id,医嘱单A0/A1/医嘱本B,PDFFileName
                objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "医嘱"
                
                Call SaveSetting("ZLSOFT", "私有模块\" & strUser & "\zl9Report\LocalSet\" & arrParOne(3), "Printer", strDeviceName)
                If UCase(arrParOne(6)) = "A0" Then '医嘱单长嘱
                    Call clsCommFun.PDFFile(CStr(arrParOne(7)))
                    If clsKernel Is Nothing Then
                        Set clsKernel = CreateObject("zlCISKernel.clsCISKernel")
                        Call clsKernel.InitCISKernel(cnMain, objFrm, arrParOne(2), "")
                    End If
                    Call clsKernel.zlPrintAdvice(objFrm, arrParOne(4), arrParOne(5), 0, 0, strDeviceName)
                ElseIf UCase(arrParOne(6)) = "A1" Then '医嘱单临嘱
                    Call clsCommFun.PDFFile(CStr(arrParOne(7)))
                    If clsKernel Is Nothing Then
                        Set clsKernel = CreateObject("zlCISKernel.clsCISKernel")
                        Call clsKernel.InitCISKernel(cnMain, objFrm, arrParOne(2), "")
                    End If
                    Call clsKernel.zlPrintAdvice(objFrm, arrParOne(4), arrParOne(5), 0, 1, strDeviceName)
                Else
                    Call clsReport.ReportOpen(cnMain, CLng(arrParOne(2)), arrParOne(3), objFrm, "病人id=" & arrParOne(4), "主页id=" & arrParOne(5), "PDF=" & arrParOne(7), 4)
                End If
            Case "R9"               '临床路径
                '系统号,FileName,病人ID,主页ID
                objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "临床路径"
                If clsPath Is Nothing Then
                    Set clsPath = CreateObject("zlCISPath.clsDockPath")
                    Dim objPath As Object
                    Set objPath = CreateObject("zlCISPath.clsCISPath")
                    Call objPath.InitCISPath(cnMain, objFrm, CLng(arrParOne(2)), "")
                End If
                Call clsPath.zlFuncPathTableOutPut(4, True, CStr(arrParOne(3)), CLng(arrParOne(4)), CLng(arrParOne(5)), strDeviceName)
            Case "R2"               '住院病历
                '系统号,FileName,ID
                objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "病历:" & "住院病历_" & arrParOne(4)
                If clsRichEPR Is Nothing Then
                    Set clsRichEPR = CreateObject("zlRichEPR.cRichEPR")
                    Call clsRichEPR.InitRichEPR(cnMain, objFrm, CLng(arrParOne(2)), False)
                    Set gclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
                End If
                
                If InStr("," & mstrPrintDocIDs, "," & Val(arrParOne(4)) & ",") = 0 Then '本次没打过
                    Call clsCommFun.PDFFile(CStr(arrParOne(3)))
                    Call gclsDockAduits.zlPrintDocument(3, 2, Val(arrParOne(4)), strDeviceName)
                End If
            Case "R3"               '护理病历
                '系统号,FileName,ID
                objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "病历:" & "护理病历_" & arrParOne(4)
                If clsRichEPR Is Nothing Then
                    Set clsRichEPR = CreateObject("zlRichEPR.cRichEPR")
                    Call clsRichEPR.InitRichEPR(cnMain, objFrm, CLng(arrParOne(2)), False)
                    Set gclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
                End If
                
                If InStr("," & mstrPrintDocIDs, "," & Val(arrParOne(4)) & ",") = 0 Then '本次没打过
                    Call clsCommFun.PDFFile(CStr(arrParOne(3)))
                    Call gclsDockAduits.zlPrintDocument(3, 2, Val(arrParOne(4)), strDeviceName)
                End If
            Case "R4"               '护理记录
                '系统号,新版N/旧版O,体温单1/护理记录单2/产程图3,FileName,病人ID,主页ID,科室ID,婴儿序号,lngKey/lngFileID,Period
                If UCase(arrParOne(3)) = "O" Then '旧版'
                    If clsRichEPR Is Nothing Then
                        Set clsRichEPR = CreateObject("zlRichEPR.cRichEPR")
                        Call clsRichEPR.InitRichEPR(cnMain, objFrm, CLng(arrParOne(2)), False)
                        Set gclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
                    End If
                    
                    If Val(arrParOne(4)) = 1 Then
                        objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "体温单"
                        Call clsCommFun.PDFFile(CStr(arrParOne(5)))
                        Call gclsDockAduits.zlRefreshTendBody(Val(arrParOne(6)), Val(arrParOne(7)), Val(arrParOne(8)), CInt(Val(arrParOne(9))))
                        Call gclsDockAduits.zlPrintDocument(1, 2, , strDeviceName)
                    Else
                        objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "护理记录单"
                        Call clsCommFun.PDFFile(CStr(arrParOne(5)))
                        Call gclsDockAduits.zlRefresh(3, Val(arrParOne(10)), Val(arrParOne(6)), Val(arrParOne(7)), Val(arrParOne(8)), CStr(arrParOne(11)), , CInt(Val(arrParOne(9))))
                        Call gclsDockAduits.zlPrintDocument(2, 2, , strDeviceName)
                    End If
                Else  '新版
                    objFrm.Label2.Caption = "开始输出" & arrParOne(0) & IIf(arrParOne(4) = 1, "体温单", IIf(arrParOne(4) = 2, "护理记录单", "产程图"))
                    If clsTendsNew Is Nothing Then
                        Set clsTendsNew = CreateObject("zl9TendFile.clsTendFile")
                        Call clsTendsNew.InitTendFile(cnMain, Val(arrParOne(2)))
                    End If
                    
                    Call clsCommFun.PDFFile(CStr(arrParOne(5)))
                    
                    Call clsTendsNew.zlPrintDocument(Val(arrParOne(6)), Val(arrParOne(7)), CInt(Val(arrParOne(9))), Val(arrParOne(8)), Val(arrParOne(10)), CByte(Val(arrParOne(4))), strDeviceName)
                End If
            Case "R6"               '医嘱报告
                '系统号,FileName,ID
                objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "医嘱报告" & arrParOne(4)
                If clsRichEPR Is Nothing Then
                    Set clsRichEPR = CreateObject("zlRichEPR.cRichEPR")
                    Call clsRichEPR.InitRichEPR(cnMain, objFrm, CLng(arrParOne(2)), False)
                    Set gclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
                End If
            
                Call clsCommFun.PDFFile(CStr(arrParOne(3)))
                Call gclsDockAduits.zlPrintDocument(4, 2, Val(arrParOne(4)), strDeviceName & "|" & arrParOne(3))
            Case "R7"               '疾病证明
                '系统号,FileName,ID
                objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "疾病证明" & arrParOne(4)
                If clsRichEPR Is Nothing Then
                    Set clsRichEPR = CreateObject("zlRichEPR.cRichEPR")
                    Call clsRichEPR.InitRichEPR(cnMain, objFrm, CLng(arrParOne(2)), False)
                    Set gclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
                End If
            
                Call clsCommFun.PDFFile(CStr(arrParOne(3)))
                Call gclsDockAduits.zlPrintDocument(3, 2, Val(arrParOne(4)), strDeviceName)
            Case "R8"               '知情文件
                '系统号,FileName,ID
                objFrm.Label2.Caption = "开始输出" & arrParOne(0) & "知情文件" & arrParOne(4)
                If clsRichEPR Is Nothing Then
                    Set clsRichEPR = CreateObject("zlRichEPR.cRichEPR")
                    Call clsRichEPR.InitRichEPR(cnMain, objFrm, CLng(arrParOne(2)), False)
                    Set gclsDockAduits = CreateObject("zlRichEPR.clsDockAduits")
                End If
            
                Call clsCommFun.PDFFile(CStr(arrParOne(3)))
                Call gclsDockAduits.zlPrintDocument(3, 2, Val(arrParOne(4)), strDeviceName)
        End Select
    Next
    PrintDocument = True
    
    Exit Function
errHand:
    If blnTrans Then cnMain.RollbackTrans
    objFrm.Label2.Caption = Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Sub gclsDockAduits_AfterEprPrint(ByVal lngRecordId As Long)
    mstrPrintDocIDs = mstrPrintDocIDs & lngRecordId & ","
End Sub


