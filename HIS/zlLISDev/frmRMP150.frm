VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRMP150 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "导入设置"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6375
   Icon            =   "frmRMP150.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chkFile 
      Alignment       =   1  'Right Justify
      Caption         =   "数据文件方式"
      Height          =   255
      Left            =   270
      TabIndex        =   6
      ToolTipText     =   "选择按项目导入还是按单个文件导入"
      Top             =   3090
      Width           =   1380
   End
   Begin VB.CheckBox chkReplace 
      Caption         =   "覆盖"
      Height          =   225
      Left            =   2490
      TabIndex        =   5
      ToolTipText     =   "不勾：文件如果已经导入过，则不导入；勾上：要据所选文件导入，而不管文件是否已经导入过。"
      Top             =   3405
      Width           =   885
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   300
      Left            =   705
      TabIndex        =   3
      Top             =   3360
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   111607811
      CurrentDate     =   39658
   End
   Begin MSScriptControlCtl.ScriptControl vbsCalce 
      Left            =   5295
      Top             =   3090
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5100
      TabIndex        =   1
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确认(&O)"
      Height          =   350
      Left            =   3870
      TabIndex        =   0
      Top             =   3270
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgList 
      Height          =   2940
      Left            =   30
      TabIndex        =   2
      Top             =   45
      Width           =   6300
      _cx             =   11112
      _cy             =   5186
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   6
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lbl日期 
      AutoSize        =   -1  'True
      Caption         =   "日期:"
      Height          =   180
      Left            =   225
      TabIndex        =   4
      Top             =   3405
      Width           =   450
   End
End
Attribute VB_Name = "frmRMP150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFilePath As String                      '文件路径
Private mdtStart As Date
Private mdtEnd As Date
Private mlngSampleNo As Long
Private mobjFile As New Scripting.FileSystemObject  '文件对像
Private mstrReturn() As String '存返回结果

Enum mCol
    选择 = 0:    项目编号: 英文: 项目名称: 检验时间: 起始标本号: 小数: CutOff公式: 阳性公式: 弱阳公式: 最小值: 文件名
End Enum
'---读写INI文件的API声明
#If Win32 Then
   Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#Else
   Private Declare Function GetPrivateProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
   Private Declare Function WritePrivateProfileString Lib "Kernel" (ByVal Appname As String, ByVal KeyName As Any, ByVal NewString As Any, ByVal Filename As String) As Integer
#End If
'----------------------

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_NOMOVE = &H2 '注释：不移动窗体
Private Const SWP_NOSIZE = &H1 '注释：不改变窗体尺寸
Private Const HWND_TOPMOST = -1         '注释：窗体总在最前面
Private Const HWND_NOTOPMOST = -2       '注释：窗体不在最前面


Public Function ShowMe(ByVal strFile As String, ByVal lngSampleNO As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    
    mstrFilePath = mobjFile.GetParentFolderName(strFile)
    If Len(mstrFilePath) = 0 Then mstrFilePath = App.Path
    mdtStart = dtStart
    mdtEnd = dtEnd
    mlngSampleNo = lngSampleNO
    ReDim mstrReturn(0) As String
    Me.Show vbModal
    ShowMe = mstrReturn
End Function

Private Sub ImpFile()
    Dim objStream As TextStream, strLine As String, str日期 As String
    Dim strFileName As String, strDataName As String
    Dim strImpedFile As String, strImpFile As String
    Dim intCount As Integer
    
    Dim lngFileCount As Long, lngLoop As Long, strTmpFile As String, bln已导入 As Boolean
    strFileName = mstrFilePath & "\RMPDAT.IDX"
    
    If chkFile = 0 Then
        '--------------------------------------------------------------------------------
        '---  自动导入
        '--------------------------------------------------------------------------------
        '打开配置文件 读取已导入文件列表
        
        If chkReplace.Value <> 0 Then
            '覆盖方式，标本号清０
            With vfgList
                For intCount = .FixedRows To .Rows - 1
                    If .TextMatrix(intCount, mCol.选择) = 1 Then
                        WriteToIni App.Path & "\RMP150.ini", Format(dtpDate.Value, "yyyy-MM-dd"), .TextMatrix(intCount, mCol.英文), 0
                    End If
                Next
            End With
        Else
            strLine = ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & Format(dtpDate.Value, "yyyy-MM-dd"), "Count")
            If Val(strLine) > 0 Then
                For intCount = 1 To Val(strLine)
                    strImpedFile = strImpedFile & "," & ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & Format(dtpDate.Value, "yyyy-MM-dd"), "File" & intCount)
                Next
            End If
        End If
        
        '打开 RMPDAT.IDX文件,找出要导入的文件,保存到待导入文件列表
        Set objStream = mobjFile.OpenTextFile(strFileName, ForReading)
        Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            If InStr(strLine, ".DAT=UNDEFINED,") > 0 And InStr(strLine, "RMP") > 0 Then
                str日期 = Split(strLine, ",")(2)
                str日期 = Right(str日期, 2) & "/" & Mid(str日期, 1, 5)
                strDataName = Replace(Split(strLine, ",")(0), "=UNDEFINED", "")
                
                If Format(CDate(str日期), "yyyy-MM-dd") = Format(dtpDate.Value, "yyyy-MM-dd") _
                   And InStr(strImpedFile, "," & strDataName) <= 0 Then
                    
                    strImpFile = strImpFile & "," & strDataName & "|" & Format(CDate(str日期), "yyyy-MM-dd")
                    
                End If
            End If
        Loop
    Else
        With vfgList
            For intCount = .FixedRows To .Rows - 1
                If .TextMatrix(intCount, mCol.选择) = 1 Then
                    
                    If chkReplace.Value = 0 Then
                        '检查是否已导入过
                        bln已导入 = False
                        lngFileCount = Val(ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & Format(dtpDate.Value, "yyyy-MM-dd"), "Count"))
                        If lngFileCount > 0 Then
                           For lngLoop = 1 To lngFileCount
                               strTmpFile = ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & Format(dtpDate.Value, "yyyy-MM-dd"), "File" & lngLoop)
                               If strTmpFile = .TextMatrix(intCount, mCol.文件名) Then
                                   bln已导入 = True
                                   Exit For
                               End If
                           Next
                        End If
                        If Not bln已导入 Then strImpFile = strImpFile & "," & .TextMatrix(intCount, mCol.文件名) & "|" & Format(dtpDate.Value, "yyyy-MM-dd")
                    Else
                        strImpFile = strImpFile & "," & .TextMatrix(intCount, mCol.文件名) & "|" & Format(dtpDate.Value, "yyyy-MM-dd")
                    End If
                End If
            Next
        End With
    End If
    '计算结果
    If strImpFile <> "" Then Call ImpItem(strImpFile)
End Sub

Private Sub ImpItem(ByVal strImpFiles As String)
    '计算结果
    Dim varItem As Variant, strFileName As String, intRow As Integer
    Dim objStream As TextStream
    
    Dim dblBC_Total As Double, dblBC As Double, intBC_Count As Integer                  '空白孔
    Dim dblNC_Total As Double, dblNC As Double, intNC_Count As Integer                  '阴性孔
    Dim dblPC_Total As Double, dblPC As Double, intPC_Count As Integer                  '阳性孔
    Dim iCount As Integer '总记录数
    Dim iRow As Integer, strType As String, strXY As String, dblOD As Double
    Dim i序号 As Integer
    Dim arrData(3, 1 To 8, 1 To 12) As String                '(0=编号;1=原始OD:2=OD;3=定性)
    Dim lngResultCount As Long
    Dim strIni标本号 As String, str日期 As String, str项目 As String
    '--- 读取要导入的项目和公式
    Dim str要导的项目 As String, int小数 As Integer, strCoutOff公式 As String, str阳性公式 As String, str弱阳性公式 As String
    Dim dblCoutOff值 As Double, bln阳性 As Boolean, bln弱阳性 As Boolean, int开始标本号 As String, strLine As String
    Dim strOD As String, dbl阴性孔最小值 As Double, str定性结果 As String
    Dim intFileCount As Integer, str公式 As String
    varItem = Split(Mid(strImpFiles, 2), ",")
    lngResultCount = -1
    For intRow = LBound(varItem) To UBound(varItem)
        strFileName = varItem(intRow)
        
        str日期 = Split(strFileName, "|")(1)
        strFileName = Split(strFileName, "|")(0)
        i序号 = 0
        '读取数据
        str项目 = ReadFromIni(mstrFilePath & "\" & strFileName, "Test Log", "Test Class")
        str项目 = Mid(str项目, InStr(str项目, " ") + 1)
        int小数 = 0: strCoutOff公式 = "": str阳性公式 = "": int开始标本号 = 0
        intFileCount = Val(ReadFromIni(App.Path & "\RMP150.ini", "ImpFile " & str日期, "Count"))
        With vfgList
            For iRow = .FixedRows To .Rows - 1
                If str项目 = .TextMatrix(iRow, mCol.英文) And .TextMatrix(iRow, mCol.选择) = 1 Then
                    int小数 = Val(.TextMatrix(iRow, mCol.小数))
                    strCoutOff公式 = .TextMatrix(iRow, mCol.CutOff公式)
                    str阳性公式 = .TextMatrix(iRow, mCol.阳性公式)
                    str弱阳性公式 = .TextMatrix(iRow, mCol.弱阳公式)
                    int开始标本号 = Val(ReadFromIni(App.Path & "\RMP150.ini", str日期, str项目))
                    dbl阴性孔最小值 = Val(.TextMatrix(iRow, mCol.最小值))
                    Exit For
                End If
            Next
        End With
        
        If strCoutOff公式 <> "" Then
            iCount = Val(ReadFromIni(mstrFilePath & "\" & strFileName, "Test Log", "Nof Results"))
            
            dblBC_Total = 0: intBC_Count = 0
            dblNC_Total = 0: intNC_Count = 0
            dblPC_Total = 0: intPC_Count = 0
            
            dblBC = 0: dblNC = 0: dblPC = 0
            If iCount > 0 Then
                For iRow = 1 To iCount
                    
                    strType = ReadFromIni(mstrFilePath & "\" & strFileName, "Result " & iRow, "Liquid Type") '类型
                    strXY = ReadFromIni(mstrFilePath & "\" & strFileName, "Result " & iRow, "Position Name") '座标
                    strOD = ReadFromIni(mstrFilePath & "\" & strFileName, "Result " & iRow, "OD Values")    '检测结果
                    
                   
                    
                    If strOD <> "" Then
                        dblOD = Val(strOD)
                        
                        If strType = "blk" Then                             '空白孔
                            dblBC_Total = dblBC_Total + dblOD
                            intBC_Count = intBC_Count + 1
                        ElseIf strType = "nc" Then                          '阴性孔
                            dblNC_Total = dblNC_Total + dblOD
                            intNC_Count = intNC_Count + 1
                        ElseIf strType = "pc" Then                          '阳性孔
                            dblPC_Total = dblPC_Total + dblOD
                            intPC_Count = intPC_Count + 1
                        ElseIf strType = "smp" Then                         'OD值
                            
                            If dblBC = 0 Then dblBC = dblBC_Total / intBC_Count ' / 1077936128 'intBC_Count
                            If dblNC = 0 Then dblNC = (dblNC_Total / intNC_Count - dblBC) ' / 1073741824
                            If dblPC = 0 Then dblPC = (dblPC_Total / intPC_Count - dblBC) '/ 1073741824
                            
                            dblOD = dblOD - dblBC
                            
                            If dblNC < dbl阴性孔最小值 Then dblNC = dbl阴性孔最小值
                            If dblOD < dbl阴性孔最小值 Then dblOD = dbl阴性孔最小值
                            
                            '根据公式计算，阴阳性
                            str公式 = strCoutOff公式
                            str公式 = Replace(str公式, "[NC]", dblNC)
                            str公式 = Replace(str公式, "[PC]", dblPC)
                            'strCoutOff公式 = Replace(strCoutOff公式, "[BC]", dblBC)
                            
                            dblCoutOff值 = vbsCalce.Eval(str公式)
                                
                            '阳性公式
                            str公式 = str阳性公式
                            str公式 = Replace(str公式, "[NC]", dblNC)
                            str公式 = Replace(str公式, "[PC]", dblPC)
                            'str阳性公式 = Replace(str阳性公式, "[BC]", dblBC)
                            str公式 = Replace(str公式, "[OD]", dblOD)
                            
                            bln阳性 = vbsCalce.Eval(str公式)
                            
                            '弱阳性公式
                            str公式 = str弱阳性公式
                            str公式 = Replace(str公式, "[NC]", dblNC)
                            str公式 = Replace(str公式, "[PC]", dblPC)
                            'str弱阳性公式 = Replace(str弱阳性公式, "[BC]", dblBC)
                            str公式 = Replace(str公式, "[OD]", dblOD)
                            
                            bln弱阳性 = vbsCalce.Eval(str公式)
                            If bln阳性 Then
                                str定性结果 = "阳性"
                            Else
                                If bln弱阳性 Then
                                    str定性结果 = "弱阳性"
                                Else
                                    str定性结果 = "阴性"
                                End If
                            End If
                            
                            i序号 = i序号 + 1
                            If mlngSampleNo = -1 Then
                                lngResultCount = lngResultCount + 1
                                ReDim Preserve mstrReturn(lngResultCount)
                                mstrReturn(lngResultCount) = str日期 & "|" & int开始标本号 + i序号 & "| |血清|0|" & str项目 & "|" & str定性结果 & "^" & Format(dblOD, "0." & String(int小数, "0")) & "^" & Format(dblCoutOff值, "0." & String(int小数, "0")) & "^" & Format(dblOD / dblCoutOff值, "0." & String(int小数, "0"))
                            Else
                                If int开始标本号 + lngResultCount = mlngSampleNo Then
                                    lngResultCount = lngResultCount + 1
                                    ReDim Preserve mstrReturn(lngResultCount)
                                    mstrReturn(lngResultCount) = str日期 & "|" & int开始标本号 + i序号 & "| |血清|0|" & str项目 & "|" & str定性结果 & "^" & Format(dblOD, "0." & String(int小数, "0")) & "^" & Format(dblCoutOff值, "0." & String(int小数, "0")) & "^" & Format(dblOD / dblCoutOff值, "0." & String(int小数, "0"))
                                End If
                            End If
                            WriteLog strFileName, "项目类型：" & strType & " 位置：" & strXY & " 原始OD=" & strOD, "BC=" & Format(dblBC, "0.000") & ",NC=" & Format(dblNC, "0.000") & ",PC=" & Format(dblPC, "0.000") & ",OD=" & Format(dblOD, "0.000")
                        End If
                        
                        
                    End If 'if strOd<>""
                Next
            End If  'iCount > 0
            WriteToIni App.Path & "\RMP150.ini", str日期, str项目, int开始标本号 + i序号
        
            intFileCount = intFileCount + 1
            WriteToIni App.Path & "\RMP150.ini", "ImpFile " & str日期, "Count", intFileCount
            WriteToIni App.Path & "\RMP150.ini", "ImpFile " & str日期, "File" & intFileCount, strFileName

        End If 'strCoutOff公式 <> ""
    Next
    
End Sub

Private Sub WriteToIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String, ByVal Value As String)
''写INI文件
    Dim buff As String * 128
    buff = Trim(Value) + Chr(0)
    WritePrivateProfileString Section, Key, buff, Filename

End Sub

Private Function ReadFromIni(ByVal Filename As String, ByVal Section As String, ByVal Key As String) As String
''读INI文件
    Dim i As Long
    Dim buff As String * 128
    GetPrivateProfileString Section, Key, "", buff, 128, Filename
    i = InStr(buff, Chr(0))
    ReadFromIni = Trim(Left(buff, i - 1))
End Function

Private Sub chkFile_Click()
    If chkFile.Value = 0 Then
        Call ShowItemList
    Else
        Call ShowFileList
    End If
End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Call ImpFile
    Unload Me
End Sub

Private Sub dtpDate_Change()
    Dim str标本号 As String, intRow As Integer
    With vfgList
        For intRow = .FixedRows To .Rows - 1
            str标本号 = ReadFromIni(App.Path & "\RMP150.ini", Format(dtpDate.Value, "yyyy-MM-dd"), .TextMatrix(intRow, mCol.英文))
            .TextMatrix(intRow, mCol.起始标本号) = Val(str标本号) + 1
        Next
    End With
End Sub

Private Sub Form_Load()

     
     dtpDate.MinDate = mdtStart
     dtpDate.MaxDate = mdtEnd
     dtpDate.Value = mdtStart
     Call ShowItemList
    '置顶显示窗体
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = mCol.起始标本号 Then
        vfgList.TextMatrix(Row, Col) = CLng(Val(vfgList.TextMatrix(Row, Col)))
        WriteToIni App.Path & "\RMP150.ini", Format(dtpDate.Value, "yyyy-MM-dd"), vfgList.TextMatrix(Row, mCol.英文), CLng(Val(vfgList.TextMatrix(Row, Col))) - 1
    End If
End Sub

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.起始标本号 Then Cancel = True
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .MouseRow <> 0 Then
            If .TextMatrix(.MouseRow, 0) = 1 Then
                .TextMatrix(.MouseRow, 0) = 0
            Else
                .TextMatrix(.MouseRow, 0) = 1
            End If
        End If
    End With
End Sub

Private Sub ShowItemList()
    '读配置文件
    Dim iCount As Integer, iLoop As Integer
    Dim strItem As String, strItems As String, varItem As Variant
    Dim str标本号 As String
    vfgList.Rows = 2: vfgList.Cols = 12
    vfgList.TextMatrix(0, mCol.选择) = "": vfgList.ColWidth(mCol.选择) = 300: vfgList.ColDataType(mCol.选择) = flexDTBoolean
    vfgList.TextMatrix(0, mCol.项目编号) = "编号": vfgList.ColWidth(mCol.项目编号) = 450
    vfgList.TextMatrix(0, mCol.英文) = "英文": vfgList.ColWidth(mCol.英文) = 1000
    vfgList.TextMatrix(0, mCol.项目名称) = "项目名称": vfgList.ColWidth(mCol.项目名称) = 2500
    vfgList.TextMatrix(0, mCol.起始标本号) = "起始标本号": vfgList.ColWidth(mCol.起始标本号) = 1200: vfgList.ColDataType(mCol.起始标本号) = flexDTLong
    vfgList.TextMatrix(0, mCol.小数) = "小数": vfgList.ColWidth(mCol.小数) = 0
    vfgList.TextMatrix(0, mCol.CutOff公式) = "CutOff公式": vfgList.ColWidth(mCol.CutOff公式) = 0
    vfgList.TextMatrix(0, mCol.阳性公式) = "阳性公式": vfgList.ColWidth(mCol.阳性公式) = 0
    vfgList.TextMatrix(0, mCol.弱阳公式) = "弱阳公式": vfgList.ColWidth(mCol.弱阳公式) = 0
    vfgList.TextMatrix(0, mCol.最小值) = "最小值": vfgList.ColWidth(mCol.最小值) = 0
    vfgList.TextMatrix(0, mCol.检验时间) = "时间": vfgList.ColWidth(mCol.检验时间) = 0
    vfgList.TextMatrix(0, mCol.文件名) = "文件名": vfgList.ColWidth(mCol.文件名) = 0
     
     
     strItems = ReadFromIni(App.Path & "\RMP150.ini", "Base", "Item")
     
     If Len(strItems) <= 0 Then '设为默认值
         Exit Sub
     Else
         varItem = Split(strItems, "|")
         For iLoop = LBound(varItem) To UBound(varItem)
             strItem = ReadFromIni(App.Path & "\RMP150.ini", varItem(iLoop), "Info")
             '编号|英文|中文|小数位数|CoutOff公式|阳性公式|弱阳公式|阴性对照最小值
             With vfgList
                 .TextMatrix(.Rows - 1, mCol.选择) = 1
                 .TextMatrix(.Rows - 1, mCol.项目编号) = Split(strItem, "|")(0)
                 .TextMatrix(.Rows - 1, mCol.英文) = varItem(iLoop)
                 .TextMatrix(.Rows - 1, mCol.项目名称) = Split(strItem, "|")(1)
                 
                  str标本号 = ReadFromIni(App.Path & "\RMP150.ini", Format(dtpDate.Value, "yyyy-MM-dd"), varItem(iLoop))
                 .TextMatrix(.Rows - 1, mCol.起始标本号) = Val(str标本号) + 1
                 
                 .TextMatrix(.Rows - 1, mCol.小数) = Split(strItem, "|")(2)
                 .TextMatrix(.Rows - 1, mCol.CutOff公式) = Split(strItem, "|")(3)
                 .TextMatrix(.Rows - 1, mCol.阳性公式) = Split(strItem, "|")(4)
                 .TextMatrix(.Rows - 1, mCol.弱阳公式) = Split(strItem, "|")(5)
                 .TextMatrix(.Rows - 1, mCol.最小值) = Split(strItem, "|")(6)
                 .Rows = .Rows + 1
             End With
             
         Next
         If vfgList.Rows > 2 Then vfgList.Rows = vfgList.Rows - 1
         vfgList.Editable = flexEDKbdMouse
     End If
     
End Sub
Private Sub ShowFileList()
    '根据日期显示待导入文件列表
    Dim objStream As TextStream
    Dim strIDX As String, strLine As String, str日期 As String, str时间 As String, str标本号  As String
    Dim str英文 As String, strDataName As String, strItem As String, strImpFile As String
    Dim strDate As String '日期
    Dim lngFileCount As Long, lngLoop As Long, strTmpFile As String
    strIDX = mstrFilePath & "\RMPDAT.IDX"
    
    vfgList.Rows = 2: vfgList.Cols = 12
    vfgList.TextMatrix(0, mCol.选择) = "": vfgList.ColWidth(mCol.选择) = 300: vfgList.ColDataType(mCol.选择) = flexDTBoolean
    vfgList.TextMatrix(0, mCol.项目编号) = "编号": vfgList.ColWidth(mCol.项目编号) = 450
    vfgList.TextMatrix(0, mCol.英文) = "英文": vfgList.ColWidth(mCol.英文) = 1000
    vfgList.TextMatrix(0, mCol.项目名称) = "项目名称": vfgList.ColWidth(mCol.项目名称) = 1500
    vfgList.TextMatrix(0, mCol.起始标本号) = "起始标本号": vfgList.ColWidth(mCol.起始标本号) = 1200: vfgList.ColDataType(mCol.起始标本号) = flexDTLong
    vfgList.TextMatrix(0, mCol.小数) = "小数": vfgList.ColWidth(mCol.小数) = 0
    vfgList.TextMatrix(0, mCol.CutOff公式) = "CutOff公式": vfgList.ColWidth(mCol.CutOff公式) = 0
    vfgList.TextMatrix(0, mCol.阳性公式) = "阳性公式": vfgList.ColWidth(mCol.阳性公式) = 0
    vfgList.TextMatrix(0, mCol.弱阳公式) = "弱阳公式": vfgList.ColWidth(mCol.弱阳公式) = 0
    vfgList.TextMatrix(0, mCol.最小值) = "最小值": vfgList.ColWidth(mCol.最小值) = 0
    vfgList.TextMatrix(0, mCol.检验时间) = "时间": vfgList.ColWidth(mCol.检验时间) = 900
    vfgList.TextMatrix(0, mCol.文件名) = "文件名": vfgList.ColWidth(mCol.文件名) = 0
    
    strDate = Format(dtpDate.Value, "yyyy-MM-dd")
    Set objStream = mobjFile.OpenTextFile(strIDX, ForReading)
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If InStr(strLine, ".DAT=UNDEFINED,") > 0 And InStr(strLine, "RMP") > 0 Then
            str日期 = Split(strLine, ",")(2)
            str日期 = Right(str日期, 2) & "/" & Mid(str日期, 1, 5)
            str时间 = Split(strLine, ",")(3)
            str英文 = Split(strLine, ",")(1)
            
            strDataName = Replace(Split(strLine, ",")(0), "=UNDEFINED", "")
            
            If Format(CDate(str日期), "yyyy-MM-dd") = strDate Then
                With vfgList
                    .TextMatrix(.Rows - 1, mCol.选择) = 0
                    .TextMatrix(.Rows - 1, mCol.英文) = Mid(str英文, InStr(str英文, " ") + 1)
                    strItem = ReadFromIni(App.Path & "\RMP150.ini", Mid(str英文, InStr(str英文, " ") + 1), "Info")
                    .TextMatrix(.Rows - 1, mCol.检验时间) = str时间
                     str标本号 = ReadFromIni(App.Path & "\RMP150.ini", strDate, .TextMatrix(.Rows - 1, mCol.英文))
                    .TextMatrix(.Rows - 1, mCol.起始标本号) = Val(str标本号) + 1
                    .TextMatrix(.Rows - 1, mCol.项目编号) = Split(strItem, "|")(0)
                    .TextMatrix(.Rows - 1, mCol.项目名称) = Split(strItem, "|")(1)
                    .TextMatrix(.Rows - 1, mCol.小数) = Split(strItem, "|")(2)
                    .TextMatrix(.Rows - 1, mCol.CutOff公式) = Split(strItem, "|")(3)
                    .TextMatrix(.Rows - 1, mCol.阳性公式) = Split(strItem, "|")(4)
                    .TextMatrix(.Rows - 1, mCol.弱阳公式) = Split(strItem, "|")(5)
                    .TextMatrix(.Rows - 1, mCol.最小值) = Split(strItem, "|")(6)
                    .TextMatrix(.Rows - 1, mCol.文件名) = strDataName
                    

                    .Rows = .Rows + 1
                End With
                
            End If
        End If
    Loop
    If vfgList.Rows > 2 Then vfgList.Rows = vfgList.Rows - 1
    
End Sub
