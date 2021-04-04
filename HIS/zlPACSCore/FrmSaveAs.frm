VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSaveAs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "另存为"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "FrmSaveAs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox TxtNameNum 
      Height          =   315
      Left            =   1230
      TabIndex        =   14
      Text            =   "1"
      Top             =   1920
      Width           =   2715
   End
   Begin VB.TextBox TxtFile 
      Height          =   315
      Left            =   1230
      TabIndex        =   13
      Text            =   "Image"
      Top             =   1485
      Width           =   2715
   End
   Begin VB.CommandButton cmdAbort 
      Cancel          =   -1  'True
      Caption         =   "停止(&T)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4272
      TabIndex        =   10
      Top             =   720
      Width           =   1100
   End
   Begin MSComctlLib.ProgressBar pgbProcessState 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   3090
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton CmdCencel 
      Caption         =   "退出(&Q)"
      Height          =   350
      Left            =   4272
      TabIndex        =   8
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "保存(&S)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4272
      TabIndex        =   7
      Top             =   210
      Width           =   1100
   End
   Begin VB.ComboBox CmbFormat 
      Height          =   300
      ItemData        =   "FrmSaveAs.frx":000C
      Left            =   1230
      List            =   "FrmSaveAs.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1065
      Width           =   2715
   End
   Begin VB.ComboBox CmbBound 
      Height          =   300
      ItemData        =   "FrmSaveAs.frx":0010
      Left            =   1230
      List            =   "FrmSaveAs.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   645
      Width           =   2715
   End
   Begin VB.CommandButton CmdPath 
      Caption         =   "…"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   245
      Width           =   336
   End
   Begin VB.TextBox TxtPath 
      Height          =   315
      Left            =   1230
      TabIndex        =   0
      Top             =   210
      Width           =   2715
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "开 始 数："
      Height          =   180
      Left            =   270
      TabIndex        =   16
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "文 件 名："
      Height          =   180
      Left            =   270
      TabIndex        =   15
      Top             =   1545
      Width           =   900
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Height          =   180
      Left            =   0
      TabIndex        =   12
      Top             =   2760
      Width           =   5415
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   840
      TabIndex        =   11
      Top             =   2400
      Width           =   3450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "格    式："
      Height          =   180
      Left            =   270
      TabIndex        =   5
      Top             =   1110
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "范    围："
      Height          =   180
      Left            =   270
      TabIndex        =   4
      Top             =   675
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "保存路径："
      Height          =   180
      Left            =   270
      TabIndex        =   2
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "FrmSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim DicImages As New DicomImages                    '用于方便保存时调用
Dim SaveIndex As Integer                            '当前保存到位置
Public f As Form                                    '要操作的窗体
Private dsetDicomDir As DicomDataSet                '保存DICOMDIR的数据集
Private mblnIsAbort As Boolean                      '是否强行终止


Private Sub ExportImageOfZlSeriesInfos(ByVal strExportDir As String, _
    ByVal strImgFormat As String, ByVal strImageName As String, ByVal lngStartIndex As Long)
'//*****************************************************
'//
'//导出ZlSeriesInfos中的图像
'//
'//
'//
'//******************************************************
    Dim i As Integer
    Dim j As Integer
    
    Dim strCurPath As String
    Dim strTransfersyntax As String
    Dim strCurExportDir As String
    Dim objCurDicomImg As DicomImage
    Dim DicAviTmp As New DicomImages
    Dim DicAVISave As New DicomImages
    Dim lngFileIndex As Long
    
    
    On Error GoTo errHandle
    
    '
    strTransfersyntax = "1.2.840.10008.1.2.1"
    strCurExportDir = strExportDir & "\"
    lngFileIndex = lngStartIndex
    
    Me.Label5 = "已处理序列<0>,总序列<" & ZLSeriesInfos.Count & ">。"
        
    For i = 1 To ZLSeriesInfos.Count
        '//一个序列一个AVI文件
        Call DicAVISave.Clear
        Call DicAviTmp.Clear
        
        For j = 1 To ZLSeriesInfos(i).ImageInfos.Count
            Set objCurDicomImg = funLoadAImage(i, j, 0)
            
            If Not objCurDicomImg Is Nothing Then
                '//取得存储路径
                strCurPath = "\" & objCurDicomImg.Name & "(" & Val(objCurDicomImg.PatientID) & ")\" & objCurDicomImg.StudyUID & "\" & objCurDicomImg.SeriesUID & "\"
                
                '//如果目录不存在，则进行创建
                If Dir(strExportDir & strCurPath, vbDirectory) = "" Then
                    Call MkLocalDir(strExportDir & strCurPath)
                End If
                
                Select Case strImgFormat
                    Case "JPG"
                        '//添加导出信息
                        Call subInitImageLabels(i, 0, objCurDicomImg, True, True, True, True)
                        '//Call subSaveLabelToImg(objCurDicomImg)
                        
                                                
                        '//保存图像
                        Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
                    Case "BMP"
                        '//添加导出信息
                        Call subInitImageLabels(i, 0, objCurDicomImg, True, True, True, True)
                        '//Call subSaveLabelToImg(objCurDicomImg)
                        
                                                
                        '//保存图像
                        Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
                    Case "AVI"
                        '//如果是多帧图像，则直接输出到AVI
                        If objCurDicomImg.FrameCount > 1 Then
                            Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
                        Else
                            '//存储AVI的单帧图像
                            Call DicAviTmp.Add(objCurDicomImg)
                        End If
                    Case "DCM"
                        '//如果是DICOM文件，则直接存储
                        Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
                    Case "DICOMDIR"
                        Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & "DCM", "DCM", strTransfersyntax)
                        
                        '//添加DICOMDIR信息
                        Call dsetDicomDir.AddToDirectory(objCurDicomImg, strCurPath & strImageName & lngFileIndex & "." & "DCM", strTransfersyntax, 0)
                End Select
                
                lngFileIndex = lngFileIndex + 1
            End If
                                    
            DoEvents
            
            If mblnIsAbort Then
                mblnIsAbort = False
                Exit Sub
            End If
            
            '//更新处理进度
            Call UpdateProcessState(j, ZLSeriesInfos(i).ImageInfos.Count)
        Next j
        
        If strImgFormat = "AVI" And DicAviTmp.Count > 0 Then
            DicAVISave.Add DicAviTmp.MakeMultiFrame(True)
            Call WriteFile(DicAVISave(1), strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
        End If
        
        Me.Label5 = "已处理序列<" & i & ">,总序列<" & ZLSeriesInfos.Count & ">。"
    Next i
    
    Me.Label5 = "图像保存结束。"
    
    '//保存AVI和DICOMDIR
    '//Select Case strImgFormat
        '//Case "AVI"
            '//DicAVISave.Add DicAviTmp.MakeMultiFrame(True)
            '//Call WriteFile(DicAVISave(1), strCurExportDir & DicAVISave(1).Name & "(" & DicAVISave(1).PatientID & ").avi", strImgFormat)
        '//Case "DICOMDIR"
            '//dsetDicomDir.WriteDirectory strCurExportDir & "DICOMDIR"
    '//End Select
    
    If strImgFormat = "DICOMDIR" Then
        Call dsetDicomDir.WriteDirectory(strCurExportDir & "DICOMDIR")
    End If
        
    Exit Sub
errHandle:
    Me.Label4 = "保存范围<" & Me.CmbBound & ">,保存格式<" & Me.CmbFormat & ">。"
    If ErrCenter = 1 Then Resume
End Sub


Private Sub UpdateProcessState(ByVal lngCurIndex As Long, ByVal lngCount As Long)
'//*****************************************************
'//
'//更新处理进度
'//
'//lngCurIndex:当前处理进度
'//
'//lngCount:
'//
'//******************************************************

    On Error Resume Next
        
        pgbProcessState.Max = lngCount
        pgbProcessState.Value = lngCurIndex
        
End Sub



'根据条件选择需要的图像
Function GetImages(ImagesFiltrate As String) As Integer
'------------------------------------------------
'功能：提取需要另存的图像
'参数： ImagesFiltrate --- 提取图像的方式：当前图像，选中图像
'返回：成功提取图像的数量
'------------------------------------------------
    Dim i, j As Integer             '临时变量
    Dim ImgTmp As DicomImage
    Dim ViewTmp As Variant
    Dim intImageCount As Integer
 
    On Error GoTo GetError
    
    DicImages.Clear                 '清空图像
       
    If ImagesFiltrate = "选中图像" Then
        For i = 1 To ZLShowSeriesInfos.Count
            For j = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
                If ZLShowSeriesInfos(i).ImageInfos(j).blnSelected Then
                    '首先判断图像是否已经装载，如果已经装载，则找到这个图像并显示出来，如果没有装载，则装载该图像
                    If ZLShowSeriesInfos(i).ImageInfos(j).blnDisplayed = False Then
                        Call funcAddAImageA(f.Viewer(i), j)
                    End If
                    DicImages.Add f.Viewer(i).Images(j)
                    '记录图像所在的Viewer的索引，为另存图像时提取图像的数据库信息做准备
                    DicImages(DicImages.Count).Tag = i
                    intImageCount = intImageCount + 1
                End If
            Next j
        Next i
    ElseIf ImagesFiltrate = "当前图像" Then
        If f.intClickImageIndex <> 0 And f.intSelectedSerial <> 0 Then
            DicImages.Add f.Viewer(f.intSelectedSerial).Images(f.intClickImageIndex)
            '记录图像所在的Viewer的索引，为另存图像时提取图像的数据库信息做准备
            DicImages(1).Tag = f.intSelectedSerial
            intImageCount = 1
        End If
    End If
    
'    '读出保存图像
'    With f
'        For Each ViewTmp In .Viewer
'            For Each ImgTmp In ViewTmp.Images
'                Select Case ImagesFiltrate
'                    Case "选中图像"
'                        If ImgTmp.Tag <> "" Then
'                            DicImages.Add ImgTmp
'                            'subLabelCopyRebuild ImgTmp, DicImages(DicImages.Count)
'
'                            GetImages = GetImages + 1
'                        End If
'                    Case "当前图像"
'
'                        If .intClickImageIndex <> 0 And GetImages < 1 Then
'                            DicImages.Add .Viewer(.intSelectedSerial).Images(.intClickImageIndex)
'                            'subLabelCopyRebuild .Viewer(.intSelectedSerial).Images(.intClickImageIndex), DicImages(DicImages.Count)
'
'                            GetImages = GetImages + 1
'                        End If
'                End Select
'            Next
'        Next
'    End With
    Exit Function
GetError:
    '发生错误时不处理
    
End Function


'保存文件
Function SaveImagesAs(FilePath As String, ImagesFormat As String) As Integer
    Dim i, j As Integer
    Dim DicAVIOUT As New DicomImages
    Dim DicImgTmp As New DicomImages
    Dim SaveCmp As Boolean
    Dim transfersyntax As String        'dicomdir的传输语法
    Dim strMiddlePath As String         'dicomdir中使用到的中间路径，结构为：DICOM\PatientName\StudyUID\
    Dim objCurDicomImg As DicomImage
    
    transfersyntax = "1.2.840.10008.1.2.1"
    '清除图像
    DicImgTmp.Clear
    DicAVIOUT.Clear
    '过滤多帧图像
    ImgAVIBount
    '处理AVI
    If ImagesFormat = "AVI" Then
        '处理只有一个多侦多像
        If DicImages.Count = DicImages(1).FrameCount Then
            WriteFile DicImages(1), FilePath & ".avi", "AVI"
            MsgBox "全部保存完成!", vbInformation, gstrSysName
            Unload Me
            Exit Function
        End If
        '处理多个多侦图像
        For i = 1 To DicImages.Count
            '多帧图像
            If DicImages(i).FrameCount > 1 Then
                If DicImages(i).Tag <> "AVIOUT" Then
                    '多帧图像直接输出到AVI
                    SaveCmp = WriteFile(DicImages(i), FilePath & j & ".avi", "AVI")
                End If
            Else
                DicImgTmp.Add DicImages(i)
            End If
        Next
        '制作多帧图像
        If DicImgTmp.Count > 0 Then
            DicAVIOUT.Add DicImgTmp.MakeMultiFrame(False)
            '输出到AVI
            SaveCmp = WriteFile(DicAVIOUT(1), FilePath & j & ".avi", "AVI")
        End If
        '提示信息
        Me.Label4 = Me.CmbBound & "有:" & DicImages.Count & "幅图像。可保存为" & ImgAVIBount & "个AVI. 已保存:" & ImgAVIBount & "个。"
        MsgBox "图像保存结束。", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    SaveIndex = SaveIndex + 1

        '自动命名
    If ImagesFormat = "DICOMDIR" Then       '对于DICOMDIR,单独处理一下文件路径，将其中的文件名去掉
        FilePath = left(FilePath, InStr(FilePath, Me.TxtFile.Text) - 1)
    End If
    
    For i = SaveIndex To DicImages.Count
        Select Case ImagesFormat
            Case "JPG"
                Set objCurDicomImg = DicImages(i)
                Call subInitImageLabels(Val(objCurDicomImg.Tag), 0, objCurDicomImg, True, True, False, True)
                
                WriteFile objCurDicomImg, FilePath & Me.TxtNameNum + i - 1 & "." & ImagesFormat, ImagesFormat
            Case "BMP"
                Set objCurDicomImg = DicImages(i)
                Call subInitImageLabels(Val(objCurDicomImg.Tag), 0, objCurDicomImg, True, True, False, True)
            
                WriteFile objCurDicomImg, FilePath & Me.TxtNameNum + i - 1 & "." & ImagesFormat, ImagesFormat
            Case "DCM"
                subSaveLabelToImg DicImages(i)
                WriteFile DicImages(i), FilePath & Me.TxtNameNum + i - 1 & ".DCM", "DCM"
            Case "DICOMDIR"
                subSaveLabelToImg DicImages(i)
                '如果目录不存在，则创建目录
                strMiddlePath = "DICOM"
                If Dir(FilePath & "\" & strMiddlePath, vbDirectory) = "" Then
                    MkDir (FilePath & "\" & strMiddlePath)
                End If
                strMiddlePath = strMiddlePath & "\" & DicImages(i).Name
                If Dir(FilePath & "\" & strMiddlePath, vbDirectory) = "" Then
                    MkDir (FilePath & "\" & strMiddlePath)
                End If
                strMiddlePath = strMiddlePath & "\" & DicImages(i).StudyUID
                If Dir(FilePath & "\" & strMiddlePath, vbDirectory) = "" Then
                    MkDir (FilePath & "\" & strMiddlePath)
                End If
                WriteFile DicImages(i), FilePath & "\" & strMiddlePath & "\" & Me.TxtFile.Text & Me.TxtNameNum + i - 1 & ".DCM", "DCM", transfersyntax
                dsetDicomDir.AddToDirectory DicImages(i), strMiddlePath & "\" & Me.TxtFile.Text & Me.TxtNameNum + i - 1 & ".DCM", transfersyntax, 0
        End Select
        Me.Label4 = Me.CmbBound & "有:" & DicImages.Count & "幅图像，" & "已保存:" & i & "幅。"
        Me.Refresh
    Next
    If ImagesFormat = "DICOMDIR" Then
        dsetDicomDir.Name = "ZLPACS"
        dsetDicomDir.WriteDirectory FilePath & "DICOMDIR"
    End If
    SaveIndex = 0
    MsgBox "图像保存结束。", vbInformation, gstrSysName
    Unload Me
End Function



'当图像选择范围发生变化时
Private Sub CmbBound_Click()
    Me.Label4 = "保存范围<" & Me.CmbBound & ">,保存格式<" & Me.CmbFormat & ">。"
    
    cmdAbort.Enabled = IIf(CmbBound.Text = "全部图像", True, False)
    
End Sub

Private Sub CmbFormat_Click()
    Me.Label4 = "保存范围<" & Me.CmbBound & ">,保存格式<" & Me.CmbFormat & ">。"
End Sub

Private Sub cmdAbort_Click()
    mblnIsAbort = True
End Sub

Private Sub CmdCencel_Click()
    Unload Me
End Sub

Private Sub CmdPath_Click()
    Dim StrTmp As String
    '得到路径
    StrTmp = BrowPath(Me.hwnd, "请选定保存的文件目录：")
    '当用新的路径时才保存
    If StrTmp <> "" And StrTmp <> Me.TxtPath Then
        Me.TxtPath = StrTmp
    End If
End Sub

Private Sub CmdSave_Click()
    Dim strPath As String
    Dim strTemp As String
    Dim intImageCount As Integer
    
    On Error GoTo errHandle
    
    '保存全部图像的情况，单独处理
    If CmbBound.Text = "全部图像" Then
        Call UpdateProcessState(0, 100)
        Call ExportImageOfZlSeriesInfos(TxtPath.Text, CmbFormat.Text, TxtFile.Text, Val(TxtNameNum.Text))
        
        MsgBox "图像保存完成。", vbInformation, gstrSysName
        
        Exit Sub
    End If
    
    '保存当前图像，选中图像的情况
    intImageCount = GetImages(Me.CmbBound)
    
    Me.Label4 = Me.CmbBound & "有:" & intImageCount & "幅图像。" & "已保存:" & SaveIndex & "幅。"
    '当路径不为根目录时加"\"
    If Len(Me.TxtPath) > 3 Then
        strTemp = "\"
    End If
    
    '当前没有可以保存图像时提示
    If DicImages.Count < 1 Then
         MsgBox "对不起！您没有选择图像！", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    If Len(Dir(Me.TxtPath, vbDirectory)) < 1 Then
        MsgBox "你的路径不正确！", vbExclamation, gstrSysName
        Me.TxtPath.SetFocus
        Exit Sub
    End If
    

    '自动
    If Me.TxtFile.Text = "" And Me.TxtNameNum.Text = "" Then
        MsgBox "请输入文件名和开始序号。", vbInformation, gstrSysName
        Exit Sub
    End If
    strPath = Me.TxtPath & strTemp & Me.TxtFile

    SaveImagesAs strPath, Me.CmbFormat
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub Form_Load()
    mblnIsAbort = False
    Set dsetDicomDir = New DicomDataSet
   '初使化路径
    Me.TxtPath = GetSetting("ZLSOFT", "私有模块\ZLHIS\" & App.EXEName & "\FrmSaveAs", "报告图另存路径", App.Path)
    '测试用
'    Me.TxtPath = "c:\66"
    '范围
    With Me.CmbBound
        .AddItem "当前图像"
        .AddItem "选中图像"
        .AddItem "全部图像"
        .ListIndex = 0
    End With
    With Me.CmbFormat
        .AddItem "JPG"
        .AddItem "BMP"
        .AddItem "DCM"
        .AddItem "AVI"
        .AddItem "DICOMDIR"
        .ListIndex = 0
    End With
    Me.Label4 = "保存范围<" & Me.CmbBound & ">,保存格式<" & Me.CmbFormat & ">。"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '清空图像
    DicImages.Clear
    
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\" & App.EXEName & "\FrmSaveAs", "报告图另存路径", TxtPath.Text
End Sub
'显示保存目录
Public Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        '设置浏览窗口
        .hWndOwner = lWindowHwnd
        '返回选中的目录
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "请选定开始搜索的文件夹："
        Else
            .lpszTitle = sTitle
        End If
    End With
    '调出浏览窗口
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '获取路径
        SHGetPathFromIDList lpIDList, sPath
        '释放内存
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
End Function
'过滤多帧图像
Function ImgAVIBount() As Integer
    Dim i, j As Integer
    For i = 1 To DicImages.Count
        For j = 1 To DicImages.Count
            If DicImages(i).InstanceUID = DicImages(j).InstanceUID And i <> j And DicImages(i).Tag <> "AVIOUT" Then
                If InStr(DicImages(j).Tag, "AVIOUT") < 1 Then
                    DicImages(j).Tag = "AVIOUT"
                End If
            End If
        Next
    Next
    For i = 1 To DicImages.Count
        If DicImages(i).Tag <> "AVIOUT" And DicImages(i).FrameCount > 1 Then
            ImgAVIBount = ImgAVIBount + 1
        End If
    Next
    '只有一个文件时累加
    If DicImages.Count <> DicImages(1).FrameCount Then
        ImgAVIBount = ImgAVIBount + 1
    End If
End Function
'写文件
Function WriteFile(img As DicomImage, Filename As String, WriteType As String, Optional strTransfersyntax As String = "") As Boolean
    Dim lngDialogState As Long
    
    On Error GoTo WriteError
            
    '发现文件存在时提示是否覆盖
    If Dir(Filename) <> "" Then
        lngDialogState = MsgBox("文件" & Filename & "已存在，是否覆盖！", vbQuestion + vbYesNo, App.EXEName)
        
        If lngDialogState = vbNo Then
            WriteFile = False
            Exit Function
        End If
    End If
    '按不同类型保存文件
    Select Case WriteType
        Case "DCM"
            If strTransfersyntax = "" Then
                img.WriteFile Filename, True
            Else
                img.WriteFile Filename, True, strTransfersyntax
            End If
        Case "AVI"
            img.WriteAVI Filename, 1, img.FrameCount, 1, "", 100, False
        Case "JPG"
            img.FileExport Filename, WriteType
        Case "BMP"
            img.FileExport Filename, WriteType
    End Select
    '保存成功
    WriteFile = True
    Exit Function
WriteError:
    '保存出错处理
    If MsgBox("文件" & Filename & "正在被使用，请关闭后选择<是>重试保存，选择<否>跳过这个文件！", vbQuestion + vbYesNo, App.EXEName) = vbYes Then
        Resume
    End If
End Function

Private Sub TxtFile_GotFocus()
    Me.TxtFile.SelStart = 0
    Me.TxtFile.SelLength = Len(Me.TxtFile)
End Sub

Private Sub TxtPath_Click()
'    Me.TxtPath.SelStart = 0
'    Me.TxtPath.SelLength = Len(Me.TxtPath)
End Sub

Private Sub txtPath_GotFocus()
    Me.TxtPath.SelStart = 0
    Me.TxtPath.SelLength = Len(Me.TxtPath)
End Sub

