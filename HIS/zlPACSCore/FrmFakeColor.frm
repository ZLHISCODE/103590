VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Begin VB.Form FrmFakeColor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "伪彩设置"
   ClientHeight    =   6870
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   9870
   DrawStyle       =   1  'Dash
   Icon            =   "FrmFakeColor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "修改伪彩预设方案"
      Height          =   6015
      Left            =   4200
      TabIndex        =   25
      Top             =   120
      Width           =   5655
      Begin VB.ListBox lstColorList 
         Height          =   3120
         Left            =   3600
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "终止"
         Height          =   975
         Left            =   3000
         TabIndex        =   28
         Top             =   4320
         Width           =   2500
         Begin VB.TextBox txtColor 
            Height          =   285
            Index           =   2
            Left            =   960
            TabIndex        =   15
            Top             =   240
            Width           =   1300
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "…"
            Height          =   255
            Index           =   2
            Left            =   2040
            TabIndex        =   29
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "位置："
            Height          =   180
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "颜色："
            Height          =   195
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   540
         End
         Begin VB.Shape shpColor 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   2
            Left            =   960
            Top             =   600
            Width           =   1100
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "起始"
         Height          =   975
         Left            =   120
         TabIndex        =   26
         Top             =   4320
         Width           =   2500
         Begin VB.CommandButton cmdColor 
            Caption         =   "…"
            Height          =   255
            Index           =   1
            Left            =   2040
            TabIndex        =   27
            Top             =   600
            Width           =   255
         End
         Begin VB.TextBox txtColor 
            Height          =   285
            Index           =   1
            Left            =   960
            TabIndex        =   12
            Top             =   240
            Width           =   1300
         End
         Begin VB.Shape shpColor 
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   255
            Index           =   1
            Left            =   960
            Top             =   600
            Width           =   1100
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "颜色："
            Height          =   195
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "位置："
            Height          =   195
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.CommandButton cmdModifyColor 
         Caption         =   "修改"
         Height          =   350
         Left            =   2400
         TabIndex        =   17
         Top             =   5520
         Width           =   1100
      End
      Begin VB.PictureBox picColorRect 
         Height          =   3375
         Left            =   120
         ScaleHeight     =   3315
         ScaleWidth      =   3315
         TabIndex        =   9
         Top             =   480
         Width           =   3375
         Begin MSComDlg.CommonDialog dlgColor 
            Left            =   2160
            Top             =   2640
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "预设伪彩方案："
      Height          =   1335
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   3975
      Begin VB.CommandButton cmdDelPalette 
         Cancel          =   -1  'True
         Caption         =   "删除"
         Height          =   350
         Left            =   2760
         TabIndex        =   3
         Top             =   840
         Width           =   1100
      End
      Begin VB.CommandButton cmdModifyPalette 
         Caption         =   "修改"
         Height          =   350
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddPalette 
         Caption         =   "增加"
         Height          =   350
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   1100
      End
      Begin VB.ComboBox cobColor 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   3672
      End
   End
   Begin VB.PictureBox picColorScale 
      Height          =   3345
      Left            =   3840
      ScaleHeight     =   3285
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   1680
      Width           =   255
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一方案(&D)"
      Height          =   350
      Left            =   3240
      TabIndex        =   21
      Top             =   6360
      Width           =   1245
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一方案(&U)"
      Height          =   350
      Left            =   1320
      TabIndex        =   20
      Top             =   6360
      Width           =   1245
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   3690
      TabIndex        =   23
      Top             =   1680
      Width           =   3756
      Begin DicomObjects.DicomViewer ViewerFackColor 
         Height          =   3240
         Left            =   30
         TabIndex        =   4
         Top             =   45
         Width           =   3630
         _Version        =   262147
         _ExtentX        =   6403
         _ExtentY        =   5715
         _StockProps     =   35
         BackColor       =   -2147483640
      End
   End
   Begin VB.Frame famOpt 
      Caption         =   "应用范围(序列内)"
      Height          =   855
      Left            =   120
      TabIndex        =   22
      Top             =   5280
      Width           =   4005
      Begin VB.OptionButton OptImage 
         Caption         =   "所有图像"
         Height          =   288
         Index           =   2
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   1116
      End
      Begin VB.OptionButton OptImage 
         Caption         =   "所选图像"
         Height          =   276
         Index           =   1
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1128
      End
      Begin VB.OptionButton OptImage 
         Caption         =   "当前图像"
         Height          =   240
         Index           =   0
         Left            =   228
         TabIndex        =   6
         Top             =   330
         Value           =   -1  'True
         Width           =   1164
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7080
      TabIndex        =   19
      Top             =   6360
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5160
      TabIndex        =   18
      Top             =   6360
      Width           =   1100
   End
End
Attribute VB_Name = "FrmFakeColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public f As frmViewer
Dim varColor As Variant
Dim varTmpColor As Variant
Dim intX As Integer
Dim intY As Integer

Private Sub cmdAddPalette_Click()
    Dim strPaletteName As String
    Dim dsData As New Recordset
    Dim strSQL As String
    Dim varNewPalette As String
    Dim intID As Integer
    On Error GoTo errh
    '检查伪彩模版名称是否存在
    strPaletteName = InputBox("请输入新的伪彩模版名称：", "保存伪彩模版")
    If strPaletteName = "" Then Exit Sub
    '检查伪彩模版名称是否重复
    If blLocalRun = True Then
        strSQL = "select 颜色 from 影像颜色清单 "
        Set dsData = cnAccess.Execute(strSQL)
    Else
        strSQL = "select 颜色 from 影像颜色清单 "
        Set dsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    End If
    While Not dsData.EOF
        If dsData!颜色 = strPaletteName Then
            MsgBox "伪彩模版名称重复，请重新输入。", vbInformation, gstrSysName
            Exit Sub
        End If
        dsData.MoveNext
    Wend
    dsData.Close
    '根据lstBox的内容，组织颜色数据块数组
    varNewPalette = funGetPalette()
    If blLocalRun = True Then
        '将新的伪彩模版保存到数据库
        strSQL = "insert into 影像颜色清单 (颜色,颜色内容) values ('" & strPaletteName & "','" & varNewPalette & "')"
        cnAccess.Execute strSQL
        strSQL = "select 序号,颜色内容 from 影像颜色清单 where 颜色 = '" & strPaletteName & "'"
        dsData.Open strSQL, cnAccess, adOpenDynamic, adLockPessimistic
        '将伪彩模版名称添加到调色板下拉列表中
        Me.cobColor.AddItem "用户方案：" & strPaletteName
        Me.cobColor.ItemData(Me.cobColor.NewIndex) = dsData!序号
    Else
        strSQL = "select max(序号) as 最大序号 from 影像颜色清单 "
        Set dsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        intID = dsData("最大序号") + 1
        strSQL = "ZL_影像颜色清单_INSERT('" & strPaletteName & "','" & varNewPalette & "',0)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Me.cobColor.AddItem "用户方案：" & strPaletteName
        Me.cobColor.ItemData(Me.cobColor.NewIndex) = intID
    End If
    Exit Sub
errh:
    If blLocalRun = False Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    Else
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    End If
End Sub

Private Function funGetPalette() As String
'------------------------------------------------
'功能： 从lstColorList中读取当前设置好的调色板数据
'参数： 无
'返回： 调色板颜色串
'------------------------------------------------
    Dim i As Integer
    Dim bytR As Byte
    Dim bytG As Byte
    Dim bytB As Byte
    Dim strColor As String
    Dim strRGB As String
    
    For i = 1 To 256
        strColor = Me.lstColorList.list(i - 1)
        strColor = Right(strColor, Len(strColor) - InStr(strColor, "|") - 2)
        bytR = strColor Mod 256
        bytG = strColor \ 256 Mod 256
        bytB = strColor \ 256 \ 256
        
        funGetPalette = funGetPalette & bytR & "," & bytG & "," & bytB & ";"
    Next i
End Function
Private Sub cmdColor_Click(Index As Integer)
    Me.dlgColor.Color = Me.shpColor(Index).FillColor
    Me.dlgColor.ShowColor
    Me.shpColor(Index).FillColor = Me.dlgColor.Color
End Sub

Private Sub cmdDelPalette_Click()
    Dim strSQL As String
    Dim rsData As Recordset
    If Me.cobColor.ListIndex = -1 Then Exit Sub
    On Error GoTo errh
    If MsgBox("是否确定删除模版：" & Me.cobColor.list(Me.cobColor.ListIndex), vbQuestion + vbOKCancel, gstrSysName) = vbOK Then
        '判断模版是否系统模版，系统模版不允许删除
        If blLocalRun = True Then
            strSQL = "select 系统方案 from 影像颜色清单 where 序号 =" & Me.cobColor.ItemData(Me.cobColor.ListIndex)
            Set rsData = cnAccess.Execute(strSQL)
        Else
            strSQL = "select 系统方案 from 影像颜色清单 where 序号 = [1]"
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Me.cobColor.ItemData(Me.cobColor.ListIndex)))
        End If
        
        If Not rsData.BOF And Not rsData.EOF Then
            If rsData!系统方案 = 1 Then
                MsgBox "当前选中模版为系统模版，不允许删除，只能删除用户自己创建的模版。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '从数据库中删除模版
        If blLocalRun = True Then
            strSQL = "delete from 影像颜色清单 where 序号=" & Me.cobColor.ItemData(Me.cobColor.ListIndex)
            cnAccess.Execute strSQL
        Else
            strSQL = "ZL_影像颜色清单_DELETE(" & Me.cobColor.ItemData(Me.cobColor.ListIndex) & ")"
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
        '删除模版下拉列表中的模版名称
        Me.cobColor.RemoveItem Me.cobColor.ListIndex
        '修正下拉列表的当前值
        Me.cobColor.ListIndex = 0
    End If
    Exit Sub
errh:
    If blLocalRun = False Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    Else
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    End If
End Sub

Private Sub cmdModifyColor_Click()
    If Me.txtColor(1).Text = "" Or Me.txtColor(2).Text = "" Then
        MsgBox "请输入颜色的开始和结束值。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Dim i As Integer
    Dim bytR As Byte
    Dim bytG As Byte
    Dim bytB As Byte
    Dim bytStartR As Integer
    Dim bytStartG As Integer
    Dim bytStartB As Integer
    Dim bytEndR As Integer
    Dim bytEndG As Integer
    Dim bytEndB As Integer
    Dim lngColor As Long
    Dim intFlagLength As Integer        '记录起始和终止调色板颜色编号之间的距离
    
    intFlagLength = Me.txtColor(2).Text - Me.txtColor(1).Text
    If intFlagLength = 0 Then intFlagLength = 1
    bytStartR = Me.shpColor(1).FillColor Mod 256
    bytStartG = Me.shpColor(1).FillColor \ 256 Mod 256
    bytStartB = Me.shpColor(1).FillColor \ 256 \ 256
    bytEndR = Me.shpColor(2).FillColor Mod 256
    bytEndG = Me.shpColor(2).FillColor \ 256 Mod 256
    bytEndB = Me.shpColor(2).FillColor \ 256 \ 256
    For i = Me.txtColor(1).Text To Me.txtColor(2).Text
        bytR = bytStartR + (bytEndR - bytStartR) * ((i - Me.txtColor(1).Text) / intFlagLength)
        bytG = bytStartG + (bytEndG - bytStartG) * (i - Me.txtColor(1).Text) / intFlagLength
        bytB = bytStartB + (bytEndB - bytStartB) * (i - Me.txtColor(1).Text) / intFlagLength
        varTmpColor(i) = bytR & "," & bytG & "," & bytB
    Next i
    picColorRect_Paint
    subFillColorList
    
End Sub

Private Sub cmdModifyPalette_Click()
    If Me.cobColor.ListIndex = -1 Then Exit Sub
    Dim strSQL As String
    Dim rsData As New Recordset
    Dim varNewPalette As String
    On Error GoTo errh
    If blLocalRun = True Then
        '检查当前被选中的调色板是否为系统模版，如果是，则提示不允许修改系统模版
        strSQL = "select 系统方案 from 影像颜色清单 where 序号=" & Me.cobColor.ItemData(Me.cobColor.ListIndex)
        Set rsData = cnAccess.Execute(strSQL)
    Else
        strSQL = "select 系统方案 from 影像颜色清单 where 序号= [1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(Me.cobColor.ItemData(Me.cobColor.ListIndex)))
    End If
    If Not rsData.BOF And Not rsData.EOF Then
        If rsData!系统方案 = 1 Then
            MsgBox "当前选中模版为系统模版，不允许修改，只能修改用户自己创建的模版。", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    rsData.Close
    '修改用户调色板模版
    varNewPalette = funGetPalette
    
    '将新的伪彩模版保存到数据库
    If blLocalRun = True Then
        strSQL = "update 影像颜色清单 set 颜色内容 = '" & varNewPalette & "' where 序号 = " & Me.cobColor.ItemData(Me.cobColor.ListIndex)
        cnAccess.Execute strSQL
    Else
        strSQL = "ZL_影像颜色清单_UPDATE(" & Me.cobColor.ItemData(Me.cobColor.ListIndex) & ",'" & varNewPalette & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    cobColor_Click
    Exit Sub
    
errh:
    If blLocalRun = False Then
        If ErrCenter() = 1 Then
            Resume
        End If
        Call SaveErrLog
    Else
        MsgBox "发生错误:" & err.Description, vbExclamation, gstrSysName
    End If
End Sub

Private Sub cmdNext_Click()
    If cobColor.ListIndex < cobColor.ListCount - 1 Then cobColor.ListIndex = cobColor.ListIndex + 1
End Sub

Private Sub cmdOK_Click()
    Dim iColorNum As Long
    Dim strTemp As String
    Dim tmpF As New Scripting.FileSystemObject
    Dim imgs As New DicomImages
    Dim NewImgs As New DicomImages
    Dim im As DicomImage, imb As New DicomImage
    Dim blnSelected As Boolean      '是否有已选择的图像
    Dim NewImg As DicomImage
    Dim i As Integer
    Dim iVieweIndex As Integer
    Dim iImageIndex As Integer
    
    iVieweIndex = f.intSelectedSerial
    '需要根据选择的情况，对“当前图像”，“所选图像”，“所有图像”做伪彩
    '如果选择了对所选图像进行伪彩处理，检查是否有已经被选择的图像
    If Me.OptImage(1) Then
        For i = 1 To ZLShowSeriesInfos(iVieweIndex).ImageInfos.Count
            If ZLShowSeriesInfos(iVieweIndex).ImageInfos(i).blnSelected = True Then
                blnSelected = True
                Exit For
            End If
        Next i
        
        If blnSelected = False Then
            MsgBox "当前没有选择任何图像，不能作操作。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '在观片站中显示图像
    '先判断图像的选择情况，如果是“所选图像”和“所有图像”需要首先把这些图像加载进来
    If OptImage(1).Value = True Then    '所选图象
        iImageIndex = 1
        For i = 1 To ZLShowSeriesInfos(iVieweIndex).ImageInfos.Count
            If ZLShowSeriesInfos(iVieweIndex).ImageInfos(i).blnSelected = True Then
                '首先判断图像是否已经装载，如果已经装载，则找到这个图像并显示出来，如果没有装载，则装载该图像
                If ZLShowSeriesInfos(iVieweIndex).ImageInfos(i).blnDisplayed = False Then
                    Call funcAddAImageA(f.Viewer(iVieweIndex), i)
                End If
                
                '查找图象的索引
                While f.Viewer(iVieweIndex).Images(iImageIndex).Tag < i And iImageIndex < f.Viewer(iVieweIndex).Images.Count
                    iImageIndex = iImageIndex + 1
                Wend
                
                If iImageIndex <= f.Viewer(iVieweIndex).Images.Count Then
                    If f.Viewer(iVieweIndex).Images(iImageIndex).Tag = i Then
                        Set im = f.Viewer(iVieweIndex).Images(iImageIndex)
                    End If
                End If
                
                If Not im Is Nothing Then
                    imgs.Add im
                End If
            End If
        Next i
    ElseIf OptImage(2).Value = True Then    '所有图像
        '确保所有图象都被加载到Viewer中
        If ZLShowSeriesInfos(iVieweIndex).ImageInfos.Count <> f.Viewer(iVieweIndex).Images.Count Then
            For i = 1 To ZLShowSeriesInfos(iVieweIndex).ImageInfos.Count
                If ZLShowSeriesInfos(iVieweIndex).ImageInfos(i).blnDisplayed = False Then
                    Call funcAddAImageA(f.Viewer(iVieweIndex), i)
                End If
            Next i
        End If
        '把Viewer中的所有图象都加载到imgs集合中
        For i = 1 To f.Viewer(iVieweIndex).Images.Count
            imgs.Add f.Viewer(iVieweIndex).Images(i)
        Next i
    Else    '当前图象
        imgs.Add f.SelectedImage
    End If
    
    '处理图像集合中的所有图象，修改成伪彩图象
    iColorNum = cobColor.ItemData(cobColor.ListIndex)
    Call GetBmpPaletteFromDB(iColorNum)
    For Each im In imgs
        im.Labels.Clear
        strTemp = App.Path & "\temp\" & tmpF.GetTempName
        im.FileExport strTemp, "BMP"
        If ChangeBmpPaletteFromDB(strTemp) = 1 Then MsgBox "该图像已经是彩色图像，不能够进行伪彩操作。", vbInformation, gstrSysName
        Set NewImg = New DicomImage
        NewImg.FileImport strTemp, "BMP"
        NewImgs.Add NewImg
        
        '装载完毕，将所有临时图象同时删除
        Kill strTemp
        DoEvents
        Me.Caption = "伪彩设置。。。。。正在处理伪彩，请稍候"
    Next
    
    '如果横向显示的Viewer个数少于2，则设置横向显示2列
    If f.intCountX < 2 Then
        f.intCountX = 2
        Call subChangeSeriesLayout(f)
    End If
    
    '加载显示NewImgs中的图像
    Call funShowTempImages(f, NewImgs, 0)
    
    Unload Me
End Sub

Private Sub cmdPrevious_Click()
    If cobColor.ListIndex > 0 Then cobColor.ListIndex = cobColor.ListIndex - 1
End Sub
Private Sub cobColor_Click()
    Dim tmpF As New Scripting.FileSystemObject
    Dim imgs As New DicomImages
    Dim iPalNum As Long
    Dim strTemp As String
    Dim DirTemp As SECURITY_ATTRIBUTES              '建立目录时需要的类型
    Dim CreateTrue As Integer                        '建立目录是否成功（非0表示成功）

    imgs.Add f.SelectedImage
    imgs(1).Labels.Clear
    iPalNum = cobColor.ItemData(cobColor.ListIndex)
    If Dir(App.Path & "\temp", vbDirectory) = "" Then
        '建立目录
        CreateTrue = CreateDirectory(App.Path & "\temp", DirTemp)
        If CreateTrue = 0 Then
            '建立目录失败时退出
            MsgBox "建立临时目录" & App.Path & "\temp" & "失败!", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    Call GetBmpPaletteFromDB(iPalNum)
    
    strTemp = App.Path & "\temp\" & tmpF.GetTempName
    imgs(1).FileExport strTemp, "BMP"
    If ChangeBmpPaletteFromDB(strTemp) Then MsgBox "该图像已经是彩色图像，不能够进行伪彩操作。", vbInformation, gstrSysName
    Dim img As New DicomImage
    img.FileImport strTemp, "BMP"
    ViewerFackColor.Images.Clear
    ViewerFackColor.Images.Add img
    ViewerFackColor.Refresh
End Sub

Private Function GetBmpPaletteFromDB(palNum As Long) As Integer
'------------------------------------------------
'功能：根据伪彩颜色编号，从数据库中读取伪彩颜色串，并且重画颜色显示列表
'参数：palNum 伪彩方案中的伪彩颜色编号
'返回：0-修改了ImgFile文件的调色板成功；1－传入图像为彩色图像，不能修改调色板。
'------------------------------------------------
    Dim vTemp As Variant
    Dim strSQL As String
    Dim rsFackColor As Recordset
    Dim strFackColor As String
    
    On Error GoTo 0
    
    If blLocalRun = True Then
        strSQL = "select 颜色内容 from 影像颜色清单 where 序号=" & palNum
        Set rsFackColor = cnAccess.Execute(strSQL)
    Else
        strSQL = "select 颜色内容 from 影像颜色清单 where 序号=[1]"
        Set rsFackColor = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, palNum)
    End If
    strFackColor = rsFackColor.Fields("颜色内容")
    vTemp = Split(strFackColor, ";")
    varColor = vTemp
    varTmpColor = varColor
    
    picColorScale_Paint '刷新伪彩标尺的显示
    picColorRect_Paint  '刷新伪彩颜色矩形的显示
    subFillColorList    '刷新并填充颜色列表的显示

End Function

Private Function ChangeBmpPaletteFromDB(ImgFile As String) As Integer
'------------------------------------------------
'功能：修改8位BMP图像的调色板，从而实现伪彩的功能
'参数：ImgFile 需要实现伪彩的8位灰度bmp图像文件名
'返回：0-修改了ImgFile文件的调色板成功；1－传入图像为彩色图像，不能修改调色板。
'上级函数或过程：FrmFakeColor.cmdOK_Click；FrmFakeColor.cobColor_Click
'下级函数或过程：无
'编制人：黄捷
'说明： 现在使用的伪彩调色板直接跟调色板名字一起保存在数据库ZLPACS.MDB中“颜色清单”表里面，
'       其中共包含51种伪彩调色板。使用WINDOWS GDI API中LOGPALETTE结构的格式存储，后面还有四个0，
'       每一个调色板实际包含的字节数为1032个字节。
'------------------------------------------------
    Dim palR As Byte
    Dim palG As Byte
    Dim palB As Byte
    Dim palFlag As Byte
    Dim i As Integer
    Dim intImageType As Integer
    Dim intRGB As String

    ChangeBmpPaletteFromDB = 0
    On Error Resume Next
    '打开图像文件
    Open ImgFile For Binary As #1
    
    '判断图像是否黑白图像
    Get #1, 29, intImageType
    If intImageType > 8 Then
        ChangeBmpPaletteFromDB = 1
        Close #1
        Exit Function
    End If
    
    On Error Resume Next
    
    For i = 0 To 255
        intRGB = varColor(i)
        '从调色板文件中读取R,G,B和标志位，共四位，组成一种颜色，本调色板总共使用了256种颜色
        palR = strGetRGB(intRGB, 0)
        palG = strGetRGB(intRGB, 1)
        palB = strGetRGB(intRGB, 2)
        palFlag = strGetRGB(intRGB, 3)
        
        '将伪彩调色板写入到图像文件相应位置中，BMP文件前14位是BITMAPFILEHEADER,
        '随后40位是BITMAPINFO，BITMAPINFO =BITMAPINFOHEADER + RGBQUAD
        '从54位后开始，就是图像的调色板，下面就是将图像原有的调色板换成伪彩调色板
        Put #1, 54 + 4 * i + 1, palB
        Put #1, 54 + 4 * i + 2, palG
        Put #1, 54 + 4 * i + 3, palR
        Put #1, 54 + 4 * i + 4, palFlag
    Next
    Close #1
    Exit Function
err:
    Close #1
End Function

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.cobColor.ListIndex = -1
    Me.picColorScale.Line -(10, 10)
    If InStr(mstrPrivs, "观片设置") <> 0 Then
        cmdAddPalette.Enabled = True
        cmdModifyPalette.Enabled = True
        cmdDelPalette.Enabled = True
        Frame2.Enabled = True
    Else
        cmdAddPalette.Enabled = False
        cmdModifyPalette.Enabled = False
        cmdDelPalette.Enabled = False
        Frame2.Enabled = False
    End If
End Sub

Private Sub lstColorList_Click()
    Dim strColor As String
    strColor = Me.lstColorList.list(Me.lstColorList.ListIndex)
    Me.txtColor(1).Text = Val(left(strColor, InStr(strColor, "|") - 3))
    Me.txtColor(2).Text = Me.txtColor(1).Text
    strColor = Right(strColor, Len(strColor) - InStr(strColor, "|") - 2)
    Me.shpColor(1).FillColor = Val(strColor)
    Me.shpColor(2).FillColor = Val(strColor)
End Sub

Private Sub picColorRect_Click()
    Dim intXColor As Integer
    Dim intYColor As Integer
    Dim intColor As Integer
    Dim strColor As String
    intXColor = intX \ (Me.picColorRect.width / 16)
    intYColor = intY \ (Me.picColorRect.height / 16)
    If intYColor >= 16 Then intYColor = 15
    If intXColor >= 16 Then intXColor = 15
    intColor = intYColor * 16 + intXColor
    
    Me.txtColor(1).Text = intColor
    Me.txtColor(2).Text = intColor
    
    strColor = Me.lstColorList.list(intColor)
    strColor = Right(strColor, Len(strColor) - InStr(strColor, "|") - 2)
    Me.shpColor(1).FillColor = Val(strColor)
    Me.shpColor(2).FillColor = Val(strColor)
End Sub

Private Sub picColorRect_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    intX = x
    intY = y
End Sub

Private Sub picColorRect_Paint()
    Dim i As Integer
    Dim j As Integer
    Dim intColor As Integer
    Dim intOldScaleWidth As Integer
    Dim intOldScaleHeight As Integer
    Dim intOldx As Integer
    Dim intOldy As Integer
    Dim intRGB As String
    
    '保存原有的设置
    intOldScaleWidth = Me.picColorRect.ScaleWidth
    intOldScaleHeight = Me.picColorRect.ScaleHeight
    intOldx = Me.picColorRect.CurrentX
    intOldy = Me.picColorRect.CurrentY
    '将画框设置成256行，10列
    Me.picColorRect.Scale (0, 0)-(160, 160)
    
    '设置伪彩标尺的单位高度
    Me.picColorRect.DrawWidth = 1
    Me.picColorRect.FillStyle = 0
    Me.picColorRect.ForeColor = vbBlack
    intColor = 0
    
    For j = 0 To 15
        For i = 0 To 15
            intRGB = varTmpColor(intColor)
            Me.picColorRect.FillColor = RGB(strGetRGB(intRGB, 0), strGetRGB(intRGB, 1), strGetRGB(intRGB, 2))
            Me.picColorRect.Line (i * 10 + 1, j * 10 + 1)-(i * 10 + 9, j * 10 + 9), , B
            intColor = intColor + 1
        Next i
    Next j
    
    '将画框还原
    Me.picColorRect.Scale (0, 0)-(intOldScaleWidth, intOldScaleHeight)
    Me.picColorRect.CurrentX = intOldx
    Me.picColorRect.CurrentY = intOldy
End Sub

Private Sub picColorScale_Paint()
    Dim i As Integer
    Dim intOldScaleWidth As Integer
    Dim intOldScaleHeight As Integer
    Dim intOldx As Integer
    Dim intOldy As Integer
    Dim intRGB As String
    
    '保存原有的设置
    intOldScaleWidth = Me.picColorScale.ScaleWidth
    intOldScaleHeight = Me.picColorScale.ScaleHeight
    intOldx = Me.picColorScale.CurrentX
    intOldy = Me.picColorScale.CurrentY
    '将画框设置成256行，10列
    Me.picColorScale.Scale (0, 0)-(9, 255)
    
    '设置伪彩标尺的单位高度
    Me.picColorScale.DrawWidth = 1
    
    For i = 0 To 255
        intRGB = varColor(i)
        Me.picColorScale.ForeColor = RGB(strGetRGB(intRGB, 0), strGetRGB(intRGB, 1), strGetRGB(intRGB, 2))
        Me.picColorScale.Line (0, i)-(10, i)
    Next i
    
    '将画框还原
    Me.picColorScale.Scale (0, 0)-(intOldScaleWidth, intOldScaleHeight)
    Me.picColorScale.CurrentX = intOldx
    Me.picColorScale.CurrentY = intOldy
End Sub

Private Sub subFillColorList()
    Dim i As Integer
    Dim intRGB As String
    Me.lstColorList.Clear  '清空listBox的原有内容
    For i = 0 To 255
        intRGB = varTmpColor(i)
        Me.lstColorList.AddItem Val(i) & "  |  " & RGB(strGetRGB(intRGB, 0), strGetRGB(intRGB, 1), strGetRGB(intRGB, 2))
    Next i
End Sub

Private Sub txtColor_GotFocus(Index As Integer)
    Me.txtColor(Index).SelStart = 0
    Me.txtColor(Index).SelLength = Len(Me.txtColor(Index).Text)
End Sub

Private Sub txtColor_LostFocus(Index As Integer)
    If Val(Me.txtColor(Index).Text) = 0 Then    '只允许输入数字
        Me.txtColor(Index).Text = 0
    End If
    '确保数值在0-255之间
    If Val(Me.txtColor(Index).Text) < 0 Then Me.txtColor(Index).Text = 0
    If Val(Me.txtColor(Index).Text) > 255 Then Me.txtColor(Index).Text = 255
    '确保两个txt之间数值的大小关系
    If Index = 1 Then
        If Val(Me.txtColor(1).Text) > Val(Me.txtColor(2).Text) Then Me.txtColor(1).Text = Me.txtColor(2).Text
    Else
        If Val(Me.txtColor(2).Text) < Val(Me.txtColor(1).Text) Then Me.txtColor(2).Text = Me.txtColor(1).Text
    End If
End Sub
Private Function strGetRGB(strRGB As String, intRGB As Integer) As Integer
    '从字串中得到RGB颜色
    Dim StrTmp As Variant
    StrTmp = Split(strRGB, ",")
    strGetRGB = StrTmp(intRGB)
End Function

