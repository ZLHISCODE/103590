VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmSaveAs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���Ϊ"
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
   StartUpPosition =   1  '����������
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
      Caption         =   "ֹͣ(&T)"
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
      Caption         =   "�˳�(&Q)"
      Height          =   350
      Left            =   4272
      TabIndex        =   8
      Top             =   1680
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "����(&S)"
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
      Caption         =   "��"
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
      Caption         =   "�� ʼ ����"
      Height          =   180
      Left            =   270
      TabIndex        =   16
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "�� �� ����"
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
      Caption         =   "��    ʽ��"
      Height          =   180
      Left            =   270
      TabIndex        =   5
      Top             =   1110
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "��    Χ��"
      Height          =   180
      Left            =   270
      TabIndex        =   4
      Top             =   675
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "����·����"
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


Dim DicImages As New DicomImages                    '���ڷ��㱣��ʱ����
Dim SaveIndex As Integer                            '��ǰ���浽λ��
Public f As Form                                    'Ҫ�����Ĵ���
Private dsetDicomDir As DicomDataSet                '����DICOMDIR�����ݼ�
Private mblnIsAbort As Boolean                      '�Ƿ�ǿ����ֹ


Private Sub ExportImageOfZlSeriesInfos(ByVal strExportDir As String, _
    ByVal strImgFormat As String, ByVal strImageName As String, ByVal lngStartIndex As Long)
'//*****************************************************
'//
'//����ZlSeriesInfos�е�ͼ��
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
    
    Me.Label5 = "�Ѵ�������<0>,������<" & ZLSeriesInfos.Count & ">��"
        
    For i = 1 To ZLSeriesInfos.Count
        '//һ������һ��AVI�ļ�
        Call DicAVISave.Clear
        Call DicAviTmp.Clear
        
        For j = 1 To ZLSeriesInfos(i).ImageInfos.Count
            Set objCurDicomImg = funLoadAImage(i, j, 0)
            
            If Not objCurDicomImg Is Nothing Then
                '//ȡ�ô洢·��
                strCurPath = "\" & objCurDicomImg.Name & "(" & Val(objCurDicomImg.PatientID) & ")\" & objCurDicomImg.StudyUID & "\" & objCurDicomImg.SeriesUID & "\"
                
                '//���Ŀ¼�����ڣ�����д���
                If Dir(strExportDir & strCurPath, vbDirectory) = "" Then
                    Call MkLocalDir(strExportDir & strCurPath)
                End If
                
                Select Case strImgFormat
                    Case "JPG"
                        '//��ӵ�����Ϣ
                        Call subInitImageLabels(i, 0, objCurDicomImg, True, True, True, True)
                        '//Call subSaveLabelToImg(objCurDicomImg)
                        
                                                
                        '//����ͼ��
                        Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
                    Case "BMP"
                        '//��ӵ�����Ϣ
                        Call subInitImageLabels(i, 0, objCurDicomImg, True, True, True, True)
                        '//Call subSaveLabelToImg(objCurDicomImg)
                        
                                                
                        '//����ͼ��
                        Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
                    Case "AVI"
                        '//����Ƕ�֡ͼ����ֱ�������AVI
                        If objCurDicomImg.FrameCount > 1 Then
                            Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
                        Else
                            '//�洢AVI�ĵ�֡ͼ��
                            Call DicAviTmp.Add(objCurDicomImg)
                        End If
                    Case "DCM"
                        '//�����DICOM�ļ�����ֱ�Ӵ洢
                        Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
                    Case "DICOMDIR"
                        Call WriteFile(objCurDicomImg, strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & "DCM", "DCM", strTransfersyntax)
                        
                        '//���DICOMDIR��Ϣ
                        Call dsetDicomDir.AddToDirectory(objCurDicomImg, strCurPath & strImageName & lngFileIndex & "." & "DCM", strTransfersyntax, 0)
                End Select
                
                lngFileIndex = lngFileIndex + 1
            End If
                                    
            DoEvents
            
            If mblnIsAbort Then
                mblnIsAbort = False
                Exit Sub
            End If
            
            '//���´������
            Call UpdateProcessState(j, ZLSeriesInfos(i).ImageInfos.Count)
        Next j
        
        If strImgFormat = "AVI" And DicAviTmp.Count > 0 Then
            DicAVISave.Add DicAviTmp.MakeMultiFrame(True)
            Call WriteFile(DicAVISave(1), strCurExportDir & strCurPath & strImageName & lngFileIndex & "." & strImgFormat, strImgFormat)
        End If
        
        Me.Label5 = "�Ѵ�������<" & i & ">,������<" & ZLSeriesInfos.Count & ">��"
    Next i
    
    Me.Label5 = "ͼ�񱣴������"
    
    '//����AVI��DICOMDIR
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
    Me.Label4 = "���淶Χ<" & Me.CmbBound & ">,�����ʽ<" & Me.CmbFormat & ">��"
    If ErrCenter = 1 Then Resume
End Sub


Private Sub UpdateProcessState(ByVal lngCurIndex As Long, ByVal lngCount As Long)
'//*****************************************************
'//
'//���´������
'//
'//lngCurIndex:��ǰ�������
'//
'//lngCount:
'//
'//******************************************************

    On Error Resume Next
        
        pgbProcessState.Max = lngCount
        pgbProcessState.Value = lngCurIndex
        
End Sub



'��������ѡ����Ҫ��ͼ��
Function GetImages(ImagesFiltrate As String) As Integer
'------------------------------------------------
'���ܣ���ȡ��Ҫ����ͼ��
'������ ImagesFiltrate --- ��ȡͼ��ķ�ʽ����ǰͼ��ѡ��ͼ��
'���أ��ɹ���ȡͼ�������
'------------------------------------------------
    Dim i, j As Integer             '��ʱ����
    Dim ImgTmp As DicomImage
    Dim ViewTmp As Variant
    Dim intImageCount As Integer
 
    On Error GoTo GetError
    
    DicImages.Clear                 '���ͼ��
       
    If ImagesFiltrate = "ѡ��ͼ��" Then
        For i = 1 To ZLShowSeriesInfos.Count
            For j = 1 To ZLShowSeriesInfos(i).ImageInfos.Count
                If ZLShowSeriesInfos(i).ImageInfos(j).blnSelected Then
                    '�����ж�ͼ���Ƿ��Ѿ�װ�أ�����Ѿ�װ�أ����ҵ����ͼ����ʾ���������û��װ�أ���װ�ظ�ͼ��
                    If ZLShowSeriesInfos(i).ImageInfos(j).blnDisplayed = False Then
                        Call funcAddAImageA(f.Viewer(i), j)
                    End If
                    DicImages.Add f.Viewer(i).Images(j)
                    '��¼ͼ�����ڵ�Viewer��������Ϊ���ͼ��ʱ��ȡͼ������ݿ���Ϣ��׼��
                    DicImages(DicImages.Count).Tag = i
                    intImageCount = intImageCount + 1
                End If
            Next j
        Next i
    ElseIf ImagesFiltrate = "��ǰͼ��" Then
        If f.intClickImageIndex <> 0 And f.intSelectedSerial <> 0 Then
            DicImages.Add f.Viewer(f.intSelectedSerial).Images(f.intClickImageIndex)
            '��¼ͼ�����ڵ�Viewer��������Ϊ���ͼ��ʱ��ȡͼ������ݿ���Ϣ��׼��
            DicImages(1).Tag = f.intSelectedSerial
            intImageCount = 1
        End If
    End If
    
'    '��������ͼ��
'    With f
'        For Each ViewTmp In .Viewer
'            For Each ImgTmp In ViewTmp.Images
'                Select Case ImagesFiltrate
'                    Case "ѡ��ͼ��"
'                        If ImgTmp.Tag <> "" Then
'                            DicImages.Add ImgTmp
'                            'subLabelCopyRebuild ImgTmp, DicImages(DicImages.Count)
'
'                            GetImages = GetImages + 1
'                        End If
'                    Case "��ǰͼ��"
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
    '��������ʱ������
    
End Function


'�����ļ�
Function SaveImagesAs(FilePath As String, ImagesFormat As String) As Integer
    Dim i, j As Integer
    Dim DicAVIOUT As New DicomImages
    Dim DicImgTmp As New DicomImages
    Dim SaveCmp As Boolean
    Dim transfersyntax As String        'dicomdir�Ĵ����﷨
    Dim strMiddlePath As String         'dicomdir��ʹ�õ����м�·�����ṹΪ��DICOM\PatientName\StudyUID\
    Dim objCurDicomImg As DicomImage
    
    transfersyntax = "1.2.840.10008.1.2.1"
    '���ͼ��
    DicImgTmp.Clear
    DicAVIOUT.Clear
    '���˶�֡ͼ��
    ImgAVIBount
    '����AVI
    If ImagesFormat = "AVI" Then
        '����ֻ��һ���������
        If DicImages.Count = DicImages(1).FrameCount Then
            WriteFile DicImages(1), FilePath & ".avi", "AVI"
            MsgBox "ȫ���������!", vbInformation, gstrSysName
            Unload Me
            Exit Function
        End If
        '����������ͼ��
        For i = 1 To DicImages.Count
            '��֡ͼ��
            If DicImages(i).FrameCount > 1 Then
                If DicImages(i).Tag <> "AVIOUT" Then
                    '��֡ͼ��ֱ�������AVI
                    SaveCmp = WriteFile(DicImages(i), FilePath & j & ".avi", "AVI")
                End If
            Else
                DicImgTmp.Add DicImages(i)
            End If
        Next
        '������֡ͼ��
        If DicImgTmp.Count > 0 Then
            DicAVIOUT.Add DicImgTmp.MakeMultiFrame(False)
            '�����AVI
            SaveCmp = WriteFile(DicAVIOUT(1), FilePath & j & ".avi", "AVI")
        End If
        '��ʾ��Ϣ
        Me.Label4 = Me.CmbBound & "��:" & DicImages.Count & "��ͼ�񡣿ɱ���Ϊ" & ImgAVIBount & "��AVI. �ѱ���:" & ImgAVIBount & "����"
        MsgBox "ͼ�񱣴������", vbInformation, gstrSysName
        Unload Me
        Exit Function
    End If
    
    SaveIndex = SaveIndex + 1

        '�Զ�����
    If ImagesFormat = "DICOMDIR" Then       '����DICOMDIR,��������һ���ļ�·���������е��ļ���ȥ��
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
                '���Ŀ¼�����ڣ��򴴽�Ŀ¼
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
        Me.Label4 = Me.CmbBound & "��:" & DicImages.Count & "��ͼ��" & "�ѱ���:" & i & "����"
        Me.Refresh
    Next
    If ImagesFormat = "DICOMDIR" Then
        dsetDicomDir.Name = "ZLPACS"
        dsetDicomDir.WriteDirectory FilePath & "DICOMDIR"
    End If
    SaveIndex = 0
    MsgBox "ͼ�񱣴������", vbInformation, gstrSysName
    Unload Me
End Function



'��ͼ��ѡ��Χ�����仯ʱ
Private Sub CmbBound_Click()
    Me.Label4 = "���淶Χ<" & Me.CmbBound & ">,�����ʽ<" & Me.CmbFormat & ">��"
    
    cmdAbort.Enabled = IIf(CmbBound.Text = "ȫ��ͼ��", True, False)
    
End Sub

Private Sub CmbFormat_Click()
    Me.Label4 = "���淶Χ<" & Me.CmbBound & ">,�����ʽ<" & Me.CmbFormat & ">��"
End Sub

Private Sub cmdAbort_Click()
    mblnIsAbort = True
End Sub

Private Sub CmdCencel_Click()
    Unload Me
End Sub

Private Sub CmdPath_Click()
    Dim StrTmp As String
    '�õ�·��
    StrTmp = BrowPath(Me.hwnd, "��ѡ��������ļ�Ŀ¼��")
    '�����µ�·��ʱ�ű���
    If StrTmp <> "" And StrTmp <> Me.TxtPath Then
        Me.TxtPath = StrTmp
    End If
End Sub

Private Sub CmdSave_Click()
    Dim strPath As String
    Dim strTemp As String
    Dim intImageCount As Integer
    
    On Error GoTo errHandle
    
    '����ȫ��ͼ����������������
    If CmbBound.Text = "ȫ��ͼ��" Then
        Call UpdateProcessState(0, 100)
        Call ExportImageOfZlSeriesInfos(TxtPath.Text, CmbFormat.Text, TxtFile.Text, Val(TxtNameNum.Text))
        
        MsgBox "ͼ�񱣴���ɡ�", vbInformation, gstrSysName
        
        Exit Sub
    End If
    
    '���浱ǰͼ��ѡ��ͼ������
    intImageCount = GetImages(Me.CmbBound)
    
    Me.Label4 = Me.CmbBound & "��:" & intImageCount & "��ͼ��" & "�ѱ���:" & SaveIndex & "����"
    '��·����Ϊ��Ŀ¼ʱ��"\"
    If Len(Me.TxtPath) > 3 Then
        strTemp = "\"
    End If
    
    '��ǰû�п��Ա���ͼ��ʱ��ʾ
    If DicImages.Count < 1 Then
         MsgBox "�Բ�����û��ѡ��ͼ��", vbExclamation, gstrSysName
         Exit Sub
    End If
    
    If Len(Dir(Me.TxtPath, vbDirectory)) < 1 Then
        MsgBox "���·������ȷ��", vbExclamation, gstrSysName
        Me.TxtPath.SetFocus
        Exit Sub
    End If
    

    '�Զ�
    If Me.TxtFile.Text = "" And Me.TxtNameNum.Text = "" Then
        MsgBox "�������ļ����Ϳ�ʼ��š�", vbInformation, gstrSysName
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
   '��ʹ��·��
    Me.TxtPath = GetSetting("ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\FrmSaveAs", "����ͼ���·��", App.Path)
    '������
'    Me.TxtPath = "c:\66"
    '��Χ
    With Me.CmbBound
        .AddItem "��ǰͼ��"
        .AddItem "ѡ��ͼ��"
        .AddItem "ȫ��ͼ��"
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
    Me.Label4 = "���淶Χ<" & Me.CmbBound & ">,�����ʽ<" & Me.CmbFormat & ">��"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    '���ͼ��
    DicImages.Clear
    
    SaveSetting "ZLSOFT", "˽��ģ��\ZLHIS\" & App.EXEName & "\FrmSaveAs", "����ͼ���·��", TxtPath.Text
End Sub
'��ʾ����Ŀ¼
Public Function BrowPath(lWindowHwnd As Long, Optional ByVal sTitle As String = "") As String
    Dim iNull As Integer, lpIDList As Long
    Dim sPath As String, udtBI As BrowseInfo
    With udtBI
        '�����������
        .hWndOwner = lWindowHwnd
        '����ѡ�е�Ŀ¼
        .ulFlags = BIF_RETURNONLYFSDIRS
        If sTitle = "" Then
            .lpszTitle = "��ѡ����ʼ�������ļ��У�"
        Else
            .lpszTitle = sTitle
        End If
    End With
    '�����������
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '��ȡ·��
        SHGetPathFromIDList lpIDList, sPath
        '�ͷ��ڴ�
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = left$(sPath, iNull - 1)
        End If
    End If
    BrowPath = sPath
End Function
'���˶�֡ͼ��
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
    'ֻ��һ���ļ�ʱ�ۼ�
    If DicImages.Count <> DicImages(1).FrameCount Then
        ImgAVIBount = ImgAVIBount + 1
    End If
End Function
'д�ļ�
Function WriteFile(img As DicomImage, Filename As String, WriteType As String, Optional strTransfersyntax As String = "") As Boolean
    Dim lngDialogState As Long
    
    On Error GoTo WriteError
            
    '�����ļ�����ʱ��ʾ�Ƿ񸲸�
    If Dir(Filename) <> "" Then
        lngDialogState = MsgBox("�ļ�" & Filename & "�Ѵ��ڣ��Ƿ񸲸ǣ�", vbQuestion + vbYesNo, App.EXEName)
        
        If lngDialogState = vbNo Then
            WriteFile = False
            Exit Function
        End If
    End If
    '����ͬ���ͱ����ļ�
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
    '����ɹ�
    WriteFile = True
    Exit Function
WriteError:
    '���������
    If MsgBox("�ļ�" & Filename & "���ڱ�ʹ�ã���رպ�ѡ��<��>���Ա��棬ѡ��<��>��������ļ���", vbQuestion + vbYesNo, App.EXEName) = vbYes Then
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

