VERSION 5.00
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmPubSampleImgs 
   BorderStyle     =   0  'None
   Caption         =   "����ͼƬ"
   ClientHeight    =   3015
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox PicPic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   0
      ScaleHeight     =   2205
      ScaleWidth      =   1035
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1035
      Begin C1Chart2D8.Chart2D chtPic 
         Height          =   705
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   150
         Width           =   615
         _Version        =   524288
         _Revision       =   7
         _ExtentX        =   1085
         _ExtentY        =   1244
         _StockProps     =   0
         ControlProperties=   "frmPubSampleImgs.frx":0000
      End
   End
End
Attribute VB_Name = "frmPubSampleImgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-22
'ģ�鹦��:չʾ�걾ͼ��
'---------------------------------------------------------------------------------------

Option Explicit

Private mObjImg As Object       'zlLisDev.clsDrawGraph

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-22
'��    ��:  ��ȡ�걾ͼ��
'��    ��:
'           cnOracle            ���Ӷ���
'           lngSampleID         �걾ID
'           intVersion          �汾25=�°�LIS��10=�ϰ�LIS
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Sub ReadImages(ByVal lngSampleID As Long, Optional ByVal intVersion As Integer = 25)

    '����               ���뵱ǰ�걾��ͼ�ε�Cht
    Dim strChart(0 To 8) As String
    Dim strErr As String
    Dim intloop As Integer


    '���Ű�
    Call ImageTypeSet(9, True)
    '����ͼ������
    If ReadSampleImage(lngSampleID, strChart, strErr, intVersion) = False Then
        MsgBox strErr, vbInformation, gSysInfo.AppName
    End If
    For intloop = 0 To 8
        If strChart(intloop) <> "" Then
            chtPic(intloop).Load (strChart(intloop))
        End If
    Next
    '����������Ű�
    Call ImageTypeSet(9)


End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With PicPic
        .Left = 0
        .Top = 0
        .Width = Me.Width
        .Height = Me.Height
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mObjImg = Nothing
End Sub

Private Sub PicPic_Resize()
    Call ImageTypeSet(9)
End Sub

Private Sub ImageTypeSet(intCount As Integer, Optional blnNoRead As Boolean = False)
          '����           ͼ���Ű�(���9��ͼ)
          Dim intloop As Integer
          '����������ͼ�����а�����

1         On Error GoTo ImageTypeSet_Error

2         For intloop = 0 To 8
3             If chtPic.Count - 1 < intloop Then
4                 Load chtPic(intloop)
5             End If
      '        chtPic(intLoop).Visible = False
6             If blnNoRead = True Then
7                 chtPic(intloop).Reset
8                 chtPic(intloop).ChartGroups(1).Data.NumPoints(1) = 0
9             End If
10            chtPic(intloop).Interior.Image.Layout = oc2dImageStretched
11            chtPic(intloop).Border.Type = oc2dBorderPlain
12            chtPic(intloop).Border.Width = 1
13            chtPic(intloop).IsBatched = False
14            chtPic(intloop).AllowUserChanges = False
15        Next

16        If intCount <= 4 Then
              '��4��ͼ��������
17            chtPic(0).Left = 25
18            chtPic(0).Top = 25
19            chtPic(0).Width = (Me.PicPic.ScaleWidth - 50) / 2
20            chtPic(0).Height = (Me.PicPic.ScaleHeight - 50) / 2

21            chtPic(1).Left = chtPic(0).Left + chtPic(0).Width + 25
22            chtPic(1).Top = 25
23            chtPic(1).Width = Me.PicPic.ScaleWidth - chtPic(1).Left - 25
24            chtPic(1).Height = chtPic(0).Height

25            chtPic(2).Left = 25
26            chtPic(2).Top = chtPic(0).Top + chtPic(0).Height + 25
27            chtPic(2).Height = chtPic(0).Height
28            chtPic(2).Width = chtPic(0).Width

29            chtPic(3).Left = chtPic(1).Left
30            chtPic(3).Top = chtPic(2).Top
31            chtPic(3).Height = chtPic(2).Height
32            chtPic(3).Width = Me.PicPic.ScaleWidth - chtPic(3).Left - 25
33        ElseIf intCount <= 6 Then
34            chtPic(0).Left = 25
35            chtPic(0).Top = 25
36            chtPic(0).Width = (Me.PicPic.ScaleWidth - 100) / 3
37            chtPic(0).Height = chtPic(0).Width

38            chtPic(1).Left = chtPic(0).Left + chtPic(0).Width + 25
39            chtPic(1).Top = 25
40            chtPic(1).Width = chtPic(0).Width
41            chtPic(1).Height = chtPic(0).Height

42            chtPic(2).Left = chtPic(1).Left + chtPic(1).Width + 25
43            chtPic(2).Top = 25
44            chtPic(2).Width = Me.PicPic.ScaleWidth - chtPic(2).Left
45            chtPic(2).Height = chtPic(0).Height

46            chtPic(3).Left = 25
47            chtPic(3).Top = chtPic(0).Top + chtPic(0).Height + 25
48            chtPic(3).Width = chtPic(0).Width
49            chtPic(3).Height = Me.PicPic.ScaleHeight - chtPic(3).Left

50            chtPic(4).Left = chtPic(3).Left + chtPic(3).Width + 25
51            chtPic(4).Top = chtPic(3).Top
52            chtPic(4).Width = chtPic(3).Width
53            chtPic(4).Height = chtPic(3).Height

54            chtPic(5).Left = chtPic(4).Left + chtPic(4).Width + 25
55            chtPic(5).Top = chtPic(3).Top
56            chtPic(5).Width = chtPic(3).Width
57            chtPic(5).Height = chtPic(3).Height
58        ElseIf intCount <= 9 Then
59            chtPic(0).Left = 25
60            chtPic(0).Top = 25
61            chtPic(0).Width = (Me.PicPic.ScaleWidth - 100) / 3
62            chtPic(0).Height = (Me.PicPic.ScaleHeight - 100) / 3

63            chtPic(1).Left = chtPic(0).Left + chtPic(0).Width + 25
64            chtPic(1).Top = 25
65            chtPic(1).Width = chtPic(0).Width
66            chtPic(1).Height = chtPic(0).Height

67            chtPic(2).Left = chtPic(1).Left + chtPic(1).Width + 25
68            chtPic(2).Top = 25
69            chtPic(2).Width = Me.PicPic.ScaleWidth - chtPic(2).Left
70            chtPic(2).Height = chtPic(0).Height

71            chtPic(3).Left = 25
72            chtPic(3).Top = chtPic(0).Top + chtPic(0).Height + 25
73            chtPic(3).Width = chtPic(0).Width
74            chtPic(3).Height = chtPic(0).Height

75            chtPic(4).Left = chtPic(3).Left + chtPic(3).Width + 25
76            chtPic(4).Top = chtPic(0).Top + chtPic(0).Height + 25
77            chtPic(4).Width = chtPic(3).Width
78            chtPic(4).Height = chtPic(3).Height

79            chtPic(5).Left = chtPic(4).Left + chtPic(4).Width + 25
80            chtPic(5).Top = chtPic(4).Top
81            chtPic(5).Width = PicPic.ScaleWidth - chtPic(5).Left
82            chtPic(5).Height = chtPic(3).Height

83            chtPic(6).Left = 25
84            chtPic(6).Top = chtPic(3).Top + chtPic(3).Height + 25
85            chtPic(6).Width = chtPic(0).Width
86            chtPic(6).Height = PicPic.ScaleHeight - chtPic(6).Top

87            chtPic(7).Left = chtPic(6).Left + chtPic(6).Width + 25
88            chtPic(7).Top = chtPic(6).Top
89            chtPic(7).Width = chtPic(6).Width
90            chtPic(7).Height = chtPic(6).Height

91            chtPic(8).Left = chtPic(7).Left + chtPic(7).Width + 25
92            chtPic(8).Top = chtPic(6).Top
93            chtPic(8).Width = Me.PicPic.ScaleWidth - chtPic(8).Left
94            chtPic(8).Height = chtPic(6).Height
95        End If

96        For intloop = 0 To 8
97            chtPic(intloop).Visible = True
98        Next



99        Exit Sub
ImageTypeSet_Error:
100       Call WriteErrLog("zl9LisInsideComm", "frmPubSampleImgs", "ִ��(ImageTypeSet)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
101       Err.Clear

End Sub

Public Function ReadSampleImage(lngSampleID As Long, strChar() As String, Optional strErr As String, Optional intVal As Integer = 25) As Boolean
          '����   ����걾��ͼ�񷵻ض���������
          '��ͼ��
          Dim strReturn As String
          Dim varTmp As Variant, strDir As String
          Dim i As Integer
          Dim objFile As File, strDelFile As String
          Dim objTxtFile As TextStream, strLine As String
          Dim blnDelOldImg As Boolean
          Dim objFSO As New FileSystemObject
       
1         On Error GoTo ReadSampleImage_Error
2         gobjLiscomlib.ShowFlash "���ڼ���ͼ��,������ҪһЩʱ��,���Ժ�...", Me
3         strErr = ""
          '����ɾ������ͼ���ļ���־
4         strDelFile = App.Path & "\DelImgFlag.log"
5         If objFSO.FileExists(strDelFile) Then
6             Set objTxtFile = objFSO.OpenTextFile(strDelFile, ForReading)
7             Do While Not objTxtFile.AtEndOfLine
8                 strLine = objTxtFile.ReadLine
9             Loop
10            If InStr(strLine, CStr(Format(Now, "yyyy-MM-dd"))) > 0 Then
11                blnDelOldImg = True
12            Else
13                objTxtFile.Close
14                Set objTxtFile = Nothing
                  
15                Set objTxtFile = objFSO.CreateTextFile(strDelFile, True)
16                objTxtFile.WriteLine Format(Now, "yyyy-MM-dd")
17            End If

18            objTxtFile.Close
19            Set objTxtFile = Nothing
          
20        Else
21            Set objTxtFile = objFSO.CreateTextFile(strDelFile, True)
22            objTxtFile.WriteLine Format(Now, "yyyy-MM-dd")
23            objTxtFile.Close
24            Set objTxtFile = Nothing
25        End If
          
          
          
26        strDir = App.Path & "\LisImage"
27        If Not objFSO.FolderExists(strDir) Then
28            Call objFSO.CreateFolder(strDir)
29        ElseIf Not blnDelOldImg Then
              '����Ƿ�����Ҫɾ���Ĺ���ͼ���ļ�
              

30            strDelFile = Dir(strDir & "\*.*")
31            Do While strDelFile <> ""
32                Set objFile = objFSO.GetFile(strDir & "\" & strDelFile)
33                If DateDiff("d", objFile.DateLastModified, Now) > 3 Then
34                    objFSO.DeleteFile strDir & "\" & strDelFile, True
35                End If
36                strDelFile = Dir
37            Loop
              
38        End If
39        If mObjImg Is Nothing Then
40            Set mObjImg = CreateObject("zlLisDev.clsDrawGraph")
41            If strErr <> "" Then
42                MsgBox strErr, vbInformation, gSysInfo.AppName
43                gobjLiscomlib.StopFlash
44                Exit Function
45            End If
46        End If
47        mObjImg.GetSampleImgExit strErr
          '�걾ID
          'ͼƬ����·��(���������Զ�����),
          '�Ƿ���ջ����ڱ��ص�ͼ���ļ�,True��ÿ�ζ������ݿ���ļ����浽����;False-��һ�ε���ʱ�����ݿ��ͼ�β���ͼƬ��֮��ֱ��ʹ��
          '��������ֵΪ�մ�ʱ�����ص���ʾ��Ϣ
          '���ص�ͼƬ�ļ���ʽ��0��cht(Ĭ��),1-jgp,2-png
          '���°�LIS�����ϰ�LIS�ڵ��ñ��������� 0-�ϰ�LIS��Ĭ�ϣ��ӡ�����ͼ��������ȡͼ�����ݣ���1-�°�LIS���ӡ����鱨��ͼ����ȡͼ�����ݣ�
48        If intVal = 25 Then
49            Call mObjImg.GetSampleImgInit(2500, gcnLisOracle, strErr)
50            strReturn = mObjImg.GetSampleImages(lngSampleID, strDir, False, strErr, 0, 1)
51        Else
52            Call mObjImg.GetSampleImgInit(100, gcnHisOracle, strErr)
53            strReturn = mObjImg.GetSampleImages(lngSampleID, strDir, False, strErr, 0, 0)
54        End If
55        If strReturn = "" Then
56            If strErr = "��ͼ�����ݣ�" Then
57                strErr = ""
58                ReadSampleImage = True
59            ElseIf strErr <> "" Then
60                MsgBox strErr, vbInformation, gSysInfo.AppName
61            Else
62                ReadSampleImage = True
63            End If
64            gobjLiscomlib.StopFlash
65            Exit Function
66        End If
          
67        varTmp = Split(strReturn, ",")

68        For i = LBound(varTmp) To UBound(varTmp)
69            If i > 8 Then Exit For
70            If Trim("" & varTmp(i)) <> "" Then
71                If Dir(strDir & "\" & Trim("" & varTmp(i))) <> "" Then strChar(i) = strDir & "\" & Trim("" & varTmp(i))
72            End If
73        Next
          
74        ReadSampleImage = True
75        gobjLiscomlib.StopFlash


76        Exit Function
ReadSampleImage_Error:
77        gobjLiscomlib.StopFlash
78        Call WriteErrLog("zl9LisInsideComm", "frmPubSampleImgs", "ִ��(ReadSampleImage)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
79        Err.Clear

End Function

