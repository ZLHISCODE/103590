VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDiseaseReportPrint 
   Caption         =   "���Ա����ӡ"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13545
   Icon            =   "frmDiseaseReportPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   13545
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picSel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   12600
      ScaleHeight     =   1170
      ScaleWidth      =   780
      TabIndex        =   8
      Top             =   3720
      Visible         =   0   'False
      Width           =   810
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ����"
         Height          =   180
         Index           =   3
         Left            =   30
         TabIndex        =   13
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ���"
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   720
      End
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����š�"
         Height          =   180
         Index           =   1
         Left            =   30
         TabIndex        =   10
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ�š�"
         Height          =   180
         Index           =   2
         Left            =   30
         TabIndex        =   9
         Top             =   600
         Width           =   720
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5745
      Left            =   390
      ScaleHeight     =   5745
      ScaleWidth      =   11205
      TabIndex        =   1
      Top             =   1500
      Width           =   11205
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   4635
         Left            =   480
         TabIndex        =   2
         Top             =   270
         Width           =   9375
         _cx             =   16536
         _cy             =   8176
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
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
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   360
      ScaleHeight     =   765
      ScaleWidth      =   12945
      TabIndex        =   0
      Top             =   570
      Width           =   12975
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   930
         TabIndex        =   17
         Top             =   60
         Width           =   2325
      End
      Begin VB.CheckBox chekPrint 
         BackColor       =   &H80000005&
         Caption         =   "��ʾ�Ѵ�ӡ"
         Height          =   225
         Left            =   10830
         TabIndex        =   14
         Top             =   105
         Width           =   1215
      End
      Begin VB.TextBox txtFind 
         Height          =   270
         Left            =   8220
         TabIndex        =   7
         Top             =   75
         Width           =   2295
      End
      Begin MSComCtl2.DTPicker DTPStart 
         Height          =   285
         Left            =   4170
         TabIndex        =   4
         Top             =   75
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   222953473
         CurrentDate     =   43161
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   285
         Left            =   5880
         TabIndex        =   5
         Top             =   75
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   222953473
         CurrentDate     =   43161
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   16
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֪����,������������ı�ǩ�����л������������������������������ʱʱ�������������Ч�ġ���ɫ��ʾͼ���ʾ������ӡ����"
         ForeColor       =   &H0000C000&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   10530
      End
      Begin VB.Label lblSel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ���"
         Height          =   180
         Left            =   7440
         TabIndex        =   12
         Top             =   120
         Width           =   720
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "--"
         Height          =   180
         Index           =   1
         Left            =   5580
         TabIndex        =   6
         Top             =   120
         Width           =   180
      End
      Begin VB.Label lblShow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   0
         Left            =   3360
         TabIndex        =   3
         Top             =   120
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgVsf 
      Left            =   11970
      Top             =   4770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportPrint.frx":6852
            Key             =   "ҽ����ӡ"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportPrint.frx":D0B4
            Key             =   "������ӡ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiseaseReportPrint.frx":13916
            Key             =   "��ֹ��ӡ"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   60
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDiseaseReportPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrDBUser As String        '�û�ID
Private mlngDeptID As Long          '��ǰѡ�����ID
Private mintDeptType As Integer     '���ҷ������
Private mrsDept As ADODB.Recordset  '��Ա���ڿ���
'Private mstrDeptIDS As String      'ʹ�á����п��ҡ�ѡ���޷���ȡ�����ӡ���ҵķ�����󣬵��²������ƴ�ӡ����

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/2
'��    ��:��ʾ����
'��    ��:
'           objFrm          ���ô���
'           lngDeptID       ����ID
'           intDeptType     �������� 0=����,1=����,2=סԺ
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Sub ShowMe(ByVal strDBUser As String)
    mstrDBUser = strDBUser
    Me.Show
End Sub

Private Sub lblDrowBorder(objLbl As Label, objPic As PictureBox)
    '��Labled�ı߿���,������ƶ���Lable��ʱ,����3DЧ��
    
    objPic.Line (objLbl.Left - 2, objLbl.Top - 2)-(objLbl.Left + objLbl.Width - 2, objLbl.Top - 2), &H8000000F '�ϱ���
    objPic.Line (objLbl.Left + objLbl.Width, objLbl.Top)-(objLbl.Left + objLbl.Width, objLbl.Top + objLbl.Height), vbBlack '�ұ���
    objPic.Line (objLbl.Left + objLbl.Width, objLbl.Top + objLbl.Height)-(objLbl.Left, objLbl.Top + objLbl.Height), vbBlack '�±���
    objPic.Line (objLbl.Left - 2, objLbl.Top + objLbl.Height)-(objLbl.Left - 2, objLbl.Top - 2), &H8000000F '�����

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/9
'��    ��:�ı�ѡ��ʱ��ȡ����ID�Ͳ��ŷ������
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub cboDept_Click()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim blnMZ As Boolean
          Dim blnZY As Boolean
          
1         On Error GoTo cboDept_Click_Error

2         If Me.Visible = False Then Exit Sub
3         With Me.cboDept
4             mlngDeptID = Val(.ItemData(.ListIndex))
5             If mlngDeptID = 0 Then Exit Sub
              
6             strSQL = "select distinct ������� from ��������˵�� where ����ID=[1] and (�������� = '�ٴ�' Or �������� = '����' Or �������� = '����')"
7             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���ŷ������", mlngDeptID)
8             Do While Not rsTmp.EOF
9                 If Val(rsTmp("�������") & "") = 1 Then
10                    blnMZ = True
11                ElseIf Val(rsTmp("�������") & "") = 2 Then
12                    blnZY = True
13                ElseIf Val(rsTmp("�������") & "") = 3 Then
14                    blnMZ = True
15                    blnZY = True
16                End If
17                rsTmp.MoveNext
18            Loop
19        End With
20        If blnMZ = True And blnZY = False Then
21            mintDeptType = 1
22        ElseIf blnMZ = False And blnZY = True Then
23            mintDeptType = 2
24        ElseIf blnMZ = True And blnZY = True Then
25            mintDeptType = 3
26        End If


27        Exit Sub
cboDept_Click_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "ִ��(cboDept_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
29        Err.Clear
          
End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim strFind As String
    
    If KeyAscii <> 13 Then Exit Sub
    With Me.cboDept
        strFind = Trim(.Text)
        If strFind = "" Then
            mrsDept.Filter = ""
        Else
            mrsDept.Filter = " ���� like '%" & strFind & "%' or ���� like '%" & strFind & "%' or ���� like '%" & strFind & "%'"
        End If
    End With
    Call setDataToCbo(mrsDept)
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_Find            '����
            Call FindData
        Case ConMenu_Browse_SelAll          'ȫѡ
            Call SelOrDelAll(1)
        Case ConMenu_Browse_ClsAll          'ȫ��
            Call SelOrDelAll(0)
        Case ConMenu_Browse_Print           '��ӡ
            Call BatchPrintReport(2)
        Case ConMenu_Browse_PrintSet        '��ӡ����
           Call BatchPrintReport(3)
        Case ConMenu_Browse_PrintView       'Ԥ��
            Call BatchPrintReport(1)
        Case ConMenu_Browse_Exit            '�˳�
            Unload Me
    End Select
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/5
'��    ��:��ӡ����
'��    ��:
'           1=Ԥ��,2=��ӡ,3=��ӡ����
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub BatchPrintReport(ByVal byRunMode As Byte)
          Dim lngRow As Long
          Dim lngSampleID As Long
          Dim strPrintCount As String
          Dim lngPrintCount As Long
          
1         On Error GoTo BatchPrintReport_Error

2         strPrintCount = ComGetPara(Sel_Lis_DB, "ҽ��վ��Ⱦ�������ӡ����", gSysInfo.SysNo, gSysInfo.ModlNo)
3         Select Case mintDeptType
              Case 1  '����
4                 lngPrintCount = Val(Split(strPrintCount, "|")(0))
5             Case 2  'סԺ
6                 lngPrintCount = Val(Split(strPrintCount, "|")(1))
7             Case 3  '�����סԺ����С��Ϊ׼
8                 If Val(Split(strPrintCount, "|")(0)) > Val(Split(strPrintCount, "|")(1)) Then
9                     lngPrintCount = Val(Split(strPrintCount, "|")(1))
10                Else
11                    lngPrintCount = Val(Split(strPrintCount, "|")(0))
12                End If
13            Case 0  '����
14                lngPrintCount = Val(Split(strPrintCount, "|")(2))
15        End Select
              
16        With Me.VSFList
17            For lngRow = 1 To .Rows - 1
18                If .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��")) = 1 Then
19                    lngSampleID = Val(.TextMatrix(lngRow, .ColIndex("ID")))
20                    Call PrintReport(lngSampleID, byRunMode, lngRow, lngPrintCount)
21                    If byRunMode = 3 Then Exit Sub
22                End If
23            Next
24        End With


25        Exit Sub
BatchPrintReport_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "ִ��(BatchPrintReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear
End Sub

Private Function PrintReport(lngSampleID As Long, Optional byRunMode As Byte = 2, Optional lngRow As Long, Optional lngPrintCount As Long) As Boolean
          '����       ��ӡ����
          Dim intCount As Integer
          Dim strNO As String
          Dim intSel As Integer
          Dim strChart(0 To 8) As String
          Dim strSQL As String
          Dim strTmp As String
          Dim rsTmp As ADODB.Recordset
          Dim rsReportFormat As ADODB.Recordset


1         On Error GoTo PrintReport_Error

2         strSQL = "select b.id ����id ,b.���� ��������,b.�������,Nvl(a.������Դ,1) ������Դ,a.����ʱ��,a.���Ա���,a.�걾���,a.ҽ��վ��ӡ from ���鱨���¼ a,����������¼ b where a.����id = b.id and a.id = [1]"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����ӡ", lngSampleID)

4         If rsTmp.RecordCount = 0 Then Exit Function

          '�Աȴ�ӡ�����Ͳ���
5         If lngPrintCount > 0 Then
6             If Val(rsTmp("ҽ��վ��ӡ") & "") > lngPrintCount Then
7                 With Me.VSFList
8                     .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
9                     .Cell(flexcpPicture, lngRow, .ColIndex("��ӡ��ʽ")) = imgVsf.ListImages("��ֹ��ӡ").ExtractIcon
10                End With
11                PrintReport = False
12                Exit Function
13            End If
14        End If

15        strSQL = "select id,����,����,���ﵥ��,סԺ����,��쵥��,Ժ�ⵥ��,�����ʽ,סԺ��ʽ,����ʽ,Ժ���ʽ,��ʽ����," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(���ﵥ��, '00000')) || '-2' ���ﵥ�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(סԺ����, '00000')) || '-2' סԺ���ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(��쵥��, '00000')) || '-2' ��쵥�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(Ժ�ⵥ��, '00000')) || '-2' Ժ�ⵥ�ݺ�," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(�����ʽ, '00000')) || '-2' �����ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(סԺ��ʽ, '00000')) || '-2' סԺ��ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(����ʽ, '00000')) || '-2' ����ʽ��," & vbNewLine & _
                      "       'ZLLISBILL' || Trim(To_Char(Ժ���ʽ, '00000')) || '-2' Ժ���ʽ��" & vbNewLine & _
                      "from ����������¼ where id = [1] "

16        Set rsReportFormat = ComOpenSQL(Sel_Lis_DB, strSQL, "���鼼ʦվ", Val(rsTmp("����ID") & ""))


17        rsReportFormat.Filter = "id=" & Val(rsTmp("����ID") & "")
18        If Val(rsTmp("�������")) = 1 Then
19            If Val(rsTmp("���Ա���") & "") = 1 Then
                  '����
20                intSel = 0
21            Else
                  '����
22                intSel = 1
23            End If
24        Else
25            intCount = GetSampleValCount(lngSampleID)
              'û�н��ʱ��ʾ
26            If intCount = 0 Then
27                Exit Function
28            End If
29            If rsReportFormat.RecordCount > 0 Then
30                If Val(rsReportFormat("��ʽ����") & "") > 0 Then
31                    If intCount > Val(rsReportFormat("��ʽ����") & "") Then
32                        intSel = 0
33                    Else
34                        intSel = 1
35                    End If
36                End If
37            Else
38                intSel = 0
39            End If

40        End If
41        Select Case Val(rsTmp("������Դ"))
              Case 1
42                If intSel = 0 Then
43                    strNO = rsReportFormat("���ﵥ�ݺ�")
44                Else
45                    strNO = rsReportFormat("�����ʽ��")
46                End If
47            Case 2
48                If intSel = 0 Then
49                    strNO = rsReportFormat("סԺ���ݺ�")
50                Else
51                    strNO = rsReportFormat("סԺ��ʽ��")
52                End If
53            Case 3
54                If intSel = 0 Then
55                    strNO = rsReportFormat("סԺ���ݺ�")
56                Else
57                    strNO = rsReportFormat("סԺ��ʽ��")
58                End If
59            Case 4
60                If intSel = 0 Then
61                    strNO = rsReportFormat("Ժ�ⵥ�ݺ�")
62                Else
63                    strNO = rsReportFormat("Ժ���ʽ��")
64                End If
65            Case Else
66                If intSel = 0 Then
67                    strNO = rsReportFormat("���ﵥ�ݺ�")
68                Else
69                    strNO = rsReportFormat("�����ʽ��")
70                End If
71        End Select
72        If byRunMode = 3 Then
73            If strNO <> "" Then
74                FunReportPrintSet gcnLisOracle, gSysInfo.SysNo, strNO, Me
75            End If
76        Else
             '��ͼ��
77            strTmp = "��ʼ����ͼ��:" & Now & vbCrLf
78            If ReadSampleImage(lngSampleID, strChart, "", 25) = False Then
79                Exit Function
80            End If
81            strTmp = strTmp & "����ͼ�����:" & Now & vbCrLf

82            FunReportOpen gcnLisOracle, gSysInfo.SysNo, strNO, Me, "�걾ID=" & lngSampleID, "ͼ��1=" & strChart(0), "ͼ��2=" & strChart(1), "ͼ��3=" & strChart(2), _
                      "ͼ��4=" & strChart(3), "ͼ��5=" & strChart(4), "ͼ��6=" & strChart(5), "ͼ��7=" & strChart(6), "ͼ��8=" & strChart(7), _
                      "ͼ��9=" & strChart(8), byRunMode
83            strTmp = strTmp & "��ӡ���:" & Now & vbCrLf

              '������˹��ı걾��ʶ
84            strSQL = "Zl_���鱨���ӡ_Edit(1," & lngSampleID & ",1)"
85            Call ComExecuteProc(Sel_Lis_DB, strSQL, "��ӡ�걾")
86            strTmp = strTmp & "��ɴ�ӡ:" & Now

87            SaveDBLog 18, 6, lngSampleID, "��ӡ", "�����ӡ", 2500, "�ٴ�ʵ���ҹ���"
88        End If

89        PrintReport = True

          '����ˢ�¿��ڸſ��Ѵ�ӡ��ǩ����
90        Call SendMessage("RefreshDeptSurvey7")


91        Exit Function
PrintReport_Error:
92        Call WriteErrLog("zlPublicHisCommLis", "frmPatientReprotBrowse", "ִ��(PrintReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
93        Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/5
'��    ��:ȫѡ/ȫ��
'��    ��:
'           intType     0=ȫ��,1=Ȩ��
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub SelOrDelAll(ByVal intType As Integer)
          Dim lngRow As Long
1         On Error GoTo SelOrDelAll_Error

2         With Me.VSFList
3             For lngRow = 1 To .Rows - 1
4                 .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��")) = intType
5             Next
6         End With


7         Exit Sub
SelOrDelAll_Error:
8         Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "ִ��(SelOrDelAll)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
9         Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/2
'��    ��:��������
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub FindData()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim lngRow As Long
          Dim lngDeptID As Long
          Dim strErr As String
          
1         On Error GoTo FindData_Error
          
          '��ȡ����ID
2         With Me.cboDept
3             If .Text <> "���п���" Then
4                 lngDeptID = Val(.ItemData(.ListIndex))
5                 If lngDeptID = 0 Then Exit Sub
6             End If
7         End With
      '    If mstrDeptIDS = "" Then Exit Sub

              
          '����ID
      '    With Me.cboDept
      '        If .Text <> "���п���" Then
      '            strSQL = "Select Distinct '0' ѡ��,decode(Sign(Nvl(a.ҽ��վ��ӡ, 0)),1,'ҽ����ӡ',decode(Sign(Nvl(a.������ӡ����, 0)),1,'������ӡ','')) ��ӡ��ʽ,a.ҽ��վ��ӡ, a.Id, a.�걾��� �걾��, a.����, decode(a.�Ա�,1,'��',2,'Ů','δ֪') �Ա�, a.����," & vbCrLf & _
      '                " a.�걾����, a.���ʱ�� ����ʱ��, a.����� ������,a.��������, a.������Ŀ" & vbCrLf & _
      '                " From ���鱨���¼ A, ����������� B, ���鱨����ӡ���� C" & vbCrLf & _
      '                " Where a.Id = b.�걾id And b.���id = c.���id And a.�Ƿ�Ⱦ�� = 1 And ����� Is Not Null And c.����id = [1] "
      '        Else
8                 strSQL = "Select /*+cardinality(d,10)*/ Distinct '0' ѡ��,decode(Sign(Nvl(a.ҽ��վ��ӡ, 0)),1,'ҽ����ӡ',decode(Sign(Nvl(a.������ӡ����, 0)),1,'������ӡ','')) ��ӡ��ʽ,a.ҽ��վ��ӡ, a.Id, a.�걾��� �걾��, a.����, decode(a.�Ա�,1,'��',2,'Ů','δ֪') �Ա�, a.����," & vbCrLf & _
                      " a.�걾����, a.���ʱ�� ����ʱ��, a.����� ������,a.��������, a.������Ŀ" & vbCrLf & _
                      " From ���鱨���¼ A, ����������� B, ���鱨����ӡ���� C" & vbCrLf & _
                      " Where a.Id = b.�걾id And b.���id = c.���id And a.�Ƿ�Ⱦ�� = 1 And ����� Is Not Null And c.����id in (Select Column_Value From Table(Cast(f_Str2list([1]) As zltools.t_strlist)) d)"
      '        End If
      '    End With
          
          '��������
9         If Trim(Me.txtFind.Text) <> "" Then
10            Select Case Me.lblSel.Caption
                  Case "��  ���"
11                    strSQL = strSQL & " and a.��������=[2]"
12                Case "����š�"
13                    strSQL = strSQL & " and a.�����=[2]"
14                Case "סԺ�š�"
15                    strSQL = strSQL & " and a.סԺ��=[2]"
16                Case "��  ����"
17                    strSQL = strSQL & " and a.���� like [2]"
18            End Select
19        End If
          
          '����ʱ��
20        If Trim(Me.txtFind.Text) = "" Then
21            strSQL = strSQL & " and a.���ʱ�� between [3] and [4]"
              '�߷�ʱ�����Ʋ�ѯ
22            If Not funCheckRushHours(2500, 2001, "���������", CDate(Format(Me.DTPStart.value, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(Me.DTPEnd.value, "yyyy-mm-dd") & " 23:59:59")) Then Exit Sub
23        End If
          
          '�Ƿ���ʾ�Ѵ�ӡ
24        If Me.chekPrint.value <> 1 Then
25            strSQL = strSQL & " and nvl(a.ҽ��վ��ӡ,0)=0 and nvl(a.������ӡ����,0)=0 "
26        End If
          
27        strSQL = strSQL & " order by a.���ʱ��"
      '    Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鱨���¼", IIf(cboDept.Text <> "���п���", lngDeptID, mstrDeptIDS), Trim(Me.txtFind.Text), CDate(Format(Me.DTPStart.value, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(Me.DTPEnd.value, "yyyy-mm-dd") & " 23:59:59"))
28        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鱨���¼", lngDeptID, IIf(Me.lblSel.Caption = "��  ����", "%" & Trim(Me.txtFind.Text) & "%", Trim(Me.txtFind.Text)), CDate(Format(Me.DTPStart.value, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(Me.DTPEnd.value, "yyyy-mm-dd") & " 23:59:59"))
29        If vfgLoadFromRecord(Me.VSFList, rsTmp, strErr) = False Then
30            MsgBox strErr
31            Exit Sub
32        End If
33        With Me.VSFList
34            .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
35            .ColWidth(.ColIndex("ѡ��")) = 500
36            .ColWidth(.ColIndex("��ӡ��ʽ")) = 250
37            .ColWidth(.ColIndex("�걾��")) = 800
38            .ColWidth(.ColIndex("����")) = 1000
39            .ColWidth(.ColIndex("�Ա�")) = 500
40            .ColWidth(.ColIndex("����")) = 800
41            .ColWidth(.ColIndex("�걾����")) = 1000
42            .ColWidth(.ColIndex("����ʱ��")) = 2000
43            .ColWidth(.ColIndex("������")) = 1000
44            .ColWidth(.ColIndex("��������")) = 1500
45            .ColWidth(.ColIndex("������Ŀ")) = 1000
46            .ExtendLastCol = True
              
47            .Cell(flexcpPicture, 0, .ColIndex("��ӡ��ʽ")) = imgVsf.ListImages("ҽ����ӡ").ExtractIcon
              
48            For lngRow = 1 To .Rows - 1
49                If Trim(.TextMatrix(lngRow, .ColIndex("��ӡ��ʽ"))) = "ҽ����ӡ" Then
50                    .Cell(flexcpPicture, lngRow, .ColIndex("��ӡ��ʽ")) = imgVsf.ListImages("ҽ����ӡ").ExtractIcon
51                End If
                  
52                If Trim(.TextMatrix(lngRow, .ColIndex("��ӡ��ʽ"))) = "������ӡ" Then
53                    .Cell(flexcpPicture, lngRow, .ColIndex("��ӡ��ʽ")) = imgVsf.ListImages("������ӡ").ExtractIcon
54                End If
55            Next
56        End With


57        Exit Sub
FindData_Error:
58        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "ִ��(FindData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
59        Err.Clear
End Sub

Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    On Error Resume Next
    With Me.picTop
        .Left = Left
        .Top = Top
        .Width = Right - Left
    End With
    With Me.picMain
        .Left = Left
        .Top = Me.picTop.Top + Me.picTop.Height
        .Width = Me.picTop.Width
        .Height = Bottom - .Top
    End With
    With picSel
        .Left = Me.picTop.Left + lblSel.Left - 10
        .Top = Me.picTop.Top + lblSel.Top + lblSel.Height
    End With
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/5
'��    ��:��ݼ������ʹ�ÿ�ݼ���Ч�����ݼ��Ƿ���������ռ��
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 1      'ȫѡ
            Call SelOrDelAll(1)
        Case 4      'ȫ��
            Call SelOrDelAll(0)
        Case 16     '��ӡ
            Call BatchPrintReport(2)
        Case 21     'Ԥ��
            Call BatchPrintReport(1)
        Case 17     '�˳�
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrMain.VisualTheme = xtpThemeOffice2003
    Me.cbrMain.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrMain.EnableCustomization False

    '-----------------------------------------------------
    '�˵�����
    Me.cbrMain.ActiveMenuBar.Title = "�˵�"
    Me.cbrMain.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Find, "����(F5)")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_SelAll, "ȫѡ(Crl+A)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_ClsAll, "ȫ��(Crl+D)")
        Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_Print, "��ӡ(Crl+P)"): cbrControl.BeginGroup = True
        cbrControl.Style = xtpButtonIconAndCaption
        With cbrControl.CommandBar.Controls
            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "��ӡ����  ")
        End With
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintView, "Ԥ��(Crl+U)")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "�˳�(Crl+Q)"): cbrControl.BeginGroup = True
    End With

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    '�����
    With Me.cbrMain.KeyBindings
        .Add 0, VK_F5, ConMenu_Browse_Find
    End With
    
    Call intData
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/5
'��    ��:��ʼ������
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub intData()
          Dim strTitle As String
          
          '��ʼ��VSF
1         On Error GoTo intData_Error

2         strTitle = "ѡ��,500,1;��ӡ��ʽ,250,1;�걾��,800,1;����,1000,1;�Ա�,500,1;����,800,1;�걾����,1000,1;" & _
                      "����ʱ��,2000,1;������,1000,1;��������,1500,1;������Ŀ,1000,1"
3         Call vfgSetting(0, Me.VSFList, strTitle)
4         With Me.VSFList
5             .ExtendLastCol = True
6              .Cell(flexcpPicture, 0, .ColIndex("��ӡ��ʽ")) = imgVsf.ListImages("ҽ����ӡ").ExtractIcon
7         End With
          
          '��ȡ������ʱ��
8         Me.DTPStart.value = Currentdate
9         Me.DTPEnd.value = Me.DTPStart.value
          
          '��ȡ�û���ǰ����
10        Call getUserDept


11        Exit Sub
intData_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "ִ��(intData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
13        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/9
'��    ��:��ȡ��Ա����
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub getUserDept()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
1         On Error GoTo getUserDept_Error

2         strSQL = "Select a.Id, a.����, a.����, a.����" & vbCrLf & _
                  " From ���ű� A, �ϻ���Ա�� B, ������Ա C" & vbCrLf & _
                  " Where a.Id = c.����id And b.��Աid = c.��Աid And b.�û��� = [1] And a.����ʱ�� > Sysdate "
3         If gUserInfo.NodeNo <> "-" Then
4             strSQL = strSQL & " And (a.վ�� = 1 Or a.վ�� Is Null)"
5         End If
6         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "��Ա����", mstrDBUser)
7         Set mrsDept = rsTmp
8         Call setDataToCbo(rsTmp)

9         Exit Sub
getUserDept_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "ִ��(getUserDept)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/9
'��    ��:�����ݰ󶨵������б���
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub setDataToCbo(ByVal rsTmp As ADODB.Recordset)
1         On Error GoTo setDataToCbo_Error

2         With Me.cboDept
3             .Clear
      '        .AddItem "���п���"
4             Do While Not rsTmp.EOF
5                 .AddItem "[" & rsTmp("����") & "]" & rsTmp("����")
6                 .ItemData(.ListCount - 1) = Val(rsTmp("ID") & "")
      '            If Val(rsTmp("ID") & "") <> 0 Then mstrDeptIDS = mstrDeptIDS & "," & rsTmp("ID")
7                 rsTmp.MoveNext
8             Loop
      '        If mstrDeptIDS <> "" Then mstrDeptIDS = Mid(mstrDeptIDS, 2)
9             If .ListCount > 0 Then .ListIndex = 0
10        End With


11        Exit Sub
setDataToCbo_Error:
12        Call WriteErrLog("zlPublicHisCommLis", "frmDiseaseReportPrint", "ִ��(setDataToCbo)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
13        Err.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
'    mstrDeptIDS = ""
    mstrDBUser = ""
    mlngDeptID = 0
    Set mrsDept = Nothing
End Sub

Private Sub lblSel_Click()
    If Me.picSel.Visible = False Then
        Me.picSel.Visible = True
    Else
        Me.picSel.Visible = False
    End If
End Sub

Private Sub lblSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblDrowBorder(lblSel, picTop)
End Sub

Private Sub lblSelect_Click(Index As Integer)
    Me.lblSel.Caption = Me.lblSelect(Index).Caption
    Me.picSel.Visible = False
End Sub

Private Sub lblSelect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call lblDrowBorder(lblSelect(Index), picSel)
End Sub

Private Sub PicMain_Resize()
    On Error Resume Next
    With Me.VSFList
        .Left = 0
        .Top = 0
        .Width = Me.picMain.Width
        .Height = Me.picMain.Height
    End With
End Sub

Private Sub picSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSel.Cls
End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picTop.Cls
End Sub

Private Sub txtFind_GotFocus()
    With Me.txtFind
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call FindData
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/5
'��    ��:ѡ��/ȡ��ѡ��
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub VSFList_Click()
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error Resume Next
    
    With Me.VSFList
        lngRow = .MouseRow
        lngCol = .MouseCol
        
        If lngRow <= 0 Or lngCol <> .ColIndex("ѡ��") Then Exit Sub
        If .Cell(flexcpChecked, lngRow, lngCol) = 1 Then
            .Cell(flexcpChecked, lngRow, lngCol) = 0
        Else
            .Cell(flexcpChecked, lngRow, lngCol) = 1
        End If
    End With
End Sub
