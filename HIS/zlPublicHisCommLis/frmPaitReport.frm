VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmPaitReport 
   Caption         =   "���˱���鿴"
   ClientHeight    =   11160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16755
   Icon            =   "frmPaitReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11160
   ScaleWidth      =   16755
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8295
      Left            =   270
      ScaleHeight     =   8265
      ScaleWidth      =   15615
      TabIndex        =   16
      Top             =   1950
      Width           =   15645
      Begin VB.PictureBox picPaitList 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5925
         Left            =   150
         ScaleHeight     =   5895
         ScaleWidth      =   3375
         TabIndex        =   22
         Top             =   390
         Width           =   3405
         Begin VSFlex8Ctl.VSFlexGrid vsfPaitList 
            Height          =   3525
            Left            =   0
            TabIndex        =   23
            Top             =   120
            Width           =   3705
            _cx             =   6535
            _cy             =   6218
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
      Begin VB.Frame fraWE 
         BorderStyle     =   0  'None
         Height          =   3645
         Left            =   3810
         MousePointer    =   9  'Size W E
         TabIndex        =   20
         Top             =   870
         Width           =   105
      End
      Begin VB.PictureBox picPaitReport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   7935
         Left            =   4050
         ScaleHeight     =   7905
         ScaleWidth      =   11295
         TabIndex        =   17
         Top             =   180
         Width           =   11325
         Begin VSFlex8Ctl.VSFlexGrid vsfScroll 
            Height          =   6315
            Left            =   180
            TabIndex        =   18
            Top             =   810
            Width           =   8745
            _cx             =   15425
            _cy             =   11139
            Appearance      =   2
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
            ForeColor       =   -2147483643
            BackColorFixed  =   -2147483643
            ForeColorFixed  =   -2147483643
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483643
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            TreeColor       =   -2147483632
            FloodColor      =   -2147483643
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
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
            BackColorFrozen =   -2147483643
            ForeColorFrozen =   -2147483643
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
            Begin VB.PictureBox picScroll 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   6015
               Left            =   420
               ScaleHeight     =   6015
               ScaleWidth      =   8055
               TabIndex        =   19
               Top             =   90
               Visible         =   0   'False
               Width           =   8055
               Begin zl9LisInsideComm.uclReport uclSampleReport 
                  Height          =   5145
                  Index           =   0
                  Left            =   60
                  TabIndex        =   21
                  Top             =   60
                  Width           =   7965
                  _extentx        =   14049
                  _extenty        =   10451
               End
            End
         End
      End
   End
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1785
      ScaleWidth      =   16455
      TabIndex        =   0
      Top             =   120
      Width           =   16485
      Begin VB.PictureBox picIDKIND 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   6960
         ScaleHeight     =   285
         ScaleWidth      =   705
         TabIndex        =   26
         Top             =   210
         Width           =   735
      End
      Begin VB.TextBox txtPaitKey 
         Height          =   315
         Left            =   960
         TabIndex        =   25
         ToolTipText     =   "��������ͷͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
         Top             =   210
         Width           =   5955
      End
      Begin VB.CheckBox chkVerifyDate 
         Height          =   255
         Left            =   7530
         TabIndex        =   5
         Top             =   960
         Width           =   300
      End
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   945
         TabIndex        =   4
         Top             =   570
         Width           =   2970
      End
      Begin VB.ComboBox cbodor 
         Height          =   300
         Left            =   4830
         TabIndex        =   3
         Top             =   570
         Width           =   2910
      End
      Begin VB.ComboBox cboDiseases 
         Height          =   300
         ItemData        =   "frmPaitReport.frx":6852
         Left            =   945
         List            =   "frmPaitReport.frx":685F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1305
         Width           =   1485
      End
      Begin VB.TextBox txtRptCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   3570
         MaxLength       =   4
         TabIndex        =   1
         Text            =   "0"
         Top             =   1335
         Width           =   510
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   300
         Left            =   2490
         TabIndex        =   6
         Top             =   930
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   117112835
         CurrentDate     =   40954
      End
      Begin MSComCtl2.DTPicker dtpVS 
         Height          =   300
         Left            =   4830
         TabIndex        =   7
         Top             =   930
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   117112835
         CurrentDate     =   40954
      End
      Begin MSComCtl2.DTPicker dtpVE 
         Height          =   300
         Left            =   6225
         TabIndex        =   8
         Top             =   930
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   117112835
         CurrentDate     =   40954
      End
      Begin MSComCtl2.DTPicker dtpS 
         Height          =   300
         Left            =   945
         TabIndex        =   9
         Top             =   930
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   117112835
         CurrentDate     =   40954
      End
      Begin MSComCtl2.UpDown upd 
         Height          =   330
         Left            =   4170
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1230
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         Value           =   1
         OrigLeft        =   3480
         OrigTop         =   420
         OrigRight       =   3735
         OrigBottom      =   690
         Max             =   99
         Min             =   1
         Enabled         =   -1  'True
      End
      Begin VB.Label ������ 
         AutoSize        =   -1  'True
         Caption         =   "�� �� ��"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label5 
         Caption         =   "��������"
         Height          =   240
         Left            =   4080
         TabIndex        =   15
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "��������"
         Height          =   240
         Left            =   60
         TabIndex        =   14
         Top             =   960
         Width           =   750
      End
      Begin VB.Label lblDor 
         AutoSize        =   -1  'True
         Caption         =   "����ҽ��"
         Height          =   180
         Left            =   4020
         TabIndex        =   13
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lblDept 
         AutoSize        =   -1  'True
         Caption         =   "������ҡ�"
         Height          =   180
         Left            =   60
         TabIndex        =   12
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "�� Ⱦ ��"
         Height          =   180
         Left            =   60
         TabIndex        =   11
         Top             =   1365
         Width           =   720
      End
      Begin VB.Line Line1 
         X1              =   3570
         X2              =   4140
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "�鿴�������          �ݱ���"
         Height          =   225
         Left            =   2520
         TabIndex        =   10
         Top             =   1350
         Width           =   2925
      End
   End
   Begin XtremeCommandBars.CommandBars cbrMain 
      Left            =   0
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPaitReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'��̬�����Ƿ���ʾ����߿�
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const const_PicRectBackColour As Long = &HE0E0E0

'��ӡPDF
Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (lpExecInfo As SHELLEXECUTEINFO) As Long

Private mblnDoctorShow As Boolean                       '�Ƿ���ҽ��վ����
Private mstrPrivs As String                             '����ԱȨ��
Private mlngPatientID As Long                           '����ID
Private mlngPatientPage As Long                         '��ҳID
Private mstrPatientGH As String                         '�Һŵ�
Private mstrThirdReport As String                       '��������
Private WithEvents mobjIDKind As VBControlExtender      'IDKind����
Attribute mobjIDKind.VB_VarHelpID = -1

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-26
'��    ��:  ���ش���
'           objfrm              ���ö���
'           intShowType         ������Դ��1=��ʦվ���ã�������ʾ�ʹ�ӡδ��˵ı��棬2=ҽ��վ���ã���ʱֻ��ʾ�Ѿ���˵ı��棬
'           lngPatientID        ����ID
'           strPrivs            ģ��Ȩ��
'           lngDept             �򿪵�ǰģ��Ŀ���
'           lngDeptDistrict     �򿪵�ǰģ��Ĳ���
'           intPatientType      ������Դ
'           lngPatientPage      ��ҳID
'           blnShowBorder       �Ƿ���ʾ����
'           blnFindData         �򿪴���ʱ�Ƿ�Ĭ�ϼ�������
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Function showMe(objFrm As Object, Optional lngPatientID As Long, Optional strPrivs As String, Optional lngDept As Long, Optional lngDeptDistrict As Long, _
                       Optional intPatientType As Integer, Optional lngPatientPage As Long, Optional strErr As String, Optional ByVal blnShowBorder As Boolean, _
                       Optional ByRef objOutFrm As Object, Optional blnDoctorShow As Boolean = True, Optional ByVal blnFindData As Boolean = True) As Boolean
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset

          '��ȡȨ��
1         On Error GoTo showMe_Error

2         mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 2001)
3         mstrPrivs = strPrivs & ";" & mstrPrivs
4         mblnDoctorShow = blnDoctorShow
5         mlngPatientPage = lngPatientPage


6         If lngPatientID <> 0 Then txtPaitKey.Text = lngPatientID

7         If lngDeptDistrict > 0 Then
              'סԺ�鿴���鱨��
8             lblDept.Caption = "���벡����"
9             Call InitDepts(1)
10            Call GetDeptDor
11            If cboDept.ListCount > 0 Then
12                CboFind cboDept, lngDeptDistrict
13            End If
14            If cbodor.ListCount > 0 Then
15                CboFind cbodor, UserInfo.ID
16            End If

              '��ѯסԺ�������Ժʱ��
17            strSQL = "select ��Ժ����,nvl(��Ժ����,sysdate) ��Ժ���� from ������ҳ where ����id=[1] and ��ҳID=[2]"
18            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "������ҳ", lngPatientID, lngPatientPage)
19            If Not rsTmp.EOF Then
20                dtpS.value = CDate(Format(rsTmp("��Ժ����") & "", "yyyy/mm/dd hh:mm:ss"))
21                dtpE.value = CDate(Format(rsTmp("��Ժ����") & "", "yyyy/mm/dd hh:mm:ss"))
22            End If
23        Else
              '����鿴���鱨��
24            lblDept.Caption = "������ҡ�"
25            Call InitDepts(0)
26            Call GetDeptDor
27            If cboDept.ListCount > 0 Then
28                CboFind cboDept, lngDept
29            End If
30        End If

31        If blnShowBorder Then
32            Me.Show  '�������ʾ����ı߿����ʾ�ô���ΪǶ��ʽ���ã����ǵ���show����
33        Else
34            Call YSystemMenu(Me.hWnd)
35        End If
          
          'Ĭ�ϼ�������
36        If blnFindData Then
37            Call GetDeptPaits
38        End If

39        Set objOutFrm = Me

40        showMe = True


41        Exit Function
showMe_Error:
42        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "ִ��(showMe)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
43        Err.Clear
End Function

Public Function BHShowMe(lngMain As Long, Optional strErr As String) As Boolean
    On Error GoTo errH
    mstrPrivs = ComGetPrivs(Sel_Lis_DB, gSysInfo.SysNo, 1013)
    

    gobjLiscomlib.ShowChildWindow Me.hWnd, lngMain
    BHShowMe = True
        

    Exit Function
errH:
    strErr = "������(ShowMe),������Ϣ:" & Err.Number & " " & Err.Description
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/5/25
'��    ��:����API��̬���ô����border
'��    ��:
'           new_Hwnd    ����ľ��
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub YSystemMenu(ByVal new_Hwnd As Long)
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 'Or WS_SYSMENU Or &H20000
End Sub

Private Function InitDepts(intDeptView As Integer, Optional strErr As String) As Boolean
      '���ܣ���ʼ��סԺ�ٴ�����
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, i As Long
          Dim strDeptIDs As String, lngPreDept As Long


1         On Error GoTo InitDepts_Error

2         If mblnDoctorShow Then
3             If intDeptView = 0 Then
                  '�����Ҷ�ȡ��ʾ
                  '�����ż���۲��ҵĲ��˻�û���ϴ�������ֻ�Դ����в��˵Ŀ��ҵ�����
4                 If InStr(";" & mstrPrivs & ";", ";ȫԺ����;") > 0 Then
5                     strDeptIDs = GetUser����IDs
6                     strSQL = _
                    " Select Distinct A.ID,A.����,A.����,a.����" & _
                             " From ���ű� A,��������˵�� B" & _
                             " Where B.����ID=A.ID And B.��������='�ٴ�'" & _
                             " And (B.������� IN(2,3) Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                             " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                             " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                             " Order by A.����"
7                 Else
                      '����Ȩ�޵Ŀ��ң��������ڿ���+�������������Ŀ���
8                     strSQL = _
                    " Select A.ID,A.����,A.����,a.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
                             " From ���ű� A,��������˵�� B,������Ա C" & _
                             " Where B.����ID=A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
                             " And (B.������� IN(2,3) Or (B.�������=1 And Exists(Select 1 From ��λ״����¼ C Where B.����ID = C.����ID)))" & _
                             " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                             " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                             " And B.��������='�ٴ�'"
9                     strSQL = strSQL & " Union " & _
                             " Select C.ID,C.����,C.����,C.����,Nvl(A.ȱʡ,0) As ȱʡ" & _
                             " From ������Ա A,�������Ҷ�Ӧ B,���ű� C" & _
                             " Where A.����ID=B.����ID And B.����ID=C.ID And A.��ԱID=[1]" & _
                             " And Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=B.����ID)" & _
                             " And Not Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=B.����ID)" & _
                             " And (C.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                             " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)"
10                    If InStr(";" & mstrPrivs & ";", ";ICU����;") > 0 Then
11                        strSQL = strSQL & " Union " & _
                                 " Select A.ID,A.����,A.����,a.����,0 As ȱʡ" & _
                                 " From ���ű� A" & _
                                 " Where Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='ICU')" & _
                                 " And Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='�ٴ�')" & _
                                 " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                                 " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
12                    End If
13                    strSQL = "Select ID,����,����,����,Max(ȱʡ) As ȱʡ From (" & strSQL & ") Group By ID,����,����,���� Order by ����"
14                End If
15            Else
                  '��������ȡ��ʾ
16                If InStr(";" & mstrPrivs & ";", ";ȫԺ����;") > 0 Then
17                    strDeptIDs = GetUser����IDs
18                    strSQL = _
                    " Select Distinct A.ID,A.����,A.����,a.����" & _
                             " From ���ű� A,��������˵�� B " & _
                             " Where A.ID=B.����ID And B.������� in(1,2,3) And B.��������='����'" & _
                             " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                             " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                             " Order by A.����"
19                Else
                      '����Ȩ������ֱ�����ڲ���+���ڿ�����������
20                    strSQL = _
                    " Select A.ID,A.����,A.����,a.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
                             " From ���ű� A,��������˵�� B,������Ա C" & _
                             " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
                             " And B.������� in(1,2,3) And B.��������='����'" & _
                             " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                             " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
21                    strSQL = strSQL & " Union " & _
                             " Select C.ID,C.����,C.����,C.����,Nvl(A.ȱʡ,0) as ȱʡ" & _
                             " From ������Ա A,�������Ҷ�Ӧ B,���ű� C" & _
                             " Where A.����ID=B.����ID And B.����ID=C.ID And A.��ԱID=[1]" & _
                             " And Exists(Select 1 From ��������˵�� Where ��������='�ٴ�' And ����ID=B.����ID)" & _
                             " And Not Exists(Select 1 From ��������˵�� Where ��������='����' And ����ID=B.����ID)" & _
                             " And (C.����ʱ�� is NULL or Trunc(C.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                             " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)"
22                    If InStr(";" & mstrPrivs & ";", ";ICU����;") > 0 Then
23                        strSQL = strSQL & " Union " & _
                                 " Select A.ID,A.����,A.����,a.����,0 As ȱʡ" & _
                                 " From ���ű� A" & _
                                 " Where Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='ICU')" & _
                                 " And Exists(Select 1 From ��������˵�� B Where A.ID=B.����ID And B.��������='����')" & _
                                 " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                                 " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
24                    End If
25                    strSQL = "Select ID,����,����,����,Max(ȱʡ) as ȱʡ From (" & strSQL & ") Group by ID,����,����,���� Order by ����"
26                End If
27            End If
28        Else
29            strSQL = "Select Distinct a.id, a.����, a.����, a.���� From ���ű� A, ��������˵�� B" & _
                     " Where a.Id = b.����id And a.����ʱ�� Is Not Null And a.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd') And" & _
                     " (b.�������� = '�ٴ�' Or b.�������� = '����' Or b.�������� = '����' Or b.�������� = '����') order by a.����"
30        End If

31        cboDept.Clear
32        If InStr(";" & mstrPrivs & ";", ";���п���;") > 0 Then cboDept.AddItem "00-���п���"
33        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, UserInfo.ID)

34        For i = 1 To rsTmp.RecordCount
35            cboDept.AddItem rsTmp!���� & "-" & rsTmp!���� & "[" & rsTmp!���� & "]"
36            cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
37            rsTmp.MoveNext
38        Next
39        If rsTmp.RecordCount > 0 Then
40            cboDept.ListIndex = 0
41        End If
42        InitDepts = True


43        Exit Function
InitDepts_Error:
44        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmShowSampleReport", "ִ��(InitDepts)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
45        Err.Clear

End Function

Public Function GetUser����IDs(Optional ByVal bln���� As Boolean, Optional strErr As String) As String
      '���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
      '�������Ƿ�ȡ���������µĿ���
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, i As Long, blnNew As Boolean

1         On Error GoTo GetUser����IDs_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
          'û��ǿ�������ٴ�,����ҽ��������
7         If blnNew Then
8             strSQL = "Select 1 as ���,����ID From ������Ա Where ��ԱID=[1] Union" & _
                     " Select Distinct 2 as ���,B.����ID From ������Ա A,�������Ҷ�Ӧ B" & _
                     " Where A.����ID=B.����ID And A.��ԱID=[1]"

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", UserInfo.ID)
10        End If
11        If bln���� = False Then
12            rsTmp.Filter = "��� = 1"
13        Else
14            rsTmp.Filter = ""
15        End If

16        For i = 1 To rsTmp.RecordCount
17            If InStr("," & GetUser����IDs & ",", "," & rsTmp!����ID & ",") = 0 Then
18                GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
19            End If
20            rsTmp.MoveNext
21        Next
22        GetUser����IDs = Mid(GetUser����IDs, 2)



23        Exit Function
GetUser����IDs_Error:
24        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmShowSampleReport", "ִ��(GetUser����IDs)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
25        Err.Clear

End Function

Public Function GetUser����IDs(Optional strErr As String) As String
      '���ܣ���ȡ����Ա�����Ĳ���(ֱ�����ڲ��������ڿ��������Ĳ���),�����ж��
          Static rsTmp As ADODB.Recordset
          Dim strSQL As String, i As Long, blnNew As Boolean

1         On Error GoTo GetUser����IDs_Error

2         If rsTmp Is Nothing Then
3             blnNew = True
4         Else
5             blnNew = (rsTmp.State = adStateClosed)
6         End If
7         If blnNew Then
8             strSQL = _
              "Select Distinct ����ID From (" & _
                     " Select A.����ID as ����ID" & _
                     " From ��������˵�� A,������Ա B" & _
                     " Where A.����ID=B.����ID And B.��ԱID=[1]" & _
                     " And A.������� in(1,2,3) And A.��������='����'" & _
                     " Union" & _
                     " Select A.����ID From �������Ҷ�Ӧ A,������Ա B" & _
                     " Where A.����ID=B.����ID And B.��ԱID=[1])"

9             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", UserInfo.ID)
10        ElseIf rsTmp.RecordCount > 0 Then
11            rsTmp.MoveFirst
12        End If
13        For i = 1 To rsTmp.RecordCount
14            GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
15            rsTmp.MoveNext
16        Next

17        GetUser����IDs = Mid(GetUser����IDs, 2)



18        Exit Function
GetUser����IDs_Error:
19        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmShowSampleReport", "ִ��(GetUser����IDs)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
20        Err.Clear

End Function

Private Sub CboFind(objcbo As ComboBox, lngID As Long)
    '����           �ҵ�cbo��Ӧ��id
    Dim intloop As Integer
    With objcbo
        For intloop = 0 To .ListCount - 1
            If .ItemData(intloop) = lngID Then
                .ListIndex = intloop
                Exit Sub
            End If
        Next
        .ListIndex = 0
    End With
End Sub

Private Sub cboDept_Click()
    '��ȡ����ҽ��
    Call GetDeptDor(cboDept.ItemData(cboDept.ListIndex))
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-26
'��    ��:  ��ȡѡ�п��Ҳ���
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Sub GetDeptPaits()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsPait As ADODB.Recordset
          Dim strPatientIDs As String
          Dim strArr() As String
          Dim i As Integer
          Dim j As Integer
          Dim lngPaitID As Long



1         On Error GoTo GetDeptPaits_Error

2         If Trim(txtPaitKey.Text) <> "" Then
              'ˢ��ͨ������ID���ң�ͨ��ID����ʱ������������
3             If Mid(Trim(txtPaitKey.Text), 1, 1) = "-" Then
4                 If IsNumeric(Mid(Trim(txtPaitKey.Text), 2)) Then lngPaitID = Mid(Trim(txtPaitKey.Text), 2)
5             Else
6                 If IsNumeric(Trim(txtPaitKey.Text)) Then lngPaitID = Trim(txtPaitKey.Text)
7             End If

8             strSQL = "Select *" & vbCrLf & _
                     "   From (Select row_number() over(Partition By a.HIS����ID Order By a.����ʱ�� Desc) ���, a.HIS����ID, a.����," & vbCrLf & _
                     "                 Decode(a.�Ա�, '1', '��', '2', 'Ů', '9', 'δ֪', '������') �Ա�, a.����," & vbCrLf & _
                     "                 Nvl(a.������, Decode(a.������Դ, 1, a.�����, 2, a.סԺ��, 3, a.������, 4, a.������, Decode(a.�Һŵ�, Null, a.�շѵ���, a.�Һŵ�))) ������,a.�Һŵ�" & vbCrLf & _
                     "          From ���鱨���¼ A" & vbCrLf & _
                     "          Where Nvl(a.�Ƿ��ʿر걾, 0) = 0 and a.HIS����ID =[1]) Where ��� = 1"
9             Set rsPait = ComOpenSQL(Sel_Lis_DB, strSQL, "�����б�", lngPaitID)
10        Else
              '���°���ȥ���Ҳ���
11            strSQL = "Select *" & vbCrLf & _
                     "   From (Select row_number() over(Partition By a.HIS����ID Order By a.����ʱ�� Desc) ���, a.HIS����ID, a.����," & vbCrLf & _
                     "                 Decode(a.�Ա�, '1', '��', '2', 'Ů', '9', 'δ֪', '������') �Ա�, a.����," & vbCrLf & _
                     "                 Nvl(a.������, Decode(a.������Դ, 1, a.�����, 2, a.סԺ��, 3, a.������, 4, a.������, Decode(a.�Һŵ�, Null, a.�շѵ���, a.�Һŵ�))) ������,a.�Һŵ�" & vbCrLf & _
                     "          From ���鱨���¼ A" & vbCrLf & _
                     "          Where Nvl(a.�Ƿ��ʿر걾, 0) = 0 and a.HIS����ID is not null And a.����ʱ�� Between [1] And [2] "


12            If Trim(Me.cboDept.Text) <> "00-���п���" Then
13                strSQL = strSQL & " and (a.�������=[3] or a.������� is null)"
14            End If

15            If Trim(Me.cbodor.Text) <> "00-����" Then
16                strSQL = strSQL & " and (a.������=[4] or a.������ is null)"
17            End If

18            Select Case Trim(cboDiseases.Text)
              Case "����"

19            Case "��Ⱦ��"
20                strSQL = strSQL & " and nvl(a.�Ƿ�Ⱦ��,0)=1"
21            Case "�Ǵ�Ⱦ��"
22                strSQL = strSQL & " and nvl(a.�Ƿ�Ⱦ��,0)=0"
23            End Select

24            If chkVerifyDate.value = 1 Then
25                strSQL = strSQL & " and a.���ʱ�� between [5] and [6]"
26            End If

27            strSQL = strSQL & " ) Where ��� = 1"

28            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����б�", CDate(Format(dtpS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpE.value, "yyyy/mm/dd 23:59:59")), Trim(Me.cbodor.Text), Trim(Me.cboDept.Text), CDate(Format(dtpVS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpVE.value, "yyyy/mm/dd 23:59:59")))
29            If rsPait Is Nothing Then
30                Set rsPait = gobjLiscomlib.CopyNewRec(rsTmp, True)
31            End If
32            Do While Not rsTmp.EOF
33                With rsPait
34                    .Filter = "HIS����ID=" & rsTmp("HIS����ID")
35                    If .RecordCount <= 0 Then
36                        .AddNew
37                        For j = 0 To .Fields.Count - 1
38                            .Fields(j).value = rsTmp.Fields(j).value
39                        Next
40                    End If
41                End With
42                rsTmp.MoveNext
43            Loop

              '���ϰ���ȥ���Ҳ���
44            strSQL = "Select *" & vbCrLf & _
                     "   From (Select row_number() over(Partition By a.����ID Order By a.����ʱ�� Desc) ���, a.����ID his����ID, a.����," & vbCrLf & _
                     "                 Decode(a.�Ա�, '1', '��', '2', 'Ů', '9', 'δ֪', '������') �Ա�, a.����, Decode(a.������Դ, 1, a.�����, 2, a.סԺ��, a.�Һŵ�) ������,a.�Һŵ�" & vbCrLf & _
                     "          From ����걾��¼ A" & vbCrLf & _
                     "          Where Nvl(a.�Ƿ��ʿ�Ʒ, 0) = 0 and a.����ID is not null And a.����ʱ�� Between [1] And [2] "

45            If Trim(Me.cboDept.Text) <> "00-���п���" Then
46                strSQL = strSQL & " and (a.�������ID=[3] or a.�������ID is null)"
47            End If

48            If Trim(Me.cbodor.Text) <> "00-����" Then
49                strSQL = strSQL & " and (a.������=[4] or a.������ is null)"
50            End If

51            Select Case Trim(cboDiseases.Text)
              Case "��Ⱦ��"
52                strSQL = strSQL & " and 0=1"
53            End Select

54            If chkVerifyDate.value = 1 Then
55                strSQL = strSQL & " and a.���ʱ�� between [5] and [6]"
56            End If

57            strSQL = strSQL & " ) Where ��� = 1"

58            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�����б�", CDate(Format(dtpS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpE.value, "yyyy/mm/dd 23:59:59")), Trim(Me.cbodor.Text), Trim(Me.cboDept.ItemData(cboDept.ListIndex)), CDate(Format(dtpVS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpVE.value, "yyyy/mm/dd 23:59:59")))
59            If rsPait Is Nothing Then
60                Set rsPait = gobjLiscomlib.CopyNewRec(rsTmp, True)
61            End If
62            Do While Not rsTmp.EOF
63                With rsPait
64                    .Filter = "HIS����ID=" & rsTmp("HIS����ID")
65                    If .RecordCount <= 0 Then
66                        .AddNew
67                        For j = 0 To .Fields.Count - 1
68                            .Fields(j).value = rsTmp.Fields(j).value
69                        Next
70                    End If
71                End With
72                rsTmp.MoveNext
73            Loop

74            If Not rsPait Is Nothing Then
75                rsPait.Filter = ""
76                If rsPait.RecordCount > 0 Then rsPait.MoveFirst
77            End If
78        End If
79        Call gobjLiscomlib.SetDataToVSF(vsfPaitList, rsPait)    '��������Ϣ���ص��б�

80        With vsfPaitList
81            .ColHidden(.ColIndex("HIS����ID")) = True
82            .ColHidden(.ColIndex("���")) = True
83            .ColHidden(.ColIndex("�Һŵ�")) = True

              'Ĭ��ѡ�е�һ��
84            If .Rows > 1 Then
85                Call vsfPaitList_AfterRowColChange(0, 0, 1, 0)
86            End If
87        End With

88        With Me.txtPaitKey
89            .SelStart = 0
90            .SelLength = Len(.Text)
91            .SetFocus
92        End With


93        Exit Sub
GetDeptPaits_Error:
94        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "ִ��(GetDeptPaits)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
95        Err.Clear

End Sub

Private Sub cboDept_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim strFind As String
    Dim lngS As Long
    Dim strTxt As String


    If KeyAscii = vbKeyReturn Then
        With cboDept
            strFind = UCase(Trim(.Text))
            '���������
            If IsNumeric(strFind) Then
                For i = 0 To .ListCount - 1
                    If .List(i) Like strFind & "-*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
            Else
                '���������
                For i = 0 To .ListCount - 1
                    lngS = InStr(.List(i), "[")
                    If lngS > 0 Then
                        strTxt = Mid(.List(i), lngS)
                    End If
                    If UCase(strTxt) = "[" & strFind & "]" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
                '�����Ʋ���
                For i = 0 To .ListCount - 1
                    If .List(i) Like "*" & strFind & "*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
            End If
        End With
    End If
End Sub

Private Sub cbodor_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim strFind As String
    Dim lngS As Long
    Dim strTxt As String
    If KeyAscii = vbKeyReturn Then
        With cbodor
            strFind = UCase(Trim(.Text))
            '���������
            If IsNumeric(strFind) Then
                For i = 0 To .ListCount - 1
                    If .List(i) Like strFind & "-*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
            Else
                '���������
                For i = 0 To .ListCount - 1
                    lngS = InStr(.List(i), "[")
                    If lngS > 0 Then
                        strTxt = Mid(.List(i), lngS)
                    End If
                    If UCase(strTxt) = "*" & strFind & "*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
                '�����Ʋ���
                For i = 0 To .ListCount - 1
                    If .List(i) Like "*" & strFind & "*" Then
                        .ListIndex = i
                        Exit Sub
                    End If
                Next
            End If
        End With
    End If
End Sub

Private Sub cbrMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case ConMenu_Browse_Find        '����
        Call GetDeptPaits
    Case ConMenu_Browse_Print       '��ӡδ��ӡ����
        Call PrintPaitReport(2, mlngPatientID, False)
    Case ConMenu_Browse_PrintAll    '��ӡ���б���
        Call PrintPaitReport(2, mlngPatientID, True)
    Case ConMenu_Browse_PrintView   'Ԥ��Ϊ��ӡ����
        Call PrintPaitReport(1, mlngPatientID, False)
    Case ConMenu_Browse_PrintViewAll    'Ԥ������
        Call PrintPaitReport(1, mlngPatientID, True)
    Case ConMenu_pop_Dept
        lblDept.Caption = "������ҡ�"
        InitDepts 0
    Case ConMenu_pop_DeptDistrict
        lblDept.Caption = "���벡����"
        InitDepts 1
    Case ConMenu_Appfor_ClincHelp   '���Ʋο�
        Call ShowClincHelp
    Case conMenu_Tool_PlugIn_Item + 1 To conMenu_Tool_PlugIn_Item + 99    '��ҹ���ִ��
'        Call ExePlugIn(Control.Parameter, mlngKey)
    Case ConMenu_Browse_Exit       '�˳�
        Unload Me
    End Select
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-30
'��    ��:  ������ӡ���˶�ݱ���
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Sub PrintPaitReport(ByVal bytType As Byte, ByVal lngPaitID As Long, ByVal blnPrintAll As Boolean)
          Dim objLisPrint As Object
          Dim objForm As Object
          
1         On Error GoTo PrintPaitReport_Error

2         If objLisPrint Is Nothing Then Set objLisPrint = CreateObject("zlPublicLIS.clsLis")
          '�ȴ�ӡ�°汨��
3         If Not objLisPrint Is Nothing Then
4             Call objLisPrint.Init(gcnHisOracle)
5         End If

          '���ش�ӡ����
6         Set objForm = objLisPrint.GetForm()

         '���ô�ӡ��Ӧ�Ĳ��˱���
7         Call objLisPrint.PrintLisReport(objForm, lngPaitID, mstrPatientGH, mlngPatientPage, 2, bytType, mblnDoctorShow, blnPrintAll)

8         Set objLisPrint = Nothing
          
          '��ӡ����΢���ﱨ��
9         Call beginPrint


10        Exit Sub
PrintPaitReport_Error:
11        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "ִ��(PrintPaitReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
12        Err.Clear
End Sub

'----------����΢���ﱨ�洦��
Private Sub beginPrint()
    Dim strFileSource As String
    Dim lng����ID As String
    Dim strArr() As String
    Dim i As Integer
    
    If mstrThirdReport <> "" Then
        If Left(mstrThirdReport, 4) = "<SP>" Then mstrThirdReport = Mid(mstrThirdReport, 5)
    Else
        Exit Sub
    End If
    strArr = Split(mstrThirdReport, "<SP>")
    For i = 0 To UBound(strArr)
        strFileSource = GetLisRptFile(strArr(i))
        lng����ID = Split(strArr(i), ";")(0)
        Call FunFastPrint(strFileSource, lng����ID)
    Next

End Sub

Private Sub FunFastPrint(ByVal strFile As String, ByVal lngRptID As Long)
'���ܣ�API���ÿ��ٴ�ӡPDF�ļ�
'������strFile �ļ�·��
    Dim RetVal As Long
    Dim strSQL As String
    Dim ShExInfo As SHELLEXECUTEINFO
    
     On Error GoTo errH
    With ShExInfo
        .cbSize = Len(ShExInfo)
        .fMask = &H40
        .hWnd = 0
        .lpVerb = "print"
        .lpFile = strFile
        .lpParameters = ""
        .lpDirectory = vbNullChar
        .nShow = 2
    End With
    RetVal = ShellExecuteEx(ShExInfo)
    If RetVal = 0 Then
        Exit Sub
    End If
'    strSQL = "Zl_ҽ����������_Print(" & lngRptID & ",0)"
'    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
   Exit Sub
errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
End Sub

Private Sub cbrMain_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    On Error Resume Next
    With picFilter
        .Left = Left
        .Top = Top
        .Width = Right - Left
    End With
    With picMain
        .Left = Left
        .Top = picFilter.Top + picFilter.Height
        .Width = Right - Left
        .Height = Bottom - .Top
    End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    mobjIDKind.object.ActiveFastKey
End Sub

Private Sub Form_Load()
      '���ܴ���������
          Dim cbrControl As CommandBarControl
          Dim cbrToolBar As CommandBar
          '-----------------------------------------------------
1         On Error GoTo Form_Load_Error

2         CommandBarsGlobalSettings.App = App
3         CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
4         CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
5         Me.cbrMain.VisualTheme = xtpThemeOffice2003
6         Me.cbrMain.Icons = frmPubIcons.imgPublic.Icons
7         With Me.cbrMain.Options
8             .ShowExpandButtonAlways = False
9             .ToolBarAccelTips = True
10            .AlwaysShowFullMenus = False
11            .IconsWithShadow = True    '����VisualTheme����Ч
12            .UseDisabledIcons = True
13            .LargeIcons = True
14            .SetIconSize True, 24, 24
15            .SetIconSize False, 16, 16
16        End With
17        Me.cbrMain.EnableCustomization False

          '-----------------------------------------------------
          '�˵�����
18        Me.cbrMain.ActiveMenuBar.Title = "�˵�"
19        Me.cbrMain.ActiveMenuBar.Visible = False
20        Set cbrToolBar = Me.cbrMain.Add("������", xtpBarTop)
21        cbrToolBar.ShowTextBelowIcons = False
22        cbrToolBar.EnableDocking xtpFlagStretched
23        With cbrToolBar.Controls

24            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Find, "����(&F5)"): cbrControl.BeginGroup = True
25            Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_Print, "��ӡδ��ӡ����(&F2)")
26            cbrControl.Style = xtpButtonIconAndCaption
27            With cbrControl.CommandBar.Controls
28                Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintAll, "��ӡ����  ")
29                Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintSet, "��ӡ����  ")
30                Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_unPrint, "���ô�ӡ  ")
31                cbrControl.Visible = InStr(mstrPrivs, "���������������ӡ����") > 0
32            End With
33            Set cbrControl = .Add(xtpControlSplitButtonPopup, ConMenu_Browse_PrintView, "Ԥ��δ��ӡ����")
34            cbrControl.Style = xtpButtonIconAndCaption
35            With cbrControl.CommandBar.Controls
36                Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_PrintViewAll, "Ԥ������  ")
37            End With
38            Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ClincHelp, "���Ʋο�")
39            Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Exit, "�˳�"): cbrControl.BeginGroup = True
40        End With

          '���������ť
41        Call CreatePlugInButton(cbrToolBar)

42        For Each cbrControl In cbrToolBar.Controls
43            If cbrControl.Type = xtpControlButton Then
44                cbrControl.Style = xtpButtonIconAndCaption
45            End If
46        Next

          '�����
47        With Me.cbrMain.KeyBindings
48            .Add 0, VK_F2, ConMenu_Browse_Print
49            .Add 0, VK_F5, ConMenu_Browse_Find
50        End With


          '��ʼ��IDKind
51        If mobjIDKind Is Nothing Then
52            Set mobjIDKind = NewControl(Me, "zlLisControl.ucLisIDKind", "ucLisIDKind", picIDKIND)
53            If mobjIDKind Is Nothing Then
54                Me.picIDKIND.Visible = False
55            End If
56            picIDKIND.BorderStyle = 0
57        End If

58        dtpE.value = gobjLiscomlib.comcurrdate
59        dtpS.value = dtpE.value - 7

60        Call gobjLiscomlib.vfgSetting(0, Me.vsfPaitList, "����,2000,1;�Ա�,800,1;����,800,1;������,1000,1")

61        txtRptCount.Text = Val(ComGetPara(Sel_Lis_DB, "���鱨��鿴����", 2500, 2500, "7"))

          '�Ƿ���ʾ��Ⱦ��ɸѡ��
62        cboDiseases.Enabled = InStr(mstrPrivs, "�鿴��Ⱦ������") > 0
63        Me.cboDiseases.ListIndex = 2

64        txtPaitKey.TabIndex = 0
65        cboDept.TabIndex = 1
66        cbodor.TabIndex = 2
67        dtpS.TabIndex = 3
68        dtpE.TabIndex = 4
69        dtpVS.TabIndex = 5
70        dtpVE.TabIndex = 6
71        cboDiseases.TabIndex = 7
72        txtRptCount.TabIndex = 8


73        Exit Sub
Form_Load_Error:
74        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "ִ��(Form_Load)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
75        Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    mstrPrivs = ""
    mstrThirdReport = ""
    
    For i = 1 To uclSampleReport.Count - 1
        Unload uclSampleReport(i)
    Next
    Set mobjIDKind = Nothing
    
    
    Call ComSetPara(Sel_Lis_DB, "���鱨��鿴����", Val(txtRptCount.Text), 2500, 2500)
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-26
'��    ��:  ����ѡ��Ĳ��˻�ȡ���˵ļ��鱨��
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Private Sub GetPaitReport(ByVal lngPaitID As Long)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsRpt As ADODB.Recordset
          Dim lngUclCount As Long
          Dim i As Integer
          Dim j As Integer
          Dim strWhere As String
          Dim strDept As String
          Dim strDor As String
          Dim lngS As Long
          Dim lngE As Long

          '�°汨��
1         On Error GoTo GetPaitReport_Error

2         gobjLiscomlib.ShowFlash "���ڼ��ر���,���Ժ�...", Me

          '��ж����һ�εĿؼ�
3         picScroll.Visible = False
4         For i = 1 To Me.uclSampleReport.Count - 1
5             Call uclSampleReport(i).UnloadCrl   'ж�ؿؼ���ʹ�õĶ���
6             Unload uclSampleReport(i)
7         Next


8         strSQL = "select * from (select a.ID,a.΢����,0 �������,a.���Ա���,25 �汾,a.���,a.��ע,a.����ʱ��,a.����ʱ�� from ���鱨���¼ A where a.HIS����ID=[1] and  a.����ʱ�� between [3] and [4] [����] and a.����� is not null order by a.����ʱ�� desc) where rownum<=[2]  "


9         If Trim(Me.cboDept.Text) <> "00-���п���" Then
10            strDept = Me.cboDept.Text
11            lngS = InStr(strDept, "-") + 1
12            lngE = InStr(strDept, "[")
13            strDept = Mid(strDept, lngS, lngE - lngS)
14            strWhere = strWhere & " and (a.�������=[5] or a.������� is null)"
15        End If

16        If Trim(Me.cbodor.Text) <> "00-����" Then
17            strDor = Me.cbodor.Text
18            lngS = InStr(strDor, "-") + 1
19            lngE = InStr(strDor, "[")
20            strDor = Mid(strDor, lngS, lngE - lngS)
21            strWhere = strWhere & " and (a.������=[6] or a.������ is null)"
22        End If

23        Select Case Trim(cboDiseases.Text)
          Case "��Ⱦ��"
24            strWhere = strWhere & " and nvl(a.�Ƿ�Ⱦ��,0)=1"
25        Case "�Ǵ�Ⱦ��"
26            strWhere = strWhere & " and nvl(a.�Ƿ�Ⱦ��,0)=0"
27        End Select

28        If chkVerifyDate.value = 1 Then
29            strSQL = strSQL & " and a.���ʱ�� between [7] and [8]"
30        End If

31        strSQL = Replace(strSQL, "[����]", strWhere)
32        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "���鱨���¼", lngPaitID, Val(txtRptCount.Text), CDate(Format(dtpS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpE.value, "yyyy/mm/dd 23:59:59")), strDept, strDor, CDate(Format(dtpVS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpVE.value, "yyyy/mm/dd 23:59:59")))
33        If rsRpt Is Nothing Then
34            Set rsRpt = gobjLiscomlib.CopyNewRec(rsTmp, True)
35        End If
36        Do While Not rsTmp.EOF
37            With rsRpt
38                .AddNew
39                For j = 0 To .Fields.Count - 1
40                    .Fields(j).value = rsTmp.Fields(j).value
41                Next
42            End With
43            rsTmp.MoveNext
44        Loop

          '�ϰ汨��
45        strSQL = "select * from (select a.ID,a.΢����걾 ΢����,a.������ �������,1 ���Ա���,10 �汾,'' ���,'' ��ע,a.����ʱ��,a.����ʱ��  from ����걾��¼ A where a.����ID=[1] and  a.����ʱ�� between [3] and [4] [����] and a.����� is not null order by a.����ʱ�� desc) where rownum<=[2]"

46        strWhere = ""
47        If Trim(Me.cboDept.Text) <> "00-���п���" Then
48            strWhere = " and (a.�������ID=[5] or a.�������ID is null)"
49        End If

50        If Trim(Me.cbodor.Text) <> "00-����" Then
51            strWhere = strWhere & " and (a.������=[6] or a.������ is null)"
52        End If


53        If chkVerifyDate.value = 1 Then
54            strWhere = strWhere & " and a.���ʱ�� between [7] and [8]"
55        End If

56        Select Case Trim(cboDiseases.Text)
          Case "��Ⱦ��"
57            strWhere = strWhere & " and 0=1"
58        End Select

59        strSQL = Replace(strSQL, "[����]", strWhere)

60        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "���鱨���¼", lngPaitID, Val(txtRptCount.Text), CDate(Format(dtpS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpE.value, "yyyy/mm/dd 23:59:59")), Trim(Me.cbodor.Text), Trim(Me.cboDept.ItemData(cboDept.ListIndex)), CDate(Format(dtpVS.value, "yyyy/mm/dd 00:00:00")), CDate(Format(dtpVE.value, "yyyy/mm/dd 23:59:59")))

61        If rsRpt Is Nothing Then
62            Set rsRpt = gobjLiscomlib.CopyNewRec(rsTmp, True)
63        End If
64        Do While Not rsTmp.EOF
65            With rsRpt
66                .AddNew
67                For j = 0 To .Fields.Count - 1
68                    .Fields(j).value = rsTmp.Fields(j).value
69                Next
70            End With
71            rsTmp.MoveNext
72        Loop
73        rsRpt.Filter = ""
74        If rsRpt.RecordCount > 0 Then
75            rsRpt.MoveFirst
76            picScroll.Visible = True
77        End If
78        rsRpt.Sort = "����ʱ�� desc"
79        mstrThirdReport = ""
80        Do While Not rsRpt.EOF
81            If lngUclCount >= Val(txtRptCount.Text) Then Exit Do
82            Call ShowPaitReport(Me, mblnDoctorShow, lngPaitID, Val(rsRpt("ID") & ""), Val(rsRpt("�汾") & ""), Val(rsRpt("΢����") & ""), Val(rsRpt("���Ա���") & ""), rsRpt("���") & "", rsRpt("��ע") & "", Val(rsRpt("�������") & ""), lngUclCount, CDate(Format(rsRpt("����ʱ��") & "", "yyyy/mm/dd hh:mm:ss")))
83            rsRpt.MoveNext
84        Loop


          '���ù�����
85        Me.vsfScroll.Rows = picScroll.Height / 225

86        gobjLiscomlib.StopFlash

87        Exit Sub
GetPaitReport_Error:
88        gobjLiscomlib.StopFlash
89        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmShowSampleReport", "ִ��(GetPaitReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
90        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-07-27
'��    ��:  ��ʾ����
'��    ��:
'           objFrm          ���ô���
'           mblnDoctorShow  �Ƿ���ҽ��վ����
'           lngPaintID      ����ID
'           lngSampleID     �걾ID
'           intVersion      ����汾��25=�°�LIS��10=�ϰ�LIS
'           intSampleType   �Ƿ���΢���ﱨ�棬0=��ͨ���棬1=΢���ﱨ��
'           intPositive     �������ͣ�1=ҩ�����棬3=PDF���棬����=���Ա���
'           strDiagnosis    ���
'           strResult       ��ע
'           intCount        �ϰ�LIS�������
'           dteSampleTime   �걾����ʱ��

'��    ��:
'��    ��:
'����Ӱ��:
'����ע��:
'---------------------------------------------------------------------------------------
Public Function ShowPaitReport(objFrm As Object, ByVal blnDoctorShow As Boolean, ByVal lngPaintID As Long, ByVal lngSampleID As Long, ByVal intVersion As Long, _
                                ByVal intSampleType As Integer, Optional ByVal intPositive As Integer, _
                                Optional ByVal strDiagnosis As String, Optional ByVal strResult As String, _
                                Optional ByVal intCount As Integer, Optional ByRef lngUclCount As Long, Optional ByVal dteSampleTime As Date) As Long
          Dim lngHeight As Long
          Dim strThirdReport As String

1         On Error GoTo ShowPaitReport_Error

2         If lngUclCount = 0 Then
              '���ر���
3             lngHeight = uclSampleReport(lngUclCount).GetSampleReport(Me, blnDoctorShow, lngPaintID, lngSampleID, intVersion, intSampleType, intPositive, strDiagnosis, strResult, intCount, dteSampleTime, mstrPrivs, strThirdReport)
4         Else
5             Load uclSampleReport(lngUclCount)
6             lngHeight = uclSampleReport(lngUclCount).GetSampleReport(Me, blnDoctorShow, lngPaintID, lngSampleID, intVersion, intSampleType, intPositive, strDiagnosis, strResult, intCount, dteSampleTime, mstrPrivs, strThirdReport)
7         End If
8         If strThirdReport <> "" Then
9             mstrThirdReport = mstrThirdReport & "<SP>" & strThirdReport
10        End If
          '���Զ��屨��ؼ�����vsf�У��Ա���֮����
          
11        With uclSampleReport(lngUclCount)
12            If lngUclCount = 0 Then
13                .Left = 0
14                .Top = 0
15                .Width = picScroll.Width
16                .Height = lngHeight
17                Me.picScroll.Height = .Height
18            Else
19                .Left = 0
20                .Top = uclSampleReport(lngUclCount - 1).Top + uclSampleReport(lngUclCount - 1).Height + 200
21                .Width = picScroll.Width
22                .Height = lngHeight
23                .Visible = True
24                Me.picScroll.Height = Me.picScroll.Height + .Height + 200
25            End If
26        End With

27        lngUclCount = lngUclCount + 1


28        Exit Function
ShowPaitReport_Error:
29        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "ִ��(ShowPaitReport)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
30        Err.Clear

End Function

Private Sub fraWE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim LeftColl As New Collection, Rightcoll As New Collection
    If Button = vbLeftButton Then
        LeftColl.Add Me.picPaitList
        Rightcoll.Add Me.picPaitReport
        Call SplitWE(LeftColl, Me.fraWE, Rightcoll, X, 1000)
        Set LeftColl = Nothing
        Set Rightcoll = Nothing
    End If
End Sub

Private Sub lblDept_Click()
    Dim objPopup As CommandBar
    Dim cbrControl As CommandBarControl
    Dim vPoint As POINTAPI
    On Error Resume Next

    Set objPopup = Me.cbrMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_Dept, "�������")
        Set cbrControl = .Add(xtpControlButton, ConMenu_pop_DeptDistrict, "���벡��")
    End With
    vPoint.X = lblDept.Left / Screen.TwipsPerPixelX
    vPoint.Y = (lblDept.Top + lblDept.Height + 30) / Screen.TwipsPerPixelY
    ClientToScreen picFilter.hWnd, vPoint

    objPopup.ShowPopup , vPoint.X * Screen.TwipsPerPixelX, vPoint.Y * Screen.TwipsPerPixelY
End Sub

Private Sub mobjIDKind_ObjectEvent(Info As EventInfo)
    Select Case Info
        Case "ReadCard"
                txtPaitKey.Text = IIf(Info.EventParameters(1).value = 0, Info.EventParameters(0).value, Info.EventParameters(1).value)
                Call GetDeptPaits
                txtPaitKey.SelStart = 0
                txtPaitKey.SelLength = Len(txtPaitKey.Text)
    End Select
End Sub

Private Sub picMain_Resize()
    On Error Resume Next
    With picPaitList
        .Left = 0
        .Top = 0
        .Height = Me.picMain.Height
    End With
    With fraWE
        .Left = picPaitList.Left + picPaitList.Width
        .Top = 0
        .Height = Me.picMain.Height
    End With
    With picPaitReport
        .Left = fraWE.Left + fraWE.Width
        .Top = 0
        .Width = Me.picMain.Width - .Left
        .Height = Me.picMain.Height
    End With
End Sub

Private Sub picPaitList_Resize()
    On Error Resume Next
    With vsfPaitList
        .Left = 0
        .Top = 0
        .Width = Me.picPaitList.Width
        .Height = Me.picPaitList.Height
    End With
End Sub

Private Sub picPaitReport_Resize()
    On Error Resume Next
    With vsfScroll
        .Left = 0
        .Top = 0
        .Width = Me.picPaitReport.Width
        .Height = Me.picPaitReport.Height
    End With
    With picScroll
        .Left = 0
        .Top = -vsfScroll.TopRow * vsfScroll.RowHeight(0)
        .Width = Me.picPaitReport.Width - 300
    End With
End Sub

Private Sub picScroll_Resize()
    Dim i As Integer
    
    For i = 0 To uclSampleReport.Count - 1
        With uclSampleReport(i)
            .Left = 0
            .Width = Me.picScroll.Width
        End With
    Next
End Sub

Private Sub txtPaitKey_GotFocus()
    With Me.txtPaitKey
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub txtPaitKey_KeyPress(KeyAscii As Integer)
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim strFind As String
          Dim strWhere As String


1         On Error GoTo txtPaitKey_KeyPress_Error

2         If KeyAscii = vbKeyReturn Then
              'ͨ����������ݲ��Ҳ���ID
3             strFind = Trim(txtPaitKey.Text)
4             If (Left(strFind, 1) = "A" Or Left(strFind, 1) = "-") And IsNumeric(Mid(strFind, 2)) Then    '����ID
5                 strWhere = " and a.HIS����ID = [2] "
6                 strFind = Mid(strFind, 2)
7             ElseIf (Left(strFind, 1) = "B" Or Left(strFind, 1) = "+") And IsNumeric(Mid(strFind, 2)) Then    'סԺ��
8                 strWhere = " and a.סԺ�� = [1] "
9                 strFind = Mid(strFind, 2)
10            ElseIf (Left(strFind, 1) = "D" Or Left(strFind, 1) = "*") And IsNumeric(Mid(strFind, 2)) Then    '�����
11                strWhere = " and a.����� = [1] "
12                strFind = Mid(strFind, 2)
13            ElseIf Left(strFind, 1) = "G" Or Left(strFind, 1) = "." Then    '�Һŵ�
14                strWhere = " and a.�Һŵ� = [1] "
15            ElseIf Left(strFind, 1) = "/" Then    '�շѵ��ݺ�
16                strWhere = " and a.�շѵ��� = [1] "
17            End If
18            strSQL = "      Select His����id" & vbNewLine & _
                     "       From �����������" & vbNewLine & _
                     "       Where His����id = [2] " & vbNewLine & _
                     "       Union All" & vbNewLine & _
                     "       Select His����id" & vbNewLine & _
                     "       From �����������" & vbNewLine & _
                     "       Where סԺ�� = [1] " & vbNewLine & _
                     "       Union All" & vbNewLine & _
                     "       Select His����id" & vbNewLine & _
                     "       From �����������" & vbNewLine & _
                     "       Where ����� = [1] " & vbNewLine & _
                     "       Union All" & vbNewLine & _
                     "       Select His����id" & vbNewLine & _
                     "       From �����������" & vbNewLine & _
                     "       Where �Һŵ� = [1]" & vbNewLine & _
                     "       Union All" & vbNewLine & _
                     "       Select His����id" & vbNewLine & _
                     "       From �����������" & vbNewLine & _
                     "       Where �������� = [1]"
19            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����ID", strFind, Val(strFind))
20            If Not rsTmp.EOF Then
21                txtPaitKey.Text = rsTmp("HIS����ID")
22                Call GetDeptPaits   '��ͨ��ID���Ҳ���
23            End If
24        End If


25        Exit Sub
txtPaitKey_KeyPress_Error:
26        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "ִ��(txtPaitKey_KeyPress)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear
End Sub

Private Sub txtRptCount_Change()
    upd.value = Val(txtRptCount.Text)
End Sub

Private Sub txtRptCount_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub upd_DownClick()
    txtRptCount.Text = upd.value
End Sub

Private Sub upd_UpClick()
    txtRptCount.Text = upd.value
End Sub

Private Sub vsfPaitList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    With Me.vsfPaitList
        If NewRow < 1 Then Exit Sub
        If .ColIndex("HIS����ID") < 0 Or .ColIndex("�Һŵ�") < 0 Then Exit Sub
        mlngPatientID = Val(.TextMatrix(NewRow, .ColIndex("HIS����ID")))
        mstrPatientGH = .TextMatrix(NewRow, .ColIndex("�Һŵ�"))
        If mstrPatientGH = "0" Then mstrPatientGH = ""
        Call GetPaitReport(Val(.TextMatrix(NewRow, .ColIndex("HIS����ID"))))
    End With
End Sub

Private Sub vsfScroll_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Me.picScroll.Top = -vsfScroll.TopRow * vsfScroll.RowHeight(0)
End Sub


Private Function GetDeptDor(Optional ByVal lngDeptID As Long) As ADODB.Recordset
      '����           ������һ������ض�Ӧ��ҽ����¼��
      '����
      '               lngDeptID ����ID����ID
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset

1         On Error GoTo GetDeptDor_Error

2         strSQL = "Select distinct b.id, b.���,b.����,b.����" & vbNewLine & _
                   "From ������Ա A, ��Ա�� B, ���ű� C,��Ա����˵�� D" & vbNewLine & _
                   "Where A.��Աid = B.Id And A.����id = C.Id And b.id=D.��ԱID And (C.����ʱ�� Is Null Or C.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) "
3         If lngDeptID <> 0 Then
4             strSQL = strSQL & "and c.id = [1] "
5         End If
          
6         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "LIS", lngDeptID)
7         With cbodor
8             .Clear
9             .AddItem "00-����"
10            .ItemData(.NewIndex) = 0
11            Do Until rsTmp.EOF
12                .AddItem rsTmp!��� & "-" & rsTmp!���� & "[" & rsTmp!���� & "]"
13                .ItemData(.NewIndex) = rsTmp!ID
14                rsTmp.MoveNext
15            Loop
16            If .ListCount > 0 Then .ListIndex = 0

17        End With


18        Exit Function
GetDeptDor_Error:
19        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "ִ��(GetDeptDor)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
20        Err.Clear

End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2019-04-19
'��    ��:  ��ʾ���Ʋο�
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Private Sub ShowClincHelp()
          Dim lngSampleID As Long
          Dim lngVer As Long

1         On Error GoTo ShowClincHelp_Error

'2         With Me.vsfLeft
'3             If .Row < 1 Then
'4                 MsgBox "��ѡ��һ�ݱ���", vbInformation, gSysInfo.AppName
'5                 Exit Sub
'6             End If
'7             If Val(.TextMatrix(.Row, .ColIndex("ID"))) = 0 Then
'8                 MsgBox "��ѡ��һ�ݱ���", vbInformation, gSysInfo.AppName
'9                 Exit Sub
'10            End If
'11            lngSampleID = Val(.TextMatrix(.Row, .ColIndex("ID")))
'12            lngVer = Val(.TextMatrix(.Row, .ColIndex("�汾")))
'13        End With
'
'14        Call funShowClincHelp(Me, lngSampleID, lngVer)


15        Exit Sub
ShowClincHelp_Error:
16        Call gobjLiscomlib.writeErrLog("zl9LisInsideComm", "frmPaitReport", "ִ��(ShowClincHelp)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear

End Sub


