VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSampleSendCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ͼ�˶�"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10470
   Icon            =   "frmSampleSendCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame FraAdvice 
      Caption         =   "ҽ����Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5445
      Left            =   60
      TabIndex        =   12
      Top             =   1680
      Width           =   10365
      Begin VB.CheckBox chkAll 
         Caption         =   "ȫѡ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8610
         TabIndex        =   15
         Top             =   420
         Width           =   870
      End
      Begin VB.TextBox txtCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1575
         TabIndex        =   4
         Top             =   375
         Width           =   3090
      End
      Begin VB.ComboBox cboCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   375
         Width           =   1440
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   4470
         Left            =   60
         TabIndex        =   13
         Top             =   840
         Width           =   10260
         _cx             =   18098
         _cy             =   7885
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
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
         Editable        =   2
         ShowComboButton =   0
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
      Begin VB.Label lblRefresh 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F56C58&
         Height          =   240
         Left            =   9615
         MouseIcon       =   "frmSampleSendCheck.frx":08CA
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   405
         Width           =   510
      End
      Begin VB.Label lblInfo 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   5265
         TabIndex        =   14
         Top             =   405
         Width           =   2955
      End
   End
   Begin VB.Frame fraPerson 
      Caption         =   "Ա����Ϣ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   10365
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3255
         TabIndex        =   2
         Top             =   945
         Width           =   720
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   855
         TabIndex        =   1
         Top             =   945
         Width           =   1605
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   855
         MaxLength       =   20
         TabIndex        =   0
         Top             =   405
         Width           =   3120
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "ȷ��(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9135
         TabIndex        =   5
         Top             =   405
         Width           =   1035
      End
      Begin VB.CommandButton cmdCnacel 
         Caption         =   "ȡ��(&C)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9135
         TabIndex        =   6
         Top             =   945
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   330
         TabIndex        =   11
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2715
         TabIndex        =   10
         Top             =   1005
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   315
         TabIndex        =   9
         Top             =   465
         Width           =   480
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   7125
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmSampleSendCheck.frx":0BD4
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13864
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   9750
      Top             =   -270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSampleSendCheck.frx":1468
            Key             =   "����"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSampleSendCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnCbo As Boolean
Private mintSendCount As Integer    '��ǰ��Ա�ͼ�걾��
Private mintCurrentCount As Integer    '��ǰ�걾��
Private mintCheckCount As Integer    '��ѡ�걾��
Private mintDays As Integer  '�����ѯ����
Private mstrSDate As String  '��ʼʱ��
Private mstrEDate As String  '����ʱ��
Private mstrAdvice As String    '�Ѻ˶�ҽ��

Private Sub cboCode_Click()
    If mblnCbo Then txtCode.SetFocus
End Sub

Private Sub chkAll_Click()
    Dim i As Integer
    With vsfList
        If .Rows > 1 Then
            For i = 1 To .Rows - 1
                .Cell(flexcpChecked, i, .ColIndex("ѡ��"), i, .ColIndex("ѡ��")) = chkAll.value
            Next
        End If
    End With
End Sub

Private Sub cmdCnacel_Click()
    Unload Me
End Sub

Public Function ShowMe(ByVal frmParent As Form, intType As Integer, intDays As Integer) As Boolean
    mintDays = intDays
    Me.Show vbModal, frmParent
End Function

Private Sub cmdOK_Click()
          Dim lngRow As Long

1         On Error GoTo cmdOK_Click_Error

2         If txtInfo(0).Tag = "" Then
3             MsgBox "����ȷ���ͼ�Ա����Ϣ��", vbInformation, "������Ϣ"
4             txtInfo(0).SetFocus
5             Exit Sub
6         End If

7         If mintCurrentCount = 0 Then
8             MsgBox "����ɨ��Ҫ����Ҫ�ͼ�˶Եı걾��", vbInformation, "������Ϣ"
9             txtCode.SetFocus
10            Exit Sub
11        End If

12        If SaveSampleNum Then
13            MsgBox "�ͼ�˶Գɹ���", vbInformation, "������Ϣ"
14            mintSendCount = 0
15            mintCheckCount = 0

16            With vsfList
17                For lngRow = .Rows - 1 To 1 Step -1
18                    If .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"), lngRow, .ColIndex("ѡ��")) = 1 Then
19                        .RemoveItem lngRow
20                    End If
21                Next
22            End With

23            txtCode.Text = ""
24            txtCode.SetFocus
25        End If

26        Call ShowInfo


27        Exit Sub
cmdOK_Click_Error:
28        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "ִ��(cmdOK_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
29        Err.Clear
End Sub

Private Function SaveSampleNum() As Boolean
      '��¼����˶�����
          Dim strSQL As String
          Dim rsTemp As Recordset
          Dim strAdvice As String
          Dim strYes As String
          Dim strNO As String
          Dim lngRow As Long
          Dim strMsg As String
          Dim intCount As Integer
          Dim strBatchNO As String
          Dim rsSampleCode As ADODB.Recordset
          Dim strSendAdivce As String
          Dim blnTre As Boolean
          Dim strErr As String

1         On Error GoTo SaveSampleNum_Error

2         lblInfo.Caption = ""

          '��ȡҽ������
3         With vsfList
4             If .Rows > 1 Then
5                 For lngRow = 1 To .Rows - 1
6                     If .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"), lngRow, .ColIndex("ѡ��")) = 1 Then
7                         strAdvice = strAdvice & ";" & .TextMatrix(lngRow, .ColIndex("ҽ��id")) & "^" & .TextMatrix(lngRow, .ColIndex("�Թܱ���"))
8                         strSendAdivce = strSendAdivce & "," & .TextMatrix(lngRow, .ColIndex("id")) & "," & .TextMatrix(lngRow, .ColIndex("ҽ��id"))
9                     End If
10                Next
11            End If
12        End With

13        If strAdvice <> "" Then
14            strAdvice = Mid(strAdvice, 2)
15            strSendAdivce = Mid(strSendAdivce, 2)
16        Else
17            MsgBox "��ѡ��Ҫ�˶Եļ�¼��", vbInformation, "������Ϣ"
18            Exit Function
19        End If


          '�ɼ�����վ
20        strSQL = "Select ��Աid, �Ǽ�����, �Ǽ���Ŀ From �걾�ͼ��¼ Where �˶�ʱ�� Is Null And ��Աid = [1] And �Ǽ�ʱ�� Between [2] And [3]"
21        Set rsTemp = ComOpenSQL(Sel_Lis_DB, strSQL, "�걾�ͼ��¼", Val(txtInfo(0).Tag), CDate(mstrSDate), CDate(mstrEDate))

22        If rsTemp.EOF Then
23            strSQL = "Zl_�걾�ͼ��¼_Edit(1," & Val(txtInfo(0).Tag) & ",'" & Trim(txtInfo(1).Text) & "'," & mintCheckCount & ",'" & strAdvice & "',To_Date('" & mstrSDate & "','yyyy-mm-dd hh24:mi:ss'),To_Date('" & mstrEDate & "','yyyy-mm-dd hh24:mi:ss'))"
24            Call ComExecuteProc(Sel_Lis_DB, strSQL, "�걾�ͼ��¼")
25            SaveSampleNum = True
26            SaveDBLog 18, 6, 0, "�ͼ�˶�", "�ͼ�걾�Ǽ�,�Ǽ���:" & Trim(txtInfo(1).Text), 1018, "�걾ǩ��"

27            gcnHisOracle.BeginTrans
              '�˶�֮���ͼ�걾
              '���ɷ�������
              '�ύ�ϰ�LIS����
28            strSQL = "select ����ҽ������_�걾��������.NEXTVAL from dual"
29            Set rsSampleCode = ComOpenSQL(Sel_His_DB, strSQL, "�걾��������", "")
30            strBatchNO = rsSampleCode(0) & ""
31            strSQL = "Zl_LisԤ������_�걾�ͳ�('" & strSendAdivce & "',0,'" & txtInfo(1).Text & "','" & strBatchNO & "')"
32            Call ComExecuteProc(Sel_His_DB, strSQL, "�걾�ͼ�")

              '�ύ�°�LIS����
33            If funSampleSendInfo(strSendAdivce, 0, txtInfo(1).Text, strErr) = False Then
34                gcnHisOracle.RollbackTrans
35                If strErr <> "" Then
36                    MsgBox strErr, vbInformation, "�ͼ�˶�"
37                End If
38                Exit Function
39            End If
40            gcnHisOracle.CommitTrans
41            blnTre = False

42        Else
43            MsgBox "��Ա��" & Trim(txtInfo(1).Text) & "������δ�˶Ե��ͼ��¼�����Ⱥ˶ԣ�", vbInformation, "������Ϣ"
44            SaveSampleNum = False
45        End If


46        Exit Function
SaveSampleNum_Error:
47        If blnTre Then gcnHisOracle.RollbackTrans
48        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "ִ��(SaveSampleNum)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
49        Err.Clear
End Function

Private Function CheckAdvice(ByVal strOldAdvice As String, ByVal strNewAdvice As String, strMsg As String, Optional blnModify As Boolean = False) As Boolean
      '�˶�ҽ������
          Dim arrOld As Variant
          Dim arrNew As Variant
          Dim strOld As String
          Dim strNew As String
          Dim strTemp As String
          Dim i As Integer

1         On Error GoTo CheckAdvice_Error

2         If strNewAdvice = "" Or strOldAdvice = "" Then Exit Function

3         arrNew = Split(strNewAdvice, ";")
4         arrOld = Split(strOldAdvice, ";")

5         If blnModify Then
              '�����˶�ʧ�ܵ��ͼ��¼
6             For i = 0 To UBound(arrOld)
7                 If InStr(";" & strNewAdvice & ";", ";" & arrOld(i) & ";") > 0 Then
                      '��������ҽ��
8                     strMsg = strMsg & ";" & arrOld(i)
9                 Else
                      '����������ҽ��
10                    strTemp = strTemp & ";" & arrOld(i)
11                End If
12            Next
13            If strMsg = "" Then Exit Function
14            strMsg = Mid(strMsg, 2) & "|" & Mid(strTemp, 2)
15        Else
              '�˶��ͼ��¼
16            If UBound(arrOld) >= UBound(arrNew) Then
                  'ȱ����Ŀ
17                For i = 0 To UBound(arrOld)
18                    If InStr(";" & strNewAdvice & ";", ";" & arrOld(i) & ";") > 0 Then
19                    Else
20                        strOld = strOld & ";" & arrOld(i)
21                    End If
22                Next
23                If strOld <> "" Then
24                    strOld = Mid(strOld, 2)
25                    strMsg = strOld & "|"
26                    Call GetMsg(strOld)
27                    strMsg = strMsg & "�˶�����" & UBound(arrNew) + 1 & "С�ڵ����ͼ�����" & UBound(arrOld) + 1 & vbCrLf & vbCrLf & "ȱ����Ŀ��" & strOld
28                    Exit Function
29                End If
30            Else
                  '������Ŀ
31                For i = 0 To UBound(arrNew)
32                    If InStr(";" & strOldAdvice & ";", ";" & arrNew(i) & ";") > 0 Then
33                    Else
34                        strNew = strNew & ";" & arrNew(i)
35                    End If
36                Next
37                If strNew <> "" Then
38                    strNew = Mid(strNew, 2)
39                    strMsg = strNew & "|"
40                    Call GetMsg(strNew)
41                    strMsg = strMsg & "�˶�����" & UBound(arrNew) + 1 & "�����ͼ�����" & UBound(arrOld) + 1 & vbCrLf & vbCrLf & "������Ŀ��" & strNew
42                    Exit Function
43                End If
44            End If
45        End If

46        CheckAdvice = True


47        Exit Function
CheckAdvice_Error:
48        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "ִ��(CheckAdvice)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
49        Err.Clear
End Function

Private Sub GetMsg(strMsg As String)
      '��ȡ�˶������ҽ����Ϣ
          Dim strSQL As String
          Dim rsTemp As Recordset
          Dim strҽ������ As String
          Dim str���� As String
          Dim str����id As String
          Dim str�걾���� As String
          Dim str�������� As String
          Dim str��������id As String
          Dim strִ�п���id As String
          Dim str�Թܱ��� As String
          Dim arrAdvice As Variant
          Dim i As Integer

1         On Error GoTo GetMsg_Error

2         If strMsg = "" Then Exit Sub

3         arrAdvice = Split(strMsg, ";")
4         strMsg = ""

5         For i = 0 To UBound(arrAdvice)
6             strSQL = "Select Distinct a.����id, a.���id ҽ��id, Decode(a.������־, 1, '����', '') ����, Decode(a.������Դ, 1, '����', 2, 'סԺ', 3, 'Ժ��', 4, '���') ������Դ," & vbNewLine & _
                     "                a.����, a.�Ա�, d.ҽ������, a.�걾��λ �걾����, b.��������, b.������, b.����ʱ��, b.�ͼ���, b.�걾�ͳ�ʱ�� �ͼ�ʱ��, a.��������id, a.ִ�п���id, c.�Թܱ���" & vbNewLine & _
                       "From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C, ����ҽ����¼ D, (Select Column_Value id From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) E" & vbNewLine & _
                       "Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And a.���id = d.Id And c.��� = 'C' And a.���id Is Not Null And b.����ʱ�� Is Null And" & vbNewLine & _
                     "      b.ִ��״̬ = 0 And a.���id = e.id And b.����ʱ�� Between [2] And [3] And c.�Թܱ��� = [4]"
7             Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "ҽ����Ϣ", Split(arrAdvice(i), "^")(0), CDate(mstrSDate), CDate(mstrEDate), Split(arrAdvice(i), "^")(1))

8             Do While Not rsTemp.EOF
9                 If strҽ������ <> "" & rsTemp!ҽ������ And str���� = "" & rsTemp!���� And str����id = "" & rsTemp!����ID And str�걾���� = "" & rsTemp!�걾���� And _
                     str�������� = "" & rsTemp!�������� And str��������id = "" & rsTemp!��������ID And strִ�п���id = "" & rsTemp!ִ�п���id And str�Թܱ��� = "" & rsTemp!�Թܱ��� Then
                      'ҽ���ϲ�
10                    strMsg = strMsg & "," & rsTemp!ҽ������
11                Else
12                    strҽ������ = "" & rsTemp!ҽ������
13                    str���� = "" & rsTemp!����
14                    str����id = "" & rsTemp!����ID
15                    str�걾���� = "" & rsTemp!�걾����
16                    str�������� = "" & rsTemp!��������
17                    str��������id = "" & rsTemp!��������ID
18                    strִ�п���id = "" & rsTemp!ִ�п���id
19                    str�Թܱ��� = "" & rsTemp!�Թܱ���

20                    strMsg = strMsg & vbCrLf & rsTemp!���� & "  " & rsTemp!ҽ������
21                End If

22                rsTemp.MoveNext
23            Loop
24        Next

25        Exit Sub
GetMsg_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "ִ��(GetMsg)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear
End Sub

Private Sub Form_Load()
    Dim intOnlyBarcode As Integer

    Me.Caption = "�걾�ͼ��¼"

    intOnlyBarcode = Val(ComGetPara(Sel_Lis_DB, "��֧������ɨ��¼������", 2500, 1018, 0))
    If intOnlyBarcode = 1 Then cboCode.Enabled = False

    mstrEDate = Format(Currentdate, "yyyy-mm-dd") & " 23:59:59"
    mstrSDate = Format(CDate(mstrEDate) - mintDays, "yyyy-mm-dd") & " 00:00:00"

    Call CreateCbo
    Call GetAdvice
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnCbo = False
    mintSendCount = 0
    mintCurrentCount = 0
    mintCheckCount = 0
    mstrAdvice = ""
End Sub

Private Sub lblRefresh_Click()
    vsfList.Rows = 1
    vsfList.Rows = 2
    mintCurrentCount = 0
    mintCheckCount = 0
    lblInfo.Caption = ""
    Call ShowInfo
End Sub

Private Sub txtCode_GotFocus()
    txtCode.SelStart = 0
    txtCode.SelLength = LenB(StrConv(Trim(txtCode.Text), vbFromUnicode))
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtCode.Text) <> "" Then
            Call GetAdvice(Trim(txtCode.Text))
            Call txtCode_GotFocus
        End If
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case cboCode.Text
        Case "����ɨ��", "�� �� ��", "ס Ժ ��"
            If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = vbKeyBack Then
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    txtInfo(Index).SelStart = 0
    txtInfo(Index).SelLength = LenB(StrConv(Trim(txtInfo(Index).Text), vbFromUnicode))
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    If Index = 0 And KeyCode = 13 Then
        For i = 1 To 2
            txtInfo(i).Text = ""
        Next
        txtInfo(0).Tag = ""
        If Trim(txtInfo(0).Text) <> "" Then
            Call GetPerson(Trim(txtInfo(0).Text))
            Call ShowInfo
        End If
    End If
End Sub

Private Sub ShowInfo()
'��ʾ�걾����
    stbThis.Panels(2).Text = ""
    If mintSendCount <> 0 Then
        stbThis.Panels(2).Text = "��Ա��" & Trim(txtInfo(1).Text) & "���ͼ�걾����" & mintSendCount & " "
    End If

    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "��ǰ�걾����" & mintCurrentCount & " "
    stbThis.Panels(2).Text = stbThis.Panels(2).Text & "�˶Ա걾����" & mintCheckCount
End Sub

Private Sub GetPerson(ByVal strCode As String)
      '��ȡ������Ϣ
          Dim rsTemp As Recordset
          Dim strSQL As String

1         On Error GoTo GetPerson_Error

2         mintSendCount = 0

3         strSQL = "Select a.id, a.����, a.�Ա� From ��Ա�� A Where  a.��� = [1]"

4         Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "��Ա��Ϣ", strCode)

5         If rsTemp.EOF Then
6             MsgBox "δ�ҵ���Ա��", vbInformation, "������Ϣ"
7             Call txtInfo_GotFocus(0)
8         Else
9             txtInfo(0).Tag = rsTemp!ID
10            txtInfo(1).Text = "" & rsTemp!����
11            txtInfo(2).Text = "" & rsTemp!�Ա�

              '��Ա�ͼ��¼���
12            strSQL = "Select ��Աid, �Ǽ�����, �Ǽ���Ŀ From �걾�ͼ��¼ Where �˶�ʱ�� Is Null And ��Աid = [1] And �Ǽ�ʱ�� Between [2] And [3]"
13            Set rsTemp = ComOpenSQL(Sel_Lis_DB, strSQL, "�걾�ͼ��¼", Val(txtInfo(0).Tag), CDate(mstrSDate), CDate(mstrEDate))

14            If rsTemp.EOF Then

15            Else
16                MsgBox "��Ա��" & Trim(txtInfo(1).Text) & "������δ�˶Ե��ͼ��¼�����Ⱥ˶ԣ�", vbInformation, "������Ϣ"
17            End If


18            txtCode.SetFocus
19        End If


20        Exit Sub
GetPerson_Error:
21        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "ִ��(GetPerson)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
22        Err.Clear
End Sub

Private Sub CreateCbo()
    cboCode.AddItem "����ɨ��"
    cboCode.AddItem "�� �� ��"
    cboCode.AddItem "ס Ժ ��"
    cboCode.AddItem "�� �� ��"
    cboCode.ListIndex = 0
    mblnCbo = True
End Sub

Private Sub GetAdvice(Optional ByVal strCode As String)
      '��ȡҽ����Ϣ
          Dim strSQL As String
          Dim rsTemp As Recordset
          Dim strTitle As String

1         On Error GoTo GetAdvice_Error

2         If cboCode.Text = "����ɨ��" Then
3             strSQL = "Select Distinct 1 ѡ��, a.����id,a.id, a.���id ҽ��id, Decode(a.������־, 1, '����', '') ����, Decode(a.������Դ, 1, '����', 2, 'סԺ', 3, 'Ժ��', 4, '���') ������Դ," & vbNewLine & _
                     "                a.����, a.�Ա�, d.ҽ������, a.�걾��λ �걾����, b.��������, b.������, b.����ʱ��, b.�ͼ���, b.�걾�ͳ�ʱ�� �ͼ�ʱ��, a.��������id, a.ִ�п���id, c.�Թܱ���" & vbNewLine & _
                       "From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C, ����ҽ����¼ D" & vbNewLine & _
                       "Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And a.���id = d.Id And c.��� = 'C' And a.���id Is Not Null And b.�걾�ͳ�ʱ�� is null And b.����ʱ�� Is Null And b.������ is not null And" & vbNewLine & _
                     "      b.ִ��״̬ = 0 And b.�������� = [1] And b.����ʱ�� Between [2] And [3]"
4         Else
5             strSQL = "Select Distinct 1 ѡ��, a.����id,a.id, a.���id ҽ��id, Decode(a.������־, 1, '����', '') ����, Decode(a.������Դ, 1, '����', 2, 'סԺ', 3, 'Ժ��', 4, '���') ������Դ," & vbNewLine & _
                     "                a.����, a.�Ա�, d.ҽ������, a.�걾��λ �걾����, b.��������, b.������, b.����ʱ��, b.�ͼ���, b.�걾�ͳ�ʱ�� �ͼ�ʱ��, a.��������id, a.ִ�п���id, c.�Թܱ���" & vbNewLine & _
                       "From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C, ����ҽ����¼ D, ������Ϣ E" & vbNewLine & _
                       "Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And a.���id = d.Id And a.����id = e.����id And c.��� = 'C' And a.���id Is Not Null And" & vbNewLine & _
                     "       b.�걾�ͳ�ʱ��  is null And b.����ʱ�� Is Null And b.������ is not null And b.ִ��״̬ = 0 And e.����id = [1] And b.����ʱ�� Between [2] And [3]"

6             If cboCode.Text = "�� �� ��" Then
7                 strSQL = Replace(strSQL, "e.����id = [1]", "e.����� = [1]")
8             ElseIf cboCode.Text = "ס Ժ ��" Then
9                 strSQL = Replace(strSQL, "e.����id = [1]", "e.סԺ�� = [1]")
10            ElseIf cboCode.Text = "�� �� ��" Then
11                strSQL = Replace(strSQL, "e.����id = [1]", "a.�Һŵ� = [1]")
12            End If
13        End If

14        If vsfList.TextMatrix(0, 1) = "" Then
              '�״ν��벻ȡ���ݣ����б��ʼ��

15            strSQL = strSQL & " And 1=0"
16            Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "ҽ����Ϣ", strCode, CDate(mstrSDate), CDate(mstrEDate))

17            Call vfgLoadFromRecord(vsfList, rsTemp, "", imgList)
18            With vsfList
19                .ExplorerBar = flexExSortShow
20                .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
21                .ColWidth(.ColIndex("ѡ��")) = 250: .TextMatrix(0, .ColIndex("ѡ��")) = ""
22                .ColWidth(.ColIndex("����")) = 250: .ColHidden(.ColIndex("����")) = False
23                .ColWidth(.ColIndex("������Դ")) = 1200: .ColHidden(.ColIndex("������Դ")) = False
24                .ColWidth(.ColIndex("����")) = 1200: .ColHidden(.ColIndex("����")) = False
25                .ColWidth(.ColIndex("�Ա�")) = 600: .ColHidden(.ColIndex("�Ա�")) = False
26                .ColWidth(.ColIndex("ҽ������")) = 3000: .ColHidden(.ColIndex("ҽ������")) = False
27                .ColWidth(.ColIndex("�걾����")) = 1200: .ColHidden(.ColIndex("�걾����")) = False
28                .ColWidth(.ColIndex("��������")) = 1600: .ColHidden(.ColIndex("��������")) = False
29                .ColWidth(.ColIndex("������")) = 1200: .ColHidden(.ColIndex("������")) = False
30                .ColWidth(.ColIndex("����ʱ��")) = 1400: .ColHidden(.ColIndex("����ʱ��")) = False
31                .ColWidth(.ColIndex("�ͼ���")) = 1200: .ColHidden(.ColIndex("�ͼ���")) = False
32                .ColWidth(.ColIndex("�ͼ�ʱ��")) = 1400: .ColHidden(.ColIndex("�ͼ�ʱ��")) = False

33                mintCurrentCount = 0
34                mintCheckCount = 0
35            End With

              '�Ѻ˶�ҽ��
36            strSQL = "Select �˶���Ŀ From �걾�ͼ��¼ Where �˶�ʱ�� Is Not Null And �Ǽ�ʱ�� Between [1] And [2]"
37            Set rsTemp = ComOpenSQL(Sel_Lis_DB, strSQL, "�걾�ͼ��¼", CDate(mstrSDate), CDate(mstrEDate))
38            mstrAdvice = ""
39            Do While Not rsTemp.EOF
40                mstrAdvice = mstrAdvice & ";" & rsTemp!�˶���Ŀ
41                rsTemp.MoveNext
42            Loop
43            If mstrAdvice <> "" Then mstrAdvice = Mid(mstrAdvice, 2)
44        Else
45            Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "ҽ����Ϣ", strCode, CDate(mstrSDate), CDate(mstrEDate))
46            If rsTemp.EOF Then MsgBox "δ�ҵ��걾����걾δ�����������ͼ죬���ѵǼǣ�", vbInformation, "������Ϣ": Exit Sub

47            With vsfList
48                Do While Not rsTemp.EOF
49                    If Not FindAdvice("" & rsTemp!ҽ��id, "" & rsTemp!�Թܱ���, "" & rsTemp!��������) Then    '�ж��Ƿ��Ѵ����б���
50                        If .TextMatrix(.Rows - 1, .ColIndex("ҽ��id")) <> "" Then
51                            .Rows = .Rows + 1
52                        End If

53                        If .TextMatrix(.Rows - 2, .ColIndex("ҽ������")) <> "" & rsTemp!ҽ������ And _
                             .TextMatrix(.Rows - 2, .ColIndex("����")) = "" & rsTemp!���� And _
                             .TextMatrix(.Rows - 2, .ColIndex("����id")) = "" & rsTemp!����ID And _
                             .TextMatrix(.Rows - 2, .ColIndex("�걾����")) = "" & rsTemp!�걾���� And _
                             .TextMatrix(.Rows - 2, .ColIndex("��������")) = "" & rsTemp!�������� And _
                             .TextMatrix(.Rows - 2, .ColIndex("��������id")) = "" & rsTemp!��������ID And _
                             .TextMatrix(.Rows - 2, .ColIndex("ִ�п���id")) = "" & rsTemp!ִ�п���id And _
                             .TextMatrix(.Rows - 2, .ColIndex("�Թܱ���")) = "" & rsTemp!�Թܱ��� Then
                              'ҽ���ϲ�
54                            .TextMatrix(.Rows - 2, .ColIndex("ҽ��id")) = .TextMatrix(.Rows - 2, .ColIndex("ҽ��id")) & "," & rsTemp!ҽ��id
55                            .TextMatrix(.Rows - 2, .ColIndex("ҽ������")) = .TextMatrix(.Rows - 2, .ColIndex("ҽ������")) & "," & rsTemp!ҽ������
56                            .Rows = .Rows - 1
57                        Else
58                            .TextMatrix(.Rows - 1, .ColIndex("id")) = "" & rsTemp!ID
59                            .TextMatrix(.Rows - 1, .ColIndex("����id")) = "" & rsTemp!����ID
60                            .TextMatrix(.Rows - 1, .ColIndex("ҽ��id")) = "" & rsTemp!ҽ��id
61                            .TextMatrix(.Rows - 1, .ColIndex("����")) = "" & rsTemp!����
62                            .TextMatrix(.Rows - 1, .ColIndex("������Դ")) = "" & rsTemp!������Դ
63                            .TextMatrix(.Rows - 1, .ColIndex("����")) = "" & rsTemp!����
64                            .TextMatrix(.Rows - 1, .ColIndex("�Ա�")) = "" & rsTemp!�Ա�
65                            .TextMatrix(.Rows - 1, .ColIndex("ҽ������")) = "" & rsTemp!ҽ������
66                            .TextMatrix(.Rows - 1, .ColIndex("�걾����")) = "" & rsTemp!�걾����
67                            .TextMatrix(.Rows - 1, .ColIndex("��������")) = "" & rsTemp!��������
68                            .TextMatrix(.Rows - 1, .ColIndex("������")) = "" & rsTemp!������
69                            .TextMatrix(.Rows - 1, .ColIndex("����ʱ��")) = "" & rsTemp!����ʱ��
70                            .TextMatrix(.Rows - 1, .ColIndex("�ͼ���")) = "" & rsTemp!�ͼ���
71                            .TextMatrix(.Rows - 1, .ColIndex("�ͼ�ʱ��")) = "" & rsTemp!�ͼ�ʱ��
72                            .TextMatrix(.Rows - 1, .ColIndex("��������id")) = "" & rsTemp!��������ID
73                            .TextMatrix(.Rows - 1, .ColIndex("ִ�п���id")) = "" & rsTemp!ִ�п���id
74                            .TextMatrix(.Rows - 1, .ColIndex("�Թܱ���")) = "" & rsTemp!�Թܱ���
75                            .Cell(flexcpChecked, .Rows - 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = "" & rsTemp!ѡ��

76                            If "" & rsTemp!���� = "����" Then
77                                .Cell(flexcpPicture, .Rows - 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = imgList.ListImages("����").ExtractIcon
78                            End If

79                            .TopRow = .Rows - 1
80                            .Row = .Rows - 1
81                        End If
82                    End If

83                    rsTemp.MoveNext
84                Loop
85            End With
86        End If


87        Exit Sub
GetAdvice_Error:
88        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "ִ��(GetAdvice)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
89        Err.Clear
End Sub

Private Sub vsfList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngRow As Long
    With vsfList
        If .Col = .ColIndex("ѡ��") Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If

        mintCheckCount = 0
        mintCurrentCount = 0
        If .Rows > 1 Then
            If .TextMatrix(1, 1) <> "" Then
                For lngRow = 1 To .Rows - 1
                    mintCurrentCount = mintCurrentCount + 1
                    If .Cell(flexcpChecked, lngRow, 0, lngRow, 0) = 1 Then
                        mintCheckCount = mintCheckCount + 1
                    End If
                Next
            End If
        End If

        Call ShowInfo
    End With
End Sub

Private Sub vsfList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Col = 0 And Row > 0 Then
        If vsfList.TextMatrix(Row, 1) <> "" Then
            If vsfList.Cell(flexcpChecked, Row, 0, Row, 0) = 1 Then
                mintCheckCount = mintCheckCount + 1
            Else
                mintCheckCount = mintCheckCount - 1
            End If
        End If
    End If
    Call ShowInfo
End Sub

Private Sub vsfList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        '���̰�ťDelete��ɾ��
        With vsfList
            If .TextMatrix(.Row, .ColIndex("ҽ��id")) <> "" And .Rows > 1 Then
                .RemoveItem .Row
                Call vsfList_AfterRowColChange(0, 1, 0, 1)
            End If
        End With
    End If
End Sub

Private Function FindAdvice(ByVal strAdvice As String, ByVal strNO As String, ByVal strCodeBar As String) As Boolean
      '�жϱ걾�Ƿ������б��У��Ƿ��Ѻ˶Թ�
          Dim lngRow As Long
          Dim arrAdvice As Variant
          Dim strTemp As String

1         On Error GoTo FindAdvice_Error

2         If mstrAdvice <> "" Then
3             arrAdvice = Split(mstrAdvice, ";")
4             For lngRow = 0 To UBound(arrAdvice)
5                 strTemp = arrAdvice(lngRow)
6                 If InStr("," & Split(strTemp, "^")(0) & ",", "," & strAdvice & ",") > 0 And Split(strTemp, "^")(1) = strNO Then
7                     FindAdvice = True
8                     Exit Function
9                 End If
10            Next
11        End If

12        With vsfList
13            If .Rows > 1 Then
14                For lngRow = 1 To .Rows - 1
15                    If InStr("," & .TextMatrix(lngRow, .ColIndex("ҽ��id")) & ",", "," & strAdvice & ",") > 0 And .TextMatrix(lngRow, .ColIndex("�Թܱ���")) = strNO And .TextMatrix(lngRow, .ColIndex("��������")) = strCodeBar Then
16                        FindAdvice = True
17                        .Row = lngRow
18                        Exit Function
19                    End If
20                Next
21            End If
22        End With


23        Exit Function
FindAdvice_Error:
24        Call WriteErrLog("zlPublicHisCommLis", "frmSampleSendCheck", "ִ��(FindAdvice)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
25        Err.Clear
End Function

