VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlanEditNew 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ҺŰ��ű༭"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11100
   Icon            =   "frmRegistPlanEditNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBaseBack 
      BorderStyle     =   0  'None
      Height          =   8940
      Left            =   0
      ScaleHeight     =   8940
      ScaleWidth      =   9210
      TabIndex        =   3
      Top             =   0
      Width           =   9210
      Begin VB.Frame Frame2 
         Caption         =   "������Ϣ"
         Height          =   1500
         Left            =   240
         TabIndex        =   23
         Top             =   120
         Width           =   8655
         Begin VB.TextBox txt�ű� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            MaxLength       =   5
            TabIndex        =   30
            Top             =   270
            Width           =   960
         End
         Begin VB.ComboBox cboItem 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4275
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   660
            Width           =   2115
         End
         Begin VB.ComboBox cboDoctor 
            Height          =   300
            Left            =   1050
            TabIndex        =   28
            Top             =   1065
            Width           =   2115
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            TabIndex        =   27
            Text            =   "cbo����"
            Top             =   660
            Width           =   2115
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�Һ�ʱ���뽨����"
            Height          =   195
            Left            =   4275
            TabIndex        =   26
            Top             =   1118
            Width           =   1845
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4275
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   270
            Width           =   2115
         End
         Begin VB.CheckBox chk��ſ��� 
            Caption         =   "��ſ���"
            Height          =   255
            Left            =   2130
            TabIndex        =   24
            Top             =   285
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�ű�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   615
            TabIndex        =   35
            Top             =   330
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   645
            TabIndex        =   34
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��Ŀ"
            Height          =   180
            Left            =   3870
            TabIndex        =   33
            Top             =   720
            Width           =   360
         End
         Begin VB.Label lblҽ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ժ��ҽ����"
            Height          =   180
            Left            =   120
            TabIndex        =   32
            Top             =   1125
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   3855
            TabIndex        =   31
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ӧ��ʱ��"
         Height          =   2550
         Left            =   240
         TabIndex        =   10
         Top             =   1800
         Width           =   8655
         Begin VB.OptionButton opt�� 
            Caption         =   "ÿ��(&D)"
            Height          =   315
            Left            =   225
            TabIndex        =   16
            Top             =   285
            Width           =   960
         End
         Begin VB.OptionButton opt�� 
            Caption         =   "ÿ��(&W)"
            Height          =   315
            Left            =   225
            TabIndex        =   15
            Top             =   630
            Width           =   930
         End
         Begin VB.ComboBox cbo�� 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   292
            Width           =   1110
         End
         Begin VB.CheckBox chk��Ч�� 
            Caption         =   "��Ч��"
            Height          =   195
            Left            =   255
            TabIndex        =   13
            Top             =   2115
            Width           =   855
         End
         Begin VB.TextBox txt��Լ 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4980
            MaxLength       =   5
            TabIndex        =   12
            Top             =   292
            Width           =   1215
         End
         Begin VB.TextBox txt�޺� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3045
            MaxLength       =   5
            TabIndex        =   11
            Top             =   292
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   1170
            TabIndex        =   17
            Top             =   2055
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   117112835
            CurrentDate     =   38091
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   3555
            TabIndex        =   18
            Top             =   2055
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   117112835
            CurrentDate     =   38091
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPlan 
            Height          =   1275
            Left            =   1155
            TabIndex        =   19
            Top             =   675
            Width           =   7200
            _cx             =   12700
            _cy             =   2249
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
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmRegistPlanEditNew.frx":000C
            ScrollTrack     =   0   'False
            ScrollBars      =   0
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
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   3315
            TabIndex        =   22
            Top             =   2115
            Width           =   180
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "��Լ"
            Height          =   180
            Left            =   4545
            TabIndex        =   21
            Top             =   345
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "�޺�"
            Height          =   180
            Left            =   2610
            TabIndex        =   20
            Top             =   352
            Width           =   360
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Ӧ������:"
         Height          =   4020
         Left            =   240
         TabIndex        =   4
         Top             =   4560
         Width           =   8640
         Begin VB.OptionButton opt���� 
            Caption         =   "������"
            Height          =   180
            Index           =   0
            Left            =   1020
            TabIndex        =   8
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ָ������"
            Height          =   180
            Index           =   1
            Left            =   2010
            TabIndex        =   7
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "��̬����"
            Height          =   180
            Index           =   2
            Left            =   3180
            TabIndex        =   6
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ƽ������"
            Height          =   180
            Index           =   3
            Left            =   4335
            TabIndex        =   5
            Top             =   0
            Width           =   1020
         End
         Begin MSComctlLib.ListView lvwDept 
            Height          =   3480
            Left            =   150
            TabIndex        =   9
            Top             =   300
            Width           =   8220
            _ExtentX        =   14499
            _ExtentY        =   6138
            View            =   2
            Arrange         =   2
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   0
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   9840
      TabIndex        =   2
      Top             =   1590
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   9840
      TabIndex        =   1
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9840
      TabIndex        =   0
      Top             =   1065
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   780
      Left            =   9240
      TabIndex        =   36
      Top             =   3720
      Width           =   1575
      _Version        =   589884
      _ExtentX        =   2778
      _ExtentY        =   1376
      _StockProps     =   64
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "Ժ��ҽ��"
         Index           =   0
      End
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "����Ԯҽ��"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmRegistPlanEditNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : frmRegistPlanEditNew
'Description :
'Author      : ��⸣
'Date        : 05-November-2012 14:31:24
'Comments    :�ҺŰ��Ź���,���ϰ汾�Ļ������޸�,����ʱ�κͰ�������Ϊһ��

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Option Explicit 'Ҫ���������


Private Enum mPageIndex
    EM_���� = 0
    EM_ʱ�� = 1
End Enum

Private Enum mPgIndex
    Pg_�ƻ����� = 1
    Pg_�ƻ�ʱ�� = 2
End Enum


Private mfrmTime As frmResistPlanTimeSet
Private mblnChangeByCode As Boolean '�Ƿ��Ǵ�����Ƹı���tabelpage����ʾҳ
Private mrsRegOldData As ADODB.Recordset '�������ݼ�����,ԭʼ�ҺŰ���
Private mrsRegNewData As ADODB.Recordset '�������ݼ����� �������ú�İ���
Private mrsRegHistory As ADODB.Recordset '���ιҺŵ����ݼ�

Private mlngModule As Long, mstrPrivs As String, mlngID As Long, mfrmMain As Form, mblnChange As Boolean
Private mrs���� As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mblnFirst As Boolean
Private mblnSucces As Boolean
Private mlngȱʡ�Һſ���ID  As Long '�ڹҺŰ���ʱ��������������ѡ��Ŀ��ҽ���ȱʡ
Private mrsʱ��� As ADODB.Recordset
Private mstr�����޸� As String '��ĳһ����߶���İ������Ƹ���
Public Enum RegEditType
    ed_���� = 0
    ed_�޸� = 1
    ed_���� = 2
End Enum
Private mEditType As RegEditType
Private mstr����ID As String
Private mblnCboClick As Boolean     '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
Private mblnOnlyԺ��ҽ�� As Boolean '��ֻ����Ժ��ҽ��


Private Type PlanInfo               '���Ÿı���Ҫ�Աȵ���Ϣ
    str�Ű�         As String       '�Ű���Ϣ
    str�޺�         As String       '�޺���Ϣ
    bln���         As Boolean      '�Ƿ���ſ���
    blnʱ���       As Boolean      '�Ƿ�������ʱ���
End Type

Private mPlanInfo     As PlanInfo 'ԭʼ�İ�����Ϣ  ��Ҫ���ڰ����޸�ʱ ��Ӧ��Ϣ�ıȽ�







Private Sub Form_Activate()
    Dim i As Integer
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitData = False Then Unload Me: Exit Sub
    If LoadCard = False Then Unload Me: Exit Sub
    Call cboDoctor_Validate(False)
    For i = 0 To opt����.UBound
        If opt����(i).Value Then Call opt����_Click(i): Exit For
    Next
    txt�ű�.SetFocus
End Sub

Private Sub Form_Load()
    Dim intTYPE As Integer
     Call InitPage
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    mblnFirst = True
    mblnOnlyԺ��ҽ�� = Val(zlDatabase.GetPara("ֻ����ѡԺ��ҽ��", glngSys, mlngModule, "0", , InStr(1, mstrPrivs, ";��������;") > 0, intTYPE)) = 1
    If mblnOnlyԺ��ҽ�� Then
        mnuViewDoctor(0).Checked = True
        mnuViewDoctor(1).Checked = False
    Else
        mnuViewDoctor(0).Checked = False
        mnuViewDoctor(1).Checked = True
    End If
    lblҽ��.Tag = IIf(mblnOnlyԺ��ҽ��, "0", "1")
    lblҽ��.Caption = IIf(mblnOnlyԺ��ҽ��, "Ժ��ҽ��", "ҽ��") & IIf(lblҽ��.Tag = "1", "��", "")
    lblҽ��.ToolTipText = IIf(mblnOnlyԺ��ҽ��, "ֻ��ѡԺ�ڽ���ҽ��", "����Ԯҽ��(���˿���ѡ��Ժ��ҽ���⣬������������Ԯҽ��)")
End Sub


Private Sub InitPage()
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:

    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_�ƻ�����, "�ƻ�����", picBaseBack.hWnd, 0)
    ObjItem.Tag = mPgIndex.Pg_�ƻ�����

    Set mfrmTime = New frmResistPlanTimeSet
    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_�ƻ�ʱ��, "ʱ������", mfrmTime.hWnd, 0)
    ObjItem.Tag = mPgIndex.Pg_�ƻ�ʱ��
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With cmdOK
        .Left = ScaleWidth - .Width - 100
        cmdCancel.Left = .Left
        cmdHelp.Left = .Left
    End With

    With tbPage
        .Top = 50
        .Height = ScaleHeight - 100
        .Left = 50
        .Width = cmdOK.Left - .Left - 100
    End With

End Sub

Public Function ShowEdit(ByVal frmMain As Form, ByVal EditType As RegEditType, _
    ByVal lngModule As Long, ByVal strPrivs As String, Optional lngID As Long = 0, _
    Optional lngȱʡ����ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-���õ�������
    '     EditType-�༭����
    '����:  `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   `   ``
    '����:
    '����:���˺�
    '����:2009-09-15 10:25:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs: mlngID = lngID: mlngȱʡ�Һſ���ID = lngȱʡ����ID
    mEditType = EditType: mblnSucces = False
    mblnChange = False
    mstr�����޸� = Get��Լ����(lngID)
    Me.Show 1, frmMain
    ShowEdit = mblnSucces

End Function

Private Function Get��Լ����(ByVal lng����ID As Long) As String
    '��ȡ�����޸ĵİ�������
    Dim strSQL As String
    Dim rsTmp   As ADODB.Recordset
    Dim strTmp  As String
    strSQL = "Select Decode(To_Char(A.ԤԼʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7'," & _
    "                             '����') As ���� " & vbCrLf & _
    "          From ���˹Һż�¼ A,�ҺŰ��š�B " & vbCrLf & _
    "        Where  A.�ű�=B.���� And A.��¼״̬ = 1 And b.ID = [1] And A.����ʱ�� > A.�Ǽ�ʱ�� And A.ԤԼʱ�� Is Not Null"

    If gintԤԼ���� = 0 Then
        strSQL = strSQL & " And A.ԤԼʱ�� > Sysdate "
    Else
        strSQL = strSQL & " And A.ԤԼʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTmp.EOF Then Exit Function

    Do While Not rsTmp.EOF
        If InStr(strTmp, Nvl(rsTmp!����)) < 0 Or strTmp = "" Then
            strTmp = strTmp & ";" & Nvl(rsTmp!����)
        End If
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then strTmp = strTmp & ";"
    Get��Լ���� = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:�ɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2009-09-15 13:14:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, rsTemp As ADODB.Recordset
    Dim bln�������� As Boolean

    Err = 0: On Error GoTo Errhand:
    gint�ų� = GetMaxLen

    strSQL = "" & _
    "   Select '    ' ʱ��� From dual Union All  " & _
    "   Select ʱ��� From ʱ���"
    Set mrsʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsPlan
        .Clear 1
        .Tag = .BuildComboList(mrsʱ���, "ʱ���")

        .ColComboList(1) = .BuildComboList(mrsʱ���, "ʱ���")
        For i = 2 To .Cols - 1
            .ColComboList(i) = .ColComboList(0)
        Next
    End With
    With cbo��
        Do While Not mrsʱ���.EOF
            cbo��.AddItem Nvl(mrsʱ���!ʱ���)
            mrsʱ���.MoveNext
        Loop
        .ListIndex = 0
    End With

   'ȡ�������ٴ�����
    Set mrs���� = GetDepartments("'�ٴ�'", "1,3", Not zlStr.IsHavePrivs(mstrPrivs, "���п���"))
    If mrs����.RecordCount = 0 Then
        MsgBox "�㲻�߱����õ��ٴ�������Ϣ����Ȩ�޲���,���ȵ����Ź����н������û���ϵͳ����Ա����Ȩ�ޣ�", vbInformation, gstrSysName
        Exit Function
    End If

    cbo����.Clear
    Do While Not mrs����.EOF
        cbo����.AddItem mrs����!����
        cbo����.ItemData(cbo����.NewIndex) = Val(Nvl(mrs����!ID))
        If mlngȱʡ�Һſ���ID = Val(Nvl(mrs����!ID)) Then cbo����.ListIndex = cbo����.NewIndex  '���˺�:���Ӵ��������д���Ŀ���
        mrs����.MoveNext
    Loop

    '�Һ���Ŀ
    strSQL = "Select ID as ���,���� From �շ���ĿĿ¼ " & _
        " Where ���='1' And (Sysdate Between ����ʱ�� And ����ʱ�� Or ����ʱ��<Sysdate And ����ʱ�� Is Null)" & _
        " And (վ��='" & gstrNodeNo & "' Or վ�� is Null)" & _
        " Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п��õĹҺ���Ŀ��Ϣ,���ȵ��Һ���Ŀ�����г�ʼ��", vbInformation, gstrSysName
        Exit Function
    End If
    cboItem.Clear
    Do While Not rsTemp.EOF
        cboItem.AddItem rsTemp!����
        cboItem.ItemData(cboItem.NewIndex) = rsTemp!���
        rsTemp.MoveNext
    Loop

    '����
    strSQL = "Select ����,����,ȱʡ��־ From ���� Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    cbo����.Clear
    Do While Not rsTemp.EOF
        cbo����.AddItem rsTemp!����
        If IIf(IsNull(rsTemp!ȱʡ��־), 0, rsTemp!ȱʡ��־) = 1 Then
            cbo����.ListIndex = cbo����.NewIndex
        End If
        rsTemp.MoveNext
    Loop

    '��������
    strSQL = "Select ����,���ơ�From �������� Where (վ��='" & gstrNodeNo & "' Or վ�� is Null) Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    lvwDept.ListItems.Clear
    Do While Not rsTemp.EOF
        lvwDept.ListItems.Add , "D" & rsTemp!����, rsTemp!����
        rsTemp.MoveNext
    Loop
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.ActiveControl Is cbo���� Then Exit Sub
    If Me.ActiveControl Is cboDoctor Then Exit Sub
    If Me.ActiveControl Is vsPlan Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mstr�����޸� = ""
End Sub



Private Sub lblҽ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 0 Then Exit Sub
        If Val(lblҽ��.Tag) = 0 Then Exit Sub

        PopupMenu mnuPopu, 2
End Sub



Private Sub lvwDept_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    If opt����(1).Value Then
        For i = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(i).Key <> Item.Key Then
                lvwDept.ListItems(i).Checked = False
            End If
        Next
    End If
    Set lvwDept.SelectedItem = Item
End Sub


Private Sub mnuViewDoctor_Click(Index As Integer)
        mnuViewDoctor(Index).Checked = True
        If Index = 0 Then
            mnuViewDoctor(1).Checked = False: mblnOnlyԺ��ҽ�� = True
        Else
            mnuViewDoctor(0).Checked = False: mblnOnlyԺ��ҽ�� = False
        End If

        lblҽ��.Caption = IIf(mblnOnlyԺ��ҽ��, "Ժ��ҽ��", "ҽ��") & "��"
        lblҽ��.ToolTipText = IIf(mblnOnlyԺ��ҽ��, "ֻ��ѡ��Ժ�ڽ���ҽ��", "����Ԯҽ��(���˿���ѡ��Ժ��ҽ���⣬������������Ԯҽ��)")
End Sub





'
'Private Sub vsPlan_EnterCell(Row As Long, Col As Long)
'    vsPlan.Active = opt��.Value
'End Sub

Private Sub opt����_Click(Index As Integer)
    Dim i As Integer, strKey As String
    If opt����(1).Value Then
        For i = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(i).Checked Then
                If strKey = "" Then
                    strKey = lvwDept.ListItems(i).Key
                Else
                    lvwDept.ListItems(i).Checked = False
                End If
            End If
        Next
        If strKey <> "" Then
            Set lvwDept.SelectedItem = lvwDept.ListItems(strKey)
            lvwDept.SelectedItem.EnsureVisible
        End If
    End If
End Sub

Private Sub opt��_Click()
    Dim i As Integer
    Dim strPlan As String

    For i = 0 To vsPlan.Cols - 1
        If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
            If strPlan = "" Then
                strPlan = vsPlan.TextMatrix(1, i)
            Else
                If vsPlan.TextMatrix(1, i) <> strPlan Then
                    strPlan = "": Exit For
                End If
            End If
        End If
    Next

    opt��.Value = -True: txt�޺�.Enabled = True: txt��Լ.Enabled = True
    cbo��.Enabled = True

    opt��.Value = False
    With vsPlan
        .Enabled = False: .TabStop = False
        For i = 1 To 7
             .TextMatrix(1, i) = ""
             .TextMatrix(2, i) = ""
             .TextMatrix(3, i) = ""
        Next
    End With

    cbo��.ListIndex = cbo.FindIndex(cbo��, strPlan, True)
    cbo��.SetFocus
End Sub

Private Sub opt��_Click()
    Dim i As Integer

    If Trim(cbo��.Text) <> "" Then
        For i = 1 To vsPlan.Cols - 1
            vsPlan.TextMatrix(1, i) = cbo��.Text
            vsPlan.TextMatrix(2, i) = txt�޺�.Text
            vsPlan.TextMatrix(3, i) = txt��Լ.Text
        Next
    End If

    opt��.Value = False
    cbo��.Enabled = False: txt�޺�.Enabled = False: txt��Լ.Enabled = False
    cbo��.ListIndex = -1

    opt��.Value = True
    vsPlan.Enabled = True: vsPlan.TabStop = True
    vsPlan.Col = 1: vsPlan.SetFocus
End Sub






Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

    If mblnChangeByCode Then Exit Sub
    PageChange Item
End Sub

Private Sub PageChange(ByVal Item As XtremeSuiteControls.ITabControlItem)

    If mblnChangeByCode Then Exit Sub

    If Item.Index = mPageIndex.EM_ʱ�� Then
       mblnChangeByCode = True
       tbPage.Item(mPageIndex.EM_����).Selected = True
        If isValied() = False Then
            mblnChangeByCode = False
            Exit Sub
        End If
        tbPage.Item(mPageIndex.EM_ʱ��).Selected = True
        mblnChangeByCode = False
        Call LoadTimePlan
    Else
        If mfrmTime.mblnChange = False Then Exit Sub
        If mfrmTime.zl_CheckMoveAssign() = False Then
             mblnChangeByCode = True
            tbPage.Item(mPageIndex.EM_ʱ��).Selected = True
             mblnChangeByCode = False
        End If
         
    End If
End Sub
Private Sub LoadTimePlan(Optional ByVal blnSaveBeforCheck As Boolean = False)
    Dim i As Long
    Dim lng�޺��� As Long
    Dim lng��Լ�� As Long
    Dim strTemp As String
    Dim str���� As String
    Dim str�Ű� As String

    If Not mrsRegNewData Is Nothing Then Set mrsRegNewData = Nothing

    If mrsRegNewData Is Nothing Then
        Set mrsRegNewData = New ADODB.Recordset
        mrsRegNewData.Fields.Append "ID", adBigInt, 18
        mrsRegNewData.Fields.Append "������Ŀ", adVarChar, 20
        mrsRegNewData.Fields.Append "�Ű�", adVarChar, 20
        mrsRegNewData.Fields.Append "�޺���", adBigInt, 10
        mrsRegNewData.Fields.Append "��Լ��", adBigInt, 18
        mrsRegNewData.Fields.Append "��ſ���", adBigInt, 18
        mrsRegNewData.CursorLocation = adUseClient
        mrsRegNewData.LockType = adLockOptimistic
        mrsRegNewData.CursorType = adOpenStatic
        mrsRegNewData.Open
     End If

     If opt��.Value = True Then
          lng�޺��� = Val(txt�޺�.Text)
          lng��Լ�� = Val(txt��Լ.Text)
          str�Ű� = Me.cbo��.Text
          For i = 0 To 6
            strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
            '��һ,�޺���,��Լ��|�ܶ�,�޺���,��Լ��|....
            str���� = str���� & "|" & strTemp & "," & lng�޺��� & "," & lng��Լ��
             With mrsRegNewData
                .AddNew
                !ID = 0
                !������Ŀ = strTemp
                !�Ű� = str�Ű�
                !�޺��� = lng�޺���
                !��Լ�� = lng��Լ��
                !��ſ��� = Me.chk��ſ���.Value
                .Update
            End With
          Next

        Else

           With vsPlan
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTemp = Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
                    lng�޺��� = Val(Trim(vsPlan.TextMatrix(2, i)))
                    lng��Լ�� = Val(Trim(vsPlan.TextMatrix(3, i)))
                    str�Ű� = Trim(vsPlan.TextMatrix(1, i))
                    str���� = str���� & "|" & strTemp & "," & lng�޺��� & "," & lng��Լ��
                    With mrsRegNewData
                        .AddNew
                        !ID = Val(mlngID)
                        !������Ŀ = strTemp
                        !�Ű� = str�Ű�
                        !�޺��� = lng�޺���
                        !��Լ�� = lng��Լ��
                        !��ſ��� = Me.chk��ſ���.Value
                        .Update
                    End With
                End If
            Next
        End With
     End If
     If str���� <> "" Then str���� = Mid(str����, 2)
'Public Enum mRegEditType
'Ed_�ƻ����� = 0
'Ed_�����޸� = 1
'Ed_����ɾ�� = 2
'Ed_������� = 3
'Ed_����ȡ�� = 4
'Ed_���Ų��� = 5
'End Enum

     mfrmTime.zlShowPagePlan str����, mrsRegNewData, mrsRegHistory, chk��ſ���.Value = 1, Switch(mEditType = ed_�ƻ�����, EM_����_����, mEditType = Ed_�����޸�, EM_����_�޸�, True, EM_����_����), mlngID, Val(0), blnSaveBeforCheck
End Sub

Private Sub txt�ű�_GotFocus()
    Call zlControl.TxtSelAll(txt�ű�)
End Sub

Private Sub txt�ű�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�޺�_GotFocus()
    Call zlControl.TxtSelAll(txt�޺�)
End Sub

Private Sub txt�޺�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�޺�_Validate(Cancel As Boolean)
    If Trim(txt�޺�.Text) = "" And Trim(txt��Լ.Text) <> "" Then
        MsgBox "��Լ�����޺�!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If

    If Trim(txt�޺�.Text) <> "" And Trim(txt��Լ.Text) <> "" And Val(txt�޺�.Text) < Val(txt��Լ.Text) Then
        MsgBox "�޺���Ӧ������Լ��!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub txt��Լ_GotFocus()
    Call zlControl.TxtSelAll(txt��Լ)
End Sub

Private Sub txt��Լ_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If Val(txt�޺�.Text) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��Լ_Validate(Cancel As Boolean)
    If Val(txt�޺�.Text) < Val(txt��Լ.Text) And _
        Trim(txt�޺�.Text) <> "" And Trim(txt��Լ.Text) <> "" Then
        MsgBox "��Լ��ӦС���޺���!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub
Private Function zlCheckRegistPlanIsValied(ByRef blnMulitNumPlan As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ������ĺ����Ƿ�Ϸ�
    '����:blnMulitNumPlan-�����Ƿ��ж����ͬ(ͬһ��Ŀ,ͬһ����,ͬһ��,��ͬ��)�İ���
    '����:�Ϸ�����,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-12-29 10:26:45
    '������ͬһ��Ŀ,ͬһ����,ͬһ��,��ͬ�ţ�:
    '     1.ͬ���ڲ����н���İ���
    '����Ŀ:35057
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strҽ�� As String
    Dim lng��Ŀid As Long, lng����ID As Long, lngҽ��ID As Long
    Dim str�ű� As String, strTemp As String, strTemp1 As String
    Dim i As Long
    On Error GoTo errHandle
    lng����ID = cbo����.ItemData(cbo����.ListIndex)
    lng��Ŀid = cboItem.ItemData(cboItem.ListIndex)
    lngҽ��ID = 0: strҽ�� = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    strSQL = "" & _
        "   Select ����,���,���� D0,��һ D1,�ܶ� D2,���� D3,���� D4,���� D5,���� D6," & _
        "           To_Char(��ʼʱ��,'YYYY-MM-DD HH24:MI:SS') ��ʼʱ��,To_Char(��ֹʱ��,'YYYY-MM-DD HH24:MI:SS') ��ֹʱ��" & _
        "   From �ҺŰ���  "

    If lngҽ��ID = 0 Then
        strSQL = strSQL & _
            "   Where ����id=[1] and  ��ĿID =[2] and ҽ������=[3] and nvl(ҽ��ID,0)=0 and ID<>" & mlngID & " Order by ���"
    Else
        strSQL = strSQL & _
        "   Where ����id=[1] and  ��ĿID =[2] and  ҽ��ID=[4] and ID<>" & mlngID & " Order by ���"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��Ŀid, strҽ��, lngҽ��ID)
    blnMulitNumPlan = Not rsTemp.EOF
    If blnMulitNumPlan = False Then zlCheckRegistPlanIsValied = True: Exit Function
    str�ű� = ""
    Do While Not rsTemp.EOF
        str�ű� = str�ű� & "," & Nvl(rsTemp!����)
        If opt��.Value Then
            If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D0)
            If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & " ��һ:" & Nvl(rsTemp!D1)
            If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & " �ܶ�:" & Nvl(rsTemp!D2)
            If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D3)
            If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D4)
            If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D5)
            If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D6)
            If strTemp <> "" Then
                strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ���������°���:" & vbCrLf & "        " & Mid(strTemp, 2)
                Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˰���.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                zlCheckRegistPlanIsValied = False: Exit Function
            End If
        Else
            With vsPlan
                For i = 0 To 6
                    strTemp1 = "��" & Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
                    If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                        '����,�϶��ظ���
                        strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                    End If
                Next
            End With
            If strTemp <> "" Then
                strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ���������°���:" & vbCrLf & "        " & Mid(strTemp, 2)
                Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˰���.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                zlCheckRegistPlanIsValied = False: Exit Function
            End If
        End If
        rsTemp.MoveNext
    Loop
    If str�ű� <> "" Then str�ű� = Mid(str�ű�, 2)
    If MsgBox("ע��:" & vbCrLf & "   ���֡�" & cboDoctor.Text & "��ҽ���Ѿ��������°���:" & vbCrLf & "    " & str�ű� & vbCrLf & "   �Ƿ�����Ը�ҽ�����а���?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        zlCheckRegistPlanIsValied = True: Exit Function
    End If
    zlCheckRegistPlanIsValied = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Function zlCheckPlanArrageIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ƻ������Ƿ���Ч
    '����:���ƻ������Ƿ������صİ���,�������صİ���,�򷵻�False,���򷵻�true
    '����:���˺�
    '����:2010-12-29 19:53:56
    '����Ŀ:35057
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strҽ�� As String
    Dim lng��Ŀid As Long, lng����ID As Long, lngҽ��ID As Long
    Dim str�ű� As String, strTemp As String, strTemp1 As String
    Dim blnCheck As Boolean
    Dim i As Long
    On Error GoTo errHandle
    lng����ID = cbo����.ItemData(cbo����.ListIndex)
    lng��Ŀid = cboItem.ItemData(cboItem.ListIndex)
    lngҽ��ID = 0: strҽ�� = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)

    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  distinct A.����,A.���� D0,A.��һ D1,A.�ܶ� D2,A.���� D3,A.���� D4,A.���� D5,A.���� D6," & _
    "           To_Char(��Чʱ��,'YYYY-MM-DD HH24:MI:SS') ��Чʱ��,To_Char(ʧЧʱ��,'YYYY-MM-DD HH24:MI:SS') ʧЧʱ��" & _
    "   From �ҺŰ��żƻ� A, �ҺŰ��� B " & _
    "   Where A.����ID=B.ID    " & _
    "      and   B.����id=[1] and  B.��ĿID =[2] and B.ҽ������=[3] and nvl(B.ҽ��ID,0)=[4] and B.ID<>" & mlngID & _
    "   Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��Ŀid, strҽ��, lngҽ��ID)
    If rsTemp.EOF Then
        zlCheckPlanArrageIsValied = True: Exit Function
    End If
    Do While Not rsTemp.EOF
        str�ű� = str�ű� & "," & Nvl(rsTemp!����)
        blnCheck = chk��Ч��.Value = 0
        If chk��Ч��.Value = 1 Then
            blnCheck = Nvl(rsTemp!��Чʱ��) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!��Чʱ��) < Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
            blnCheck = blnCheck Or Nvl(rsTemp!ʧЧʱ��) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!ʧЧʱ��) < Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
            blnCheck = blnCheck Or Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!��Чʱ��) And Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!ʧЧʱ��)
            blnCheck = blnCheck Or Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!��Чʱ��) And Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!ʧЧʱ��)

        End If
        If blnCheck Then
            If opt��.Value Then
                If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D0)
                If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & " ��һ:" & Nvl(rsTemp!D1)
                If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & " �ܶ�:" & Nvl(rsTemp!D2)
                If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D3)
                If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D4)
                If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D5)
                If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D6)
                If strTemp <> "" Then
                    strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ���������¼ƻ�����:" & vbCrLf & "        " & Mid(strTemp, 2)
                    Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˰���.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckPlanArrageIsValied = False: Exit Function
                End If
            Else
                With vsPlan
                    For i = 0 To 6
                        strTemp1 = "��" & Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
                        If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                            '����,�϶��ظ���
                            strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                        End If
                    Next
                End With
                If strTemp <> "" Then
                    strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ���������¼ƻ�����:" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & "  ��Чʱ��:" & IIf(Nvl(rsTemp!��Чʱ��) = "1901-01-01", "����", Nvl(rsTemp!��Чʱ��) & "-" & Nvl(rsTemp!ʧЧʱ��)) & vbCrLf
                    Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˰���.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckPlanArrageIsValied = False: Exit Function
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
    zlCheckPlanArrageIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Sub vsPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If mEditType = edt_���� Then Cancel = True: Exit Sub
        If Not opt��.Value = True Then Cancel = True: Exit Sub
    End With
End Sub


Private Sub vsPlan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������صĸ�ʽ
    '����:���˺�
    '����:2011-11-11 11:33:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsPlan
       If Row = 1 Then
              If Trim(.EditText) = "" Then
               .TextMatrix(2, Col) = ""
               .TextMatrix(3, Col) = ""
            End If
            Exit Sub
        End If
        .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), "###;;;")
    End With
    Exit Sub
End Sub
Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String
    Call zl_VsGridRowChange(vsPlan, OldRow, NewRow, OldCol, NewCol)
    vsPlan.ColComboList(NewCol) = ""

    If mstr�����޸� <> "" Then
        strTmp = ";��" & vsPlan.TextMatrix(0, NewCol) & ";"
        vsPlan.Editable = flexEDKbdMouse
        'If InStr(mstr�����޸�, strTmp) > 0 Then vsPlan.Editable = flexEDNone
    End If

    If OldRow = 1 And Trim(vsPlan.TextMatrix(1, OldCol)) = "" Then
        vsPlan.TextMatrix(2, OldCol) = ""
        vsPlan.TextMatrix(3, OldCol) = ""
    End If
    If NewRow <> 1 Then Exit Sub
    vsPlan.ColComboList(NewCol) = vsPlan.Tag
End Sub
Private Sub vsPlan_GotFocus()
    Call zl_VsGridGotFocus(vsPlan)
End Sub
Private Sub vsPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    With vsPlan
        If KeyCode = vbKeyDelete Then
            .TextMatrix(.Row, .Col) = ""
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub

    With vsPlan
        If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If .Row < 3 Then
            .Row = .Row + 1
        Else
            .Row = 1
            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
         End If
    End With
End Sub

Private Sub vsPlan_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '�༭����
    Dim intCol As Integer, strKey As String, lngRow As Long

    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPlan
            If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If .Row < 3 Then
            .Row = .Row + 1
        Else
            .Row = 1
            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
         End If
    End With
End Sub
Private Sub vsPlan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub vsPlan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsPlan
        If Row <= 1 Then Exit Sub
        VsFlxGridCheckKeyPress vsPlan, Row, Col, KeyAscii, m����ʽ
    End With
End Sub
Private Sub vsPlan_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsPlan)
End Sub

Private Sub vsPlan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, strTemp As String, strTmp As String
    '������֤
    With vsPlan
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        If .Row <= 1 Then Exit Sub
        If zlCommFun.DblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
            Cancel = True: Exit Sub
        End If
        strKey = Format(Abs(Val(strKey)), "####;;;")
         If mstr�����޸� <> "" Then
               strTmp = "��" & vsPlan.TextMatrix(0, Col)
               'vsPlan.Editable = flexEDKbdMouse
               If InStr(mstr�����޸�, ";" & strTmp & ";") > 0 Then
                   Cancel = Val(strKey) < Val(.TextMatrix(Row, Col))
               End If
        End If
        If Cancel Then Exit Sub
        If Row = 2 Then
            If Val(strKey) < Val(.TextMatrix(3, Col)) Then
                If MsgBox("�޺���С������Լ��,�Ƿ������Լ��?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True: Exit Sub
                .TextMatrix(3, Col) = ""
            End If
        ElseIf Row = 3 Then
            If Val(strKey) > Val(.TextMatrix(2, Col)) Then
                Call MsgBox("�޺���С������Լ��,���ܼ���", vbOKOnly, gstrSysName)
                Cancel = True: Exit Sub
            End If
        End If

        .EditText = strKey
    End With
End Sub


Private Function Checkʱ��() As Boolean
    '----------------------------------
    '�ж��Ƿ��ʱ��
    '----------------------------------
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset

    If mEditType = edt_���� Or mEditType = edt_���� Then Exit Function

    On Error GoTo Hd
    strSQL = _
    "   Select 1 As Hdata From �ҺŰ���ʱ�� Where ����id =[1] And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
     Checkʱ�� = Not rsTmp.EOF
    Set rsTmp = Nothing
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Private Function LoadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���سɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2009-09-15 12:14:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL          As String
    Dim rsTemp          As New ADODB.Recordset
    Dim i               As Long
    Dim strTemp         As String
    Dim rs�޺�          As ADODB.Recordset
    Dim blnÿ��         As Boolean
    Dim bln�޺�         As Boolean
    Dim str�޺�         As String
    Dim bln��Լ         As Boolean
    Dim str��Լ         As String
    Err = 0: On Error GoTo Errhand:

    For i = 1 To lvwDept.ListItems.Count
        lvwDept.ListItems(i).Checked = False
    Next

    If mEditType = edt_���� Then
        txt�ű�.Text = GetNext�ű�
        txt�޺�.Text = ""
        txt��Լ.Text = ""
        chk����.Value = 0

        If cbo����.ListIndex >= 0 Then
            If mlngȱʡ�Һſ���ID <> cbo����.ItemData(cbo����.ListIndex) Then
                cbo����.ListIndex = -1
                cboItem.ListIndex = -1
                cboDoctor.Text = ""
            End If
        Else
            cbo����.ListIndex = -1
            cboItem.ListIndex = -1
            cboDoctor.Text = ""
        End If
        dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = CDate("3000-01-01")

        opt��.Value = True
        cbo��.Enabled = True
        cbo��.ListIndex = cbo.FindIndex(cbo��, "ȫ��", True)
        If cbo��.ListIndex = -1 Then cbo��.ListIndex = 0
        opt��.Value = False
        vsPlan.Enabled = False
        LoadCard = True
        opt����(0).Value = True
        Exit Function
    End If
    '�޸Ļ�鿴
    strSQL = " " & _
    "   Select A.Id as ����ID,0 as �ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id," & _
    "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,A.Ĭ��ʱ�μ��, " & _
    "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ���� " & _
    "   From �ҺŰ��� A,�շ���ĿĿ¼ B,���ű� D " & _
    "   Where A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
    "         And A.Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)

    If rsTemp.EOF Then
        ShowMsgbox "δ�ҵ�ָ���ĺű�,����!"
        Exit Function
    End If
    strSQL = "Select ������Ŀ,�޺���,  ��Լ�� From  �ҺŰ������� where ����ID=[1]       "
   Set rs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)

    cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(rsTemp!����), True)
    txt�ű�.Text = Nvl(rsTemp!����)

    cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(rsTemp!����), True)
    cboItem.ListIndex = cbo.FindIndex(cboItem, Nvl(rsTemp!��Ŀ), True)

    cboDoctor.ListIndex = cbo.FindIndex(cboDoctor, Nvl(rsTemp!ҽ������), True)
    If cboDoctor.ListIndex = -1 Then cboDoctor.Text = Nvl(rsTemp!ҽ������)


    chk����.Value = IIf(Val(Nvl(rsTemp!��������)) = 1, 1, 0)

    chk��ſ���.Value = IIf(Val(Nvl(rsTemp!��ſ���)) = 1, 1, 0):     chk��ſ���.Tag = chk��ſ���.Value
    '��ȡ�޸�ǰ�İ����Ƿ���ſ���
    mPlanInfo.bln��� = IIf(Val(Nvl(rsTemp!��ſ���)) = 1, True, False)
    '��Чʱ�䷶Χ
    dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = CDate("3000-01-01")
    If Not IsNull(rsTemp!��ʼʱ��) Then
        chk��Ч��.Value = 1
        dtpBegin.Value = CDate(Format(rsTemp!��ʼʱ��, "yyyy-mm-dd HH:MM:SS"))
        If Not IsNull(rsTemp!��ֹʱ��) Then
            dtpEnd.Value = CDate(Format(rsTemp!��ֹʱ��, "yyyy-mm-dd HH:MM:SS"))
        End If
    End If

     '����ԭʼ���ݵ����ݼ�
     With mrsRegOldData
        Set mrsRegOldData = New ADODB.Recordset
        mrsRegOldData.Fields.Append "ID", adBigInt, 18
        mrsRegOldData.Fields.Append "������Ŀ", adVarChar, 20
        mrsRegOldData.Fields.Append "�޺���", adBigInt, 10
        mrsRegOldData.Fields.Append "��Լ��", adBigInt, 18
        mrsRegOldData.Fields.Append "��ſ���", adBigInt, 18
        mrsRegOldData.CursorLocation = adUseClient
        mrsRegOldData.LockType = adLockOptimistic
        mrsRegOldData.CursorType = adOpenStatic
        mrsRegOldData.Open


        rs�޺�.Filter = 0
        If rs�޺�.RecordCount > 0 Then rs�޺�.MoveFirst
        Do While Not rs�޺�.EOF
            With mrsRegOldData
                .AddNew
                !ID = mlngID
                !������Ŀ = Nvl(rs�޺�!������Ŀ)
                !�޺��� = Val(Nvl(rs�޺�!�޺���))
                !��Լ�� = Val(Nvl(rs�޺�!��Լ��))
                !��ſ��� = Val(Nvl(rsTemp!��ſ���))
                .Update
            End With
            rs�޺�.MoveNext
        Loop
    End With

    Call LoadRegHistory

    '---------------------------------------------------
    '�ж� ÿ�հ��� �޺��� ��Լ�� ���Ƿ�һ��
    '---------------------------------------------------
    blnÿ�� = Nvl(rsTemp!����) <> Nvl(rsTemp!��һ) Or Nvl(rsTemp!����) <> Nvl(rsTemp!�ܶ�) _
        Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) _
        Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) Or Nvl(rsTemp!����) <> Nvl(rsTemp!����)

    If blnÿ�� = False Then
             rs�޺�.Filter = "������Ŀ='����'"
             If Not rs�޺�.EOF Then
                str�޺� = Nvl(rs�޺�!�޺���)
                str��Լ = Nvl(rs�޺�!��Լ��)
             End If
            For i = 1 To 6
                strTemp = Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
                rs�޺�.Filter = "������Ŀ='" & "��" & strTemp & "'"
                If Not rs�޺�.EOF Then
                    bln�޺� = Nvl(rs�޺�!�޺���) = str�޺�
                    bln��Լ = Nvl(rs�޺�!��Լ��) = str��Լ
                    If bln��Լ = False Or bln�޺� = False Then Exit For
                End If
            Next
          blnÿ�� = True
         If bln�޺� And bln��Լ Then blnÿ�� = False

    End If

   If blnÿ�� Or mrsRegHistory.RecordCount > 0 Then
        'ÿ��
        opt��.Value = True
        With vsPlan
            For i = 1 To .Cols - 1
                strTemp = Switch(i - 1 = 0, "��", i - 1 = 1, "һ", i - 1 = 2, "��", i - 1 = 3, "��", i - 1 = 4, "��", i - 1 = 5, "��", True, "��")
                .TextMatrix(1, i) = Nvl(rsTemp.Fields("��" & strTemp))
                rs�޺�.Filter = "������Ŀ='" & "��" & strTemp & "'"
                If Not rs�޺�.EOF Then
                    .TextMatrix(2, i) = Nvl(rs�޺�!�޺���)
                    .TextMatrix(3, i) = Nvl(rs�޺�!��Լ��)
                End If
                If InStr(mstr�����޸�, ";��" & strTemp & ";") > 0 Then
                    .Cell(flexcpForeColor, 2, i, 3, i) = vbBlue
                End If
            Next
        End With
        opt��.Value = False: cbo��.Enabled = False: txt�޺�.Enabled = False: txt��Լ.Enabled = False
        vsPlan.Enabled = True: chk��ſ���.Enabled = mstr�����޸� = ""
    Else
        'ÿ��
        opt��.Value = True:  cbo��.ListIndex = cbo.FindIndex(cbo��, Nvl(rsTemp!����), True)
        If cbo��.ListIndex = -1 Then cbo��.ListIndex = 0:
        opt��.Value = False: vsPlan.Enabled = False
        If rs�޺�.RecordCount <> 0 Then rs�޺�.MoveFirst
        If rs�޺�.EOF = False Then
            txt�޺�.Text = Nvl(rs�޺�!�޺���)
            txt��Լ.Text = Nvl(rs�޺�!��Լ��)
        End If
    End If

    '------------------------------
    '��ȡ�޸�ǰ�� ʱ��κ� �޺���
    '�����ڱ���ʱ �Ա��޺���Լ�Լ�ʱ����Ƿ����˱仯
    '��������˱仯����Ҫ��ʾ  ����Ա��������ʱ����Ϣ
    '------------------------------
   mPlanInfo.str�Ű� = ""
   mPlanInfo.str�޺� = ""

    If blnÿ�� Or mrsRegHistory.RecordCount > 0 Then
         For i = 1 To vsPlan.Cols - 1
            mPlanInfo.str�Ű� = mPlanInfo.str�Ű� & "'" & Trim(vsPlan.TextMatrix(1, i)) & "',"

                mPlanInfo.str�޺� = mPlanInfo.str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                     mPlanInfo.str�޺� = mPlanInfo.str�޺� & ",0,0"
                Else
                     mPlanInfo.str�޺� = mPlanInfo.str�޺� & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Val(Trim(vsPlan.TextMatrix(3, i)))
                End If
        Next


    Else
         For i = 1 To 7
             mPlanInfo.str�Ű� = mPlanInfo.str�Ű� & "'" & Trim(cbo��.Text) & "',"
             mPlanInfo.str�޺� = mPlanInfo.str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
             mPlanInfo.str�޺� = mPlanInfo.str�޺� & "," & Val(txt�޺�.Text) & "," & Val(txt��Լ.Text)
        Next
    End If
    If mPlanInfo.str�޺� <> "" Then mPlanInfo.str�޺� = Mid(mPlanInfo.str�޺�, 2)
    '-------------------------------

     Select Case Val(Nvl(rsTemp!���﷽ʽ))     '0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
        Case 0  '"������"
            opt����(0).Value = True
        Case 1  ' "ָ������"
            opt����(1).Value = True
        Case 2 '"��̬����"
            opt����(2).Value = True
        Case 3 ' "ƽ������"
            opt����(3).Value = True
    End Select

    strSQL = "Select �ű�ID,�������ҡ�From �ҺŰ������� Where �ű�ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    Do While Not rsTemp.EOF
        For i = 1 To lvwDept.ListItems.Count
            If rsTemp!�������� = lvwDept.ListItems(i).Text Then
                lvwDept.ListItems(i).Checked = True
            End If
        Next
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    If mstr�����޸� <> "" Then opt��.Enabled = False
    '������޸�ʱ ��ȡԭ���İ����Ƿ��Ѿ�������ʱ��
    If mEditType = edt_�޸� Then mPlanInfo.blnʱ��� = Checkʱ��
    If mrsRegHistory.RecordCount > 0 Then opt��.Enabled = False
    LoadCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Sub cboDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        cboDoctor.ListIndex = GetCboIndex(cboDoctor, cboDoctor)
'    End If
End Sub

Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    If KeyAscii <> 13 Then Exit Sub
    If cboDoctor.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If mrsDoctor Is Nothing Then Exit Sub
    If Trim(cboDoctor.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub

    If zlPersonSelect(Me, mlngModule, cboDoctor, mrsDoctor, cboDoctor.Text, True, "") = False Then
        If mblnOnlyԺ��ҽ�� = False Then
                zlCommFun.PressKey vbKeyTab
        End If
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cboDoctor_Validate(Cancel As Boolean)
      If mblnOnlyԺ��ҽ�� Then
           If cboDoctor.ListIndex < 0 Then cboDoctor.Text = ""
      End If

    'ָ��ҽ��ʱ����ָ���������
    If Trim(cboDoctor.Text) <> "" Then
        opt����(2).Enabled = False
        opt����(3).Enabled = False
        If opt����(2).Value Or opt����(3).Value Then opt����(0).Value = True
    Else
        opt����(2).Enabled = True
        opt����(3).Enabled = True
    End If
End Sub

Private Sub cbo����_Click()
    mblnCboClick = True
    If cbo����.ListIndex = -1 Then Exit Sub
    Call LoadDoctor
End Sub

Private Sub LoadDoctor()
    Set mrsDoctor = GetDoctor(Val(cbo����.ItemData(cbo����.ListIndex)), "")
    cboDoctor.Clear
    Do While Not mrsDoctor.EOF
        cboDoctor.AddItem mrsDoctor!����
        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
        mrsDoctor.MoveNext
    Loop
End Sub

Private Sub cbo����_GotFocus()
    zlControl.TxtSelAll cbo����
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo����.Text = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If cbo����.ListIndex >= 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        mblnCboClick = True
        If Select����(Me, mlngModule, mrs����, cbo����, cbo����.Text) = True Then
            mblnCboClick = False
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        If cbo����.Enabled Then cbo����.SetFocus
        mblnCboClick = False
        zlControl.TxtSelAll cbo����
    Else
       ' Call zlControl.CboSetIndex(cbo����.hWnd, zlControl.CboMatchIndex(cbo����.hWnd, KeyAscii))
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
 '�����cbo��keypress�¼������˵����б�ĵ�API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
    If Not mblnCboClick Then cbo����_Click
    mblnCboClick = False
End Sub

Private Sub chk��Ч��_Click()
    dtpBegin.Enabled = chk��Ч��.Value = 1
    dtpEnd.Enabled = chk��Ч��.Value = 1

    If Visible And dtpBegin.Enabled Then
        dtpBegin.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Function GetDoctorPlan(lngҽ��ID As Long, strҽ������ As String) As ADODB.Recordset
'����:����ָ��ҽ��ID�����������кű��ʱ����Ϣ
'   ���ڼ���������޸ĵĺű��Ƿ������еĺű���ʱ�����ظ�
    Dim strSQL As String

    On Error GoTo errH
    strSQL = "Select ����,���� D0,��һ D1,�ܶ� D2,���� D3,���� D4,���� D5,���� D6," & _
            " To_Char(��ʼʱ��,'YYYY-MM-DD HH24:MI:SS') ��ʼʱ��,To_Char(��ֹʱ��,'YYYY-MM-DD HH24:MI:SS') ��ֹʱ��" & _
            " From �ҺŰ��� Where (��ֹʱ�� is null or ��ֹʱ��>sysdate) And " & IIf(lngҽ��ID <> 0, " ҽ��ID=[1]", " ҽ������=[1]") & _
            IIf(mEditType = edt_����, "", " And ID<>[2]")
    Set GetDoctorPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(lngҽ��ID <> 0, lngҽ��ID, strҽ������), mlngID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckExistsBooking() As Boolean
'����:��鵱ǰʱ���֮���Ƿ����ԤԼ�Һŵ�
    Dim rsTemp As ADODB.Recordset, rsBooking As ADODB.Recordset, strSQL As String
    Dim i As Long, strʱ��� As String

    On Error GoTo errH
    If opt��.Value Then
        strʱ��� = _
               "Select 1 From ʱ��� b Where b.ʱ��� = [2] And (" & _
               " ('3000-01-10 '||To_Char(a.����ʱ��,'HH24:MI:SS')" & _
               " Between" & _
               " Decode(Sign(b.��ʼʱ��-b.��ֹʱ��),1,'3000-01-09 '||To_Char(b.��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.��ʼʱ��,'HH24:MI:SS'))" & _
               " And" & _
               " '3000-01-10 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'))" & _
               " Or" & _
               " ('3000-01-10 '||To_Char(a.����ʱ��,'HH24:MI:SS')" & _
               " Between" & _
               " '3000-01-10 '||To_Char(b.��ʼʱ��,'HH24:MI:SS')" & _
               " And" & _
               " Decode(Sign(b.��ʼʱ��-b.��ֹʱ��),1,'3000-01-11 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'))))"

        strSQL = "Select  /*+ Rule*/ Min(����ʱ��) ʱ��" & vbNewLine & _
            "From ������ü�¼ a" & vbNewLine & _
            "Where ��¼���� = 4 And ��¼״̬ In (0, 1) And ���㵥λ = [1] And ����ʱ�� > �Ǽ�ʱ��"
        If gintԤԼ���� = 0 Then
            strSQL = strSQL & " And ����ʱ�� > Sysdate"
        Else
            strSQL = strSQL & " And ����ʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
        End If
        strSQL = strSQL & " And Not Exists (" & strʱ��� & ")"

        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt�ű�.Text, Trim(cbo��.Text))
        CheckExistsBooking = Not IsNull(rsTemp!ʱ��)
    Else
        strSQL = "Select /*+ Rule*/ ����ʱ��,To_Char(����ʱ��,'D') ���� From ������ü�¼ a Where ��¼���� = 4 and ��¼״̬ In(0,1) And ���㵥λ = [1] And ����ʱ�� > �Ǽ�ʱ��"
        If gintԤԼ���� = 0 Then
            strSQL = strSQL & " And ����ʱ�� > Sysdate"
        Else
            strSQL = strSQL & " And ����ʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
        End If

        Set rsBooking = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt�ű�.Text)
        For i = 1 To rsBooking.RecordCount
            strʱ��� = Trim(vsPlan.TextMatrix(1, rsBooking!���� - 1))
            If strʱ��� = "" Then
                CheckExistsBooking = True
            Else
               strSQL = _
                    "Select Count(*) cnt From ʱ��� b Where b.ʱ��� = [2] And (" & _
                    " ('3000-01-10 '||To_Char([1],'HH24:MI:SS')" & _
                    " Between" & _
                    " Decode(Sign(b.��ʼʱ��-b.��ֹʱ��),1,'3000-01-09 '||To_Char(b.��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.��ʼʱ��,'HH24:MI:SS'))" & _
                    " And" & _
                    " '3000-01-10 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'))" & _
                    " Or" & _
                    " ('3000-01-10 '||To_Char([1],'HH24:MI:SS')" & _
                    " Between" & _
                    " '3000-01-10 '||To_Char(b.��ʼʱ��,'HH24:MI:SS')" & _
                    " And" & _
                    " Decode(Sign(b.��ʼʱ��-b.��ֹʱ��),1,'3000-01-11 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'))))"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(rsBooking!����ʱ��), strʱ���)
                CheckExistsBooking = rsTemp!cnt = 0
            End If

            If CheckExistsBooking Then Exit Function
            rsBooking.MoveNext
        Next
    End If

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function isValied() As Boolean
     Dim i As Integer, intCount As Integer, j As Integer
    Dim strʱ��� As String, str���� As String, str�޺� As String
    Dim lngNextID As Long, lngҽ��ID As Long
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strInfo As String, strTmp As String, strOld As String, strNew As String
    Dim str�ű� As String
    Dim rsDoctorPlan As ADODB.Recordset
    Dim rsNewDate As ADODB.Recordset
    Dim rsOldDate As ADODB.Recordset
    Dim rsSNState As ADODB.Recordset
    Dim blnMulitNumPlan As Boolean  '�Ƿ��ΰ���
    Dim blnChange       As Boolean '�Ƿ�ı��� ʱ�䰲��
    Dim strMsg          As String

    If opt��.Value Then
        If cbo��.ListIndex = -1 Then
            MsgBox "�úű�ÿ���Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
            cbo��.SetFocus: Exit Function
        End If

        If Val(txt�޺�.Text) = 0 And Val(txt��Լ.Text) = 0 Then
            MsgBox "��������ʱ��ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
            txt�޺�.SetFocus: Exit Function
        End If
        '�޺���Լ����
        If Trim(txt�޺�.Text) <> "" Then
            If Trim(txt��Լ.Text) <> "" And Val(txt�޺�.Text) < Val(txt��Լ.Text) Then
                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
                txt��Լ.SetFocus: Exit Function
            End If
        ElseIf Trim(txt��Լ.Text) <> "" Then
            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
            txt�޺�.SetFocus: Exit Function
        End If
    Else
        With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))

                        If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                            MsgBox "��������ʱ��ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Function
                        End If

                        '�޺���Լ����
                        If Val(.TextMatrix(2, i)) <> 0 Then
                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
                                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
                                .Row = 2: .Col = i
                                .SetFocus: Exit Function
                            End If
                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
                            
                            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Function
                        End If
                End If
            Next
            If strTmp = "" Then
                MsgBox "�úű�ÿ�ܵ�Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
                vsPlan.SetFocus: Exit Function
            End If
        End With
    End If
    isValied = True
End Function
Private Sub cmdOK_Click()
    Dim i As Integer, intCount As Integer, j As Integer
    Dim strʱ��� As String, str���� As String, str�޺� As String
    Dim lngNextID As Long, lngҽ��ID As Long
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strInfo As String, strTmp As String, strOld As String, strNew As String
    Dim cllPro As Collection
    Dim str�ű� As String
    Dim rsDoctorPlan As ADODB.Recordset
    Dim rsNewDate As ADODB.Recordset
    Dim rsOldDate As ADODB.Recordset
    Dim rsSNState As ADODB.Recordset
    Dim blnMulitNumPlan As Boolean  '�Ƿ��ΰ���
    Dim blnChange       As Boolean '�Ƿ�ı��� ʱ�䰲��
    Dim strMsg          As String
    If mEditType = edt_���� Then Unload Me: Exit Sub
    If Me.tbPage.Item(mPageIndex.EM_����).Selected = False Then
        mblnChangeByCode = True
        tbPage.Item(mPageIndex.EM_����).Selected = True
        mblnChangeByCode = False
    End If
    If mblnOnlyԺ��ҽ�� Then
        If cboDoctor.ListIndex < 0 And cboDoctor.Text <> "" Then
                MsgBox "��ѡ���ҽ��������,����������ҽ��!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                If cboDoctor.Enabled Then cboDoctor.SetFocus
                Exit Sub
        End If
    End If
    '�����Լ��
    If Trim(txt�ű�) = "" Then
        MsgBox "�ű���Ϊ�գ�", vbInformation, gstrSysName
        txt�ű�.SetFocus: Exit Sub
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "δ���úű�����Ӧ�Ŀ��ң�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    If cboItem.ListIndex = -1 Then
        MsgBox "δ���úű�����Ӧ�ĹҺ���Ŀ��", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Sub
    End If

    If dtpBegin.Enabled And dtpEnd.Enabled Then
        If dtpBegin.Value >= dtpEnd.Value Then
            MsgBox "��ʼʱ��Ӧ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
            dtpBegin.SetFocus: Exit Sub
        End If
    End If

    If opt��.Value Then
        If cbo��.ListIndex = -1 Then
            MsgBox "�úű�ÿ���Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
            cbo��.SetFocus: Exit Sub
        End If
        If chk��ſ���.Value = 1 Then
            If Val(txt�޺�.Text) = 0 And Val(txt��Լ.Text) = 0 Then
                MsgBox "ʹ����ſ���ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
                txt�޺�.SetFocus: Exit Sub
            End If
        End If
        '�޺���Լ����
        If Trim(txt�޺�.Text) <> "" Then
            If Trim(txt��Լ.Text) <> "" And Val(txt�޺�.Text) < Val(txt��Լ.Text) Then
                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
                txt��Լ.SetFocus: Exit Sub
            End If
        ElseIf Trim(txt��Լ.Text) <> "" Then
            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
            txt�޺�.SetFocus: Exit Sub
        End If
    Else
        With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))
                    If chk��ſ���.Value = 1 Then
                          If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                              MsgBox "ʹ����ſ���ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
                              .Row = 2: .Col = i
                              .SetFocus: Exit Sub
                          End If
                      End If
                        '�޺���Լ����
                        If Val(.TextMatrix(2, i)) <> 0 Then
                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
                                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
                                .Row = 2: .Col = i
                                .SetFocus: Exit Sub
                            End If
                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
                            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Sub
                        End If
                End If
            Next
            If strTmp = "" Then
                MsgBox "�úű�ÿ�ܵ�Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
                vsPlan.SetFocus: Exit Sub
            End If
        End With
    End If
    '�����ж�
    If opt����(1).Value Or opt����(2).Value Or opt����(3).Value Then
        intCount = 0
        For i = 1 To lvwDept.ListItems.Count
            If lvwDept.ListItems(i).Checked Then intCount = intCount + 1
        Next
        If opt����(1).Value Then
            If intCount = 0 Then
                MsgBox "ָ������ʱ����ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
                lvwDept.SetFocus: Exit Sub
            ElseIf intCount > 1 Then
                MsgBox "ָ������ʱֻ��ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
                lvwDept.SetFocus: Exit Sub
            End If
        ElseIf opt����(2).Value Or opt����(3).Value Then
            If intCount < 2 Then
                MsgBox "��̬�����ƽ������ʱ����Ҫѡ��������Ӧ���������ң�", vbInformation, gstrSysName
                lvwDept.SetFocus: Exit Sub
            End If
        End If
    End If

    '��Ŀ�۸��ж�
    If ReadRegistPrice(cboItem.ItemData(cboItem.ListIndex), False, False) = 0 Then
        MsgBox "��Ŀ""" & cboItem.Text & """δ������Ч�۸�,���ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Sub
    End If

    'ȡҽ��ID
    If cboDoctor.ListIndex <> -1 Then lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
'    '����:����һ��ҽ�����Լ����ظ�����
'    If zlCheckPlanArrageIsValied = False Then
'        If cboDoctor.Enabled Then cboDoctor.SetFocus
'        Exit Sub
'    End If
'
'    If zlCheckRegistPlanIsValied(blnMulitNumPlan) = False Then
'        If cboDoctor.Enabled Then cboDoctor.SetFocus
'        Exit Sub
'    End If
    '�Ƿ�ͬһҽ���İ���ʱ����Ƿ��ظ��򽻲�
    If Trim(cboDoctor.Text) <> "" Then
        Set rsDoctorPlan = GetDoctorPlan(lngҽ��ID, cboDoctor.Text)
        If rsDoctorPlan.RecordCount > 0 Then
            strSQL = "Select ʱ���, ��ʼʱ��, Decode(Sign(��ֹʱ�� - ��ʼʱ��), 1, ��ֹʱ�� , ��ֹʱ��+ 1) ��ֹʱ�� From ʱ���"
            Set rsNewDate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            Set rsOldDate = rsNewDate.Clone
        End If

        strInfo = ""
        For j = 1 To rsDoctorPlan.RecordCount
            strTmp = ""
            For i = 0 To IIf(opt��.Value, 6, vsPlan.Cols - 2)
               strOld = "" & rsDoctorPlan.Fields("D" & i).Value
               If opt��.Value Then
                   strNew = cbo��.Text
               Else
                   strNew = Trim(vsPlan.TextMatrix(1, i + 1))
               End If

               rsNewDate.Filter = "ʱ���='" & strNew & "'"
               rsOldDate.Filter = "ʱ���='" & strOld & "'"
               If rsNewDate.RecordCount > 0 And rsOldDate.RecordCount > 0 Then
                    If rsNewDate!��ʼʱ�� >= rsOldDate!��ʼʱ�� And rsNewDate!��ʼʱ�� <= rsOldDate!��ֹʱ�� Or rsNewDate!��ֹʱ�� >= rsOldDate!��ʼʱ�� And rsNewDate!��ֹʱ�� <= rsOldDate!��ֹʱ�� Or rsNewDate!��ʼʱ�� <= rsOldDate!��ʼʱ�� And rsNewDate!��ֹʱ�� >= rsOldDate!��ֹʱ�� Then
                    'ʱ�佻��,���ж�Ч���Ƿ񽻲�
                         If chk��Ч��.Value = 0 Then
                             strTmp = strTmp & "," & "����" & Choose(i + 1, "��", "һ", "��", "��", "��", "��", "��") & ":" & strOld
                         Else
                             'Ϊ���ж�,�ٶ����ݰ��淶����,��ʼʱ��ͽ���ʱ��,Ҫô����,Ҫô��û��,���Խ��Կ�ʼʱ�����ж�����
                             If IsNull(rsDoctorPlan!��ʼʱ��) Then
                                 strTmp = strTmp & "," & "����" & Choose(i + 1, "��", "һ", "��", "��", "��", "��", "��") & ":" & strOld
                             Else
                                 If dtpBegin.Value >= CDate(rsDoctorPlan!��ʼʱ��) And dtpBegin.Value <= CDate(Nvl(rsDoctorPlan!��ֹʱ��, "3000-01-01")) Or dtpEnd.Value >= CDate(rsDoctorPlan!��ʼʱ��) And dtpEnd.Value <= CDate(Nvl(rsDoctorPlan!��ֹʱ��, "3000-01-01")) Or dtpBegin.Value <= CDate(rsDoctorPlan!��ʼʱ��) And dtpEnd.Value >= CDate(Nvl(rsDoctorPlan!��ֹʱ��, "3000-01-01")) Then
                                    strTmp = strTmp & "," & "����" & Choose(i + 1, "��", "һ", "��", "��", "��", "��", "��") & ":" & strOld
                                 End If
                             End If
                         End If
                    End If
               End If
            Next
            If strTmp <> "" Then
                strInfo = strInfo & vbCrLf & "�ںű� [" & rsDoctorPlan!���� & "] ���������°���:" & vbCrLf & "        " & Mid(strTmp, 2)
                If Not IsNull(rsDoctorPlan!��ʼʱ��) Then
                    strInfo = strInfo & vbCrLf & "        ��Ч��:" & rsDoctorPlan!��ʼʱ�� & "~" & rsDoctorPlan!��ֹʱ��
                Else
                    strInfo = strInfo & vbCrLf & "        ��Ч��:����"
                End If
            End If
            rsDoctorPlan.MoveNext
        Next
        If strInfo <> "" Then
            If blnMulitNumPlan Then
                '��ΰ���ʱ,���ܴ��ڽ���
                Call MsgBox("����" & cboDoctor.Text & "ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ���" & vbCrLf & strInfo & vbCrLf & vbCrLf & "���ܰ���!", vbInformation + vbOKOnly, gstrSysName)
                Exit Sub
            Else
                If MsgBox("����" & cboDoctor.Text & "ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ���" & vbCrLf & strInfo & vbCrLf & vbCrLf & "ȷʵҪ���浱ǰ�ű���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If

    If Not mEditType = edt_���� Then
        If CheckExistsBooking() Then
            If MsgBox("�úű�ǰ���ŵ�ʱ���֮�����ԤԼ�Һŵ�,�Ƿ�Ҫ����?", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    '�ȼ��
    'ȡʱ���
    str�޺� = ""
    If opt��.Value Then 'ÿ��
        For i = 1 To 7
            strʱ��� = strʱ��� & "'" & Trim(cbo��.Text) & "',"
            str�޺� = str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
            str�޺� = str�޺� & "," & Val(txt�޺�.Text) & "," & Val(txt��Լ.Text)
        Next
    Else
        For i = 1 To vsPlan.Cols - 1
            strʱ��� = strʱ��� & "'" & Trim(vsPlan.TextMatrix(1, i)) & "',"

                str�޺� = str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                    str�޺� = str�޺� & ",0,0"
                Else
                    str�޺� = str�޺� & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Val(Trim(vsPlan.TextMatrix(3, i)))
                End If
        Next
    End If
    If str�޺� <> "" Then str�޺� = Mid(str�޺�, 2)


    'ȡ�Һ�����
    For i = 1 To lvwDept.ListItems.Count
        If lvwDept.ListItems(i).Checked Then
            str���� = str���� & ";" & lvwDept.ListItems(i).Text
        End If
    Next
    str���� = Mid(str����, 2)


    'ȡ���﷽ʽ
    intCount = 0
    For i = 0 To opt����.UBound
        If opt����(i).Value Then intCount = i: Exit For
    Next

    'ȡ��ʼʱ�䷶Χ
    strBegin = "NULL": strEnd = "NULL"
    If chk��Ч��.Value = 1 Then
        strBegin = "To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strEnd = "To_Date('" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    End If

      '�鿴�Ƿ�ı����Ű���� �ı��� �޺��� ��Լ�� ������ſ���
    blnChange = (str�޺� <> mPlanInfo.str�޺�) Or (strʱ��� <> mPlanInfo.str�Ű�)
    blnChange = blnChange Or (chk��ſ���.Value <> IIf(mPlanInfo.bln���, 1, 0))
    str�޺� = "'" & str�޺� & "',"
    Set cllPro = New Collection
    'ȡID
    If mEditType = edt_���� Then

        '����
        lngNextID = zlDatabase.GetNextId("�ҺŰ���")

        strSQL = "zl_�ҺŰ���_INSERT(" & _
            lngNextID & ",'" & Trim(txt�ű�.Text) & "','" & cbo����.Text & "'," & _
            cbo����.ItemData(cbo����.ListIndex) & "," & _
            cboItem.ItemData(cboItem.ListIndex) & ",'" & Trim(cboDoctor.Text) & "'," & _
            lngҽ��ID & "," & _
            chk����.Value & "," & strʱ��� & str�޺� & intCount & "," & _
            "'" & str���� & "'," & strBegin & "," & strEnd & ",1," & chk��ſ���.Value & ",0," & 5 & ")"
    Else
'
' Zl_�ҺŰ���_Insert
'(
'  Id_In       �ҺŰ���.ID%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����id_In   �ҺŰ���.����id%Type,
'  ��Ŀid_In   �ҺŰ���.��Ŀid%Type,
'  ҽ��_In     �ҺŰ���.ҽ������%Type,
'  ҽ��id_In   �ҺŰ���.ҽ��id%Type,
'  ��������_In �ҺŰ���.��������%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ��һ_In     �ҺŰ���.��һ%Type,
'  �ܶ�_In     �ҺŰ���.�ܶ�%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����_In     �ҺŰ���.����%Type,
'  �޺ſ���_In Varchar2,
'  ���﷽ʽ_In �ҺŰ���.���﷽ʽ%Type,
'  ����_In     Varchar2,
'  ��ʼʱ��_In �ҺŰ���.��ʼʱ��%Type,
'  ��ֹʱ��_In �ҺŰ���.��ֹʱ��%Type,
'  ����_In     Number,
'  ��ſ���_In �ҺŰ���.��ſ���%Type,
'  ��������_In Number:=0,
'  Ĭ��ʱ�μ��_In �ҺŰ���.Ĭ��ʱ�μ��%Type
') As
'  -----------------------------------------------------------
'  --������
'  --  ����_IN=��';'�ŷָ��Ķ����������
'  --  �޺ſ���_IN:|��һ,22(�޺�),13(��Լ)|�ܶ�,20(�޺�),11(��Լ)....
'  --  ��������_IN:�޸İ���ʱ ��ʱ�����ݵĴ��� 0--������ 1--ɾ��ʱ����Ϣ
        '�޸�

        lngNextID = mlngID
        strSQL = "    " & vbNewLine & "zl_�ҺŰ���_INSERT("
        strSQL = strSQL & vbNewLine & lngNextID
        strSQL = strSQL & vbNewLine & ",'" & (txt�ű�.Text) & "','" & cbo����.Text & "',"
        strSQL = strSQL & vbNewLine & cbo����.ItemData(cbo����.ListIndex) & ","
        strSQL = strSQL & vbNewLine & cboItem.ItemData(cboItem.ListIndex) & ",'" & Trim(cboDoctor.Text) & "',"
        strSQL = strSQL & vbNewLine & lngҽ��ID & "," & chk����.Value & ","
        strSQL = strSQL & vbNewLine & strʱ��� & str�޺� & intCount & ","
        strSQL = strSQL & vbNewLine & "'" & str���� & "'," & strBegin & "," & strEnd & ",0," & chk��ſ���.Value & ","
        strSQL = strSQL & vbNewLine & 0 & ","
        strSQL = strSQL & vbNewLine & 5 & ")"


    End If

    On Error GoTo errH
    zlAddArray cllPro, strSQL

   ' Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    LoadTimePlan True

    mfrmTime.zlSaveData lngNextID, cllPro
    zlExecuteProcedureArrAy cllPro, Me.Caption
    On Error GoTo 0
    mblnSucces = True

    If mEditType <> edt_���� Then Unload Me: Exit Sub
    Call LoadCard
    mblnChangeByCode = True
    tbPage.Item(mPageIndex.EM_����).Selected = True
    mblnChangeByCode = False
    Call mfrmTime.ClearCustomData
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadRegHistory() As Boolean
    Dim strSQL As String
    strSQL = " Select Decode(To_Char(a.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',"
    strSQL = strSQL & vbCrLf & "                       '7', '����') As ������Ŀ, Max(Nvl(a.����, 0)) As ������, Count(1) As ͳ��,to_char(Max(����ʱ��),'hh24:mi:ss') as ����ʱ��"
    strSQL = strSQL & vbCrLf & " From ���˹Һż�¼ a, �ҺŰ��� b"
    strSQL = strSQL & vbCrLf & " Where a.��¼״̬ = 1 And a.����ʱ�� Between Sysdate And Sysdate + " & IIf(gintԤԼ���� = 0, 15, gintԤԼ����) & " And a.�ű� = b.���� And b.Id=[1]"
    strSQL = strSQL & vbCrLf & " Group By Decode(To_Char(a.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',"
    strSQL = strSQL & vbCrLf & "                             '7', '����')"

    On Error GoTo Hd:
    Set mrsRegHistory = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    LoadRegHistory = True
Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
