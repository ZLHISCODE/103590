VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmResistPlanTimeSet 
   BorderStyle     =   0  'None
   Caption         =   "��ʱ������"
   ClientHeight    =   7935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdOther 
      Caption         =   "������������(&T)"
      Height          =   350
      Left            =   3840
      TabIndex        =   14
      ToolTipText     =   "������¼���ʱ��"
      Top             =   0
      Width           =   1515
   End
   Begin VB.Frame fraӦ���� 
      Caption         =   "Ӧ���ڡ�"
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   7320
      Width           =   7755
      Begin VB.OptionButton optӦ���� 
         Caption         =   "��ҽ��(����)"
         Height          =   255
         Index           =   1
         Left            =   2115
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "������"
         Height          =   255
         Index           =   0
         Left            =   795
         TabIndex        =   12
         Top             =   255
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "������(�ڿ�)"
         Height          =   255
         Left            =   3870
         TabIndex        =   10
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "���кű�"
         Height          =   255
         Left            =   5685
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.PictureBox picPage 
      BorderStyle     =   0  'None
      Height          =   3540
      Index           =   0
      Left            =   795
      ScaleHeight     =   3540
      ScaleWidth      =   2535
      TabIndex        =   7
      Top             =   600
      Width           =   2535
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   4875
      Left            =   525
      TabIndex        =   6
      Top             =   2010
      Width           =   2535
      _Version        =   589884
      _ExtentX        =   4471
      _ExtentY        =   8599
      _StockProps     =   64
   End
   Begin VB.CommandButton cmd����ʱ�� 
      Caption         =   "��������(&F)"
      Height          =   350
      Left            =   2385
      TabIndex        =   0
      ToolTipText     =   "������¼���ʱ��"
      Top             =   0
      Width           =   1150
   End
   Begin MSComCtl2.UpDown udTime 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   30
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtTimeOut"
      BuddyDispid     =   196618
      OrigLeft        =   2025
      OrigTop         =   3
      OrigRight       =   2280
      OrigBottom      =   348
      Max             =   1440
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTime 
      Height          =   7545
      Index           =   0
      Left            =   3465
      TabIndex        =   2
      Top             =   900
      Width           =   5100
      _cx             =   8996
      _cy             =   13309
      Appearance      =   0
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmResistPlanTimeSet.frx":0000
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
      Begin VB.CommandButton cmdɾ�� 
         Caption         =   "ɾ"
         Height          =   255
         Index           =   0
         Left            =   4320
         TabIndex        =   13
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdԤԼ 
         Caption         =   "Ԥ"
         Height          =   255
         Index           =   0
         Left            =   2685
         TabIndex        =   3
         Top             =   2535
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   300
      Left            =   1335
      MaxLength       =   4
      TabIndex        =   5
      Text            =   "10"
      Top             =   30
      Width           =   465
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "ʱ����(��)"
      Height          =   180
      Left            =   225
      TabIndex        =   4
      Top             =   85
      Width           =   1080
   End
End
Attribute VB_Name = "frmResistPlanTimeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit 'Ҫ���������
'Public Enum gPlanEditType
'    EM_����_���� = 0
'    EM_����_�޸�
'    EM_����_����
'    EM_�ƻ�_���� = 11
'    EM_�ƻ�_�޸�
'    EM_�ƻ�_����
'End Enum
Private mEditType         As gPlanEditType
Private mstr���� As String  '��һ,�޺���,��Լ��|�ܶ�,�޺���,��Լ��|....
Private mbln��ſ��� As Boolean
Private mlngSelIndex As Long '�����ڱ༭״̬������
Private mblnOnChange As Boolean  '�Ƿ���봥��tbpage��SelectedChanged�¼�
Private mlng����ID As Long '
Private mlng�ƻ�Id As Long '
Private mblnInit As Boolean  '�Ƿ�����˳�ʼ���ĵ���
'Private mrsTime          As ADODB.Recordset
Private mrs�޺�          As ADODB.Recordset
Private mrs�ϰ�ʱ���    As ADODB.Recordset
Private mrs����          As ADODB.Recordset
Private mrsRegPlan       As ADODB.Recordset ' �޸�
Private mrsAssign        As ADODB.Recordset '�ѷ������
Private mblnCellChange   As Boolean
Private mstrKey         As String
Public mblnChange       As Boolean '�Ƿ�ı�������
Private mblnReload      As Boolean '�ڹҺŰ��Ź���ҳ����� ShowMe�Ժ� �Ƿ���Ҫˢ��
Private mstr�����޸� As String '��ĳһ����߶���İ������Ƹ���
Private mrsHistory As ADODB.Recordset 'ԤԼ�Һ�ͳ����Ϣ
Private WithEvents mfrmOtherCalc As frmRegistPlanTimeOther  '
Attribute mfrmOtherCalc.VB_VarHelpID = -1
'�����ϰ�ʱ��
Private Type t_�ϰ�ʱ��
  dat_�����ϰ� As Date
  dat_�����°� As Date
  dat_�����ϰ� As Date
  dat_�����°� As Date
End Type
Private t_ʱ�� As t_�ϰ�ʱ��
Private Const strMaskKey As String = "09:00-09:00"
Private mstrӦ��ʱ�� As String
Public Event zlSaveTimePageSelected(ByVal str���� As String)
Private mblnNotBrush As Boolean '���ǽ���ˢ�²���

Private Sub LoadPageControl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҳ�ؼ�
    '����:���˺�
    '����:2012-06-15 13:33:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    For i = 1 To 6
        Load picPage(i): Load vsTime(i)
        Load cmdԤԼ(i): Load cmdɾ��(i)
       ' cmdԤԼ(i).Visible = True
        Set cmdԤԼ(i).Container = vsTime(i)
        Set cmdɾ��(i).Container = vsTime(i)
        'cmdɾ��(i).Visible = True
        picPage(i).Visible = True: vsTime(i).Visible = True
        Set vsTime(i).Container = picPage(i)
    Next
    Set vsTime(0).Container = picPage(0)
    Call LoadPage
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : ClearCustomData
'Description : ��ձ�����Ϣ
'Author      : ��⸣
'Date        : 05-November-2012 14:58:54
'Input       :
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Public Sub ClearCustomData()
     mstr���� = ""
     mbln��ſ��� = False
     mlngSelIndex = 0
     mblnOnChange = False
     mlng����ID = 0
     mlng�ƻ�Id = 0
     mblnInit = False
     Set mrs�޺� = Nothing
     Set mrsRegPlan = Nothing
     Set mrsAssign = Nothing
     mstrKey = ""
     mblnChange = False
     mstr�����޸� = ""
     Set mrsHistory = Nothing
End Sub

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : LoadPage
'Description : ����ҳ����Ϣ
'Author      : ��⸣
'Date        : 05-November-2012 14:59:21
'Input       :
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-

Private Function LoadPage() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҳ
    '����:���˺�
    '����:2012-06-15 13:37:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Dim strTemp As String
    On Error GoTo errHandle
    
    tbPage.RemoveAll
    For i = 0 To 6
        strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
        Set ObjItem = tbPage.InsertItem(i + 1, strTemp, picPage(i).hwnd, 0)
        ObjItem.Tag = strTemp
    Next
     With tbPage
         
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionBottom
    End With
    LoadPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
 Private Sub ShowPage()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾҳ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-11-26 15:21:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, i As Long
    Dim j As Long, lngIndex As Long, p As Long, strTemp As String
    
    For j = 0 To tbPage.ItemCount - 1
         tbPage(j).Visible = False: tbPage(j).Enabled = False
         tbPage(j).Selected = False
    Next
    
    On Error GoTo errHandle
    varData = Split(mstr����, "|")
    lngIndex = -1: mlngSelIndex = -1
    For i = 0 To UBound(varData)
        ''��һ,�޺���,��Լ��|�ܶ�,�޺���,��Լ��|....
        varTemp = Split(varData(i) & ",,,,", ",")
        If varTemp(0) <> "" Then
            For j = 0 To tbPage.ItemCount - 1
                If tbPage(j).Tag = varTemp(0) Then
                    If lngIndex < 0 Then lngIndex = j
                    tbPage(j).Visible = True: tbPage(j).Enabled = True
                    p = GetVsGridIndex(varTemp(0))
                    vsTime(p).Tag = varTemp(1) & "," & varTemp(2)
                    If mlngSelIndex = -1 Then mlngSelIndex = j: tbPage(j).Selected = True
                End If
            Next
        End If
    Next
    If mlngSelIndex = -1 Then mlngSelIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub
 
 Private Function GetVsGridIndex(ByVal str���� As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������
    '����:���˺�
    '����:2012-06-15 14:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    str���� = Switch(str���� = "����", 0, str���� = "��һ", 1, str���� = "�ܶ�", 2, str���� = "����", 3, str���� = "����", 4, str���� = "����", 5, str���� = "����", 6, True, 0)
    GetVsGridIndex = Val(str����)
 End Function
 
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : GetVsGridCaption
'Description : ����������ȡ������Ŀ
'Author      : ��⸣
'Date        : 05-November-2012 15:02:14
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'nIndex            Integer           ByVal                .����ֵ
'Output      :     ��Ӧ������
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
 
 Private Function GetVsGridCaption(ByVal nIndex As Integer) As String
    '����:����������ȡ������Ŀ
    Dim str���� As String
    str���� = Switch(nIndex = 0, "����", nIndex = 1, "��һ", nIndex = 2, "�ܶ�", nIndex = 3, "����", nIndex = 4, "����", nIndex = 5, "����", nIndex = 6, "����", True, "")
    GetVsGridCaption = str����
 End Function
 
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : zlShowPagePlan
'Description : ��ʾҳ��
'Author      : ��⸣
'Date        : 05-November-2012 15:03:02
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'str�������       String            ByVal                .
'rsRegPlan         ADODB.Recor...    ByVal                .
'rsHistory         ADODB.Recor...    ByRef                .
'bln��ſ���       Boolean           ByVal                .
'bytType           gPlanEditType     ByVal                .
'lng����ID         Long              ByVal                .
'lng�ƻ�ID         Long              ByVal                .
'blnBeforCheck     Boolean = F...    ByVal                .
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
 Public Sub zlShowPagePlan(ByVal str������� As String, ByVal rsRegPlan As ADODB.Recordset, ByRef rsHistory As ADODB.Recordset, _
                        ByVal bln��ſ��� As Boolean, ByVal BytType As gPlanEditType, Optional ByVal lng����ID As Long, _
                        Optional ByVal lng�ƻ�ID As Long, Optional ByVal blnBeforCheck As Boolean = False, Optional ByVal strӦ��ʱ�� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾҳ��
    '����:���˺�
    '����:2012-06-15 13:49:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    mstr���� = str�������
    mstrӦ��ʱ�� = strӦ��ʱ��
                                                                          
    Set mrsRegPlan = rsRegPlan
    If bln��ſ��� <> mbln��ſ��� And Not mrsAssign Is Nothing Then
         mrsAssign.Filter = 0
         Do While Not mrsAssign.EOF
            mrsAssign.Delete
            mrsAssign.MoveNext
         Loop
         If blnBeforCheck Then Exit Sub
    End If
'    mlngSelIndex = -1
     mEditType = BytType: mlng����ID = lng����ID: mlng�ƻ�Id = lng�ƻ�ID
    Set mrsHistory = rsHistory
    If Not blnBeforCheck Then Call ShowPage
    If mblnInit Then
        Call AssignManage
    End If
    mblnInit = True
    Call InitRs(mbln��ſ��� = bln��ſ���)
    mbln��ſ��� = bln��ſ���
    If blnBeforCheck Then Exit Sub
    For i = 0 To 6
       If tbPage.Item(i).Selected Then
            Call tbPage_SelectedChanged(tbPage.Item(i))
            Exit For
       End If
    Next
 End Sub
 
 
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : SavePlanData
'Description : ����ʱ�����ݱ���
'Author      : ��⸣
'Date        : 05-November-2012 15:04:05
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'lngID             Long              ByVal                .����ID
'cllPro            Collection        ByRef                .������ر������ݵ�SQL
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
 Private Function SavePlanData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
    Dim i As Long, str���� As String, lng��� As String, strSQL As String
    Dim str���s As String, BytType As Byte 'Ӧ����
    Dim bytRowStep As Byte, bytStepCol As Byte
    Dim intPage As Integer, cllPage As Collection
    Dim strʱ�� As String
    Dim strProc As String
    Dim strTmp As String
    Dim strTemp As String
    Dim p As Integer, j As Long
   
    On Error GoTo errHandle
    
    Call AssignManage  '��ŷ��䴦��
    If cllPro Is Nothing Then
        Set cllPro = New Collection
    End If
    strSQL = "Zl_�Һżƻ�ʱ��_Delete(" & lngID & ")"
    zlAddArray cllPro, strSQL
    For i = 0 To 6
        strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
        mrsAssign.Filter = "������Ŀ='" & strTemp & "'"
        If mrsAssign.RecordCount > 0 Then
            Do While Not mrsAssign.EOF
    '            ���,��ʼʱ��,����ʱ��,��������,ԤԼ��־|...
                strTmp = mrsAssign!���
                strTmp = strTmp & "," & mrsAssign!��ʼʱ�� & "," & mrsAssign!����ʱ�� & "," & mrsAssign!�������� & "," & mrsAssign!�Ƿ�ԤԼ
                If Len(strʱ�� & "|" & strTmp) > 4000 Then
                    strʱ�� = Mid(strʱ��, 2)
                    strSQL = "  Zl_�Һżƻ�ʱ��_Insert("
                    '  ����id_In �ҺŰ���ʱ��.����id%Type,
                    strSQL = strSQL & lngID & ","
                    '  ����_In   �ҺŰ���ʱ��.����%Type,
                    strSQL = strSQL & "'" & strTemp & "',"
                    '  ʱ��_In   Varchar2,
                    strSQL = strSQL & "'" & strʱ�� & "'"
                    strSQL = strSQL & "" & ")"
                    zlAddArray cllPro, strSQL
                    strʱ�� = ""
                End If
                strʱ�� = strʱ�� & "|" & strTmp
                mrsAssign.MoveNext
            Loop
            If strʱ�� <> "" Then
                 
                strʱ�� = Mid(strʱ��, 2)
                strSQL = "  Zl_�Һżƻ�ʱ��_Insert("
                '  ����id_In �ҺŰ���ʱ��.����id%Type,
                strSQL = strSQL & lngID & ","
                '  ����_In   �ҺŰ���ʱ��.����%Type,
                strSQL = strSQL & "'" & strTemp & "',"
                '  ʱ��_In   Varchar2,
                strSQL = strSQL & "'" & strʱ�� & "'"
                strSQL = strSQL & "" & ")"
                zlAddArray cllPro, strSQL
                strʱ�� = ""
            End If
        
        End If
    Next
    SavePlanData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

 End Function

Private Function SaveData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ�����ݱ���
    '���:lngID-����ID
    '����:cllPro-������ر������ݵ�SQL
    '����:���˺�
    '����:2012-06-15 13:18:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str���� As String, lng��� As String, strSQL As String
    Dim str���s As String, BytType As Byte 'Ӧ����
    Dim bytRowStep As Byte, bytStepCol As Byte
    Dim intPage As Integer
    Dim strʱ�� As String
    Dim strProc As String
    Dim strTmp As String
    Dim strTemp As String
    Dim p As Integer, j As Long
     
   
    On Error GoTo errHandle
      
    Call AssignManage  '��ŷ��䴦��
    If cllPro Is Nothing Then
        Set cllPro = New Collection
    End If
    strSQL = "Zl_�ҺŰ���ʱ��_Delete(" & lngID & ")"
    zlAddArray cllPro, strSQL
    For i = 0 To 6
        strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
        mrsAssign.Filter = "������Ŀ='" & strTemp & "'"
        If mrsAssign.RecordCount > 0 Then
            Do While Not mrsAssign.EOF
    '            ���,��ʼʱ��,����ʱ��,��������,ԤԼ��־|...
                strTmp = mrsAssign!���
                strTmp = strTmp & "," & mrsAssign!��ʼʱ�� & "," & mrsAssign!����ʱ�� & "," & mrsAssign!�������� & "," & mrsAssign!�Ƿ�ԤԼ
                If Len(strʱ�� & "|" & strTmp) > 4000 Then
                    strʱ�� = Mid(strʱ��, 2)
                    strSQL = "  Zl_�ҺŰ���ʱ��_Insert("
                    '  ����id_In �ҺŰ���ʱ��.����id%Type,
                    strSQL = strSQL & lngID & ","
                    '  ����_In   �ҺŰ���ʱ��.����%Type,
                    strSQL = strSQL & "'" & strTemp & "',"
                    '  ʱ��_In   Varchar2,
                    strSQL = strSQL & "'" & strʱ�� & "'"
                    strSQL = strSQL & "" & ")"
                    zlAddArray cllPro, strSQL
                    strʱ�� = ""
                End If
                strʱ�� = strʱ�� & "|" & strTmp
                mrsAssign.MoveNext
            Loop
            If strʱ�� <> "" Then
                 
                strʱ�� = Mid(strʱ��, 2)
                strSQL = "  Zl_�ҺŰ���ʱ��_Insert("
                '  ����id_In �ҺŰ���ʱ��.����id%Type,
                strSQL = strSQL & lngID & ","
                '  ����_In   �ҺŰ���ʱ��.����%Type,
                strSQL = strSQL & "'" & strTemp & "',"
                '  ʱ��_In   Varchar2,
                strSQL = strSQL & "'" & strʱ�� & "'"
                strSQL = strSQL & "" & ")"
                zlAddArray cllPro, strSQL
                strʱ�� = ""
            End If
        
        End If
    Next
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Public Function zlSaveData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '���:
    '����:cllPro-������ر������ݵ�SQL
    '����:���˺�
    '����:2012-06-15 13:18:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mblnInit Then Exit Function
    If zl_CheckMoveAssign() = False Then Exit Function
    If VsTimeValidate(-1) = False Then Exit Function
    
    If mEditType = EM_����_�޸� Or mEditType = EM_����_���� Then
        If SaveData(lngID, cllPro) = False Then Exit Function
    Else
        If SavePlanData(lngID, cllPro) = False Then Exit Function
    End If
    
    zlSaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
'
'
'
'Private Sub cmd����ʱ��_Click()
''�ԹҺŰ���ʱ�ν�������
'    Dim str����         As String
'
'    cmdԤԼ.Visible = False
'    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
'    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
'    mrsTime.Filter = "����='" & str���� & "'"
'    If mrsTime.RecordCount > 0 Then
'      '****************************************************************
'      '�����йҺŰ���ʱ�ε������
'      '��ʾ����Ա �Ƿ���Ҫ���¼���ʱ��
'      '****************************************************************
'        If MsgBox("�˰�����" & str���� & "�Ѿ�����ʱ�� " & vbCrLf & "�Ƿ����¼���ʱ��?", vbDefaultButton2 + vbYesNo, Me.Caption) = vbNo Then
'            mrsTime.Filter = 0
'            Exit Sub
'        End If
'    End If
'    Select Case chk��ſ���.Value = 1
'    Case True:
'        Setר�Һ�ʱ��
'        setVsFlexBgColor (True)
'    Case False:
'        Set��ͨ��ʱ��
'        setVsFlexBgColor
'    End Select
'
'    mblnChange = True
'End Sub
'Private Sub Set��ͨ��ʱ��()
'    Dim strSQL      As String
'    Dim str����     As String
'    Dim strʱ��     As String
'    Dim lng�޺�     As Long
'    Dim lng��Լ     As Long
'    Dim lng���     As Long
'    Dim dblDatCount As Long '��ʱ����
'    Dim datʱ��     As Date 'ÿ��ʱ��ε�
'    Dim blnȫ��     As Boolean  '�Ƿ���ȫ�춼����Һ� �����ȫ�����Ϊ���������
'    Dim datStart    As Date
'    Dim datEnd      As Date
'    Dim i           As Long
'    Dim j           As Long
'    Dim lngRow      As Long
'    Dim lngCol      As Long
'    Dim strData     As String
'    Dim strTime     As String
'    Dim strList()   As String
'    Dim blnExit     As Boolean
'    Dim lngIndex    As Long
'    Dim lngStart    As Long
'    On Error GoTo Hd
'    If mrs�ϰ�ʱ��� Is Nothing Then Exit Sub
'    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
'    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
'    mrs�޺�.Filter = "����='" & str���� & "'"
'    If mrs�޺�.RecordCount = 0 Then
'        MsgBox "��ǰ�ű���" & str���� & ",û�ж�Ӧ�ĹҺŰ�������" & vbCrLf & "�뵽�ҺŰ���������!", vbOKOnly, Me.Caption
'        Exit Sub '����ҺŰ�����û�����ô������Ϣ �Ͳ���������
'    End If
'    lng�޺� = Nvl(mrs�޺�!�޺���, 0): lng��Լ = Nvl(mrs�޺�!��Լ��, 0)
'    If lng�޺� = 0 Then
'        MsgBox "��ǰ�ű���" & str���� & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
'        Exit Sub
'    End If
'    Me.txt�޺�.Text = lng�޺�
'    Me.txt��Լ.Text = lng��Լ
'    If lng��Լ = 0 Then lng��Լ = lng�޺� '�����ԤԼû����������Ϊ�����Լ�����޺�����ͬ
'    strʱ�� = Nvl(mrs����(str����).Value)
'    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
'
'    '*********************************
'    '��ʱ�ξ��崦�� ��Ϊȫ��ͷ�ȫ��
'    'ȫ���Ϊ���������
'    '*********************************
'
'    lng��� = Val(txtTimeOut.Text)
'
'    With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 400: .RowHeightMin = 400
'        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'    End With
'    '*************************************
'    '��ͨ��
'    '*************************************
'    With vsTime
'        .Cols = 8: .FixedCols = 0
'        .Rows = 1: .FixedRows = 1
'        For i = 0 To .Cols - 1 Step 2
'           .TextMatrix(0, i) = "ʱ���"
'        Next
'        For i = 1 To .Cols - 1 Step 2
'           .TextMatrix(0, i) = "ԤԼ����"
'        Next
'        lngRow = 1: lngCol = -1
'        j = 1: lngStart = 1
'        Do While Not mrs�ϰ�ʱ���.EOF
'            If blnExit Then Exit Do
'            datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00"))
'            For i = j To lng�޺�
'                If lngStart > lng�޺� Then
'                    blnExit = True
'                    Exit For
'                End If
'
'                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
'                    j = i
'                    Exit For
'                End If
'
'                lngCol = lngCol + 1
'                If lngCol * 2 > .Cols - 2 Then lngRow = lngRow + 1: lngCol = 0
'                strData = IIf(lng��Լ >= i, 1, 0)
'                strTime = Format(datʱ��, "HH:mm") & "-" & _
'                      IIf(Format(DateAdd("n", lng���, datʱ��), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
'                      Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
'
'                If lngRow > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(lngRow, lngCol * 2) = strTime
'                .TextMatrix(lngRow, lngCol * 2 + 1) = strData
'                lngStart = lngStart + 1
'                datʱ�� = DateAdd("n", lng���, datʱ��)
'            Next
'            mrs�ϰ�ʱ���.MoveNext
'        Loop
'
'
'         For i = 0 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'         Next
'         .Redraw = flexRDBuffered
'    End With
'
'Exit Sub
'Hd:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Sub
'Private Sub Setר�Һ�ʱ��()
'    Dim strSQL      As String
'    Dim str����     As String
'    Dim strʱ��     As String
'    Dim lng�޺�     As Long
'    Dim lng��Լ     As Long
'    Dim lng���     As Long
'    Dim dblDatCount As Long '��ʱ����
'    Dim datʱ��     As Date 'ÿ��ʱ��ε�
'    Dim strʱ��     As String
'    Dim blnȫ��     As Boolean  '�Ƿ���ȫ�춼����Һ� �����ȫ�����Ϊ���������
'    Dim datStart    As Date
'    Dim datEnd      As Date
'    Dim i           As Long
'    Dim j           As Long
'    Dim lngRow      As Long
'    Dim lngCol      As Long
'    Dim strData     As String
'    Dim strTime     As String
'    Dim strList()   As String
'    Dim blnExit     As Boolean
'    Dim lngIndex    As Long
'    Dim lngStart    As Long
'    On Error GoTo Hd
'    If mrs�ϰ�ʱ��� Is Nothing Then Exit Sub
'    If mrs�޺� Is Nothing Then
'        strSQL = _
'        "Select ����id, ������Ŀ as ���� , �޺���, ��Լ�� From �ҺŰ������� Where ����id = [1]"
'        Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Nvl(txt�ű�.Tag))
'        If mrsTime.RecordCount = 0 Then
'        MsgBox "��ǰ�ű�û�ж�Ӧ�ĹҺŰ�������" & vbCrLf & "�뵽�ҺŰ���������!", vbOKOnly, Me.Caption
'        Set mrs�޺� = Nothing
'        Exit Sub '����ҺŰ�����û�����ô������Ϣ �Ͳ���������
'    End If
'    End If
'    If tbWeekTime.SelectedItem Is Nothing Then Exit Sub
'    str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
'    mrs�޺�.Filter = "����='" & str���� & "'"
'    If mrs�޺�.RecordCount = 0 Then
'        MsgBox "��ǰ�ű���" & str���� & ",û�ж�Ӧ�ĹҺŰ�������" & vbCrLf & "�뵽�ҺŰ���������!", vbOKOnly, Me.Caption
'        Exit Sub '����ҺŰ�����û�����ô������Ϣ �Ͳ���������
'    End If
'    lng�޺� = Nvl(mrs�޺�!�޺���, 0): lng��Լ = Nvl(mrs�޺�!��Լ��, 0)
'    If lng�޺� = 0 Then
'        MsgBox "��ǰ�ű���" & str���� & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
'        Exit Sub
'    End If
'    Me.txt�޺�.Text = lng�޺�
'    Me.txt��Լ.Text = lng��Լ
'    lng��Լ = lng�޺�
'    strʱ�� = Nvl(mrs����(str����).Value)
'    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
'
''*************************************************************
''ʱ�������� ���õļ��
''*************************************************************
'      lng��� = Val(Me.txtTimeOut.Text)
'     ' datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!��ʼʱ��, "00:00:00"))
'
'      With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 400: .RowHeightMin = 400
'        .Rows = 0: .Cols = 2:   .Clear: lngRow = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'      End With
'    '*************************************
'    'ר�Һ�
'    '���������
'    '���� ʱ��α��е� ���°�ʱ�����ж�
'    '���� ȫ���������  ��Ϊ���������
'    '*************************************
'
'    With vsTime
'         .Cols = 2
'         lngRow = -1: lngCol = 0
'         j = 1
'         lngStart = 1
'         Do While Not mrs�ϰ�ʱ���.EOF
'            If blnExit Then Exit Do
'
'            datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!�ϰ�, "00:00:00"))
'             For i = j To lng��Լ
'                If lngStart > lng��Լ Then
'                    blnExit = True
'                    Exit For
'                End If
'
'                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss") Then
'                    j = i
'                    Exit For
'                 End If
'                lngCol = lngCol + 1
'                If strʱ�� <> Format(datʱ��, "HH") & ":00" Then lngRow = lngRow + 2: lngCol = 1
'                If lngCol = 1 Then
'                     If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'                     strʱ�� = Format(datʱ��, "HH") & ":00"
'                     vsTime.TextMatrix(lngRow - 1, 0) = strʱ��
'                     vsTime.TextMatrix(lngRow, 0) = strʱ��
'
'                End If
'                strData = lngStart
'                lngStart = lngStart + 1
'                strTime = Format(datʱ��, "HH:mm") & "-" & _
'                           IIf(Format(DateAdd("n", lng���, datʱ��), "yyyy-MM-dd hh:mm:ss") > Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "yyyy-MM-dd hh:mm:ss"), _
'                           Format(CDate(Nvl(mrs�ϰ�ʱ���!�°�, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
'
'                If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
'                vsTime.TextMatrix(lngRow - 1, lngCol) = strData
'                vsTime.TextMatrix(lngRow, lngCol) = strTime
'                '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
'
'                datʱ�� = DateAdd("n", lng���, datʱ��)
'             Next
'             mrs�ϰ�ʱ���.MoveNext
'         Loop
''         '***********************
''         '������
''         '**********************
''         For i = 1 To lng��Լ
''            If Format(datʱ��, "dd:mm:ss") >= Format(CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), "dd:mm:ss") Then Exit For
''            lngCol = lngCol + 1
''            If strʱ�� <> Format(datʱ��, "HH") & ":00" Then lngRow = lngRow + 2: lngCol = 1
''            If lngCol = 1 Then
''                 If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
''                 strʱ�� = Format(datʱ��, "HH") & ":00"
''                 vsTime.TextMatrix(lngRow - 1, 0) = strʱ��
''                 vsTime.TextMatrix(lngRow, 0) = strʱ��
''
''            End If
''            strData = i
''            strTime = Format(datʱ��, "HH:mm") & "-" & _
''                       IIf(DateAdd("n", lng���, datʱ��) > CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), _
''                       Format(CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
''
''            If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
''            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
''            vsTime.TextMatrix(lngRow, lngCol) = strTime
''            '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
''
''            datʱ�� = DateAdd("n", lng���, datʱ��)
''         Next
''         If blnȫ�� Then
''             mrs�ϰ�ʱ���.Filter = "ʱ���='����'"
''            datʱ�� = CDate(Nvl(mrs�ϰ�ʱ���!��ʼʱ��, "00:00:00"))
''         End If
''         j = i
''         For i = j To lng��Լ
''            If Format(datʱ��, "dd:mm:ss") >= CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")) Then Exit For
''            lngCol = lngCol + 1
''            If lngCol > vsTime.Cols - 1 Then lngRow = lngRow + 2: lngCol = 1
''            strData = i
''            strTime = Format(datʱ��, "HH:mm") & "-" & _
''                       IIf(DateAdd("n", lng���, datʱ��) > CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), _
''                       Format(CDate(Nvl(mrs�ϰ�ʱ���!��ֹʱ��, "00:00:00")), "HH:mm"), Format(DateAdd("n", lng���, datʱ��), "HH:mm"))
''            If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
''            If lngRow < 0 Then vsTime.Rows = vsTime.Rows + 2: lngRow = lngRow + 2
''            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
''            vsTime.TextMatrix(lngRow, lngCol) = strTime
''
''            '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
''            If lngCol = 1 Then
''                 vsTime.TextMatrix(lngRow - 1, 0) = Format(datʱ��, "HH:mm")
''                 vsTime.TextMatrix(lngRow, 0) = Format(datʱ��, "HH:mm")
''            End If
''            datʱ�� = DateAdd("n", lng���, datʱ��)
''         Next
'         For i = 1 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'         Next
'         .ColWidth(0) = 1200
'         .FixedAlignment(0) = flexAlignRightTop
'         .ColAlignment(0) = flexAlignRightTop
'         If .Rows > 0 Then
'            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
'            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
'         End If
'         .Redraw = flexRDBuffered
'    End With
'
'Exit Sub
'Hd:
'    If ErrCenter() = 1 Then
'         Resume
'    End If
'    SaveErrLog
'End Sub
'
'Private Sub cmdԤԼ_Click()
'    '��ʱ����ܷ�ԤԼ��������
'    If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Then Exit Sub
'    If mViewMode = ViewMode.ViewItem Or vsTime.TextMatrix(vsTime.MouseRow, vsTime.MouseCol) = "" Then Exit Sub
'    With vsTime
'        If .CellForeColor = vbBlue Then
'            .Cell(flexcpForeColor, .Row, .Col, .Row + 1, .Col) = &H80000008
'            .Cell(flexcpFontBold, .Row, .Col, .Row + 1, .Col) = False
'         Else
'            .Cell(flexcpForeColor, .Row, .Col, .Row + 1, .Col) = vbBlue
'            .Cell(flexcpFontBold, .Row, .Col, .Row + 1, .Col) = True
'        End If
'    End With
'    mblnChange = True
'End Sub
'
'Private Sub Form_Activate()
'    Me.Icon = frmRegistPlan.Icon
'End Sub
'
'Private Sub Form_Load()
'    Initʱ���
'End Sub
'
'Private Sub Form_Resize()
'  On Error Resume Next
'  '********************************************
'  '�������� �������С��Ⱥ���С�߶�
'  '********************************************
'  If Me.Width < 701 * Screen.TwipsPerPixelX Then Me.Width = 701 * Screen.TwipsPerPixelX
'  If Me.Height < 511 * Screen.TwipsPerPixelY Then Me.Height = 511 * Screen.TwipsPerPixelY
'  '********************************************
'  '�ҺŰ��Ż�����Ϣ λ�ò��ƶ��ƶ�
'  '���ƶ� ʱ������
'  '********************************************
'  With fraDate
'     .Width = Me.ScaleWidth - 2 * .Left
'     .Height = Me.ScaleHeight - Me.fraInfo.Top - Me.fraInfo.Height - 65 * Screen.TwipsPerPixelY
'  End With
'
'  With picTime
'     .Width = fraDate.Width - 2 * .Left
'     .Height = fraDate.Height - .Top * 2
'  End With
'  With Me.tbWeekTime
'    .Width = picTime.ScaleWidth - 2 * .Left
'  End With
'  With Me.vsTime
'    .Width = picTime.ScaleWidth - 2 * .Left
'    .Height = picTime.ScaleHeight - .Top - cmd����ʱ��.Top
'  End With
'  '-------------------------------------------
'  'Ӧ���� λ�õĵ���
'  '-------------------------------------------
'  With Me.fraӦ����
'       .Left = .Left
'       .Top = Me.fraDate.Top + Me.fraDate.Height + 5 * Screen.TwipsPerPixelY
'
'  End With
'
'  '********************************************
'  'ȷ����ť��ȡ����ť���ƶ�
'  '********************************************
'
'  With Me.cmdCancel
'       .Left = Me.ScaleWidth - 40 * Screen.TwipsPerPixelX - .Width
'       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
'  End With
'  With Me.cmdOK
'       .Left = cmdCancel.Left - 20 * Screen.TwipsPerPixelX - .Width
'       .Top = Me.ScaleHeight - .Height - 15 * Screen.TwipsPerPixelY
'  End With
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'     mlngPre����ID = -1
'     mblnChange = False
'     Set mrsTime = Nothing
'     mstr�����޸� = ""
'     Set mrs�޺� = Nothing
'     Set mrs�ϰ�ʱ��� = Nothing
'     Set mrs���� = Nothing
'End Sub
'
'
'
'
'
'Private Sub tbWeekTime_Click()
'    Dim i       As Integer
'    If mblnChange Then
'        mblnChange = False
'        If MsgBox("��ǰ�ҺŰ�����" & mstrKey & "��ʱ���Ѹı�!�Ƿ񱣴�?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption) = vbYes Then
'            cmdOK_Click
'         For i = 1 To tbWeekTime.Tabs.Count
'            If tbWeekTime.Tabs(i).Key = "K" & mstrKey Then
'                tbWeekTime.Tabs(i).Selected = True
'                Exit For
'            End If
'         Next
'        End If
'    End If
'    mstrKey = Mid(tbWeekTime.SelectedItem.Key, 2)
'     If mstr�����޸� <> "" Then
'        vsTime.Editable = flexEDKbdMouse: cmd����ʱ��.Enabled = True
'        If InStr(mstr�����޸�, ";" & mstrKey & ";") > 0 Then vsTime.Editable = flexEDNone: cmd����ʱ��.Enabled = False
'    End If
'    Select Case mViewMode
'        Case ViewMode.ViewItem:
'             Call LoadTimePlan(mlng����ID, Me.chk��ſ���.Value = 1)
'        Case ViewMode.Edit:
'            cmdԤԼ.Visible = False
'            Call LoadEditTimePlan(mlng����ID, Me.chk��ſ���.Value = 1)
'    End Select
'     setVsFlexBgColor (Me.chk��ſ���.Value = 1)
'End Sub
'
'
'
'
'Private Sub txtTimeOut_KeyPress(KeyAscii As Integer)
'
'    '���Ʒ���������
'    If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
'    If txtTimeOut.Text = "" And KeyAscii = Asc(0) Then KeyAscii = 0
'End Sub
'
'Private Sub txtTimeOut_Validate(Cancel As Boolean)
'    If Val(txtTimeOut.Text) < 1 Then Cancel = True
'End Sub
'
'
'
'Private Sub udTime_DownClick()
'    If Val(txtTimeOut.Text) < 2 Then Exit Sub
'    txtTimeOut.Text = Val(txtTimeOut.Text) - 1
'End Sub
'
'Private Sub udTime_UpClick()
'  txtTimeOut.Text = Val(txtTimeOut.Text) + 1
'End Sub
'
'
'
'
''Private Sub vsTime_Click()
''  Select Case mViewMode
''    Case ViewMode.Edit, ViewMode.NewItem:
''       If vsTime.MouseRow < 0 Or vsTime.MouseCol < 0 Or (chk��ſ���.Value = 0 And vsTime.MouseRow < 1) Then Exit Sub
''       Select Case chk��ſ���.Value = 1
''            Case True:
''            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
''            Case False:
''            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
''       End Select
''        If vsTime.MouseRow < 0 Or vsTime.MouseCol < 1 Then Exit Sub
''
''        If chk��ſ���.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
''            cmdԤԼ.Left = vsTime.MouseCol * 1200 + 20
''            cmdԤԼ.Top = vsTime.MouseRow * 400 + 20
''            cmdԤԼ.Visible = True
''        End If
''
''    Case ViewMode.ViewItem:
''         vsTime.Editable = flexEDNone
''  End Select
''End Sub
'
'Public Function ShowMe(lng����ID As Long, mode As ViewMode) As Boolean
'    mViewMode = mode: mlng����ID = lng����ID
'    If InitData() = False Then
'        '���عҺŰ��Ż�����Ϣ
'         Exit Function
'    End If
'    Select Case mViewMode
'         Case ViewMode.ViewItem:
'                vsTime.Editable = flexEDNone
'                Me.txtTimeOut.Enabled = False
'                Me.cmd����ʱ��.Enabled = False
'               '�鿴
'              Call LoadTimePlan(mlng����ID, chk��ſ���.Value = 1, False)
'         Case ViewMode.Edit
'              If LoadEditTimePlan(mlng����ID, chk��ſ���.Value = 1, False) = False Then
'               Exit Function
'              End If
'    End Select
'    setVsFlexBgColor (chk��ſ���.Value = 1)
'    Me.Show 1
'    ShowMe = mblnReload
'End Function
''------------------------------------------------------------------------
''ҳ����ù����뷽��
''------------------------------------------------------------------------
'Public Function InitData() As Boolean
'    Dim strSQL          As String
'    Dim lng����ID       As Long
'    If mlng����ID = -1 Then Exit Function
'     lng����ID = mlng����ID
'     On Error GoTo Hd
'     strSQL = " " & _
'        "   Select A.Id as ����ID,0 as �ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id," & _
'        "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,nvl(A.Ĭ��ʱ�μ��,5) As Ĭ��ʱ�μ��, " & _
'        "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ���� " & _
'        "   From �ҺŰ��� A,�շ���ĿĿ¼ B,���ű� D " & _
'        "   Where A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
'        "         And A.Id=[1]"
'         Set mrs���� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
'
'         If mrs����.EOF Then
'              ShowMsgbox "δ�ҵ�ָ���ĺű�,����!"
'             Exit Function
'        End If
'        strSQL = "Select ������Ŀ,�޺���,  ��Լ��,������Ŀ as ���� From  �ҺŰ������� where ����ID=[1]  Order BY ������Ŀ      "
'        Set mrs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
'        cbo����.Text = Nvl(mrs����!����)
'        txt�ű�.Tag = Nvl(mrs����!����id)
'        txtTimeOut.Tag = Val(Nvl(mrs����!Ĭ��ʱ�μ��, 0))
'        txtTimeOut.Text = txtTimeOut.Tag
'        txt�ű�.Text = Nvl(mrs����!����)
'        cbo����.Text = Nvl(mrs����!����)
'        cboItem.Text = Nvl(mrs����!��Ŀ)
'        cboDoctor.Text = Nvl(mrs����!ҽ������)
'        chk����.Value = IIf(Val(Nvl(mrs����!��������)) = 1, 1, 0)
'       chk��ſ���.Value = IIf(Val(Nvl(mrs����!��ſ���)) = 1, 1, 0):  chk��ſ���.Tag = chk��ſ���.Value
'        strSQL = "" & _
'        "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
'        "               ��������,�Ƿ�ԤԼ" & _
'        "   From  �ҺŰ���ʱ�� " & _
'        "   Where ����ID=[1]" & _
'        "   Order by ����,ʱ��,���"
'        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
'        mstr�����޸� = Get��Լ����(mlng����ID)
'       InitData = True
'Exit Function
'Hd:
'     If ErrCenter() = 1 Then Resume
'     SaveErrLog
'End Function
'
'
'Private Function LoadEditTimePlan(ByVal lng����ID As Long, ByVal bln��ſ��� As Boolean, _
'    Optional bln�ƻ� As Boolean = False) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:
'    '���:
'    '����:
'    '����:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL           As String
'    Dim rsTemp           As ADODB.Recordset
'    Dim str����          As String
'    Dim i                As Long
'    Dim r                As Integer
'    Dim lngRow           As Long
'    Dim lngCol           As Integer
'    Dim strʱ��          As String
'    Dim strTime          As String
'    Dim strData          As String
'    Dim strKey           As String
'
'    On Error GoTo errHandle
'    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
'    If mrsTime Is Nothing Then
'        mlngPre����ID = -1
'    ElseIf mrsTime.State <> 1 Then
'         mlngPre����ID = -1
'    End If
'    If mlngPre����ID <> lng����ID Then
'        mlngPre����ID = lng����ID
'        tbWeekTime.Tabs.Clear
'        With tbWeekTime
'            If Not mrs�޺�.EOF Then
'                mrs�޺�.Filter = "����='��һ'"
'                If mrs�޺�.RecordCount > 0 Then
'                '�޺���,  ��Լ��,������Ŀ
'                    If Nvl(mrs�޺�!�޺���, 0) > 0 Then
'                        tbWeekTime.Tabs.Add , _
'                            "K��һ", "��һ" & IIf(Nvl(mrs����!��һ) = "", "", "(" & Nvl(mrs����!��һ) & ")")
'                    End If
'                End If
'                mrs�޺�.Filter = "����='�ܶ�'"
'                If mrs�޺�.RecordCount > 0 Then
'                   If Nvl(mrs�޺�!�޺���, 0) > 0 Then
'                    tbWeekTime.Tabs.Add , _
'                        "K�ܶ�", "�ܶ�" & IIf(Nvl(mrs����!�ܶ�) = "", "", "(" & Nvl(mrs����!�ܶ�) & ")")
'                    End If
'                End If
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then
'                     If Nvl(mrs�޺�!�޺���, 0) > 0 Then
'                    tbWeekTime.Tabs.Add , _
'                        "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
'                    End If
'                 End If
'
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then
'                  If Nvl(mrs�޺�!�޺���, 0) > 0 Then
'                    tbWeekTime.Tabs.Add , _
'                      "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
'                  End If
'                End If
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then
'                     If Nvl(mrs�޺�!�޺���, 0) > 0 Then
'                        tbWeekTime.Tabs.Add , _
'                            "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
'                     End If
'                End If
'
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then
'                   If Nvl(mrs�޺�!�޺���, 0) > 0 Then
'                        tbWeekTime.Tabs.Add , _
'                          "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
'                   End If
'                End If
'                mrs�޺�.Filter = "����='����'"
'                If mrs�޺�.RecordCount > 0 Then
'                    If Nvl(mrs�޺�!�޺���, 0) > 0 Then
'                        tbWeekTime.Tabs.Add , _
'                            "K����", "����" & IIf(Nvl(mrs����!����) = "", "", "(" & Nvl(mrs����!����) & ")")
'                    End If
'                End If
'                mrs�޺�.Filter = 0
'            End If
'            .Visible = tbWeekTime.Tabs.Count <> 0
'            If .Tabs.Count > 0 Then
'                .Tabs(1).Selected = True
'            Else
'                MsgBox "�ð���û�����ö�Ӧ���޺�������Լ��,����!", vbOKOnly, Me.Caption
'                Exit Function
'            End If
'
'        End With
'    End If
'    str���� = "": strTime = ""
'    If Not tbWeekTime.SelectedItem Is Nothing Then
'        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
'    End If
'    mrsTime.Filter = "����='" & str���� & "'"
'    mrs�޺�.Filter = "����='" & str���� & "'"
'    txt�޺�.Text = ""
'    txt��Լ.Text = ""
'    If mrs�޺�.RecordCount <> 0 Then
'        Me.txt�޺�.Text = Nvl(mrs�޺�!�޺���, 0)
'        Me.txt��Լ.Text = Nvl(mrs�޺�!��Լ��, 0)
'    End If
'     strʱ�� = ""
'    With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 400: .RowHeightMin = 400
'        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'        If Not bln��ſ��� Then
'             .Cols = 8: .FixedCols = 0
'             .Rows = 1: .FixedRows = 1
'             For i = 0 To .Cols - 1 Step 2
'                .TextMatrix(0, i) = "ʱ���"
'             Next
'             For i = 1 To .Cols - 1 Step 2
'                .TextMatrix(0, i) = "ԤԼ����"
'             Next
'
'             r = 1: i = -1
'            Do While Not mrsTime.EOF
'                i = i + 1
'                If i * 2 > .Cols - 2 Then r = r + 1: i = 0
'                strData = Val(Nvl(mrsTime!��������))
'                strTime = mrsTime!ʱ�䷶Χ
'                If r > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(r, i * 2) = strTime
'                .TextMatrix(r, i * 2 + 1) = strData
'                mrsTime.MoveNext
'            Loop
'            For i = 0 To .Cols - 1
'                .ColAlignment(i) = flexAlignCenterCenter
'                .ColWidth(i) = 1200
'            Next
'            .Redraw = flexRDBuffered
'            LoadEditTimePlan = True
'            Exit Function
'        End If
'        .Cols = 7: .FixedCols = 1
'        .Rows = 0: .FixedRows = 0
'        i = 1: r = -1
'        lngRow = -1: lngCol = 1
'        '******************************************
'        With vsTime
'         .Cols = 2
'         lngRow = -1: lngCol = 0
'         '***********************
'         '������
'         '**********************
'         r = mrsTime.RecordCount
'         For i = 1 To r
'            If mrsTime.EOF Then Exit For
'            lngCol = lngCol + 1
'            If strʱ�� <> Nvl(mrsTime!ʱ��) Then lngRow = lngRow + 2: lngCol = 1
'             If lngCol = 1 Then
'                strʱ�� = Nvl(mrsTime!ʱ��)
'                If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'                vsTime.TextMatrix(lngRow - 1, 0) = strʱ��
'                vsTime.TextMatrix(lngRow, 0) = strʱ��
'             End If
'            strData = mrsTime!���
'            strTime = mrsTime!ʱ�䷶Χ
'            If lngCol > vsTime.Cols - 1 Then vsTime.Cols = vsTime.Cols + 1
'            'If lngRow > vsTime.Rows - 1 Then vsTime.Rows = vsTime.Rows + 2
'            vsTime.TextMatrix(lngRow - 1, lngCol) = strData
'            vsTime.TextMatrix(lngRow, lngCol) = strTime
'            '�ǵ�һ��ʱ ��д ��ʼʱ�䵽����
'            If lngCol = 1 Then
'            End If
'            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
'                .Cell(flexcpForeColor, lngRow - 1, lngCol, lngRow, lngCol) = vbBlue
'                .Cell(flexcpFontBold, lngRow - 1, lngCol, lngRow, lngCol) = True
'            End If
'            mrsTime.MoveNext
'         Next
'
'         End With
'        '******************************************
''        Do While Not mrsTime.EOF
''            If i = 1 Then
''                r = r + 2
''                strʱ�� = Nvl(mrsTime!ʱ��)
''                If r > .Rows - 1 Then .Rows = .Rows + 2
''                .TextMatrix(r, 0) = strʱ��
''                .TextMatrix(r - 1, 0) = strʱ��
''            End If
''            i = i + 1
''            strData = mrsTime!���
''            strTime = mrsTime!ʱ�䷶Χ
''            If i >= .Cols - 1 Then i = 1
''            If r > .Rows - 1 Then .Rows = .Rows + 2
''            .TextMatrix(r, i) = strTime
''            .TextMatrix(r - 1, i) = strData
''
''        Loop
'
'
'        For i = 1 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'        Next
'        .ColWidth(0) = 1200
'        .FixedAlignment(0) = flexAlignRightTop
'        .ColAlignment(0) = flexAlignRightTop
'        If .Rows > 0 Then
'            .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
'            .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
'        End If
'        .MergeCellsFixed = flexMergeRestrictColumns
'        .MergeCol(0) = True
'        .Redraw = flexRDBuffered
'    End With
'    LoadEditTimePlan = True
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Function

'Private Sub LoadEditTimePlantext(ByVal lng����ID As Long, ByVal bln��ſ��� As Boolean, _
'    Optional bln�ƻ� As Boolean = False)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:
'    '���:
'    '����:
'    '����:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL           As String
'    Dim rsTemp           As ADODB.Recordset
'    Dim str����          As String
'    Dim i                As Long
'    Dim r                As Integer
'    Dim strʱ��          As String
'    Dim strTime          As String
'    Dim strData          As String
'    Dim strKey           As String
'
'    On Error GoTo errHandle
'    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
'    If mrsTime Is Nothing Then
'        mlngPre����ID = -1
'    ElseIf mrsTime.State <> 1 Then
'         mlngPre����ID = -1
'    End If
'    If mlngPre����ID <> lng����ID Then
'        mlngPre����ID = lng����ID
'        tbWeekTime.Tabs.Clear
'        With mrsTime
'            strTime = ""
'            Do While Not .EOF
'                If strTime <> Nvl(mrsTime!����) Then
'                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!����), Nvl(mrsTime!����)
'                    strTime = Nvl(mrsTime!����)
'                End If
'                .MoveNext
'            Loop
'            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
'            If tbWeekTime.Tabs.Count > 0 Then
'                tbWeekTime.Tabs(1).Selected = True
'            End If
'            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
'        End With
'    End If
'    str���� = "": strTime = ""
'    If Not tbWeekTime.SelectedItem Is Nothing Then
'        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
'    End If
'    mrsTime.Filter = "����='" & str���� & "'"
'    mrs�޺�.Filter = "����='" & str���� & "'"
'    txt�޺�.Text = ""
'    txt��Լ.Text = ""
'    If mrs�޺�.RecordCount <> 0 Then
'        Me.txt�޺�.Text = Nvl(mrs�޺�!�޺���, 0)
'        Me.txt��Լ.Text = Nvl(mrs�޺�!��Լ��, 0)
'    End If
'     strʱ�� = ""
'    With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 400: .RowHeightMin = 400
'        .Rows = 0: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'        If Not bln��ſ��� Then
'             .Cols = 8: .FixedCols = 0
'             .Rows = 1: .FixedRows = 1
'             For i = 0 To .Cols - 1 Step 2
'                .TextMatrix(0, i) = "ʱ���"
'             Next
'             For i = 1 To .Cols - 1 Step 2
'                .TextMatrix(0, i) = "ԤԼ����"
'             Next
'
'             r = 1: i = -1
'            Do While Not mrsTime.EOF
'                If i * 2 > .Cols - 2 Then r = r + 1: i = -1
'                i = i + 1
'                strData = Val(Nvl(mrsTime!��������))
'                strTime = mrsTime!ʱ�䷶Χ
'                If r > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(r, i * 2) = strTime
'                .TextMatrix(r, i * 2 + 1) = strData
'                mrsTime.MoveNext
'            Loop
'            For i = 0 To .Cols - 1
'                .ColAlignment(i) = flexAlignCenterCenter
'                .ColWidth(i) = 1200
'            Next
'            .Redraw = flexRDBuffered
'             Exit Sub
'        End If
'        Do While Not mrsTime.EOF
'            If strʱ�� <> Nvl(mrsTime!ʱ��) Then
'                r = r + 2
'                strʱ�� = Nvl(mrsTime!ʱ��)
'                If r > .Rows - 1 Then .Rows = .Rows + 2
'                .TextMatrix(r, 0) = strʱ��
'                .TextMatrix(r - 1, 0) = strʱ��
'                i = 0
'            End If
'            i = i + 1
'            strData = mrsTime!���
'            strTime = mrsTime!ʱ�䷶Χ
'            If i > .Cols - 1 Then .Cols = .Cols + 1
'            If r > .Rows - 1 Then .Rows = .Rows + 1
'            .TextMatrix(r, i) = strTime
'            .TextMatrix(r - 1, i) = strData
'            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
'
'                .Cell(flexcpForeColor, r - 1, i, r, i) = vbBlue
'                .Cell(flexcpFontBold, r - 1, i, r, i) = True
'            End If
'            mrsTime.MoveNext
'        Loop
'        For i = 1 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'        Next
'        .ColWidth(0) = 1200
'        .FixedAlignment(0) = flexAlignRightTop
'        .ColAlignment(0) = flexAlignRightTop
'        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
'        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
'        .MergeCellsFixed = flexMergeRestrictColumns
'        .MergeCol(0) = True
'        .Redraw = flexRDBuffered
'    End With
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Sub
'
'
'Private Sub LoadTimePlan(ByVal lng����ID As Long, ByVal bln��ſ��� As Boolean, _
'    Optional bln�ƻ� As Boolean = False)
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:
'    '���:
'    '����:
'    '����:
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL           As String
'    Dim rsTemp           As ADODB.Recordset
'    Dim str����          As String
'    Dim i                As Long
'    Dim r                As Integer
'    Dim strʱ��          As String
'    Dim strTime          As String
'    Dim strKey           As String
'    On Error GoTo errHandle
'    '���ظùҺ���Ŀ�ĵ�ͣ��ʱ����Ϣ
'    If mrsTime Is Nothing Then
'         mlngPre����ID = -1
'    ElseIf mrsTime.State <> 1 Then
'         mlngPre����ID = -1
'    End If
'    If mlngPre����ID <> lng����ID Then
'        mlngPre����ID = lng����ID
''        strSQL = "" & _
''        "   Select decode(����,'����',1,'��һ',2,'�ܶ�',3,'����',4,'����',5,'����',6,7) as ����,����,to_char(��ʼʱ��,'HH24')||':00' as ʱ��,���,to_char(��ʼʱ��,'hh24:mi')||'-' ||to_char(����ʱ��,'hh24:mi') as ʱ�䷶Χ, " & _
''        "               ��������,�Ƿ�ԤԼ" & _
''        "   From  �ҺŰ���ʱ�� " & _
''        "   Where ����ID=[1]" & _
''        "   Order by ����,ʱ��,���"
''        Set mrsTime = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
'        tbWeekTime.Tabs.Clear
'        With mrsTime
'            strTime = ""
'            Do While Not .EOF
'                If strTime <> Nvl(mrsTime!����) Then
'                    tbWeekTime.Tabs.Add , "K" & Nvl(mrsTime!����), Nvl(mrsTime!����)
'                    strTime = Nvl(mrsTime!����)
'                End If
'                .MoveNext
'            Loop
'            tbWeekTime.Visible = tbWeekTime.Tabs.Count <> 0
'            If tbWeekTime.Tabs.Count > 0 Then
'                tbWeekTime.Tabs(1).Selected = True
'            End If
'            If mrsTime.RecordCount <> 0 Then mrsTime.MoveFirst
'        End With
'        If tbWeekTime.Tabs.Count = 0 Then
'            MsgBox "�ð���û�����ö�Ӧ��ʱ��,����!"
'        End If
'    End If
'    str���� = "": strTime = ""
'    If Not tbWeekTime.SelectedItem Is Nothing Then
'        str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
'    End If
'    mrsTime.Filter = "����='" & str���� & "'"
'    mrs�޺�.Filter = "����='" & str���� & "'"
'    txt�޺�.Text = ""
'    txt��Լ.Text = ""
'    If mrs�޺�.RecordCount <> 0 Then
'        Me.txt�޺�.Text = Nvl(mrs�޺�!�޺���, 0)
'        Me.txt��Լ.Text = Nvl(mrs�޺�!��Լ��, 0)
'    End If
'     strʱ�� = ""
'    With vsTime
'        .Redraw = flexRDNone: .SelectionMode = flexSelectionFree
'        .RowHeightMax = 800: .RowHeightMin = 800
'        .Rows = 1: .Cols = 2:   .Clear: r = -1: i = 0: .FixedCols = 1:
'        .FixedRows = 0
'        If Not bln��ſ��� Then
'             .Cols = 8: .FixedCols = 0
'             r = 0: i = 0
'            Do While Not mrsTime.EOF
'                i = i + 1
'                If i > .Cols - 1 Then r = r + 1: i = 0
'                strTime = "ԤԼ" & Val(Nvl(mrsTime!��������)) & "��" & vbCrLf & vbCrLf
'                strTime = strTime & mrsTime!ʱ�䷶Χ
'                If r > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(r, i) = strTime
'                mrsTime.MoveNext
'            Loop
'            For i = 0 To .Cols - 1
'                .ColAlignment(i) = flexAlignCenterCenter
'                .ColWidth(i) = 1200
'            Next
'            .Redraw = flexRDBuffered
'             Exit Sub
'        End If
'        Do While Not mrsTime.EOF
'            If strʱ�� <> Nvl(mrsTime!ʱ��) Then
'                r = r + 1
'                strʱ�� = Nvl(mrsTime!ʱ��)
'                If r > .Rows - 1 Then .Rows = .Rows + 1
'                .TextMatrix(r, 0) = strʱ��
'                i = 0
'            End If
'            i = i + 1
'            strTime = mrsTime!��� & vbCrLf & vbCrLf
'            strTime = strTime & mrsTime!ʱ�䷶Χ
'            If i > .Cols - 1 Then .Cols = .Cols + 1
'            If r > .Rows - 1 Then .Rows = .Rows + 1
'            .TextMatrix(r, i) = strTime
'            If Val(Nvl(mrsTime!�Ƿ�ԤԼ)) = 1 Then
'                .Cell(flexcpForeColor, r, i, r, i) = vbBlue
'                .Cell(flexcpFontBold, r, i, r, i) = True
'            End If
'            mrsTime.MoveNext
'        Loop
'        For i = 1 To .Cols - 1
'            .ColAlignment(i) = flexAlignCenterCenter
'            .ColWidth(i) = 1200
'        Next
'        .ColWidth(0) = 1200
'        .FixedAlignment(0) = flexAlignRightTop
'        .ColAlignment(0) = flexAlignRightTop
'        .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
'        .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
'        .Redraw = flexRDBuffered
'    End With
'    Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'End Sub
'
'Private Sub vsTime_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
' If vsTime.Row < 0 Or vsTime.Col < 0 Or (chk��ſ���.Value = 0 And vsTime.Row < 1) Then cmdԤԼ.Visible = False: mblnCellChange = False: Exit Sub
'    cmdԤԼ.Visible = False
'    Select Case mViewMode
'    Case ViewMode.Edit, ViewMode.NewItem:
'       Select Case chk��ſ���.Value = 1
'            Case True:
'            vsTime.Editable = IIf(vsTime.Row Mod 2 <> 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'            '******************************************
'            '�������������ʽ
'            '******************************************
'            If vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
'            Case False:
'            vsTime.Editable = IIf(vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "", flexEDKbdMouse, flexEDNone)
'            '******************************************
'            '�������������ʽ
'            '******************************************
'            If NewCol Mod 2 = 0 And vsTime.Editable = flexEDKbdMouse Then vsTime.ColEditMask(vsTime.Col) = strMaskKey
'       End Select
'        If vsTime.Row < 0 Or vsTime.Col < 1 Then Exit Sub
'
'        If chk��ſ���.Value = 1 And vsTime.Row Mod 2 = 0 And vsTime.TextMatrix(vsTime.Row, vsTime.Col) <> "" Then
'            mblnCellChange = True
'        Else
'           mblnCellChange = False
'        End If
'
'    Case ViewMode.ViewItem:
'         mblnCellChange = False
'         vsTime.Editable = flexEDNone
'  End Select
'   If mstr�����޸� <> "" Then
'        vsTime.Editable = flexEDKbdMouse
'        If InStr(mstr�����޸�, ";" & mstrKey & ";") > 0 Then vsTime.Editable = flexEDNone
'
'   End If
'End Sub
'
'Private Sub vsTime_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button <> 1 Or Not mblnCellChange Then Exit Sub
'    If Nvl(txt��Լ.Text) = 0 Then Exit Sub
'    If InStr(mstr�����޸�, mstrKey) > 0 Then Exit Sub
'    cmdԤԼ.Visible = True
'    cmdԤԼ.Left = X - X Mod 1200 + 20
'    cmdԤԼ.Top = Y - Y Mod 400 + 20
'    mblnCellChange = False
'End Sub
'
'Private Sub vsTime_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
'    '**************************************************************
'    '������Ա �϶�������ʱ �� ԤԼ��ť ����
'    '**************************************************************
'    Me.cmdԤԼ.Visible = False
'End Sub
'
'Private Sub vsTime_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'    If mViewMode = ViewItem Then Exit Sub
'    Select Case chk��ſ���.Value = 1
'        Case True:
'            '******************************************
'            'ר�Һ�ʱ ��������
'            '******************************************
'            If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
'               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
'        Case False:
'            '******************************************
'            '��ͨ��ʱ ��������
'            '******************************************
'            If Col Mod 2 = 0 Then
'                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
'               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
'            Else
'                If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
'               Or KeyAscii = 13) Then KeyAscii = 0: Exit Sub
'            End If
'    End Select
'End Sub
'
'Private Function isValied() As Boolean
'    '***************************************
'    '��֤�û��ԹҺŰ���ʱ�ε��޸�
'    '***************************************
'     Dim i          As Long
'     Dim j          As Long
'     Dim lngԤԼ    As Long
'     Dim lng��Լ    As Long
'     Dim lng�޺�    As Long
'     Dim str����    As String
'     If tbWeekTime.SelectedItem Is Nothing Then Exit Function
'      str���� = Mid(tbWeekTime.SelectedItem.Key, 2)
'     lng�޺� = Val(txt�޺�.Text)
'     lng��Լ = Val(txt��Լ.Text)
'     If lng��Լ = 0 Then lng��Լ = lng�޺�
'     Select Case chk��ſ���.Value = 1
'     Case True:
'     '*************************************
'     'ר�Һż����Լ���Ƿ�����޺���
'     '*************************************
'        With vsTime
'            For i = 0 To .Rows - 1 Step 2
'                For j = 1 To .Cols - 1
'                    If .Cell(flexcpForeColor, i, j, i, j) = vbBlue And .TextMatrix(i, j) <> "" Then
'                        lngԤԼ = lngԤԼ + 1
'                    End If
'                Next
'            Next
'        End With
'     Case False:
'     '*************************************
'     '��ͨ�ż����Լ���Ƿ�����޺���
'     '*************************************
'        With vsTime
'            For i = 1 To .Rows - 1
'                For j = 1 To .Cols - 1 Step 2
'                    If .TextMatrix(i, j) <> "" Then
'                        lngԤԼ = lngԤԼ + Val(.TextMatrix(i, j))
'                    End If
'                Next
'            Next
'        End With
'     End Select
'     If lngԤԼ > lng��Լ Then
'        MsgBox "��" & str���� & "���õ�ԤԼ��" & lngԤԼ & "������" & IIf(lng�޺� = lng��Լ, "�޺���" & lng��Լ, "��Լ��" & lng��Լ) & ",����!", vbOKOnly, Me.Caption
'        Exit Function
'     End If
'    isValied = True
'    Exit Function
'End Function
'
'Private Sub vsTime_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'  Dim i         As Long
'  Dim j         As Long
'  Dim lng�޺�   As Long
'  Dim lng��Լ   As Long
'  Dim lngԤԼ�� As Long
'  If mViewMode = ViewItem Then Exit Sub
'
'  '*************************************
'  'ʱ�������֤ ������ʱ�䷶Χ
'  '**************************************
'  If vsTime.Editable = flexEDKbdMouse And vsTime.ColEditMask(vsTime.Col) = strMaskKey Then
'    Validateʱ�� Row, Col, Cancel
'    If Not Cancel Then mblnChange = True
'    Exit Sub
'  End If
'  '****************************************
'  '����ͨ�� ��ʱ�� �����������ԤԼ����������
'  '****************************************
'   If chk��ſ���.Value = 0 And vsTime.ColEditMask(vsTime.Col) <> strMaskKey And vsTime.Editable = flexEDKbdMouse Then
'        If vsTime.EditText = "" Then vsTime.EditText = "0"
'        mblnChange = True
'   End If
'End Sub
'
'Private Sub Validateʱ��(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'  Dim i         As Long
'  Dim j         As Long
'  Dim lng�޺�   As Long
'  Dim lng��Լ   As Long
'  Dim lngԤԼ�� As Long
'
'  Dim strʱ��()  As String
'  If mViewMode = ViewItem Then Exit Sub
'
'  '*************************************
'  '��֤ʱ��
'  '**************************************
'  strʱ�� = Split(vsTime.EditText, "-")
'  If UBound(strʱ��) <> 1 Then Cancel = True: Exit Sub
'   If Not IsDate(strʱ��(0)) Then Cancel = True: Exit Sub
'   If Not IsDate(strʱ��(1)) Then Cancel = True: Exit Sub
'   If CDate(strʱ��(0)) >= CDate(strʱ��(1)) Then
'        MsgBox "��ʼʱ�����С�ڽ���ʱ��!����!", vbOKOnly, Me.Caption
'        Cancel = True
'   End If
'
'End Sub
'
'Private Sub setVsFlexBgColor(Optional ByVal bln��ſ��� As Boolean = False)
'    '**************************************************************
'    '��ʱ������ü������
'    '**************************************************************
'     Dim i           As Long
'     If (bln��ſ��� And vsTime.Rows = 0) Or (bln��ſ��� = False And vsTime.Rows = 1) Then Exit Sub
'     For i = IIf(bln��ſ���, 0, 1) To vsTime.Rows - 1 Step 2
'            vsTime.Cell(flexcpBackColor, i, IIf(bln��ſ���, 1, 0), i, vsTime.Cols - 1) = &HE0E0D3
'     Next
'End Sub
'



'Private Sub Initʱ���()
'  '--------------------------------
'  '����:��ȡ���°�ʱ���
'  '--------------------------------
'    Dim strTmp      As String
'    Dim strSQL      As String
'    Dim rsTmp       As ADODB.Recordset
'    Dim strDat      As String
'    On Error GoTo Hd
'    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "07:00:00 AND 12:00:00")
'    strDat = Split(strTmp, "AND")(0)
'    If IsDate(strDat) Then
'        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
'    Else
'        t_ʱ��.dat_�����ϰ� = CDate("08:00:00")
'    End If
'
'    strDat = Split(strTmp, "AND")(1)
'    If IsDate(strDat) Then
'        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
'    Else
'        t_ʱ��.dat_�����°� = CDate("1900-01-01 12:00:00")
'    End If
'    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "14:00:00 AND 18:00:00")
'
'     strDat = Split(strTmp, "AND")(0)
'    If IsDate(strDat) Then
'        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
'    Else
'        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 14:00:00")
'    End If
'    strDat = Split(strTmp, "AND")(1)
'    If IsDate(strDat) Then
'        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
'    Else
'        t_ʱ��.dat_�����°� = CDate("1900-01-01 18:00:00")
'    End If
'    With t_ʱ��
'         If .dat_�����ϰ� > .dat_�����°� Then
'            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
'         End If
'         If .dat_�����ϰ� > .dat_�����°� Then
'            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
'         End If
'    End With
'    strSQL = _
'    "       Select ʱ���, �ϰ�, �°� " & vbNewLine & _
'    "       From (" & vbNewLine & _
'    "           With Tb As (Select ʱ���,To_Date('1900-01-01 ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ʼʱ��," & vbNewLine & _
'    "                               To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), -1, '1900-01-01 ', '1900-01-02 ') ||To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ֹʱ��," & _
'    "                               Sign(��ʼʱ�� - ��ֹʱ��) As ����, " & vbNewLine & _
'    "                                To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��, " & vbNewLine & _
'    "                                To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��, " & vbNewLine & _
'    "                                 To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��," & vbNewLine & _
'    "                                 To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��"
'    strSQL = strSQL & vbNewLine & _
'    "                       From ʱ��� )" & vbNewLine & _
'    "           Select ʱ���, '��' As ��ǩ, 0 As ��־, ��ʼʱ�� As �ϰ�, ��ֹʱ�� As �°�, ��ʼʱ��, ��ֹʱ��," & _
'    "                  �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ��" & vbNewLine & _
'    "            From Tb  Where (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) And " & _
'    "                      (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
'    "           Union All" & vbNewLine & _
'    "           Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & vbNewLine & _
'    "                        Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) As �°�, ��ʼʱ��, ��ֹʱ��, " & _
'    "                        �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
'    "           From Tb a Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
'    "           Union All " & vbNewLine & _
'    "            Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & _
'    "                   Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) As �°�, ��ʼʱ��, ��ֹʱ��, �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
'    "         From Tb a   Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��)" & vbNewLine & _
'    "            ) b" & vbNewLine & _
'    "         Order By ʱ���,�ϰ�"
'     Set mrs�ϰ�ʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
'    Exit Sub
'Hd:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Sub
'Private Function Get��Լ����(ByVal lng����ID As Long) As String
'    '��ȡ�����޸ĵİ�������
'    Dim strSQL As String
'    Dim rsTmp   As ADODB.Recordset
'    Dim strTmp  As String
'    strSQL = "Select Decode(To_Char(A.ԤԼʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7'," & _
'    "                             '����') As ���� " & vbCrLf & _
'    "          From ���˹Һż�¼ A,�ҺŰ��š�B " & vbCrLf & _
'    "        Where  A.�ű�=B.���� And A.��¼״̬ = 1 And b.ID = [1] And A.����ʱ�� > A.�Ǽ�ʱ�� And A.ԤԼʱ�� Is Not Null"
'
'    If gintԤԼ���� = 0 Then
'        strSQL = strSQL & " And A.ԤԼʱ�� > Sysdate "
'    Else
'        strSQL = strSQL & " And A.ԤԼʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
'    End If
'    On Error GoTo errH
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
'    If rsTmp.EOF Then Exit Function
'
'    Do While Not rsTmp.EOF
'        If InStr(strTmp, Nvl(rsTmp!����)) < 0 Or strTmp = "" Then
'            strTmp = strTmp & ";" & Nvl(rsTmp!����)
'        End If
'        rsTmp.MoveNext
'    Loop
'    If strTmp <> "" Then strTmp = strTmp & ";"
'    Get��Լ���� = strTmp
'    Exit Function
'errH:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Function
 
 

Private Sub cmdOther_Click()
    Dim str���� As String
    
    If Not mbln��ſ��� Then Exit Sub
    Set mfrmOtherCalc = New frmRegistPlanTimeOther
    Call mfrmOtherCalc.zlShowMe(Me, tbPage.Item(mlngSelIndex).Caption, Val(txtTimeOut.Text))
    If Not mfrmOtherCalc Is Nothing Then Unload mfrmOtherCalc
    Set mfrmOtherCalc = Nothing '
End Sub

Private Sub mfrmOtherCalc_zlRefreshCon(ByVal VarTimes As Variant)
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng�޺� As Long
    Dim lng��Լ As Long
    Dim dat��ʼʱ�� As Date
    Dim dat����ʱ�� As Date
    Dim lng��� As Long
    Dim strTmp As String
    Dim strʱ�� As String
    Dim str����ʱ�� As String
    Dim lngĬ�ϼ�� As Long
    Dim lng������� As Long
    Dim lng�̶����� As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim datʱ�� As Date
    Dim str�ֶμ�� As String
    Dim str������Ŀ As String
    Dim cllPro As Collection
    Dim varTemp As Variant
    Dim strStart As String
    Dim strEnd As String
    Dim int���� As Integer
    Dim strʱ�� As String
    Dim lngʱ���� As Long
    Dim varData As Variant
    If Not mbln��ſ��� Then Exit Sub
    If VarTimes Is Nothing Then Exit Sub
    If VarTimes("ʱ����") <> "" Then
        txtTimeOut.Text = Val(VarTimes("ʱ����"))
        Call cmd����ʱ��_Click
        Exit Sub
    End If

    str�ֶμ�� = VarTimes("�ֶμ��")
    If Trim(str�ֶμ��) = "" Then Exit Sub


    If mrs�ϰ�ʱ��� Is Nothing Then
        Call Initʱ���
    End If
    str������Ŀ = GetVsGridCaption(mlngSelIndex)


    If mrs�ϰ�ʱ��� Is Nothing Then Exit Sub
    mrsRegPlan.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mrsRegPlan.RecordCount = 0 Then mrsRegPlan.Filter = 0: Exit Sub
    lng�޺� = Nvl(mrsRegPlan!�޺���, 0): lng��Լ = Nvl(mrsRegPlan!��Լ��, 0)
    If lng��Լ = 0 Then lng��Լ = lng�޺�
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str������Ŀ & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Sub
    End If


    strʱ�� = mrsRegPlan!�Ű�
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 0 Then
        MsgBox "������ʱ��Ϊ[" & strʱ�� & "]�����°�ʱ��,����!", vbOKOnly, Me.Caption
        Exit Sub

    End If

    Set cllPro = New Collection
    varData = Split(str�ֶμ��, ";")

    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ",")
        int���� = Val(varTemp(1))
        varTemp = Split(varTemp(0), "��")
        strStart = varTemp(0)
        strEnd = varTemp(1)
        cllPro.Add int����, "K" & Replace(strStart, ":", "_")
        cllPro.Add strStart, "K" & Replace(strStart, ":", "_") & "_Start"
        cllPro.Add strEnd, "K" & Replace(strStart, ":", "_") & "_End"
    Next

    mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "' And ��ʹ��=0"
    Do While Not mrsAssign.EOF
        mrsAssign.Delete adAffectCurrent
        mrsAssign.MoveNext
    Loop
    mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mrsAssign.RecordCount <> 0 Then
        lng�̶����� = mrsAssign.RecordCount
        lngĬ�ϼ�� = Val(Nvl(mrsAssign!ʱ����, lngʱ����))
        lngʱ���� = lngĬ�ϼ��
        Do While Not mrsAssign.EOF
            lng������� = lng������� + Val(Nvl(mrsAssign!��������))
            mrsAssign.MoveNext
        Loop
    End If
    mrsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs�ϰ�ʱ���.EOF
        dat��ʼʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss"))
        If Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss") > Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss") Then
            dat����ʱ�� = CDate("1900-01-02 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        Else
            dat����ʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        End If



        If blnExit Then Exit Do
        datʱ�� = dat��ʼʱ��
        mrs�ϰ�ʱ���.MoveNext



        For i = j To lng�޺�
            ' If lngStart > lng��Լ Then blnExit = True: Exit For
            If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                j = i
                Exit For
            End If
            If strʱ�� <> Format(datʱ��, "HH:00") Then
                strʱ�� = Format(datʱ��, "HH:00")

                If InStr("," & str�ֶμ��, strʱ��) > 0 Then
                    lngʱ���� = Val(cllPro("K" & Replace(strʱ��, ":", "_")))
                Else
                    lngʱ���� = lngĬ�ϼ��
                End If
            End If

            If i > lng�̶����� Then
                With mrsAssign
                    .AddNew
                    !������Ŀ = str������Ŀ
                    !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                    !ʱ�� = Format(datʱ��, "hh:00:00")
                    !����ʱ�� = Format(DateAdd("n", lngʱ����, datʱ��), "hh:mm:00")
                    !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lngʱ����, datʱ��), "hh:mm")
                    !ʱ���� = lngʱ����
                    !�������� = IIf(lng������� >= lng�޺�, 0, 1)
                    !�Ƿ�ԤԼ = 0
                    !��� = i
                    !��ʹ�� = 0
                    .Update
                    lng������� = lng������� + IIf(lng������� >= lng�޺�, 0, 1)
                End With
            Else
                mrsAssign.Filter = "���=" & i
                If mrsAssign.RecordCount > 0 Then
                    lngĬ�ϼ�� = Nvl(mrsAssign!ʱ����, lngĬ�ϼ��)
                Else
                    lngĬ�ϼ�� = lngʱ����
                End If
            End If
            datʱ�� = DateAdd("n", IIf(i > lng�̶�����, lngʱ����, lngĬ�ϼ��), datʱ��)
        Next


        If i > lng�޺� And mbln��ſ��� Then
            blnExit = True
        End If
    Loop


    Call tbPage_SelectedChanged(tbPage(mlngSelIndex))





End Sub
Private Sub cmdɾ��_Click(Index As Integer)
    Dim blnDel As Boolean
    Dim lngSelX As Long
    Dim lngSelY As Long
    Dim i As Long, j As Long
    Dim lngCurrSn As Long
    Dim lngStartCol As Long
    With vsTime(Index)
        If .Col < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(.Row, .Col + 1)) = ""
        Else
                blnDel = True
        End If
        blnDel = blnDel And Trim(.TextMatrix(.Row, .Col)) <> "" And Not .Cell(flexcpFontUnderline, .Row, .Col)
        If Not blnDel Then Exit Sub
        If mbln��ſ��� Then
          lngSelX = .Row - (.Row Mod 2): lngSelY = .Col
          lngCurrSn = Val(.TextMatrix(lngSelX, lngSelY))
          .TextMatrix(lngSelX, lngSelY) = ""
          .TextMatrix(lngSelX + 1, lngSelY) = ""
          
          For i = lngSelX To .Rows - 1 Step 2
            lngStartCol = 1
            If i = lngSelX Then lngStartCol = lngSelY
            For j = lngStartCol To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    .TextMatrix(i, j) = lngCurrSn
                     lngCurrSn = lngCurrSn + 1
                End If
            Next
         Next
        End If
        cmdɾ��(Index).Visible = False
        cmdԤԼ(Index).Visible = False
        .SetFocus
    End With
End Sub

Private Sub cmd����ʱ��_Click()
    If AssignReapportion(Val(Me.txtTimeOut.Text), tbPage.Item(mlngSelIndex).Caption) = False Then Exit Sub
    Call tbPage_SelectedChanged(tbPage.Item(mlngSelIndex))
End Sub

Private Sub cmdԤԼ_Click(Index As Integer)
    If Not mbln��ſ��� Or mlngSelIndex < 0 Then Exit Sub
    If mlngSelIndex <> Index Then Exit Sub
    With vsTime(mlngSelIndex)
        If .MouseRow < 0 Or .MouseCol < 0 Then Exit Sub
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If .Cell(flexcpForeColor, .Row, .Col) = vbBlue Then
           .Cell(flexcpForeColor, .Row - (.Row Mod 2), .Col, .Row + (.Row + 1) Mod 2, .Col) = &H80000008
            .Cell(flexcpFontBold, .Row - (.Row Mod 2), .Col, .Row + (.Row + 1) Mod 2, .Col) = False
        Else
            .Cell(flexcpForeColor, .Row - (.Row Mod 2), .Col, .Row + (.Row + 1) Mod 2, .Col) = vbBlue
            .Cell(flexcpFontBold, .Row - (.Row Mod 2), .Col, .Row + (.Row + 1) Mod 2, .Col) = True
        End If
        mblnChange = True
        .SetFocus
    End With
End Sub

Private Sub Form_Load()
    Call LoadPageControl
    Call LoadPage
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With fraӦ����
        .Top = ScaleHeight - .Height - 50
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Visible = True
    End With
    With tbPage
        .Top = txtTimeOut.Top + txtTimeOut.Height + 50
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Height = fraӦ����.Top - .Top - 100
    End With
End Sub

Private Sub picPage_Resize(Index As Integer)
    Err = 0: On Error Resume Next

    With picPage(Index)
        vsTime(Index).Left = .ScaleLeft
        vsTime(Index).Top = .ScaleTop
        vsTime(Index).Width = .ScaleWidth
        vsTime(Index).Height = .ScaleHeight
    End With
End Sub

Private Sub InitRs(Optional ByVal blnInitRs As Boolean = True)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    If Not mrsAssign Is Nothing Then Exit Sub
    With mrsAssign
        Set mrsAssign = New ADODB.Recordset
        mrsAssign.Fields.Append "������Ŀ", adVarChar, 20
        mrsAssign.Fields.Append "��ʼʱ��", adVarChar, 20
        mrsAssign.Fields.Append "ʱ��", adVarChar, 20
        mrsAssign.Fields.Append "����ʱ��", adVarChar, 20
        mrsAssign.Fields.Append "ʱ���", adVarChar, 50
        mrsAssign.Fields.Append "ʱ����", adBigInt, 4
        mrsAssign.Fields.Append "��������", adBigInt, 10
        mrsAssign.Fields.Append "�Ƿ�ԤԼ", adBigInt, 18
        mrsAssign.Fields.Append "���", adBigInt, 18
        mrsAssign.Fields.Append "��ʹ��", adBigInt, 2
        mrsAssign.CursorLocation = adUseClient
        mrsAssign.LockType = adLockOptimistic
        mrsAssign.CursorType = adOpenStatic
        mrsAssign.Open
    End With
    If blnInitRs Then Call InitAssignRs
End Sub

Private Function InitAssignRs() As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng�̶� As Long  '�̶�����Ų��������
    Dim i As Long
    '��ʼ���ѷ������ݼ���
    If mEditType = EM_����_���� Then Exit Function
     On Error GoTo Hd
    If mEditType = EM_����_���� Or mEditType = EM_����_�޸� Or mEditType = EM_�ƻ�_���� Then
        strSQL = "Select ���, ���� As ������Ŀ, To_Char(��ʼʱ��, 'hh24:mi:ss') As ��ʼʱ��, To_Char(����ʱ��, 'hh24:mi:ss') As ����ʱ��,"
        strSQL = strSQL & vbCrLf & "         �Ƿ�ԤԼ , ��������,To_Char(��ʼʱ��, 'hh24') || ':00:00' As ʱ��,To_Char(��ʼʱ��, 'hh24:mi') || '-' || To_Char(����ʱ��, 'hh24:mi') As ʱ���"
        strSQL = strSQL & vbCrLf & " From �ҺŰ���ʱ�� Where ����ID=[1] "
        strSQL = strSQL & vbCrLf & " Order By ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    ElseIf mEditType = EM_�ƻ�_���� Or EM_�ƻ�_�޸� Then
        strSQL = "Select ���, ���� As ������Ŀ, To_Char(��ʼʱ��, 'hh24:mi:ss') As ��ʼʱ��, To_Char(����ʱ��, 'hh24:mi:ss') As ����ʱ��,"
        strSQL = strSQL & vbCrLf & "         �Ƿ�ԤԼ , ��������, To_Char(��ʼʱ��, 'hh24') || ':00:00' As ʱ��,To_Char(��ʼʱ��, 'hh24:mi') || '-' || To_Char(����ʱ��, 'hh24:mi') As ʱ���"
        strSQL = strSQL & vbCrLf & " From �Һżƻ�ʱ�� Where �ƻ�ID=[1] "
        strSQL = strSQL & vbCrLf & " Order By ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�ƻ�Id)
    End If
    Do While Not rsTmp.EOF
            With mrsAssign
                .AddNew
                !������Ŀ = Nvl(rsTmp!������Ŀ)
                !��ʼʱ�� = Nvl(rsTmp!��ʼʱ��, "00:00:00")
                !����ʱ�� = Nvl(rsTmp!����ʱ��, "00:00:00")
                !ʱ��� = Nvl(rsTmp!ʱ���, "__:__-__:__")
                !ʱ���� = DateDiff("n", CDate(!��ʼʱ��), CDate(!����ʱ��))
                !�������� = Val(Nvl(rsTmp!��������))
                !�Ƿ�ԤԼ = Val(Nvl(rsTmp!�Ƿ�ԤԼ))
                !ʱ�� = Nvl(rsTmp!ʱ��, "00:00:00")
                !��� = Val(Nvl(rsTmp!���))
                lng�̶� = 0
                If Not mrsHistory Is Nothing Then
                mrsHistory.Filter = "������Ŀ='" & Nvl(rsTmp!������Ŀ) & "'"
                    If mrsHistory.RecordCount > 0 Then
                        If CStr(mrsHistory!����ʱ��) >= CStr(Nvl(rsTmp!��ʼʱ��, "00:00:00")) Then
                            lng�̶� = 1
                        End If
                    End If
                End If
                !��ʹ�� = lng�̶�
                .Update
                
            End With
        rsTmp.MoveNext
    Loop
    Call AssignManage
'    If mblnInit Then
'        For i = 0 To 6
'            If tbPage.Item(i).Visible And tbPage.Item(i).Enabled Then
'                tbPage.Item(i).Selected = True
'                Call tbPage_SelectedChanged(tbPage.Item(i))
'                Exit For
'            End If
'        Next
'    End If
    InitAssignRs = True
Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   Dim str������Ŀ As String
   If Not mblnInit Then Exit Sub
   
   If Item.Index <> mlngSelIndex And mlngSelIndex <> -1 Then '
     If mlngSelIndex <> -1 And mblnChange Then
        If VsTimeValidate(mlngSelIndex) = False Then
            mblnOnChange = True
            tbPage.Item(mlngSelIndex).Selected = True
            mblnOnChange = False
            Exit Sub
        End If
     End If
     
     str������Ŀ = GetVsGridCaption(mlngSelIndex)
     If MoveAssign(str������Ŀ) = False Then
        If mlngSelIndex <> -1 Then tbPage.Item(mlngSelIndex).Selected = True
        Exit Sub
     End If
   End If
   
   If mblnOnChange Then Exit Sub
   mlngSelIndex = Item.Index
   SetStyle mbln��ſ���, Item.Index
   
   LoadTimePlan Item.Caption
   setVsGridSNStyle Item.Index
End Sub

Private Sub Initʱ���()
  '--------------------------------
  '����:��ȡ���°�ʱ���
  '--------------------------------
    Dim strTmp      As String
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim strDat      As String
    On Error GoTo Hd
    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "07:00:00 AND 12:00:00")
    strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����ϰ� = CDate("08:00:00")
    End If
   
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����°� = CDate("1900-01-01 12:00:00")
    End If
    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "14:00:00 AND 18:00:00")
    
     strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 14:00:00")
    End If
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����°� = CDate("1900-01-01 18:00:00")
    End If
    With t_ʱ��
         If .dat_�����ϰ� > .dat_�����°� Then
            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
         End If
         If .dat_�����ϰ� > .dat_�����°� Then
            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
         End If
    End With
    strSQL = _
    "       Select ʱ���, �ϰ�, �°� " & vbNewLine & _
    "       From (" & vbNewLine & _
    "           With Tb As (Select ʱ���,To_Date('1900-01-01 ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ʼʱ��," & vbNewLine & _
    "                               To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), -1, '1900-01-01 ', '1900-01-02 ') ||To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ֹʱ��," & _
    "                               Sign(��ʼʱ�� - ��ֹʱ��) As ����, " & vbNewLine & _
    "                                To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��, " & vbNewLine & _
    "                                To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��, " & vbNewLine & _
    "                                 To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��," & vbNewLine & _
    "                                 To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��"
    strSQL = strSQL & vbNewLine & _
    "                       From ʱ��� )" & vbNewLine & _
    "           Select ʱ���, '��' As ��ǩ, 0 As ��־, ��ʼʱ�� As �ϰ�, ��ֹʱ�� As �°�, ��ʼʱ��, ��ֹʱ��," & _
    "                  �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ��" & vbNewLine & _
    "            From Tb  Where (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) And " & _
    "                      (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
    "           Union All" & vbNewLine & _
    "           Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & vbNewLine & _
    "                        Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) As �°�, ��ʼʱ��, ��ֹʱ��, " & _
    "                        �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
    "           From Tb a Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
    "           Union All " & vbNewLine & _
    "            Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & _
    "                   Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) As �°�, ��ʼʱ��, ��ֹʱ��, �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
    "         From Tb a   Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��)" & vbNewLine & _
    "            ) b" & vbNewLine & _
    "         Order By ʱ���,�ϰ�"
     Set mrs�ϰ�ʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : AssignReapportion
'Description : ������·���
'Author      : ��⸣
'Date        : 05-November-2012 14:53:16
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'lng���ʱ��           Long              ByVal                .ʱ����
'str������Ŀ           String            ByVal                .����
'Output      :  �����Ƿ�ɹ�
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function AssignReapportion(ByVal lng���ʱ�� As Long, ByVal str������Ŀ As String) As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng�޺� As Long
    Dim lng��Լ As Long
    Dim dat��ʼʱ�� As Date
    Dim dat����ʱ�� As Date
    Dim lng��� As Long
    Dim strTmp As String
    Dim strʱ�� As String
    Dim str����ʱ�� As String
    Dim lngĬ�ϼ�� As Long
    Dim lng������� As Long
    Dim lng�̶����� As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim datʱ�� As Date
    If mrs�ϰ�ʱ��� Is Nothing Then
        Call Initʱ���
    End If

    If mrs�ϰ�ʱ��� Is Nothing Then Exit Function
    mrsRegPlan.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mrsRegPlan.RecordCount = 0 Then mrsRegPlan.Filter = 0: Exit Function
    lng�޺� = Nvl(mrsRegPlan!�޺���, 0): lng��Լ = Nvl(mrsRegPlan!��Լ��, 0)
    If lng��Լ = 0 Then lng��Լ = lng�޺�
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str������Ŀ & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Function
    End If


    strʱ�� = mrsRegPlan!�Ű�
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 0 Then
        MsgBox "������ʱ��Ϊ[" & strʱ�� & "]�����°�ʱ��,����!", vbOKOnly, Me.Caption
        Exit Function

    End If

    mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "' And ��ʹ��=0"
    Do While Not mrsAssign.EOF
        mrsAssign.Delete adAffectCurrent
        mrsAssign.MoveNext
    Loop
    mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mrsAssign.RecordCount <> 0 Then
        lng�̶����� = mrsAssign.RecordCount
        lngĬ�ϼ�� = Val(Nvl(mrsAssign!ʱ����, lng���ʱ��))
        Do While Not mrsAssign.EOF
            lng������� = lng������� + Val(Nvl(mrsAssign!��������))
            mrsAssign.MoveNext
        Loop
    End If
    mrsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs�ϰ�ʱ���.EOF
        dat��ʼʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss"))
        If Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss") > Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss") Then
            dat����ʱ�� = CDate("1900-01-02 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        Else
            dat����ʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        End If

        If blnExit Then Exit Do
        datʱ�� = dat��ʼʱ��
        mrs�ϰ�ʱ���.MoveNext

        If mbln��ſ��� Then

            For i = j To lng�޺�
                ' If lngStart > lng��Լ Then blnExit = True: Exit For
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                End If
                If i > lng�̶����� Then
                    With mrsAssign
                        .AddNew
                        !������Ŀ = str������Ŀ
                        !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                        !ʱ�� = Format(datʱ��, "hh:00:00")
                        !����ʱ�� = Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm:00")
                        !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm")
                        !ʱ���� = lng���ʱ��
                        !�������� = IIf(lng������� >= lng�޺�, 0, 1)
                        !�Ƿ�ԤԼ = 0
                        !��� = i
                        !��ʹ�� = 0
                        .Update
                        lng������� = lng������� + IIf(lng������� >= lng�޺�, 0, 1)
                    End With
                Else
                    mrsAssign.Filter = "���=" & i
                    If mrsAssign.RecordCount > 0 Then
                        lngĬ�ϼ�� = Nvl(mrsAssign!ʱ����, lngĬ�ϼ��)
                    Else
                        lngĬ�ϼ�� = lng���ʱ��
                    End If
                End If
                datʱ�� = DateAdd("n", IIf(i > lng�̶�����, lng���ʱ��, lngĬ�ϼ��), datʱ��)
            Next

        Else    '����ſ���

            Do While Not Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss")

                ' If lngStart > lng��Լ Then blnExit = True: Exit For
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then Exit Do

                If i > lng�̶����� Then
                    With mrsAssign
                        .AddNew
                        !������Ŀ = str������Ŀ
                        !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                        !ʱ�� = Format(datʱ��, "hh:00:00")
                        !����ʱ�� = Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm:00")
                        !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm")
                        !ʱ���� = lng���ʱ��
                        !�������� = IIf(lng������� >= lng��Լ, 0, 1)
                        !�Ƿ�ԤԼ = 1
                        !��� = i
                        !��ʹ�� = 0
                        .Update
                        lng������� = lng������� + IIf(lng������� >= lng��Լ, 0, 1)
                    End With
                Else
                    mrsAssign.Filter = "���=" & i
                    If mrsAssign.RecordCount > 0 Then
                        lngĬ�ϼ�� = Nvl(mrsAssign!ʱ����, lngĬ�ϼ��)
                    Else
                        lngĬ�ϼ�� = lng���ʱ��
                    End If
                End If
                datʱ�� = DateAdd("n", IIf(i > lng�̶�����, lng���ʱ��, lngĬ�ϼ��), datʱ��)
                i = i + 1
            Loop


        End If
        If i > lng�޺� And mbln��ſ��� Then
            blnExit = True
        End If
    Loop
    AssignReapportion = True
End Function


'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : AssignReapportion
'Description : ������·���
'Author      : ��⸣
'Date        : 05-November-2012 14:53:16
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'lng���ʱ��           Long              ByVal                .ʱ����
'str������Ŀ           String            ByVal                .����
'Output      :  �����Ƿ�ɹ�
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function AssignReapportion1(ByVal cllTime As Collection, ByVal str������Ŀ As String) As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng�޺� As Long
    Dim lng��Լ As Long
    Dim dat��ʼʱ�� As Date
    Dim dat����ʱ�� As Date
    Dim lng��� As Long
    Dim strTmp As String
    Dim strʱ�� As String
    Dim str����ʱ�� As String
    Dim lngĬ�ϼ�� As Long
    Dim lng���ʱ�� As Long
    Dim lng��� As Long
    Dim lng������� As Long
    Dim lng�̶����� As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim datʱ��  As Date
    Dim strPreʱ�� As String
    
    If Not mbln��ſ��� Then Exit Function
    
    If mrs�ϰ�ʱ��� Is Nothing Then
       Call Initʱ���
    End If
    
    If mrs�ϰ�ʱ��� Is Nothing Then Exit Function
    
    mrsRegPlan.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mrsRegPlan.RecordCount = 0 Then mrsRegPlan.Filter = 0: Exit Function
    lng�޺� = Nvl(mrsRegPlan!�޺���, 0): lng��Լ = Nvl(mrsRegPlan!��Լ��, 0)
    
    If lng��Լ = 0 Then lng��Լ = lng�޺�
    
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str������Ŀ & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    
    strʱ�� = mrsRegPlan!�Ű�
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 0 Then
        MsgBox "������ʱ��Ϊ[" & strʱ�� & "]�����°�ʱ��,����!", vbOKOnly, Me.Caption
        Exit Function
    
    End If
    
    mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "' And ��ʹ��=0"
    Do While Not mrsAssign.EOF
        mrsAssign.Delete adAffectCurrent
        mrsAssign.MoveNext
    Loop
    mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mrsAssign.RecordCount <> 0 Then
            lng�̶����� = mrsAssign.RecordCount
            lngĬ�ϼ�� = Val(Nvl(mrsAssign!ʱ����, lng���ʱ��))
            Do While Not mrsAssign.EOF
                lng������� = lng������� + Val(Nvl(mrsAssign!��������))
                mrsAssign.MoveNext
            Loop
    End If
    mrsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs�ϰ�ʱ���.EOF
        dat��ʼʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss"))
        If Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss") > Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss") Then
            dat����ʱ�� = CDate("1900-01-02 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        Else
            dat����ʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        End If
        
        If blnExit Then Exit Do
        datʱ�� = dat��ʼʱ��
        mrs�ϰ�ʱ���.MoveNext
        
       
        
            For i = j To lng�޺�
               ' If lngStart > lng��Լ Then blnExit = True: Exit For
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                   j = i
                   Exit For
                End If
                If i > lng�̶����� Then
                    With mrsAssign
                        .AddNew
                        !������Ŀ = str������Ŀ
                        !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                        !ʱ�� = Format(datʱ��, "hh:00:00")
                        !����ʱ�� = Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm:00")
                        !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm")
                        !ʱ���� = lng���ʱ��
                        !�������� = IIf(lng������� >= lng�޺�, 0, 1)
                        !�Ƿ�ԤԼ = 0
                        !��� = i
                        !��ʹ�� = 0
                        .Update
                         lng������� = lng������� + IIf(lng������� >= lng�޺�, 0, 1)
                    End With
                Else
                    mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "' And ���=" & i
                    If mrsAssign.RecordCount = 0 Then
                        lng��� = lngĬ�ϼ��
                    Else
                        lng��� = Val(Nvl(mrsAssign!ʱ����, lngĬ�ϼ��))
                    End If
                End If
                datʱ�� = DateAdd("n", IIf(i > lng�̶�����, lng���ʱ��, lng���), datʱ��)
            Next
           
       
        If i > lng�޺� And mbln��ſ��� Then
                blnExit = True
        End If
    Loop
    AssignReapportion1 = True
End Function


'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : AssignManage
'Description : ���Ѿ�����ĺ�������޺�����Լ���Ĺ�����д���
'Author      : ��⸣
'Date        : 05-November-2012 14:48:05
'Input       :
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function AssignManage() As Boolean
    Dim varData As Variant, varTemp As Variant, i As Long
    Dim j As Long, lngIndex As Long, p As Long, strTemp As String
    Dim lng�޺��� As Long, lng��Լ�� As Long, lng�������� As Long
    Dim lng����ԤԼ As Long, lngTmp  As Long, lngTemp As Long
    Dim str���ʱ�� As String, blnChange As Boolean
     
    varData = Split(mstr����, "|")
    lngIndex = -1
    For i = 0 To 6
        strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
        '�������Ӧ��ʱ��ı�
        If InStr("|" & mstr����, "|" & strTemp & ",") = 0 Or InStr("|" & mstrӦ��ʱ�� & "|", "|" & strTemp & "|") = 0 Then
            mrsAssign.Filter = "������Ŀ='" & strTemp & "'"
            Do While Not mrsAssign.EOF
                mrsAssign.Delete adAffectCurrent
                mrsAssign.Update
                mrsAssign.MoveNext
            Loop
        End If
    Next
    For i = 0 To UBound(varData)
        ''��һ,�޺���,��Լ��|�ܶ�,�޺���,��Լ��|....
        varTemp = Split(varData(i) & ",,,,", ",")
        If varTemp(0) <> "" Then
            lng�޺��� = Val(varTemp(1)): lng��Լ�� = Val(varTemp(2))
            If lng��Լ�� = 0 Then lng��Լ�� = lng�޺���
            str���ʱ�� = ""
            If Not mrsHistory Is Nothing Then
                mrsHistory.Filter = "������Ŀ='" & varTemp(0) & "'"
                If mrsHistory.RecordCount = 0 Then
                   str���ʱ�� = ""
                Else
                   str���ʱ�� = Nvl(mrsHistory!����ʱ��)
                End If
            End If
            mrsAssign.Filter = "������Ŀ='" & varTemp(0) & "'"
            mrsAssign.Sort = "���"
             
         
'             If str���ʱ�� <> "" Then
'                Do While Not mrsAssign.EOF
'                   If str���ʱ�� > Nvl(mrsAssign!��ʼʱ��) Then mrsAssign.Delete adAffectCurrent
'                   mrsAssign.MoveNext
'                Loop
'             End If
             
              lng�������� = 0
              blnChange = False
             Do While Not mrsAssign.EOF
                If lng�������� + Val(Nvl(mrsAssign!��������)) > IIf(mbln��ſ���, lng�޺���, lng��Լ��) Then
                    blnChange = True
                    If Val(Nvl(mrsAssign!��ʹ��)) = 0 Then
                        lngTmp = Val(mrsAssign!��������)
                        lngTemp = lng�������� + lngTmp - IIf(mbln��ſ���, lng�޺���, lng��Լ��)
                        If lngTmp <= lngTemp Then
                            lngTmp = 0
                        Else
                            lngTmp = lngTmp - lngTemp
                            lng�������� = lng�޺���
                        End If
                        mrsAssign!�������� = lngTmp
                        mrsAssign.Update
                        If mbln��ſ��� Then
                            mrsAssign.Delete adAffectCurrent
                        End If
                    End If
                Else
                    lng�������� = lng�������� + Val(Nvl(mrsAssign!��������))
                End If
                mrsAssign.MoveNext
             Loop
             If blnChange Then
                mrsAssign.Filter = "������Ŀ='" & varTemp(0) & "' And ��������>0"
                lng�������� = 0
                If mrsAssign.RecordCount = 0 Then mrsAssign.Filter = 0: AssignManage = True: Exit Function
                mrsAssign.Sort = "��� desc"
                mrsAssign.MoveFirst
                'lng��������
                Do While Not mrsAssign.EOF
                   lng�������� = lng�������� + Val(Nvl(mrsAssign!��������))
                   mrsAssign.MoveNext
                Loop
                mrsAssign.MoveFirst
                If lng�������� > IIf(mbln��ſ���, lng�޺���, lng��Լ��) Then
                   Do While Not mrsAssign.EOF
                      If Val(Nvl(mrsAssign!��ʹ��)) = 0 Then
                           lngTmp = Val(Nvl(mrsAssign!��������))
                           lngTemp = lng�������� - lng�޺���
                           If lngTemp >= lngTmp Then
                               mrsAssign!�������� = 0
                               mrsAssign.Update
                               lng�������� = lng�������� - lngTmp
                           Else
                               lngTmp = lngTmp - lngTemp
                               mrsAssign!�������� = lngTmp
                               mrsAssign.Update
                               lng�������� = lng�������� - lngTemp
                           End If
                      End If
                      If lng�������� <= lng�޺��� Then Exit Do
                      mrsAssign.MoveNext
                   Loop
                End If
             End If
        End If
    Next
    mrsAssign.Filter = 0
    If Not mrsHistory Is Nothing Then mrsHistory.Filter = 0
    AssignManage = True
End Function

Private Function VsTimeValidate(ByVal lngIndex As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤���õ���Լ���Ƿ����Ҫ��
    '���:lngIndex-ָ����ҳ��(���ڶ�Ӧ������):-1ʱ,��ʾ�����е�ҳ����м��
    '����:
    '����:У�Գɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-11-15 10:17:37
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngStep As Long, i As Long, j  As Long
    Dim lngԤԼ��   As Long, lng�޺��� As Long, lng��Լ�� As Long, lng���� As Long
    Dim str����   As String, str������Ŀ As String
    Dim lngPage As Long, lngPages As Long, lngStartPage As Long
    Dim blnNotSetTime As Boolean '��������ʱ���
    Dim blnAllowNums As Boolean '�����޺�����һ��
    Dim blnAllowYYNums As Boolean '����ԤԼ�������õ�ԤԼ����һ��
    Dim strCommand As String, blnʱ�� As Boolean '�ж�������ʱ�ε�,��Ҫ�������ʱ��ҳ�Ƿ�����
    On Error GoTo errHandle
        
    lngStartPage = 0: lngPages = tbPage.ItemCount - 1
    If lngIndex <> -1 Then lngStartPage = lngIndex: lngPages = lngIndex
    blnʱ�� = False
    For lngPage = lngStartPage To lngPages
        If mbln��ſ��� Then
            With vsTime(lngPage)
                For i = 0 To .Rows - 1 Step 2
                    For j = 1 To .Cols - 1
                       If .TextMatrix(i, j) <> "" Then
                           blnʱ�� = True: Exit For
                       End If
                    Next
                Next
            End With
        Else
                With vsTime(lngPage)
                    For i = 1 To .Rows - 1
                        For j = 1 To .Cols - 1 Step 2
                            If .TextMatrix(i, j) <> "" Then
                               blnʱ�� = True: Exit For
                            End If
                        Next
                    Next
                End With
        End If
    Next
    'δ����ʱ��
    If blnʱ�� = False Then VsTimeValidate = True: Exit Function
    
    For lngPage = lngStartPage To lngPages
        str������Ŀ = GetVsGridCaption(lngPage)
        mrsRegPlan.Filter = "������Ŀ='" & str������Ŀ & "'"
        If mrsRegPlan.RecordCount = 0 Then
            mrsRegPlan.Filter = 0
        Else
                lng�޺��� = Val(Nvl(mrsRegPlan!�޺���)): lng��Լ�� = Val(Nvl(mrsRegPlan!��Լ��))
                If lng��Լ�� = 0 Then lng��Լ�� = lng�޺���
                lng���� = 0: lngԤԼ�� = 0
                
                If mbln��ſ��� Then
                    'ר�Һż����Լ���Ƿ�����޺���
                    With vsTime(lngPage)
                        For i = 0 To .Rows - 1 Step 2
                            For j = 1 To .Cols - 1
                               If .TextMatrix(i, j) <> "" Then
                                     If .Cell(flexcpForeColor, i, j, i, j) = vbBlue Then
                                         lngԤԼ�� = lngԤԼ�� + 1
                                     End If
                                     lng���� = lng���� + 1
                               End If
                            Next
                        Next
                    End With
                    If lng���� < lng�޺��� Then
                        If lng���� = 0 Then
                           If lngIndex = -1 Then
                                If blnNotSetTime = False And blnʱ�� Then
                                        strCommand = zlCommFun.ShowMsgbox("����", "    �ڷ�ʱ��ҳ����δ���á�" & str������Ŀ & "����ʱ��,��ȷ��������ʱ���?" & vbCrLf & vbCrLf & _
                                         "���ǡ�:��ʾ��������ʱ��ν��б���" & vbCrLf & vbCrLf & _
                                         "�����ԡ�:��ʾ�������Ƶ�δ����ʱ��ε�����������,��������ʾ��" & vbCrLf & vbCrLf & _
                                         "����:��ʾ����������ʱ���,������������" & vbCrLf, "��(&O),����(&I),��(&C)", Me, vbQuestion)
                                        Select Case strCommand
                                        Case "��"
                                        Case "����"
                                             blnNotSetTime = True
                                         Case Else
                                            RaiseEvent zlSaveTimePageSelected(str������Ŀ)
                                            mblnNotBrush = True
                                            tbPage.Item(lngPage).Selected = True
                                            If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                            mblnNotBrush = False
                                            Exit Function
                                         End Select
                                End If
                           Else
                                If MsgBox("�ڷ�ʱ��ҳ����δ���á�" & str������Ŀ & "����ʱ��,��ȷ��������ʱ���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                    If lngIndex = -1 Then
                                        RaiseEvent zlSaveTimePageSelected(str������Ŀ)
                                        mblnNotBrush = True
                                        tbPage.Item(lngPage).Selected = True
                                        If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                        mblnNotBrush = False
                                    End If
                                    Exit Function
                                End If
                            End If
                        Else
                                If lngIndex = -1 Then
                                        If blnAllowNums = False Then
                                        
                                                strCommand = zlCommFun.ShowMsgbox("����", "    �ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "��������ʱ��εĺ���(" & lng���� & ")���޺���(" & lng�޺��� & ") ����,��ȷ������ǰ���õ�ʱ�α���?" & vbCrLf & vbCrLf & _
                                                 "���ǡ�:��ʾ�����޺����������һ��" & vbCrLf & vbCrLf & _
                                                 "�����ԡ�:��ʾ�����޺����������һ�£��������Ƶ�����,������ʾ��" & vbCrLf & vbCrLf & _
                                                 "����:��ʾ�������޺����������һ��,������������" & vbCrLf, "��(&O),����(&I),��(&C)", Me, vbQuestion)
                                                Select Case strCommand
                                                 Case "��"
                                                 Case "����"
                                                     blnAllowNums = True
                                                 Case Else
                                                    RaiseEvent zlSaveTimePageSelected(str������Ŀ)
                                                    mblnNotBrush = True
                                                    tbPage.Item(lngPage).Selected = True
                                                    If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                                    mblnNotBrush = False
                                                     Exit Function
                                                 End Select
                                        End If
                                   Else
                                        If MsgBox("�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "��������ʱ��εĺ���(" & lng���� & ")���޺���(" & lng��Լ�� & ") ����,��ȷ������ǰ���õ�ʱ�α���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                End If
                        End If
                    ElseIf lng���� > lng�޺��� Then
                        Call MsgBox("�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "��������ʱ��εĺ���(" & lng���� & ")�������޺���(" & lng��Լ�� & ") �㲻�ܰ���ǰ���õ�ʱ�α���!", vbQuestion + vbOKOnly + vbDefaultButton2, gstrSysName)
                        If lngIndex = -1 Then
                            RaiseEvent zlSaveTimePageSelected(str������Ŀ)
                            mblnNotBrush = True
                            tbPage.Item(lngPage).Selected = True
                            If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                            mblnNotBrush = False
                        End If
                        Exit Function
                    End If
                Else
                     '��ͨ�ż����Լ���Ƿ�����޺���
                    With vsTime(lngPage)
                        For i = 1 To .Rows - 1
                            For j = 1 To .Cols - 1 Step 2
                                If .TextMatrix(i, j) <> "" Then
                                    lngԤԼ�� = lngԤԼ�� + Val(.TextMatrix(i, j))
                                End If
                            Next
                        Next
                    End With
                End If
                If lngԤԼ�� > lng��Լ�� Then
                   MsgBox "�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "�������õ�ԤԼ��(" & lngԤԼ�� & ")������" & IIf(lng�޺��� = lng��Լ��, "�޺���(" & lng��Լ�� & ")", "��Լ��(" & lng��Լ�� & ")") & ",�㲻�ܰ���ǰ���ñ���!", vbOKOnly, Me.Caption
                    If lngIndex = -1 Then
                        RaiseEvent zlSaveTimePageSelected(str������Ŀ)
                        mblnNotBrush = True
                        tbPage.Item(lngPage).Selected = True
                        If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                        mblnNotBrush = False
                    End If
                   Exit Function
                End If
                If lngԤԼ�� < lng��Լ�� And lngԤԼ�� <> 0 Then
                    If lngIndex = -1 Then
                           If blnAllowYYNums = False Then
                                   strCommand = zlCommFun.ShowMsgbox("����", "    �ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "�������õ�ʵ��ԤԼ��(" & lngԤԼ�� & ") ����Լ��(" & lng��Լ�� & ") ����,��ȷ������ǰ���õ�ʱ�α���?" & vbCrLf & vbCrLf & _
                                    "���ǡ�:��ʾ������Լ����ԤԼ����һ��" & vbCrLf & vbCrLf & _
                                    "�����ԡ�:��ʾ������Լ����ԤԼ����һ�£��������Ƶ�����,������ʾ��" & vbCrLf & vbCrLf & _
                                    "����:��ʾ��������Լ����ԤԼ����һ��,������������" & vbCrLf, "��(&O),����(&I),��(&C)", Me, vbQuestion)
                                    Select Case strCommand
                                    Case "��"
                                    Case "����"
                                        blnAllowYYNums = True
                                    Case Else
                                       RaiseEvent zlSaveTimePageSelected(str������Ŀ)
                                       mblnNotBrush = True
                                       tbPage.Item(lngPage).Selected = True
                                       If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                       mblnNotBrush = False
                                        Exit Function
                                    End Select
                           End If
                      Else
                            If MsgBox("�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "�������õ�ʵ��ԤԼ��(" & lngԤԼ�� & ") ����Լ��(" & lng��Լ�� & ") ����,��ȷ������ǰ���õ�ʱ�α���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
        End If
    Next
    VsTimeValidate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : MoveAssign
'Description : ���������ű��浽�������ݼ�����
'Author      : ��⸣
'Date        : 05-November-2012 15:06:42
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'str������Ŀ           String            ByVal                .����
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function MoveAssign(ByVal str������Ŀ As String) As Boolean
    '�����������ŵ����ݼ���
    Dim nIndex As Long
    Dim lng��� As Long
    Dim i As Long, j As Long
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim lng���� As Long
    Dim blnԤԼ As Boolean
    Dim str���ʱ�� As String
    If Not mblnChange Then MoveAssign = True: Exit Function
    
    nIndex = GetVsGridIndex(str������Ŀ)
    
    'ɾ��û��ʹ�ò���
    mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "' and ��ʹ��=0"
    If mrsAssign.RecordCount > 0 Then
        Do While Not mrsAssign.EOF
            mrsAssign.Delete
            mrsAssign.MoveNext
        Loop
    End If
    
    If Not mbln��ſ��� Then
        With vsTime(nIndex)
          lng��� = 0
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                   If .TextMatrix(i, j) <> "" Then
                    
                    str��ʼʱ�� = Split(.TextMatrix(i, j), "-")(0)
                    str����ʱ�� = Split(.TextMatrix(i, j), "-")(1)
                    lng���� = Val(.TextMatrix(i, j + 1))
                    lng��� = lng��� + 1
                    blnԤԼ = True
                    
                    str���ʱ�� = ""
                    If Not mrsHistory Is Nothing Then
                        mrsHistory.Filter = "������Ŀ='" & str������Ŀ & "'"
                        If mrsHistory.RecordCount = 0 Then
                            str���ʱ�� = ""
                            mrsHistory.Filter = 0
                        Else
                            str���ʱ�� = Nvl(mrsHistory!����ʱ��)
                            mrsHistory.Filter = 0
                        End If
                    End If
                    
                    If (str���ʱ�� <> "" And str��ʼʱ�� > str���ʱ��) Or str���ʱ�� = "" Then
                        With mrsAssign
                            .AddNew
                            !������Ŀ = str������Ŀ
                            !��ʼʱ�� = str��ʼʱ��
                            !����ʱ�� = str����ʱ��
                            !ʱ��� = str��ʼʱ�� & "-" & str����ʱ��
                            !�������� = lng����
                            !��� = lng���
                            !��ʹ�� = 0
                            !�Ƿ�ԤԼ = 1
                            .Update
                        End With
                    End If
                   End If
                Next
            Next
        End With
        mblnChange = False
        MoveAssign = True
        Exit Function
    End If
    
    
    '��ſ���
    
    With vsTime(nIndex)
        For i = 0 To .Rows - 1 Step 2
            For j = 1 To .Cols - 1
                If Trim(.TextMatrix(i, j)) <> "" Then
                        str��ʼʱ�� = Split(.TextMatrix(i + 1, j) & "-", "-")(0)
                        str����ʱ�� = Split(.TextMatrix(i + 1, j) & "-", "-")(0)
                        lng��� = Val(.TextMatrix(i, j))
                        lng���� = 1
                        blnԤԼ = .Cell(flexcpForeColor, i, j) = vbBlue
                    If .Cell(flexcpFontUnderline, i, j) = False Then
                       
                        With mrsAssign
                            .AddNew
                            !������Ŀ = str������Ŀ
                            !��ʼʱ�� = str��ʼʱ��
                            !����ʱ�� = str����ʱ��
                            !ʱ�� = Format(str��ʼʱ��, "hh:00:00")
                            !ʱ��� = str��ʼʱ�� & "-" & str����ʱ��
                            !�������� = lng����
                            !��� = lng���
                            !��ʹ�� = 0
                            !�Ƿ�ԤԼ = IIf(blnԤԼ, 1, 0)
                            .Update
                        End With
                    ElseIf .Cell(flexcpFontUnderline, i, j) Then
                        ' �̶�����Ϣ,���ܸı��Ƿ�ԤԼ,����Ҳֻ�ɸı��Ƿ�ԤԼ
                        With mrsAssign
                            .Filter = "���=" & lng��� & " And ��ʼʱ��='" & Format(str��ʼʱ��, "hh:mm:00") & "'"
                            If .RecordCount > 0 Then
                                !�Ƿ�ԤԼ = IIf(blnԤԼ, 1, 0)
                                .Update
                            End If
                        End With
                    End If
                End If
            Next
        Next
    End With
    mblnChange = False
    MoveAssign = True
    Exit Function
End Function
Private Function ConvertToDate(ByVal strDate As String, Optional ByVal haveYear = False) As String
    '**********************************************************
    '���ַ���ת����oracle���ݿ��ܹ�ʶ�������
    '**********************************************************
    Select Case haveYear
    Case True:
        ConvertToDate = "To_Date('" & strDate & "', 'YYYY-MM-DD HH24:MI:SS')"
    Case False:
        ConvertToDate = "To_Date('" & strDate & "', 'HH24:MI:SS')"
    End Select
End Function

Private Sub SetStyle(ByVal bln��ſ��� As Boolean, ByVal lngIndex As Long)
    '����
    Dim i As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    If lngIndex > vsTime.UBound Then Exit Sub
    If Not mblnInit Then Exit Sub
    With vsTime(lngIndex)
        If bln��ſ��� Then
             
            If .Cols <= 1 Then Exit Sub
            .Rows = 0
            .FixedCols = 1
            .MergeCellsFixed = flexMergeFree
            .MergeCol(0) = True
            .FixedAlignment(0) = flexAlignRightTop
            .ColAlignment(0) = flexAlignRightTop
            lngWidth = 1275
'            lngHeight = 800
'            For i = 1 To .Cols - 1
'                .ColWidth(i) = lngWidth
'                .ColAlignment(i) = 4
'            Next
'            .ColAlignment(0) = 3
'            .ColWidth(0) = lngWidth
'            For i = 0 To .Rows - 1
'                 .RowHeight(i) = lngHeight
'            Next
'           If .Rows > 0 And .Cols > 0 Then
'                .Cell(flexcpFontBold, 0, 1, .Rows - 1, .Cols - 1) = True
'                .Cell(flexcpFontSize, 0, 1, .Rows - 1, .Cols - 1) = 9
'                .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 18
'           End If
           
        Else
             .Clear
             .Cols = 8: .Rows = 1
             .MergeCol(0) = False
            .FixedCols = 0
            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedRows = 1
            
            .RowHeightMax = 400: .RowHeightMin = 400
            For i = 0 To .Cols - 1 Step 2
              .TextMatrix(0, i) = "ʱ���"
            Next
            For i = 1 To .Cols - 1 Step 2
              .TextMatrix(0, i) = "ԤԼ����"
            Next
            For i = 0 To .Cols - 1
               .ColAlignment(i) = flexAlignCenterCenter
               .ColWidth(i) = 1200
            Next
        End If
'        If blnʱ�� Then
'            .Clear
'            .FixedCols = 1
'            .FixedRows = 0
'            .Rows = 1
'        Else
'
'        End If
    End With
End Sub

Private Sub setVsGridSNStyle(ByVal lngIndex As Long)
 '�����ʱ����vsFex����������ݺ���Ҫ�������ñ����ʽ
 '****************************************
'�Ա����ʽ��������
'****************************************
    Dim i           As Long
    Dim lngWidth    As Long
    Dim X           As Long
    Dim Y           As Long
    Dim j           As Long
    Dim lngHeight   As Long
   
    If vsTime(lngIndex).Cols <= 1 Then Exit Sub
    If mbln��ſ��� Then
        With vsTime(lngIndex)
            For i = 1 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
             Next
             .ColWidth(0) = 1200
             .FixedAlignment(0) = flexAlignRightTop
             .ColAlignment(0) = flexAlignRightTop
             If .Rows > 0 Then
                .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
                .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
             End If
    '��ʱ������ü������
    
     
         End With
    Else
    
    End If
    With vsTime(lngIndex)
         If (mbln��ſ��� And .Rows = 0) Or (mbln��ſ��� = False And .Rows = 1) Then Exit Sub
         For i = IIf(mbln��ſ���, 0, 1) To .Rows - 1 Step 2
             .Cell(flexcpBackColor, i, IIf(mbln��ſ���, 1, 0), i, .Cols - 1) = &HE0E0D3
         Next
    End With

End Sub

 
 
 
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : LoadTimePlan
'Description :
'Author      : ��⸣
'Date        : 05-November-2012 14:41:41
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'str������Ŀ           String            ByVal             ����           .
'Output      :  ����ʱ����Ƿ�ɹ�
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Function LoadTimePlan(ByVal str������Ŀ As String) As Boolean
    Dim nIndex As Integer
    Dim i As Long, r As Long
    Dim strTime As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strʱ�� As String
    Dim strData As String
    If mrsAssign Is Nothing Then Exit Function
    nIndex = GetVsGridIndex(str������Ŀ)
    cmdԤԼ(nIndex).Visible = False
    cmdɾ��(nIndex).Visible = False
    If Not mbln��ſ��� Then
        With vsTime(nIndex)
             
             mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
            mrsAssign.Sort = "��� asc "
               r = 1: i = -1
            Do While Not mrsAssign.EOF
                i = i + 1
                If i * 2 > .Cols - 2 Then r = r + 1: i = 0
                strData = Val(Nvl(mrsAssign!��������))
                strTime = mrsAssign!ʱ���
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                If Val(Nvl(mrsAssign!��ʹ��)) = 1 Then
                    .Cell(flexcpFontUnderline, r, i * 2, r, i * 2 + 1) = True
                Else
                   '������ɫ����
                End If
                mrsAssign.MoveNext
            Loop
             mrsAssign.Filter = 0
        End With
        LoadTimePlan = True
        Exit Function
    End If
    
    '-��ſ���
    With vsTime(nIndex)
        .Cols = 1: .FixedCols = 1
        .Rows = 0: .FixedRows = 0
        .Cols = 2: .Clear
        lngRow = -1: lngCol = 0
        mrsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
        If mrsAssign.RecordCount = 0 Then mrsAssign.Filter = 0: Exit Function
        i = 1
        mrsAssign.Sort = "��� asc "
        Do While Not mrsAssign.EOF
             lngCol = lngCol + 1
             If strʱ�� <> Nvl(mrsAssign!ʱ��) Then lngRow = lngRow + 2: lngCol = 1
             If lngCol = 1 Then
                strʱ�� = Nvl(mrsAssign!ʱ��)
                If lngRow > .Rows - 1 Then .Rows = .Rows + 2
                 .TextMatrix(lngRow - 1, 0) = Format(strʱ��, "hh:mm")
                 .TextMatrix(lngRow, 0) = Format(strʱ��, "hh:mm")
             End If
             strData = mrsAssign!���
             strTime = mrsAssign!ʱ���
            If lngCol > .Cols - 1 Then .Cols = .Cols + 1
            If lngRow > .Rows - 1 Then .Rows = .Rows + 2
             .TextMatrix(lngRow - 1, lngCol) = strData
             .TextMatrix(lngRow, lngCol) = strTime
            If Val(Nvl(mrsAssign!�Ƿ�ԤԼ)) = 1 Then
                .Cell(flexcpForeColor, lngRow - 1, lngCol, lngRow, lngCol) = vbBlue
                .Cell(flexcpFontBold, lngRow - 1, lngCol, lngRow, lngCol) = True
            End If
            If Val(Nvl(mrsAssign!��ʹ��)) = 1 Then
                    .Cell(flexcpFontUnderline, lngRow - 1, lngCol, lngRow, lngCol) = True
            Else
               '������ɫ����
            End If
            mrsAssign.MoveNext
        Loop
        If .Rows = 0 Then .Rows = 1
    End With
End Function
 
 

Private Sub txtTimeOut_Change()
    If Val(txtTimeOut.Text) > 1440 Then txtTimeOut.Text = 1440
End Sub
 
Private Sub vsTime_BeforeRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Not mbln��ſ��� Then
        vsTime(Index).Editable = IIf(NewCol Mod 2 = 1, flexEDKbd, flexEDNone)
          cmdԤԼ(Index).Visible = False: Exit Sub
    End If
    If NewRow < 0 Or NewCol < 0 Then Exit Sub
    
    SetCtrlMove Index, NewRow - (NewRow) Mod 2, NewCol
    If mbln��ſ��� Then vsTime(Index).Editable = flexEDNone: Exit Sub
    
    With vsTime(Index)
        .Editable = IIf(NewCol Mod 2 = 1, flexEDKbd, flexEDNone)
    End With
End Sub

Private Sub SetCtrlMove(ByVal Index As Integer, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnDel As Boolean
    With vsTime(Index)
        If mbln��ſ��� Then
            If Trim(.TextMatrix(NewRow, NewCol)) = "" Then
                cmdɾ��(Index).Visible = False
                cmdԤԼ(Index).Visible = False
                Exit Sub
            End If
            cmdɾ��(Index).Left = .Cell(flexcpLeft, NewRow, NewCol) + .Cell(flexcpWidth, NewRow, NewCol) - cmdɾ��(Index).Width
            If .Row Mod 2 <> 0 Then
                cmdɾ��(Index).Top = .Cell(flexcpTop, NewRow, NewCol)
            Else
                cmdɾ��(Index).Top = .Cell(flexcpTop, NewRow, NewCol)
            End If
            cmdԤԼ(Index).Left = .Cell(flexcpLeft, NewRow, NewCol)
            cmdԤԼ(Index).Top = cmdɾ��(Index).Top
            If NewCol < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(NewRow, NewCol + 1)) = ""
            Else
                blnDel = True
            End If
             
            blnDel = blnDel And Trim(.TextMatrix(NewRow, NewCol)) <> "" And Not .Cell(flexcpFontUnderline, NewRow, NewCol)
            cmdɾ��(Index).Visible = blnDel And mbln��ſ���
            cmdԤԼ(Index).Visible = True 'Val(txt��Լ.Text) <> 0
        Else
            cmdԤԼ(Index).Left = .Cell(flexcpTop, NewRow, NewCol)
            cmdԤԼ(Index).Top = .Cell(flexcpLeft, NewRow, NewCol)
            cmdԤԼ(Index).Visible = False
'            cmdԤԼ.Visible = Val(txt��Լ.Text) <> 0
        End If
    End With
End Sub

 

'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : vsTime_KeyDown
'Description : ���� �����¼�,��Ҫ����,��ſ��� ��ʱ��,��
'Author      : ��⸣
'Date        : 09-11-2012 05:58:34
'Input       :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Parameter Name    Parameter Type    Parameter Passing    Parameter Description
'Index             Integer           ByRef                .
'KeyCode           Integer           ByRef                .
'Shift             Integer           ByRef                .
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Private Sub vsTime_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'
    If Not mbln��ſ��� Then Exit Sub
     
    With vsTime(Index)
           
        If (.Row < 0 Or .Col < 1) Or (.Row > .Rows - 1 Or .Col > .Cols - 1) Then Exit Sub 'û����Ч��Ԫ����
        If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Sub
        If KeyCode = 13 Then
            Call cmdԤԼ_Click(Index)
            Exit Sub
        End If
        
        If KeyCode = 46 Then 'delete
            '�����:51429
            If cmdɾ��(Index).Visible = False Then Exit Sub
            If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Sub
            Call cmdɾ��_Click(Index)
        End If
     End With
End Sub

Private Sub vsTime_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13) Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vsTime_LostFocus(Index As Integer)
 '
 If Trim(vsTime(Index).EditText) <> "" Then
    With vsTime(Index)
        .TextMatrix(.Row, .Col) = .EditText
        mblnChange = True
    End With
 End If
'If mblnChange Then Stop
End Sub

Private Sub vsTime_ValidateEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mblnChange = True
End Sub


Public Property Get IsInit() As Boolean
Attribute IsInit.VB_Description = "�Ƿ񾭹��˳�ʼ��"
    IsInit = mblnInit
End Property



'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
'Name        : zl_CheckMoveAssign
'Description : ����Ƿ���ŷ����Ƿ�ı�,����Ѹı���,���ú���,���·������
'Author      : ��⸣
'Date        : 14-11-2012 10:53:40
'Input       :
'Output      :
'-=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=--=-
Public Function zl_CheckMoveAssign(Optional ByVal lngIndex As Long = -1) As Boolean
    Dim str������Ŀ As String
    If lngIndex = -1 Then lngIndex = mlngSelIndex
    If lngIndex = -1 Then zl_CheckMoveAssign = True: Exit Function
    If Not mblnChange Then zl_CheckMoveAssign = True: Exit Function
    
    If lngIndex < 0 Or lngIndex > 6 Then Exit Function
    If Not VsTimeValidate(lngIndex) Then Exit Function
    
    str������Ŀ = GetVsGridCaption(lngIndex)
    zl_CheckMoveAssign = MoveAssign(str������Ŀ)
End Function

Public Property Get ��ſ���() As Boolean
        ��ſ��� = mbln��ſ���
End Property

Public Property Let ��ſ���(ByVal vNewValue As Boolean)
        mbln��ſ��� = vNewValue
End Property
