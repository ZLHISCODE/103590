VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO0EA1~1.OCX"
Begin VB.Form frmEPRAuditPati 
   Caption         =   "���˲�����д���"
   ClientHeight    =   6930
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   10455
   Icon            =   "frmEPRAuditPati.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   10455
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6555
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRAuditPati.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15558
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
   Begin VSFlex8Ctl.VSFlexGrid vfgPati 
      Height          =   5655
      Left            =   105
      TabIndex        =   1
      Top             =   720
      Width           =   3630
      _cx             =   6403
      _cy             =   9975
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
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgAudit 
      Height          =   2400
      Left            =   3945
      TabIndex        =   2
      Top             =   750
      Width           =   6285
      _cx             =   11086
      _cy             =   4233
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
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   250
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
      WordWrap        =   -1  'True
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   1755
      Top             =   15
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRAuditPati.frx":70E4
      Left            =   525
      Top             =   75
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRAuditPati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'����
'-----------------------------------------------------
Private Enum mCol
    ��־ = 0: �¼�Ե��: Ӧд����: ����: ����ʱ��: Ҫ��ʱ��: ���ʱ��: ��ɼ�¼id: ��ǰʱ��: ��ע˵��
End Enum

Const conPane_Pati = 1
Const conPane_Audit = 2
Const conPane_Word = 3

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mlngDeptId As Long      '����id
Private mstrDeptName As String  '������
Private mintKind As Integer     '��������
Private mstrDateFrom As String  '��ʼ����
Private mstrDateTo As String    '��������
Private mstrEvent As String     '�����¼���Χ

Private WithEvents mfrmWord As frmDockEPRContent     '�������ݴ���
Attribute mfrmWord.VB_VarHelpID = -1

'-----------------------------------------------------
'��ʱ����
'-----------------------------------------------------
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim cbrToolBar As CommandBar

Dim rsTemp As New ADODB.Recordset
Dim strSQL As String
Dim lngCount As Long, lngRow As Long, lngCol As Long

Public Sub ShowMe(frmParent As Object, lngDeptId As Long, strDeptName As String, _
    intKind As Integer, strDateFrom As String, strDateTo As String, _
    Optional strEvent As String)
    mlngDeptId = lngDeptId: mstrDeptName = strDeptName
    mintKind = intKind: mstrDateFrom = strDateFrom: mstrDateTo = strDateTo
    mstrEvent = strEvent
    Me.Caption = Me.Caption & " - " & mstrDeptName
    
    Call RefreshData
    Me.Show vbModal, frmParent
End Sub

Private Sub RefreshData()
    Select Case mintKind
    Case 1
        Select Case mstrEvent
        Case "����"
            strSQL = "Select ����id, ID, �����, ����, �Ա�, To_Char(ִ��ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, ִ���� As ҽ��" & vbNewLine & _
                    "From ���˹Һż�¼" & vbNewLine & _
                    "Where ִ�в���id + 0 = [1] And Nvl(ִ��״̬, 0) <> 0 And Nvl(����, 0) <> 1 And" & vbNewLine & _
                    "      �Ǽ�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By ִ��ʱ��"
        Case "����"
            strSQL = "Select ����id, ID, �����, ����, �Ա�, To_Char(ִ��ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, ִ���� As ҽ��" & vbNewLine & _
                    "From ���˹Һż�¼" & vbNewLine & _
                    "Where ִ�в���id + 0 = [1] And Nvl(ִ��״̬, 0) <> 0 And Nvl(����, 0) = 1 And" & vbNewLine & _
                    "      �Ǽ�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By ִ��ʱ��"
        Case Else
            strSQL = "Select ����id, ID, �����, ����, �Ա�, To_Char(ִ��ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��, ִ���� As ҽ��" & vbNewLine & _
                    "From ���˹Һż�¼" & vbNewLine & _
                    "Where ִ�в���id + 0 = [1] And Nvl(ִ��״̬, 0) <> 0 And" & vbNewLine & _
                    "      �Ǽ�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By ִ��ʱ��"
        End Select
    Case 2
        Select Case mstrEvent
        Case "��Ժ"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.��Ժʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select ����id, ��ҳid, To_Char(Max(��ʼʱ��), 'yyyy-mm-dd hh24:mi') As ��Ժʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ʼԭ�� In (1, 2, 9) And ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By ����id, ��ҳid) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.��Ժʱ��"
        Case "ת��"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.ת��ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid, To_Char(��ʼʱ��, 'yyyy-mm-dd hh24:mi') As ת��ʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ʼԭ�� = 3 And ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.ת��ʱ��"
        Case "��Ժ"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'yyyy-mm-dd hh24:mi') As ��Ժ����" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P" & vbNewLine & _
                    "Where I.����id = P.����id And P.��Ժ����id + 0 = [1] And P.��Ժ��ʽ <> '����' And" & vbNewLine & _
                    "      P.��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.��Ժ����"
        Case "����"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'yyyy-mm-dd hh24:mi') As ��������" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P" & vbNewLine & _
                    "Where I.����id = P.����id And P.��Ժ����id + 0 = [1] And P.��Ժ��ʽ = '����' And" & vbNewLine & _
                    "      P.��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.��Ժ����"
        Case "ת��"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.ת��ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid, To_Char(��ֹʱ��, 'yyyy-mm-dd hh24:mi') As ת��ʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ֹԭ�� = 3 And ��ֹʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.ת��ʱ��"
        Case "����"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, ����ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select R.����id, R.��ҳid, To_Char(S.�״�ʱ��, 'yyyy-mm-dd hh24:mi') As ����ʱ��" & vbNewLine & _
                    "       From ����ҽ����¼ R, ����ҽ������ S" & vbNewLine & _
                    "       Where R.ID = S.ҽ��id And R.������� = 'F' And R.���id Is Null And R.ҽ����Ч = 1 And" & vbNewLine & _
                    "             (R.ҽ��״̬ = 8 Or R.ҽ��״̬ = 9) And R.���˿���id + 0 = [1] And" & vbNewLine & _
                    "             S.�״�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.����ʱ��"
        Case Else
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, P.��Ժ����" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id = [1] And" & vbNewLine & _
                    "             (��ʼԭ�� In (1, 2, 3, 9) And ��ʼʱ�� <= To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or" & vbNewLine & _
                    "             ��ֹԭ�� In (1, 3, 10) And (��ֹʱ�� >= To_Date([2], 'yyyy-mm-dd') Or ��ֹʱ�� Is Null))) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By P.��Ժ����"
        End Select
    Case 4
        Select Case mstrEvent
        Case "��Ժ"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.��Ժʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select ����id, ��ҳid, To_Char(Max(��ʼʱ��), 'yyyy-mm-dd hh24:mi') As ��Ժʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ʼԭ�� In (1, 2, 9) And ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By ����id, ��ҳid) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.��Ժʱ��"
        Case "ת��"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.ת��ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid, To_Char(��ʼʱ��, 'yyyy-mm-dd hh24:mi') As ת��ʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ʼԭ�� = 3 And ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.ת��ʱ��"
        Case "��Ժ"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'yyyy-mm-dd hh24:mi') As ��Ժ����" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P" & vbNewLine & _
                    "Where I.����id = P.����id And P.��ǰ����id + 0 = [1] And P.��Ժ��ʽ <> '����' And" & vbNewLine & _
                    "      P.��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.��Ժ����"
        Case "����"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'yyyy-mm-dd hh24:mi') As ��������" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P" & vbNewLine & _
                    "Where I.����id = P.����id And P.��ǰ����id + 0 = [1] And P.��Ժ��ʽ = '����' And" & vbNewLine & _
                    "      P.��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "Order By P.��Ժ����"
        Case "ת��"
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, L.ת��ʱ��" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid, To_Char(��ֹʱ��, 'yyyy-mm-dd hh24:mi') As ת��ʱ��" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id + 0 = [1] And ��ֹԭ�� = 3 And ��ֹʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By L.ת��ʱ��"
        Case Else
            strSQL = "Select P.����id, P.��ҳid, P.סԺ��, I.����, I.�Ա�, P.��Ժ����" & vbNewLine & _
                    "From ������Ϣ I, ������ҳ P," & vbNewLine & _
                    "     (Select Distinct ����id, ��ҳid" & vbNewLine & _
                    "       From ���˱䶯��¼" & vbNewLine & _
                    "       Where ����id = [1] And" & vbNewLine & _
                    "             (��ʼԭ�� In (1, 2, 3, 9) And ��ʼʱ�� <= To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400 Or" & vbNewLine & _
                    "             ��ֹԭ�� In (1, 3, 10) And (��ֹʱ�� >= To_Date([2], 'yyyy-mm-dd') Or ��ֹʱ�� Is Null))) L" & vbNewLine & _
                    "Where I.����id = P.����id And P.����id = L.����id And P.��ҳid = L.��ҳid" & vbNewLine & _
                    "Order By P.��Ժ����"
        End Select
    End Select
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo)
    With Me.vfgPati
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(1) = 0: .ColHidden(1) = True
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
    End With
    Call vfgPati_RowColChange
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    If Me.ActiveControl.Name = Me.vfgPati.Name Then
        Set objPrint.Body = Me.vfgPati
        objPrint.Title.Text = mstrDeptName & mstrEvent & "�����嵥"
    Else
        Set objPrint.Body = Me.vfgAudit
        objPrint.Title.Text = "���˲���ʱ�ޱ���"
        Set objAppRow = New zlTabAppRow
        Call objAppRow.Add(Me.vfgPati.TextMatrix(Me.vfgPati.FixedRows - 1, 2) & ":" & Me.vfgPati.TextMatrix(Me.vfgPati.Row, 2))
        Call objAppRow.Add("����:" & Me.vfgPati.TextMatrix(Me.vfgPati.Row, 3))
        Call objPrint.UnderAppRows.Add(objAppRow)
    End If
    
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    Me.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.Tag = ""
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_File_Open:
        Dim f As New frmEPRView
        f.ShowMe Me, CLng(Me.vfgAudit.TextMatrix(Me.vfgAudit.Row, mCol.��ɼ�¼id)), True
    Case conMenu_File_PrintSet: Call zlPrintSet
    Case conMenu_File_Preview:  Call zlRptPrint(0)
    Case conMenu_File_Print:    Call zlRptPrint(1)
    Case conMenu_File_Excel:    Call zlRptPrint(3)
    Case conMenu_File_Exit:     Unload Me
    Case conMenu_View_ToolBar_Button
        Me.cbsThis(2).Visible = Not Me.cbsThis(2).Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Text
        For Each cbrControl In Me.cbsThis(2).Controls
            cbrControl.Style = IIf(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
        Next
        Me.cbsThis.RecalcLayout
    Case conMenu_View_ToolBar_Size
        Me.cbsThis.Options.LargeIcons = Not Me.cbsThis.Options.LargeIcons
        Me.cbsThis.RecalcLayout
    Case conMenu_View_StatusBar
        Me.stbThis.Visible = Not Me.stbThis.Visible
        Me.cbsThis.RecalcLayout
    Case conMenu_View_Refresh: Call RefreshData
    Case conMenu_View_Jump
    Case conMenu_Tool_Monitor
        Dim lngRecordId As Long
        With Me.vfgAudit
            lngRecordId = Val(.TextMatrix(.Row, mCol.��ɼ�¼id))
        End With
        Call frmEPRAuditMonitor.zlRefList(lngRecordId)
    Case conMenu_Help_Help:     Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
    Case conMenu_Help_Web_Home: Call zlHomePage(Me.hWnd)
    Case conMenu_Help_Web_Forum '������̳
        Call zlWebForum(Me.hWnd)

    Case conMenu_Help_Web_Mail: Call zlMailTo(Me.hWnd)
    Case conMenu_Help_About:    Call ShowAbout(Me, App.Title, App.ProductName, App.major & "." & App.minor & "." & App.Revision)
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then Bottom = Me.stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_File_Open, conMenu_Tool_Monitor
        With Me.vfgAudit
            Control.Enabled = (Val(.TextMatrix(.Row, mCol.��ɼ�¼id)) > 0)
        End With
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Me.ActiveControl.Name = Me.vfgPati.Name Then
            Control.Enabled = (Me.vfgPati.Rows > Me.vfgPati.FixedRows)
        Else
            Control.Enabled = (Me.vfgAudit.Rows > Me.vfgAudit.FixedRows)
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_Pati
        Item.Handle = Me.vfgPati.hWnd
    Case conPane_Audit
        Item.Handle = Me.vfgAudit.hWnd
    Case conPane_Word
        If mfrmWord Is Nothing Then Set mfrmWord = New frmDockEPRContent
        Item.Handle = mfrmWord.hWnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Me.cbsThis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '�˵�����
    Me.cbsThis.ActiveMenuBar.Title = "�˵�"
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagStretched
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    cbrMenuBar.ID = conMenu_FilePopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��(&O)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "��ӡ����(&S)��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "�����&Excel��")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)"): cbrControl.BeginGroup = True
    End With

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    cbrMenuBar.ID = conMenu_ViewPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "������(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Jump, "��ת(&J)"): cbrControl.BeginGroup = True
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ToolPopup, "����(&T)", -1, False)
    cbrMenuBar.ID = conMenu_ToolPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "���ݼ��(&T)")
    End With
    
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    cbrMenuBar.ID = conMenu_HelpPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False  '����
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��"): cbrControl.BeginGroup = True
    End With
    
    With Me.cbsThis.ActiveMenuBar.Controls
        Select Case mintKind
        Case 1: Set cbrControl = .Add(xtpControlLabel, 0, "���ﲡ��(" & mstrDateFrom & "��" & mstrDateTo & ")")
        Case 2: Set cbrControl = .Add(xtpControlLabel, 0, "סԺ����(" & mstrDateFrom & "��" & mstrDateTo & ")")
        Case 4: Set cbrControl = .Add(xtpControlLabel, 0, "������(" & mstrDateFrom & "��" & mstrDateTo & ")")
        End Select
        cbrControl.flags = xtpFlagRightAlign
    End With
    
    '�����
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("O"), conMenu_File_Open
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F6, conMenu_View_Jump
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '���ò����ò˵�
    With Me.cbsThis.Options
        .AddHiddenCommand conMenu_File_Excel
        .AddHiddenCommand conMenu_View_Jump
    End With
    
    '-----------------------------------------------------
    '����������
    Set cbrToolBar = Me.cbsThis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Open, "��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "Ԥ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "��ӡ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Tool_Monitor, "���ݼ��"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With
    For Each cbrControl In cbrToolBar.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '-----------------------------------------------------
    '������ʾͣ������
    If mfrmWord Is Nothing Then Set mfrmWord = New frmDockEPRContent
    
    Dim panThis As Pane, panChild As Pane
    Set panThis = dkpMan.CreatePane(conPane_Pati, 300, 400, DockLeftOf, Nothing)
    panThis.Title = mstrEvent & "�����б�"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_Audit, 700, 100, DockRightOf, Nothing)
    panThis.Title = "����ʱ�����"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panChild = dkpMan.CreatePane(conPane_Word, 700, 300, DockBottomOf, panThis)
    panChild.Title = "��������"
    panChild.Options = PaneNoCaption

    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    
    '-----------------------------------------------------
    '����ָ�
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload mfrmWord: Set mfrmWord = Nothing
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub vfgAudit_GotFocus()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub vfgAudit_RowColChange()
    Dim lngRecordId As Long
    
    Err = 0: On Error Resume Next
    With Me.vfgAudit
        lngRecordId = Val(.TextMatrix(.Row, mCol.��ɼ�¼id))
    End With
    Err = 0: On Error GoTo 0
    If Me.Tag <> "" Then Exit Sub
    Call mfrmWord.zlRefresh(lngRecordId, "", True)
End Sub

Private Sub vfgPati_GotFocus()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub vfgPati_RowColChange()
    Dim lngPatiID As Long, lngPageId As Long
    Dim lngBalance As Long
    
    If Me.Tag <> "" Then Exit Sub
    lngPatiID = Me.vfgPati.TextMatrix(Me.vfgPati.Row, 0)
    lngPageId = Me.vfgPati.TextMatrix(Me.vfgPati.Row, 1)
    
    gstrSQL = "Zl_����ʱ�޼��_Neaten(" & lngPatiID & "," & lngPageId & "," & mintKind & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '��ȡʱ�޼������
    gstrSQL = "Select To_Char(�¼�ʱ��, 'yyyy-mm-dd hh24:mi ') || �䶯�¼� As �¼�Ե��, ������� || '-' || �������� As Ӧд����," & _
            "        Decode(Ψһ, 1, '��д', '��' || ���ں� || '����д') As ����, ����ʱ��, Ҫ��ʱ��, ���ʱ��, ��ɼ�¼id, Sysdate As ��ǰʱ��, Null As ��ע˵��" & _
            " From ����ʱ�޼��" & _
            " Where ����id = [1] And ��ҳid = [2] And (�������� = [3] Or �������� in (5,6) And [3]<>4) And Ҫ��ʱ�� - Sysdate < 2" & _
            " Order By ��������,�¼�ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngPatiID, lngPageId, mintKind)
    With Me.vfgAudit
        .Clear
        Set .DataSource = rsTemp
        
        .MergeCells = flexMergeFree: .MergeCol(mCol.�¼�Ե��) = True: .MergeCol(mCol.Ӧд����) = True
        .ColWidth(mCol.��־) = 250: .ColWidth(mCol.����ʱ��) = 1100: .ColWidth(mCol.Ҫ��ʱ��) = 1100: .ColWidth(mCol.���ʱ��) = 1100
        .ColWidth(mCol.��ɼ�¼id) = 0: .ColWidth(mCol.��ǰʱ��) = 0: .ColWidth(mCol.��ע˵��) = 2200
        
        .FixedAlignment(mCol.��־) = flexAlignCenterCenter
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .ColAlignment(lngCount) = flexAlignLeftTop
        Next
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, mCol.���ʱ��) = "" Then
                If .TextMatrix(lngCount, mCol.��ɼ�¼id) = "" Then
                    .TextMatrix(lngCount, mCol.��ע˵��) = "δ��д"
                Else
                    .TextMatrix(lngCount, mCol.��ע˵��) = "������д"
                End If
                lngBalance = Int((CDate(.TextMatrix(lngCount, mCol.��ǰʱ��)) - CDate(.TextMatrix(lngCount, mCol.Ҫ��ʱ��))) * 24)
                .TextMatrix(lngCount, mCol.��־) = "��"
                If lngBalance >= 0 Then
                    .Cell(flexcpForeColor, lngCount, mCol.��־, lngCount, mCol.��־) = RGB(255, 0, 0)
                    .TextMatrix(lngCount, mCol.��ע˵��) = .TextMatrix(lngCount, mCol.��ע˵��) & IIf(lngBalance = 0, "", ",�ѳ���" & lngBalance & "Сʱ")
                    .Cell(flexcpForeColor, lngCount, mCol.��ע˵��, lngCount, mCol.��ע˵��) = RGB(255, 0, 0)
                Else
                    If Abs(lngBalance) < 4 Then
                        .Cell(flexcpForeColor, lngCount, mCol.��־, lngCount, mCol.��־) = RGB(128, 128, 0)
                        .TextMatrix(lngCount, mCol.��ע˵��) = .TextMatrix(lngCount, mCol.��ע˵��) & ",ʣ��" & Abs(lngBalance) & "Сʱ,�뾡�����"
                    Else
                        .Cell(flexcpForeColor, lngCount, mCol.��־, lngCount, mCol.��־) = RGB(0, 0, 255)
                        .TextMatrix(lngCount, mCol.��ע˵��) = .TextMatrix(lngCount, mCol.��ע˵��) & ",ʣ��" & Abs(lngBalance) & "Сʱ,�밴ʱ���"
                    End If
                End If
            Else
                lngBalance = Int((CDate(.TextMatrix(lngCount, mCol.���ʱ��)) - CDate(.TextMatrix(lngCount, mCol.Ҫ��ʱ��))) * 24)
                If lngBalance > 0 Then
                    .TextMatrix(lngCount, mCol.��־) = "��"
                    .Cell(flexcpForeColor, lngCount, mCol.��־, lngCount, mCol.��־) = RGB(255, 0, 0)
                    .TextMatrix(lngCount, mCol.��ע˵��) = "���,������" & lngBalance & "Сʱ"
                    .Cell(flexcpForeColor, lngCount, mCol.��ע˵��, lngCount, mCol.��ע˵��) = RGB(255, 0, 0)
                Else
                    .TextMatrix(lngCount, mCol.��ע˵��) = "�������"
                End If
            End If
            .TextMatrix(lngCount, mCol.����ʱ��) = Format(.TextMatrix(lngCount, mCol.����ʱ��), "MM-dd hh:mm")
            .TextMatrix(lngCount, mCol.Ҫ��ʱ��) = Format(.TextMatrix(lngCount, mCol.Ҫ��ʱ��), "MM-dd hh:mm")
            .TextMatrix(lngCount, mCol.���ʱ��) = Format(.TextMatrix(lngCount, mCol.���ʱ��), "MM-dd hh:mm")
        Next
    End With
    Call vfgAudit_RowColChange
End Sub
