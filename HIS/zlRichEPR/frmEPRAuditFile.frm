VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO373F~1.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "CO0EA1~1.OCX"
Begin VB.Form frmEPRAuditFile 
   Caption         =   "�����ļ��������"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   Icon            =   "frmEPRAuditFile.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9630
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5835
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmEPRAuditFile.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14102
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
   Begin VSFlex8Ctl.VSFlexGrid vfgFile 
      Height          =   1980
      Left            =   75
      TabIndex        =   1
      Top             =   750
      Width           =   4500
      _cx             =   7937
      _cy             =   3492
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
   Begin VSFlex8Ctl.VSFlexGrid vfgEPRs 
      Height          =   2925
      Left            =   60
      TabIndex        =   2
      Top             =   2835
      Width           =   4500
      _cx             =   7937
      _cy             =   5159
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   2235
      Top             =   75
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmEPRAuditFile.frx":0E1C
      Left            =   525
      Top             =   150
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmEPRAuditFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------
'����
'-----------------------------------------------------
Const conPane_File = 1
Const conPane_EPRs = 2
Const conPane_Word = 3

'-----------------------------------------------------
'�������
'-----------------------------------------------------
Private mlngDeptId As Long      '����id
Private mstrDeptName As String  '������
Private mintKind As Integer     '��������
Private mstrDateFrom As String  '��ʼ����
Private mstrDateTo As String    '��������

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

Public Sub ShowMe(frmParent As Object, lngDeptId As Long, strDeptName As String, intKind As Integer, strDateFrom As String, strDateTo As String)
    mlngDeptId = lngDeptId: mstrDeptName = strDeptName
    mintKind = intKind: mstrDateFrom = strDateFrom: mstrDateTo = strDateTo
    Me.Caption = Me.Caption & " - " & mstrDeptName
    
    Call RefreshData
    Me.Show vbModal, frmParent
End Sub

Private Sub RefreshData()
    Dim intOut24h As Byte    '�Ƿ�����24Сʱ��Ժ��������0-������,1-���֣������Ƿ���24Сʱ�¼���Ӧ����ȷ��
    Select Case mintKind
    Case 1
        strSQL = "Select F.ID, F.���, F.����, F.�¼� || 'ʱ��д' As Ҫ��, P.�˴� As Ӧд��, W.�����, W.����д" & vbNewLine & _
                "From (Select F.ID, F.���, F.����, F.�¼�, F.Ψһ, F.��дʱ��" & vbNewLine & _
                "       From (Select F.ID, F.���, F.����, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "              From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "              Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 1) F" & vbNewLine & _
                "       Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [1]) F," & vbNewLine & _
                "     (Select E.�¼�, Decode(E.�¼�, '����', ����, '����', ����, '����', ����, '����', ����) As �˴�" & vbNewLine & _
                "       From (Select Sum(Decode(����, 1, 0, 1)) As ����, Sum(Decode(����, 1, 0, Decode(����, 1, 0, 1))) As ����," & vbNewLine & _
                "                     Sum(Decode(����, 1, 0, Decode(����, 1, 1, 0))) As ����, Sum(Decode(����, 1, 1, 0)) As ����" & vbNewLine & _
                "              From ���˹Һż�¼" & vbNewLine & _
                "              Where ִ�в���id = [1] And Nvl(ִ��״̬, 0) <> 0 And �Ǽ�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "                    To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                "            (Select Decode(Rownum, 1, '����', 2, '����', 3, '����', 4, '����') As �¼� From ������д�¼� Where Rownum < 5) E) P," & vbNewLine & _
                "     (Select �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����, Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "       From ���Ӳ�����¼" & vbNewLine & _
                "       Where �������� = 1 And ����id + 0 = [1] And ����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By �ļ�id) W" & vbNewLine & _
                "Where F.�¼� = P.�¼� And P.�˴� > 0 And F.ID = W.�ļ�id(+)" & vbNewLine & _
                "Order By F.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo)
    Case 2
        strSQL = "Select Sign(Nvl(Count(*), 0))" & vbNewLine & _
                "From (Select F.ID, F.ͨ��, A.����id" & vbNewLine & _
                "       From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "       Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And Q.�¼� In ('24Сʱ��Ժ', '24Сʱ����') And F.���� = 2) F" & vbNewLine & _
                "Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId)
        If rsTemp.RecordCount <= 0 Then
            intOut24h = 0
        Else
            intOut24h = rsTemp.Fields(0).Value
        End If
        strSQL = "Select F.ID, F.���, F.����, F.�¼� || Decode(Sign(F.��дʱ��), -1, 'ǰ', '��') || '��д' As Ҫ��," & vbNewLine & _
                "       Decode(F.Ψһ, 1, To_Char(P.�˴�), '<ѭ��>') As Ӧд��, W.�����, W.����д" & vbNewLine & _
                "From (Select F.ID, F.���, F.����, F.�¼�, F.Ψһ, F.��дʱ��" & vbNewLine & _
                "       From (Select F.ID, F.���, F.����, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "              From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "              Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 2) F" & vbNewLine & _
                "       Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [1]) F," & vbNewLine
        If intOut24h = 1 Then
            strSQL = strSQL & "     (Select E.�¼�, '��' As ʱ��, Decode(E.�¼�, '��Ժ', ��Ժ, '�״���Ժ', �״���Ժ, '�ٴ���Ժ', �ٴ���Ժ) As �˴�" & vbNewLine & _
                    "       From (Select Count(*) As ��Ժ, Sum(Decode(����Ժ, 1, 0, 1)) As �״���Ժ," & vbNewLine & _
                    "                     Sum(Decode(����Ժ, 1, '�ٴ���Ժ', 0)) As �ٴ���Ժ" & vbNewLine & _
                    "              From ������ҳ" & vbNewLine & _
                    "              Where ��Ժ����id + 0 = 36 And Nvl(��Ժ����, Sysdate + 1) - ��Ժ���� > 1 And" & vbNewLine & _
                    "                    ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                    "            (Select Decode(Rownum, 1, '��Ժ', 2, '�״���Ժ', 3, '�ٴ���Ժ') As �¼� From ������д�¼� Where Rownum < 4) E" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(Sign(��Ժ���� - ��Ժ���� - 1), -1, Decode(��Ժ��ʽ, '����', '24Сʱ����', '24Сʱ��Ժ')," & vbNewLine & _
                    "                      Decode(��Ժ��ʽ, '����', '����', '��Ժ')) As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                    "       From ������ҳ" & vbNewLine & _
                    "       Where ��Ժ����id + 0 = [1] And ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By Decode(Sign(��Ժ���� - ��Ժ���� - 1), -1, Decode(��Ժ��ʽ, '����', '24Сʱ����', '24Сʱ��Ժ')," & vbNewLine & _
                    "                        Decode(��Ժ��ʽ, '����', '����', '��Ժ'))" & vbNewLine
        Else
            strSQL = strSQL & "     (Select E.�¼�, '��' As ʱ��, Decode(E.�¼�, '��Ժ', ��Ժ, '�״���Ժ', �״���Ժ, '�ٴ���Ժ', �ٴ���Ժ) As �˴�" & vbNewLine & _
                    "       From (Select Count(*) As ��Ժ, Sum(Decode(����Ժ, 1, 0, 1)) As �״���Ժ," & vbNewLine & _
                    "                     Sum(Decode(����Ժ, 1, '�ٴ���Ժ', 0)) As �ٴ���Ժ" & vbNewLine & _
                    "              From ������ҳ" & vbNewLine & _
                    "              Where ��Ժ����id + 0 = 36 And" & vbNewLine & _
                    "                    ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400) P," & vbNewLine & _
                    "            (Select Decode(Rownum, 1, '��Ժ', 2, '�״���Ժ', 3, '�ٴ���Ժ') As �¼� From ������д�¼� Where Rownum < 4) E" & vbNewLine & _
                    "       Union All" & vbNewLine & _
                    "       Select Decode(��Ժ��ʽ, '����', '����', '��Ժ') As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                    "       From ������ҳ" & vbNewLine & _
                    "       Where ��Ժ����id + 0 = [1] And ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                    "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                    "       Group By Decode(��Ժ��ʽ, '����', '����', '��Ժ')" & vbNewLine
        End If
        strSQL = strSQL & "       Union All" & vbNewLine & _
                "       Select Decode(��ʼԭ��, 3, 'ת��', 7, '����') As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ���˱䶯��¼" & vbNewLine & _
                "       Where ����id + 0 = [1] And ��ʼԭ�� In (3, 7) And Nvl(���Ӵ�λ, 0) = 0 And" & vbNewLine & _
                "             ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��ʼԭ��, 3, 'ת��', 7, '����')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(��ֹԭ��, 3, 'ת��', 7, '����') As �¼�, 'ǰ' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ���˱䶯��¼" & vbNewLine & _
                "       Where ����id + 0 = [1] And ��ֹԭ�� In (3, 7) And Nvl(���Ӵ�λ, 0) = 0 And" & vbNewLine & _
                "             ��ֹʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��ֹԭ��, 3, 'ת��', 7, '����')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select R.�¼�, E.ʱ��, Decode(E.ʱ��, 'ǰ', R.ǰ�˴�, '��', R.���˴�) As �˴�" & vbNewLine & _
                "       From (Select Decode(R.�������, 'F', '����', Decode(I.��������, '7', '����', '����')) As �¼�," & vbNewLine & _
                "                     Sum(Decode(R.���˿���id, [1], 1, 0)) As ǰ�˴�, Sum(Decode(R.ִ�п���id, [1], 1, 0)) As ���˴�" & vbNewLine & _
                "              From ����ҽ����¼ R, ������ĿĿ¼ I, ����ҽ������ S" & vbNewLine & _
                "              Where R.ID = S.ҽ��id And R.������Ŀid = I.ID And" & vbNewLine & _
                "                    (R.������� = 'F' Or R.������� = 'Z' And I.�������� In ('7', '8')) And R.���id Is Null And" & vbNewLine & _
                "                    R.ҽ����Ч = 1 And (R.ҽ��״̬ = 8 Or R.ҽ��״̬ = 9) And" & vbNewLine & _
                "                    (R.���˿���id + 0 = [1] Or R.ִ�п���id + 0 = [1]) And S.�״�ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "                    To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "              Group By Decode(R.�������, 'F', '����', Decode(I.��������, '7', '����', '����'))) R," & vbNewLine & _
                "            (Select Decode(Rownum, 1, 'ǰ', 2, '��') As ʱ�� From ������д�¼� Where Rownum < 3) E) P," & vbNewLine
        
        strSQL = strSQL & "     (Select �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����, Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "       From ���Ӳ�����¼" & vbNewLine & _
                "       Where �������� = 2 And ����id + 0 = [1] And ����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By �ļ�id) W" & vbNewLine & _
                "Where F.�¼� = P.�¼� And Decode(Sign(F.��дʱ��), -1, 'ǰ', '��') = P.ʱ�� And P.�˴� > 0 And F.ID = W.�ļ�id(+)" & vbNewLine & _
                "Order By F.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo)
    Case 4
        strSQL = "Select F.ID, F.���, F.����, F.�¼� || Decode(Sign(F.��дʱ��), -1, 'ǰ', '��') || '��д' As Ҫ��," & vbNewLine & _
                "       Decode(F.Ψһ, 1, To_Char(P.�˴�), '<ѭ��>') As Ӧд��, W.�����, W.����д" & vbNewLine & _
                "From (Select F.ID, F.���, F.����, F.�¼�, F.Ψһ, F.��дʱ��" & vbNewLine & _
                "       From (Select F.ID, F.���, F.����, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "              From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "              Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 4) F" & vbNewLine & _
                "       Where F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = [1]) F," & vbNewLine
        strSQL = strSQL & "     (Select '��Ժ' As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ������ҳ" & vbNewLine & _
                "       Where ��Ժ����id + 0 = [1] And ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(��Ժ��ʽ, '����', '����', '��Ժ') As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ������ҳ" & vbNewLine & _
                "       Where ��ǰ����id + 0 = [1] And ��Ժ���� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��Ժ��ʽ, '����', '����', '��Ժ')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(��ʼԭ��, 3, 'ת��', 8, '����') As �¼�, '��' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ���˱䶯��¼" & vbNewLine & _
                "       Where ����id + 0 = [1] And ��ʼԭ�� In (3, 8) And Nvl(���Ӵ�λ, 0) = 0 And" & vbNewLine & _
                "             ��ʼʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��ʼԭ��, 3, 'ת��', 8, '����')" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select Decode(��ֹԭ��, 3, 'ת��', 8, '����') As �¼�, 'ǰ' As ʱ��, Count(*) As �˴�" & vbNewLine & _
                "       From ���˱䶯��¼" & vbNewLine & _
                "       Where ����id + 0 = [1] And ��ֹԭ�� In (3, 8) And Nvl(���Ӵ�λ, 0) = 0 And" & vbNewLine & _
                "             ��ֹʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By Decode(��ֹԭ��, 3, 'ת��', 8, '����')) P," & vbNewLine
        strSQL = strSQL & "     (Select �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����, Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "       From ���Ӳ�����¼" & vbNewLine & _
                "       Where �������� = 4 And ����id + 0 = [1] And ����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "             To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "       Group By �ļ�id) W" & vbNewLine & _
                "Where F.�¼� = P.�¼� And Decode(Sign(F.��дʱ��), -1, 'ǰ', '��') = P.ʱ�� And P.�˴� > 0 And F.ID = W.�ļ�id(+)" & vbNewLine & _
                "Order By F.���"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo)
    End Select
    With Me.vfgFile
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 0: .ColHidden(0) = True
        .ColAlignment(4) = flexAlignRightCenter
        For lngCount = 1 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
    End With
    Call vfgFile_RowColChange
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow
    
    If Me.ActiveControl.Name = Me.vfgFile.Name Then
        Set objPrint.Body = Me.vfgFile
        objPrint.Title.Text = mstrDeptName & "�����ļ��б�"
    Else
        Set objPrint.Body = Me.vfgEPRs
        objPrint.Title.Text = mstrDeptName & Me.vfgFile.TextMatrix(Me.vfgFile.Row, 2) & "�嵥"
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
    Case conMenu_File_Open
        Dim f As New frmEPRView
        f.ShowMe Me, CLng(Me.vfgEPRs.TextMatrix(Me.vfgFile.Row, 0)), True
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
        Call frmEPRAuditMonitor.zlRefList(Val(Me.vfgEPRs.TextMatrix(Me.vfgEPRs.Row, 0)))
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
        Control.Enabled = (Val(Me.vfgEPRs.TextMatrix(Me.vfgEPRs.Row, 0)) > 0)
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        If Me.ActiveControl.Name = Me.vfgFile.Name Then
            Control.Enabled = (Me.vfgFile.Rows > Me.vfgFile.FixedRows)
        Else
            Control.Enabled = (Me.vfgEPRs.Rows > Me.vfgEPRs.FixedRows)
        End If
    Case conMenu_View_ToolBar_Button: Control.Checked = Me.cbsThis(2).Visible
    Case conMenu_View_ToolBar_Text:   Control.Checked = Not (Me.cbsThis(2).Controls(1).Style = xtpButtonIcon)
    Case conMenu_View_ToolBar_Size:   Control.Checked = Me.cbsThis.Options.LargeIcons
    Case conMenu_View_StatusBar: Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case conPane_File
        Item.Handle = Me.vfgFile.hWnd
    Case conPane_EPRs
        Item.Handle = Me.vfgEPRs.hWnd
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
    '���ôʾ���ʾͣ������
    If mfrmWord Is Nothing Then Set mfrmWord = New frmDockEPRContent
    
    Dim panThis As Pane, panChild As Pane
    Set panThis = dkpMan.CreatePane(conPane_File, 400, 100, DockLeftOf, Nothing)
    panThis.Title = "�����ļ��嵥"
    panThis.Options = PaneNoCaption
    
    Set panChild = dkpMan.CreatePane(conPane_EPRs, 400, 300, DockBottomOf, panThis)
    panChild.Title = "������д��¼"
    panChild.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_Word, 600, 400, DockRightOf, Nothing)
    panThis.Title = "��������"
    panThis.Options = PaneNoCloseable Or PaneNoFloatable 'Or PaneNoHideable

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

Private Sub vfgEPRs_GotFocus()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub vfgEPRs_RowColChange()
    Dim lngRecordId As Long
    
    Err = 0: On Error Resume Next
    With Me.vfgEPRs
        lngRecordId = Val(.TextMatrix(.Row, 0))
    End With
    Err = 0: On Error GoTo 0
    If Me.Tag <> "" Then Exit Sub
    Call mfrmWord.zlRefresh(lngRecordId, "", True)
End Sub

Private Sub vfgFile_GotFocus()
    Me.cbsThis.RecalcLayout
End Sub

Private Sub vfgFile_RowColChange()
    Dim lngFileID As Long       '�����ļ�id
    
    If Me.Tag <> "" Then Exit Sub
    lngFileID = Val(Me.vfgFile.TextMatrix(Me.vfgFile.Row, 0))
    
    Select Case mintKind
    Case 1
        strSQL = "Select W.ID, P.����id, P.�����, P.����, P.�Ա�, To_Char(P.ִ��ʱ��, 'mm-dd hh24:mi') As ��������, W.������ As ��д��," & vbNewLine & _
                "       To_Char(W.���ʱ��, 'mm-dd hh24:mi') As ���ʱ��" & vbNewLine & _
                "From ���Ӳ�����¼ W, ���˹Һż�¼ P" & vbNewLine & _
                "Where W.��ҳid = P.ID And W.�������� = 1 And W.����id + 0 = [1] And W.�ļ�id + 0 = [4] And" & vbNewLine & _
                "      W.����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By P.ִ��ʱ��"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo, lngFileID)
    Case 2
        strSQL = "Select W.ID, I.����id, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'mm-dd hh24:mi') As ��Ժ����, W.������ As ��д��," & vbNewLine & _
                "       To_Char(W.���ʱ��, 'mm-dd hh24:mi') As ���ʱ��" & vbNewLine & _
                "From ���Ӳ�����¼ W, ������ҳ P, ������Ϣ I" & vbNewLine & _
                "Where I.����id = P.����id And P.����id = W.����id And P.��ҳid = W.��ҳid And W.�������� = 2 And W.����id + 0 = [1] And" & vbNewLine & _
                "      W.�ļ�id + 0 = [4] And W.����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "      To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By ��Ժ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo, lngFileID)
    Case 4
        strSQL = "Select W.ID, I.����id, P.סԺ��, I.����, I.�Ա�, To_Char(P.��Ժ����, 'mm-dd hh24:mi') As ��Ժ����, W.������ As ��д��," & vbNewLine & _
                "       To_Char(W.���ʱ��, 'mm-dd hh24:mi') As ���ʱ��" & vbNewLine & _
                "From ���Ӳ�����¼ W, ������ҳ P, ������Ϣ I" & vbNewLine & _
                "Where I.����id = P.����id And P.����id = W.����id And P.��ҳid = W.��ҳid And W.�������� = 4 And W.����id + 0 = [1] And" & vbNewLine & _
                "      W.�ļ�id + 0 = [4] And W.����ʱ�� Between To_Date([2], 'yyyy-mm-dd') And" & vbNewLine & _
                "      To_Date([3], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "Order By ��Ժ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDeptId, mstrDateFrom, mstrDateTo, lngFileID)
    End Select
    
    With Me.vfgEPRs
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 0: .ColHidden(0) = True
        For lngCount = 1 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
    End With
    Call vfgEPRs_RowColChange
End Sub
