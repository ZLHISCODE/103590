VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Begin VB.Form frmEPRAuditOutline 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   10635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picPane 
      BorderStyle     =   0  'None
      Height          =   3810
      Index           =   2
      Left            =   525
      ScaleHeight     =   3810
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   1140
      Width           =   4500
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   2985
         Left            =   0
         TabIndex        =   1
         Top             =   270
         Width           =   3405
         _cx             =   6006
         _cy             =   5265
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   2
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
         WallPaper       =   "frmEPRAuditOutline.frx":0000
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmEPRAuditOutline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmMain As Object
Private mintKind As Integer     '��������
Private mstrDateFrom As String  '��ʼ����
Private mstrDateTo As String    '��������
Private mlngMoual As Long
Private mblnShowAll As Boolean

'######################################################################################################################

Public Function zlInitData(ByVal frmMain As Object, ByVal lngMoual As Long) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    mlngMoual = lngMoual
    Set mfrmMain = frmMain
    
    If ExecuteCommand("��ʼ�ؼ�") = False Or ExecuteCommand("��ʼ����") = False Then Exit Function
    
End Function


Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
        
    Select Case Control.ID
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Preview

        Call RptPrint(2)
    
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Print

        Call RptPrint(1)
        
    '------------------------------------------------------------------------------------------------------------------
    Case conMenu_File_Excel

        Call RptPrint(3)
        
    End Select
    
End Sub


Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    With vfgThis
        Select Case Control.ID
        '--------------------------------------------------------------------------------------------------------------
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel               'Ԥ��,��ӡ,�����Excel
        
            Control.Enabled = (.Rows > .FixedRows + 1)
        
        End Select
        
    End With
    
End Sub

Public Function zlRefreshData(ByVal intKind As Integer, ByVal strDateFrom As String, ByVal strDateTo As String) As Boolean
    '******************************************************************************************************************
    '����:������鷶Χ��֯��ʾ�������
    '******************************************************************************************************************
    Dim lngTotal As Long
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    mstrDateFrom = strDateFrom
    mstrDateTo = strDateTo
    mintKind = intKind
    On Error GoTo errHand
    
    Select Case intKind
    '------------------------------------------------------------------------------------------------------------------
    Case 1  '���ﲡ��
        
        strSQL = "Select D.ID, D.����, D.����, W.�����, W.����д, P.�����˴�, W.�������, P.�����˴�, W.�������" & vbNewLine & _
                " From ���ű� D, ��������˵�� M," & vbNewLine & _
                "      (Select ִ�в���id, Sum(Decode(����, 1, 0, 1)) As �����˴�, Sum(Decode(����, 1, 1, 0)) As �����˴�" & vbNewLine & _
                "        From ���˹Һż�¼" & vbNewLine & _
                "        Where Nvl(ִ��״̬, 0) <> 0 And ��¼����=1 And ��¼״̬=1 And �Ǽ�ʱ�� Between To_Date([1], 'yyyy-mm-dd') And" & vbNewLine & _
                "              To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ִ�в���id) P," & vbNewLine & _
                "      (Select W.����id, Sum(W.�����) As �����, Sum(W.����д) As ����д," & vbNewLine & _
                "               Sum(Decode(F.�¼�, '����', W.�����, Null)) As �������," & vbNewLine & _
                "               Sum(Decode(F.�¼�, '����', W.�����, Null)) As �������" & vbNewLine & _
                "        From (Select F.ID, F.ͨ��, A.����id, Q.�¼�" & vbNewLine & _
                "               From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "               Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 1) F," & vbNewLine & _
                "             (Select ����id, �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����," & vbNewLine & _
                "                      Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "               From ���Ӳ�����¼" & vbNewLine & _
                "               Where �������� = 1 And ����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And" & vbNewLine & _
                "                     To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By ����id, �ļ�id) W" & vbNewLine & _
                "        Where F.ID = W.�ļ�id And (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = W.����id)" & vbNewLine & _
                "        Group By W.����id) W" & vbNewLine & _
                " Where D.ID = M.����id And M.�������� = '�ٴ�' And M.������� In (1, 3) And D.ID = P.ִ�в���id(+) And ( TO_CHAR (D.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or D.����ʱ�� is null) And" & vbNewLine & _
                "       D.ID = W.����id(+)" & vbNewLine & _
                " Order By D.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDateFrom, strDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "����": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "������д���": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "����": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "����": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            
            .TextMatrix(1, 1) = "����": .TextMatrix(1, 2) = "����"
            .TextMatrix(1, 3) = "�����": .TextMatrix(1, 4) = "����д"
            .TextMatrix(1, 5) = "�˴�": .TextMatrix(1, 6) = "��ɲ���"
            .TextMatrix(1, 7) = "�˴�": .TextMatrix(1, 8) = "��ɲ���"
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case 2  'סԺ����
        
        
        strSQL = "Select D.ID, D.����, D.����, W.�����, W.����д, I.��Ժ�˴�, W.��Ժ����, E.ת���˴�, W.ת�벡��, O.��Ժ�˴�," & vbNewLine & _
                "        W.��Ժ����, O.�����˴�, W.��������, G.ת���˴�, W.ת������, S.�����˴�, W.��������" & vbNewLine & _
                " From ���ű� D, ��������˵�� M," & vbNewLine & _
                "      (Select W.����id, Sum(�����) As �����, Sum(����д) As ����д," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '��Ժ', 1, '�״���Ժ', 1, '�ٴ���Ժ', 1, 0), 0) * �����) As ��Ժ����," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, 'ת��', Decode(Sign(F.��дʱ��), -1, 0, 1), 0), 0) * �����) As ת�벡��," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '��Ժ', 1, '24Сʱ��Ժ', 1, 0), 0) * �����) As ��Ժ����," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '����', 1, '24Сʱ����', 1, 0), 0) * �����) As ��������," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, 'ת��', Decode(Sign(F.��дʱ��), -1, 1, 0), 0), 0) * �����) As ת������," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '����', 1, 0), 0) * �����) As ��������" & vbNewLine & _
                "        From (Select F.ID, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "               From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "               Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 2) F," & vbNewLine & _
                "             (Select ����id, �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����," & vbNewLine & _
                "                      Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "               From ���Ӳ�����¼" & vbNewLine & _
                "               Where �������� = 2 And ����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By ����id, �ļ�id) W" & vbNewLine & _
                "        Where F.ID = W.�ļ�id And (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = W.����id)" & vbNewLine & _
                "        Group By W.����id) W," & vbNewLine
        strSQL = strSQL & "      (Select ��Ժ����id, Count(*) As ��Ժ�˴�" & vbNewLine & _
                "        From ������ҳ" & vbNewLine & _
                "        Where ��Ժ���� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ��Ժ����id) I," & vbNewLine & _
                "      (Select ����id, Count(*) As ת���˴�" & vbNewLine & _
                "        From ���˱䶯��¼" & vbNewLine & _
                "        Where ��ʼԭ�� = 3 And ��ʼʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ����id) E," & vbNewLine & _
                "      (Select ��Ժ����id, Sum(Decode(��Ժ��ʽ, '����', 0, 1)) As ��Ժ�˴�, Sum(Decode(��Ժ��ʽ, '����', 1, 0)) As �����˴�" & vbNewLine & _
                "        From ������ҳ" & vbNewLine & _
                "        Where ��Ժ���� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ��Ժ����id) O," & vbNewLine & _
                "      (Select ����id, Count(*) As ת���˴�" & vbNewLine & _
                "        From ���˱䶯��¼" & vbNewLine & _
                "        Where ��ֹԭ�� = 3 And ��ֹʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ����id) G," & vbNewLine & _
                "      (Select R.ִ�п���id, Count(*) As �����˴�" & vbNewLine & _
                "        From ����ҽ����¼ R" & vbNewLine & _
                "        Where R.������� = 'F' And R.����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By R.ִ�п���id) S" & vbNewLine & _
                " Where D.ID = M.����id And M.�������� = '�ٴ�' And ������� In (2, 3) And D.ID = W.����id(+) And D.ID = I.��Ժ����id(+) And" & vbNewLine & _
                "       D.ID = E.����id(+) And D.ID = O.��Ժ����id(+) And D.ID = G.����id(+) And D.ID = S.ִ�п���id(+) And ( TO_CHAR (D.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or D.����ʱ�� is null)" & vbNewLine & _
                " Order By D.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDateFrom, strDateTo)
        
        With vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "����": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "������д���": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "��Ժ": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "ת��": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            .TextMatrix(0, 9) = "��Ժ": .TextMatrix(0, 10) = .TextMatrix(0, 9)
            .TextMatrix(0, 11) = "����": .TextMatrix(0, 12) = .TextMatrix(0, 11)
            .TextMatrix(0, 13) = "ת��": .TextMatrix(0, 14) = .TextMatrix(0, 13)
            .TextMatrix(0, 15) = "����": .TextMatrix(0, 16) = .TextMatrix(0, 15)
            
            .TextMatrix(1, 1) = "����": .TextMatrix(1, 2) = "����"
            .TextMatrix(1, 3) = "�����": .TextMatrix(1, 4) = "����д"
            .TextMatrix(1, 5) = "�˴�": .TextMatrix(1, 6) = "��ɲ���"
            .TextMatrix(1, 7) = "�˴�": .TextMatrix(1, 8) = "��ɲ���"
            .TextMatrix(1, 9) = "�˴�": .TextMatrix(1, 10) = "��ɲ���"
            .TextMatrix(1, 11) = "�˴�": .TextMatrix(1, 12) = "��ɲ���"
            .TextMatrix(1, 13) = "�˴�": .TextMatrix(1, 14) = "��ɲ���"
            .TextMatrix(1, 15) = "�˴�": .TextMatrix(1, 16) = "��ɲ���"
        End With
    '------------------------------------------------------------------------------------------------------------------
    Case 4  '������
        strSQL = "Select D.ID, D.����, D.����, W.�����, W.����д, I.��Ժ�˴�, W.��Ժ����, E.ת���˴�, W.ת�벡��, O.��Ժ�˴�," & vbNewLine & _
                "        W.��Ժ����, O.�����˴�, W.��������, G.ת���˴�, W.ת������" & vbNewLine & _
                " From ���ű� D, ��������˵�� M," & vbNewLine & _
                "      (Select W.����id, Sum(�����) As �����, Sum(����д) As ����д," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '��Ժ', 1, '�״���Ժ', 1, '�ٴ���Ժ', 1, 0), 0) * �����) As ��Ժ����," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, 'ת��', Decode(Sign(F.��дʱ��), -1, 0, 1), 0), 0) * �����) As ת�벡��," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '��Ժ', 1, '24Сʱ��Ժ', 1, 0), 0) * �����) As ��Ժ����," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, '����', 1, '24Сʱ����', 1, 0), 0) * �����) As ��������," & vbNewLine & _
                "               Sum(Decode(F.Ψһ, 1, Decode(F.�¼�, 'ת��', Decode(Sign(F.��дʱ��), -1, 1, 0), 0), 0) * �����) As ת������" & vbNewLine & _
                "        From (Select F.ID, F.ͨ��, A.����id, Q.�¼�, Q.Ψһ, Q.��дʱ��" & vbNewLine & _
                "               From �����ļ��б� F, ����Ӧ�ÿ��� A, ����ʱ��Ҫ�� Q" & vbNewLine & _
                "               Where F.ID = A.�ļ�id(+) And F.ID = Q.�ļ�id And F.���� = 4) F," & vbNewLine & _
                "             (Select ����id, �ļ�id, Sum(Decode(���ʱ��, Null, 0, 1)) As �����," & vbNewLine & _
                "                      Sum(Decode(���ʱ��, Null, 1, 0)) As ����д" & vbNewLine & _
                "               From ���Ӳ�����¼" & vbNewLine & _
                "               Where �������� = 4 And ����ʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "               Group By ����id, �ļ�id) W" & vbNewLine & _
                "        Where F.ID = W.�ļ�id And (F.ͨ�� = 1 Or F.ͨ�� = 2 And F.����id = W.����id)" & vbNewLine & _
                "        Group By W.����id) W," & vbNewLine
        strSQL = strSQL & "      (Select ��Ժ����id, Count(*) As ��Ժ�˴�" & vbNewLine & _
                "        From ������ҳ" & vbNewLine & _
                "        Where ��Ժ���� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ��Ժ����id) I," & vbNewLine & _
                "      (Select ����id, Count(*) As ת���˴�" & vbNewLine & _
                "        From ���˱䶯��¼" & vbNewLine & _
                "        Where ��ʼԭ�� = 3 And ��ʼʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ����id) E," & vbNewLine & _
                "      (Select ��ǰ����id, Sum(Decode(��Ժ��ʽ, '����', 0, 1)) As ��Ժ�˴�, Sum(Decode(��Ժ��ʽ, '����', 1, 0)) As �����˴�" & vbNewLine & _
                "        From ������ҳ" & vbNewLine & _
                "        Where ��Ժ���� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ��ǰ����id) O," & vbNewLine & _
                "      (Select ����id, Count(*) As ת���˴�" & vbNewLine & _
                "        From ���˱䶯��¼" & vbNewLine & _
                "        Where ��ֹԭ�� = 3 And ��ֹʱ�� Between To_Date([1], 'yyyy-mm-dd') And To_Date([2], 'yyyy-mm-dd') + 1 - 1 / 86400" & vbNewLine & _
                "        Group By ����id) G" & vbNewLine & _
                " Where D.ID = M.����id And M.�������� = '����' And ������� In (2, 3) And D.ID = W.����id(+) And D.ID = I.��Ժ����id(+) And" & vbNewLine & _
                "       D.ID = E.����id(+) And D.ID = O.��ǰ����id(+) And D.ID = G.����id(+) And ( TO_CHAR (D.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' or D.����ʱ�� is null)" & vbNewLine & _
                " Order By D.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strDateFrom, strDateTo)
        
        With Me.vfgThis
            .Clear
            Set .DataSource = rsTemp
            
            .ColWidth(0) = 0: .ColHidden(0) = True
            .MergeCells = flexMergeFree
            .MergeRow(0) = True
            .TextMatrix(0, 1) = "����": .TextMatrix(0, 2) = .TextMatrix(0, 1)
            .TextMatrix(0, 3) = "������д���": .TextMatrix(0, 4) = .TextMatrix(0, 3)
            .TextMatrix(0, 5) = "��Ժ": .TextMatrix(0, 6) = .TextMatrix(0, 5)
            .TextMatrix(0, 7) = "ת��": .TextMatrix(0, 8) = .TextMatrix(0, 7)
            .TextMatrix(0, 9) = "��Ժ": .TextMatrix(0, 10) = .TextMatrix(0, 9)
            .TextMatrix(0, 11) = "����": .TextMatrix(0, 12) = .TextMatrix(0, 11)
            .TextMatrix(0, 13) = "ת��": .TextMatrix(0, 14) = .TextMatrix(0, 13)
            
            .TextMatrix(1, 1) = "����": .TextMatrix(1, 2) = "����"
            .TextMatrix(1, 3) = "�����": .TextMatrix(1, 4) = "����д"
            .TextMatrix(1, 5) = "�˴�": .TextMatrix(1, 6) = "��ɲ���"
            .TextMatrix(1, 7) = "�˴�": .TextMatrix(1, 8) = "��ɲ���"
            .TextMatrix(1, 9) = "�˴�": .TextMatrix(1, 10) = "��ɲ���"
            .TextMatrix(1, 11) = "�˴�": .TextMatrix(1, 12) = "��ɲ���"
            .TextMatrix(1, 13) = "�˴�": .TextMatrix(1, 14) = "��ɲ���"
        End With
    End Select
    
    
    '��ϼ�
    '------------------------------------------------------------------------------------------------------------------
    Dim lngCount As Long
    Dim lngCol As Long
    Dim lngRow As Long
    Dim blnData As Boolean
    
    With Me.vfgThis
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 2) = "�ϼ�"
        For lngCol = 3 To .Cols - 1
            lngTotal = 0
            For lngRow = .FixedRows To .Rows - 2
                lngTotal = lngTotal + Val(.TextMatrix(lngRow, lngCol))
            Next
            .TextMatrix(.Rows - 1, lngCol) = lngTotal
        Next
        .Row = .FixedRows: .Col = 1
        Call .AutoSize(1, .Cols - 1)
        
        .Redraw = False
        If mblnShowAll Then
            For lngRow = .FixedRows To .Rows - 2
                .RowHeight(lngRow) = .RowHeightMin
                .RowHidden(lngRow) = False
            Next
        Else
            For lngRow = .FixedRows To .Rows - 2
                blnData = False
                For lngCol = 3 To .Cols - 1
                    If Val(.TextMatrix(lngRow, lngCol)) <> 0 Then blnData = True: Exit For
                Next
                If blnData = False Then
                    .RowHeight(lngRow) = 0
                    .RowHidden(lngRow) = True
                End If
            Next
        End If
        .Redraw = True
    End With
    
    '��ʾ�����ؿ���
'    Call chkNoData_Click
'    Me.stbThis.Panels(2).Text = "�����չ��(Ctrl+O)����ϸ��鵱ǰ���Ҳ��˲����������������д�����"
    
    If Me.Visible Then Me.vfgThis.SetFocus
    Exit Function
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function InitCommandBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim objExtendedBar As CommandBar

    '------------------------------------------------------------------------------------------------------------------
    '��ʼ����
    Call CommandBarInit(cbsThis)
    Set cbsThis.Icons = frmPubResource.imgApp.Icons
    cbsThis.Options.LargeIcons = False
    
    '------------------------------------------------------------------------------------------------------------------
    '�˵�����:�����������ݣ����xtpControlPopup���͵�����ID���¸�ֵ

    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsThis.ActiveMenuBar.Visible = False
    
            
    '���Ź�����
    '------------------------------------------------------------------------------------------------------------------
    Set objBar = cbsThis.Add("��׼", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagStretched
    
    Set objControl = NewToolBar(objBar, xtpControlButton, 10, "��ʾ��ҵ�����", , , xtpButtonIconAndCaption)
    objControl.Checked = True
        
End Function

Private Sub RptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    '******************************************************************************************************************
    Dim objPrint As New zlPrint1Grd, objAppRow As zlTabAppRow

    Set objPrint.Body = vfgThis
    
    Select Case mintKind
    Case 1
        objPrint.Title.Text = "���ﲡ��(" & mstrDateFrom & "��" & mstrDateTo & ")��д���"
    Case 2
        objPrint.Title.Text = "סԺ����(" & mstrDateFrom & "��" & mstrDateTo & ")��д���"
    Case 4
        objPrint.Title.Text = "������(" & mstrDateFrom & "��" & mstrDateTo & ")��д���"
    End Select
    
    Set objPrint.Title.Font = vfgThis.Font

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objPrint.UnderAppRows.Add(objAppRow)

    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)

    Me.vfgThis.Tag = "Printing"
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Me.vfgThis.Tag = ""
End Sub

Private Function ExecuteCommand(strCommand As String, ParamArray varParam() As Variant) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Integer
    Dim intRow As Integer
    Dim rs As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim strTmp As String
    Dim strSQL As String
    Dim strNow As String
    Dim strNote As String
    
    On Error GoTo errHand

    Call SQLRecord(rsSQL)

    Select Case strCommand
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ�ؼ�"
                
        Call InitCommandBar
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��ʼ����"
                        
        
        If Val(zlDatabase.GetPara("��ʾ��ҵ�����", glngSys, mlngMoual, "1")) = 1 Then
            mblnShowAll = True
        Else
            mblnShowAll = False
        End If


    '------------------------------------------------------------------------------------------------------------------
    Case "ˢ������"
        
        

    End Select

    ExecuteCommand = True

    GoTo endHand
    
    '------------------------------------------------------------------------------------------------------------------
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    
endHand:

End Function

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 10
        
        mblnShowAll = Not mblnShowAll
        Control.Checked = mblnShowAll
        
    
        Dim blnData As Boolean
        Dim lngRow As Long
        Dim lngCol As Long
        
        
        With vfgThis
            .Redraw = False
            If mblnShowAll Then
                For lngRow = .FixedRows To .Rows - 2
                    .RowHeight(lngRow) = .RowHeightMin
                    .RowHidden(lngRow) = False
                Next
            Else
                For lngRow = .FixedRows To .Rows - 2
                    blnData = False
                    For lngCol = 3 To .Cols - 1
                        If Val(.TextMatrix(lngRow, lngCol)) <> 0 Then blnData = True: Exit For
                    Next
                    If blnData = False Then
                        .RowHeight(lngRow) = 0
                        .RowHidden(lngRow) = True
                    End If
                Next
            End If
            .Redraw = True
        End With


    End Select
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long
    Dim lngScaleTop  As Long
    Dim lngScaleRight  As Long
    Dim lngScaleBottom  As Long
    
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    picPane(2).Move lngScaleLeft, lngScaleTop, lngScaleRight - lngScaleLeft, lngScaleBottom - lngScaleTop
    
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case 10
        
        Control.IconId = IIf(mblnShowAll, 12, 10)
        
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Call SetPara("��ʾ��ҵ�����", IIf(mblnShowAll, 1, 0), mlngMoual)
    
End Sub

Private Sub picPane_Resize(Index As Integer)
    On Error Resume Next
    
    Select Case Index
    Case 2
        vfgThis.Move 0, 0, picPane(Index).Width, picPane(Index).Height
    End Select
End Sub
