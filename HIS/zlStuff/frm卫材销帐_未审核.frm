VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frm��������_δ��� 
   BorderStyle     =   0  'None
   Caption         =   "��������δ���"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   12615
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2490
      Width           =   12615
   End
   Begin VB.PictureBox picBatHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   -330
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   12735
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4470
      Width           =   12735
   End
   Begin VB.CheckBox chkȫѡ 
      Caption         =   "ȫ��(&A)"
      Height          =   225
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   1500
   End
   Begin VSFlex8Ctl.VSFlexGrid vsHeadGrid 
      Height          =   2055
      Left            =   75
      TabIndex        =   1
      Tag             =   "������"
      Top             =   435
      Width           =   10800
      _cx             =   19050
      _cy             =   3625
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm��������_δ���.frx":0000
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
      ExplorerBar     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsDetail 
      Height          =   1935
      Left            =   15
      TabIndex        =   2
      Tag             =   "��ϸ"
      Top             =   2535
      Width           =   10800
      _cx             =   19050
      _cy             =   3413
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm��������_δ���.frx":00DD
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
      ExplorerBar     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsBatch 
      Height          =   1815
      Left            =   15
      TabIndex        =   3
      Tag             =   "��ϸ"
      Top             =   4575
      Width           =   10800
      _cx             =   19050
      _cy             =   3201
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm��������_δ���.frx":0242
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
      ExplorerBar     =   1
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
Attribute VB_Name = "frm��������_δ���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '��������
Private mrsDetail As ADODB.Recordset
Private mrsBatch As ADODB.Recordset
Private mint��˱�־ As Integer
Private mintUnit As Integer '0-ɢװ��λ,1-��װ��λ
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
'ҽ���ӿ�
Private gclsInsure As New clsInsure
Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Public Enum ҽԺҵ��
    support����Ԥ�� = 0
    
    support�����˷� = 1
    supportԤ���˸����ʻ� = 2
    
    support�շ��ʻ�ȫ�Է� = 4       '�����շѺ͹Һ��Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�ȫ�Էѣ�ָͳ�����Ϊ0�Ľ��򳬳��޼۵Ĵ�λ�Ѳ���
    support�շ��ʻ������Ը� = 5     '�����շѺ͹Һ��Ƿ��ø����ʻ�֧�������Ը����֡������Ը�����1-ͳ�������* ���
    
    support�����ʻ�ȫ�Է� = 6       'סԺ���������������Ƿ��ø����ʻ�֧��ȫ�ԷѲ��֡�
    support�����ʻ������Ը� = 7     'סԺ���������������Ƿ��ø����ʻ�֧�������Ը����֡�
    support�����ʻ����� = 8         'סԺ���������������Ƿ��ø����ʻ�֧�����޲��֡�
    
    support����ʹ�ø����ʻ� = 9     '����ʱ��ʹ�ø����ʻ�֧��
    supportδ�����Ժ = 10          '�����˻���δ�����ʱ��Ժ
    
    support���ﲿ�����ֽ� = 11      'ֻ��������ҽ����֧���˷Ѳ�ʹ�ñ�������Ҳ����˵�����ֽ�ʱ�ſ��ǲ�������񣬶��˻ص������ʻ���ҽ�������������˷ѡ�
    support��������ҽ����Ŀ = 12  '�ڽ���ʱ�����Ը��շ�ϸĿ�Ƿ�����ҽ����Ŀ���м��
    
    support������봫����ϸ = 13    '�����շѺ͹Һ��Ƿ���봫����ϸ
    
    support�����ϴ� = 14            'סԺ���ʷ�����ϸʵʱ����
    support���������ϴ� = 15        'סԺ�����˷�ʵʱ����

    support��Ժ���˽������� = 16    '�����Ժ���˽�������
    support������Ժ = 17            '���������˳�Ժ
    support����¼�������� = 18    '������Ժ���Ժʱ������¼�������
    support������ɺ��ϴ� = 19      'Ҫ���ϴ��ڼ��������ύ���ٽ���
    support��Ժ��������Ժ = 20    '���˽���ʱ���ѡ���Ժ���ʣ��ͼ������Ժ�ſ��Խ���
    
    support�Һ�ʹ�ø����ʻ� = 21    'ʹ��ҽ���Һ�ʱ�Ƿ�ʹ�ø����ʻ�����֧��

    support���������շ� = 22        '�����������֤�󣬿ɽ��ж���շѲ���
    support�����շ���ɺ���֤ = 23  '�������շ���ɣ��Ƿ��ٴε��������֤
    
    supportҽ���ϴ� = 24            'ҽ����������ʱ�Ƿ�ʵʱ����
    support�ֱҴ��� = 25            'ҽ�������Ƿ���ֱ�
    support��;������������ϴ����� = 26 '�ṩ�����ϴ��������ݵĽ��㹦��
    support��������ѽ��ʵļ��ʵ��� = 27 '�Ƿ�����������ʵ��ݣ�����õ����Ѿ�����
    
    support�����ݳ������� = 28
    support��Ժ��ʵ�ʽ��� = 29       '��Ժ�ӿ����Ƿ�Ҫ��ӿ��̽��н���
    support�����ֳ�����ϸ = 32    '�������סԺ���ʴ�����ÿ����ϸ���в��ֳ���
    support����������� = 33        'ҽ���Ƿ�֧������������ϣ���֧��ֻ�и������ʻ�ԭ����,�����ҽ�����㷽ʽ��Ϊ�ֽ�,֧�ֵ����ж�ÿһ�ֽ��㷽ʽ�Ƿ������˻�
    supportסԺ�������� = 34        'HISʼ����ΪסԺ֧�ֽ������ϣ������֧����ҽ���ӿ��ڲ��������ؼټ��ɣ����Ӹò�����Ϊ�����GetCapability�����������ֽ��㷽ʽ�Ƿ�֧��ȫ��
    support�������� = 35            '�Ƿ����������ʣ�����Ա����Ҫӵ�и������ʵ�Ȩ�ޡ��˲���ȱʡΪ�棬��֧�ֵĽӿ��赥������
    support����_ָ��סԺ���� = 36   '�Ƿ�֧��ָ��סԺ��������ҽ������
    support����_ָ�����ڷ�Χ = 37   '�Ƿ�֧��ָ���������ڷ�Χ����ҽ������
    support����_����Ӥ�������� = 38 '�Ƿ���������Ӥ��������
    
    support������� = 41            '�Ƿ�֧������ҽ�����˵ļ��ʷ���ʹ��������������
End Enum

Private mobjPlugIn As Object             '��ҽӿڶ���

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
 
Private Sub initPara()
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    With vsHeadGrid
      '  .Editable = flexEDKbdMouse
    End With
    
End Sub
Private Function RefreshData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���»�ȡ����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-03 21:09:18
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As ADODB.Recordset, strFields As String, strWere As String, lngRow As Long
    Dim str����ID As String
    Dim strNOS As String
    
    On Error GoTo ErrHandle
    Call InitRsStruct
    vsHeadGrid.Rows = 1
    mint��˱�־ = 1
        
    ''''1����ȡ��������
    '��λ����װ����
    Select Case mintUnit
    Case 0
        strFields = "X.���㵥λ ��λ,1 ����ϵ��,A.���� As �������� "
    Case Else
        strFields = "D.��װ��λ ��λ,d.����ϵ�� ����ϵ��,A.���� As �������� "
    End Select
    If CDate(mArrFilter("���ڷ�Χ")(0)) <= CDate("1949-02-01") Then
        strWere = strWere & " And A.����� Is Null And A.״̬ = 0  "
    Else
        strWere = strWere & " And A.����� Is Null And A.״̬ = 0 And A.����ʱ�� Between [3] And [4] "
    End If
    
    '����/ҽ������
    If Val(mArrFilter("�������ID")) > 0 Then strWere = strWere & " And A.���벿��id = [2] "
    '������
    If Trim(mArrFilter("������")) <> "" Then strWere = strWere & " And A.������=[7] "
    '��������
    If Trim(mArrFilter("��������")) <> "" Then strWere = strWere & " And nvl(F.����,B.����)=[8] "
    strWere = strWere & IIf(Val(mArrFilter("סԺ��")) = 0, "", "             AND b.��ʶ��=[9] and b.�����־=2 ")
    strWere = strWere & IIf(Val(mArrFilter("����ID")) = 0, "", "             AND b.����iD=[10]  ")
    strWere = strWere & IIf(Trim(mArrFilter("����")) = "", "", "             AND b.����=[11]  ")

    '˵��:
    '1.��������Ժ����: F.״̬:0-����סԺ��1-��δ��ƣ�2-����ת�ƣ�3-��Ԥ��Ժ
    '2.��Ҫ���˵�δ���ϲ��ݵ�����:ԭ�����ڼ��ʵ���������ʱ,�ǲ����Ժ��δ�����ֵ�����.
    gstrSQL = "" & _
    "   Select Distinct A.�շ�ϸĿid,'['||X.����||']'||X.���� as ��������, X.���,A.��׼���� as ���ۼ�, " & strFields & _
    "   From (  Select A.�շ�ϸĿid, Sum(A.����) As ����,B.��׼����  " & _
    "           From סԺ���ü�¼ B, ������ҳ F, ���˷������� A " & _
    "           Where A.�������=1 And A.����id = B.ID And A.��˲���id = [1] And B.����id = F.����id  " & _
    "                 And B.��ҳid = F.��ҳid  And F.��Ժ����  Is Null And F.״̬  <> 3 " & _
    "                 " & vbCrLf & strWere & _
    "                 And Exists (Select 1 From ҩƷ�շ���¼ C  Where C.����id = A.����id And C.����� Is Not Null And (C.��¼״̬ = 1 Or Mod(C.��¼״̬, 3) = 0))" & _
    "           Group By A.�շ�ϸĿid,b.��׼����) A,�������� D, �շ���Ŀ���� E, �շ���ĿĿ¼ X " & _
    " Where A.�շ�ϸĿid = D.����id And A.�շ�ϸĿid = X.ID And X.ID = E.�շ�ϸĿid(+) And E.����(+) = 3 " & _
    " Order By ��������"
    
    '[5],[6]����ȡ��
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", _
        Val(mArrFilter("���ϲ���id")), Val(mArrFilter("�������id")), _
        CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
        CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
        Trim(mArrFilter("������")), Trim(mArrFilter("��������")), _
        Val(mArrFilter("סԺ��")), Val(mArrFilter("����ID")), Trim(mArrFilter("����")))
    
    With vsHeadGrid
        .Clear 1
        .Rows = 2
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 0
        Do While Not rsTemp.EOF
            lngRow = lngRow + 1
            .TextMatrix(lngRow, .ColIndex("���")) = "��"
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(rsTemp!��������)
            .TextMatrix(lngRow, .ColIndex("���")) = NVL(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(Val(NVL(rsTemp!��������)) / rsTemp!����ϵ��, mFMT.FM_����)
            .TextMatrix(lngRow, .ColIndex("���ʽ��")) = Format(Val(NVL(rsTemp!��������)) * rsTemp!���ۼ�, mFMT.FM_���)
            .TextMatrix(lngRow, .ColIndex("��λ")) = NVL(rsTemp!��λ)
            .Cell(flexcpData, lngRow, .ColIndex("��������")) = NVL(rsTemp!�շ�ϸĿid)
            .Cell(flexcpData, lngRow, .ColIndex("��������")) = NVL(rsTemp!��������)
            rsTemp.MoveNext
        Loop

    End With
    
    
    ''''2����ȡ��ϸ����
    '��λ�ִ�
    Select Case mintUnit
    Case 0
        strFields = "X.���㵥λ ��λ,1 ����ϵ��, A.���� "
    Case Else
        strFields = "D.��װ��λ ��λ,D.����ϵ��, A.���� "
    End Select
    
    gstrSQL = "" & _
        "   Select ����, NO, ҩƷID as ����ID, ����ʱ��, ��ʶ��, ����, ����, ��λ, ����ϵ��,��������, Sum(����) As ��������,���ۼ� " & _
        "   From (  Select Distinct C.����, C.NO, C.ҩƷID, A.����ʱ��, B.��ʶ��,B.��׼���� as ���ۼ�, nvl(F.����,B.����) ����, B.����,P.���� ��������, " & strFields & " " & _
        "           From ���˷������� A, סԺ���ü�¼ B,ҩƷ�շ���¼ C, �������� D, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E " & _
        "           Where   A.�������=1 And A.����id = B.ID And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.����id And B.�շ�ϸĿid = X.ID  " & _
        "                   And  B.����id = F.����id And B.��ҳid = F.��ҳid And F.��Ժ���� Is Null And F.״̬ <> 3 " & _
        "                   And A.���벿��id = E.ID And B.ִ�в���id = [1]  " & _
        "                   And C.����� Is Not Null And C.���� In (24, 25, 26) And (C.��¼״̬ = 1 Or Mod(C.��¼״̬, 3) = 0) " & strWere & ")" & _
        "           Group By ����, NO, ҩƷID, ����ʱ��, ��ʶ��, ����, ����, ��λ, ����ϵ��,���ۼ�,�������� "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", _
          Val(mArrFilter("���ϲ���id")), Val(mArrFilter("�������id")), _
          CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
          CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
          Trim(mArrFilter("������")), Trim(mArrFilter("��������")), _
          Val(mArrFilter("סԺ��")), Val(mArrFilter("����ID")), Trim(mArrFilter("����")))
      
    Do While Not rsTemp.EOF
        With mrsDetail
            .AddNew
    
            !���� = rsTemp!����
            !NO = rsTemp!NO
            !����ID = rsTemp!����ID
            !����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            !��ʶ�� = rsTemp!��ʶ��
            !���� = rsTemp!����
            !���� = rsTemp!����
            !�������� = rsTemp!��������
            !���ʽ�� = rsTemp!�������� * rsTemp!���ۼ�
            !����ϵ�� = rsTemp!����ϵ��
            !��λ = rsTemp!��λ
            !�������� = rsTemp!��������
            .Update
            
            If InStr(1, strNOS, rsTemp!NO) = 0 Then
                strNOS = IIf(strNOS = "", "", strNOS & ",") & rsTemp!NO
            End If
            rsTemp.MoveNext
        End With
    Loop
     
    ''''3����ȡ������ϸ����
    '��λ����װ����
    Select Case mintUnit
    Case 0
        strFields = "X.���㵥λ ��λ,1 ����ϵ��,C.ʵ������ As ׼������,A.���� As ��������"
    Case Else
        strFields = "D.��װ��λ ��λ,D.����ϵ�� ,C.ʵ������ As ׼������,A.���� As ��������"
    End Select
    
    ' 'Having Sum(ʵ������) > 0
    gstrSQL = "Select /*+ Rule*/ C.ID As �շ�ID, C.ҩƷID, C.����, C.NO, C.��� As �շ����, C.����, C.����, C.Ч��, F.����, P.���� As ��������,B.��׼���� as ���ۼ�, " & _
        " A.����id, B.��� As �������, B.��¼����, B.��ҳID, A.����ʱ��, C.���ۼ� As ����, " & strFields & " " & _
        " From ���˷������� A, סԺ���ü�¼ B, " & _
        " (Select A.ID, A.����, A.NO, A.���, A.ҩƷid, A.����, A.����, A.Ч��, A.����id, B.ʵ������, A.���ۼ� " & _
        " From ҩƷ�շ���¼ A, " & _
        " (Select a.����, a.NO, a.���, a.ҩƷid, Sum(Nvl(a.����, 1) * a.ʵ������) As ʵ������ " & _
        " From ҩƷ�շ���¼ a ,Table(Cast(f_Str2list([12]) As zlTools.t_Strlist)) b " & _
        " Where a.���� In (24, 25,26) And a.������� Is Not Null And a.No=b.Column_Value "
        
    gstrSQL = gstrSQL & " Group By ����, NO, ���, ҩƷid " & _
        " ) B" & _
        " Where A.NO = B.NO And A.���� = B.���� And A.ҩƷid + 0 = B.ҩƷid And A.��� = B.��� And A.����� Is Not Null " & _
        " And (A.��¼״̬ = 1 Or Mod(A.��¼״̬, 3) = 0))C, " & _
        " �������� D, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E " & _
        " Where A.�������=1 And A.����id = B.ID And B.No = C.No And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.����id And B.�շ�ϸĿid = X.ID And B.����id = F.����id And B.��ҳid = F.��ҳid And A.���벿��id = E.ID " & _
        " And B.ִ�в���id = [1] " & strWere

    gstrSQL = gstrSQL & " Order By A.����ʱ��, C.����, C.NO, C.��� Desc "
    
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", _
          Val(mArrFilter("���ϲ���id")), Val(mArrFilter("�������id")), _
          CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
          CDate(mArrFilter("���ڷ�Χ")(0)), CDate(mArrFilter("���ڷ�Χ")(1)), _
          Trim(mArrFilter("������")), Trim(mArrFilter("��������")), _
          Val(mArrFilter("סԺ��")), Val(mArrFilter("����ID")), Trim(mArrFilter("����")), strNOS)
          
    Do While Not rsTemp.EOF
        With mrsBatch
            .AddNew
            !���� = rsTemp!����
            !NO = rsTemp!NO
            !����ID = rsTemp!ҩƷID
            !����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            !�շ���� = rsTemp!�շ����
            !���� = rsTemp!����
            !���� = rsTemp!����
            !Ч�� = rsTemp!Ч��
            !׼������ = rsTemp!׼������
            !�������� = rsTemp!��������
            !���ʽ�� = rsTemp!�������� * rsTemp!���ۼ�
            !����ϵ�� = rsTemp!����ϵ��
            !��λ = rsTemp!��λ
            !�շ�ID = rsTemp!�շ�ID
            !��ҳid = Val(NVL(rsTemp!��ҳid))
            !������� = rsTemp!�������
            !���� = rsTemp!����
            !����ID = rsTemp!����ID
            !��¼���� = rsTemp!��¼����
            !��˱�־ = 1
            .Update
            rsTemp.MoveNext
        End With
    Loop
    Call AutoExpendQuantity
    ''''''4����λ�����ܵ�һ�У�����ȡ��һ����ϸ����
    With vsHeadGrid
        If .Rows > 1 Then
            .Row = 1: .TopRow = 1
            Call LoadDetailList(Val(vsHeadGrid.TextMatrix(1, vsHeadGrid.ColIndex("��������"))))
        End If
    End With
    With vsDetail
        '��ȡ��һ��������ϸ����
        If .Rows > 1 Then
            Call LoadBatchList(Val(.TextMatrix(1, .ColIndex("����"))), .TextMatrix(1, .ColIndex("NO")), .RowData(1), .TextMatrix(1, .ColIndex("����ʱ��")), False)
        End If
    End With
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function InitRsStruct() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ���ڲ���¼��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-03 21:08:22
    '-----------------------------------------------------------------------------------------------------------
    Set mrsDetail = New ADODB.Recordset
    With mrsDetail
        If .State = 1 Then .Close
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ʶ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���ʽ��", adDouble, 18, adFldIsNullable
        .Fields.Append "����ϵ��", adDouble, 18, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set mrsBatch = New ADODB.Recordset
    With mrsBatch
        If .State = 1 Then .Close
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�շ����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "׼������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���ʽ��", adDouble, 18, adFldIsNullable
        .Fields.Append "����ϵ��", adDouble, 18, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҳID", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "��˱�־", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Function

Private Sub chkȫѡ_Click()
    Dim n As Integer
    With vsHeadGrid
        If .Rows > 1 Then
            If Val(.Cell(flexcpData, 1, .ColIndex("��������"))) = 0 Then Exit Sub
        End If
        For n = 1 To .Rows - 1
            If Val(.Cell(flexcpData, n, .ColIndex("��������"))) <> 0 Then
                .TextMatrix(n, .ColIndex("���")) = IIf(chkȫѡ.Value = 1, "��", "")
            End If
        Next
    End With
    With mrsBatch
        .Filter = 0
        If mrsBatch.RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            !��˱�־ = IIf(chkȫѡ.Value = 1, 1, 0)
            .Update
            .MoveNext
        Loop
    End With
End Sub
Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsHeadGrid, Me.Caption, "����δ��_Head"
    zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "����δ��_Detail"
    zl_vsGrid_Para_Restore mlngModule, vsBatch, Me.Caption, "����δ��_Batch"
    
    Call initPara
End Sub
Private Sub AutoExpendQuantity()
    '���ǵ�ͬһ����ID��Ӧ����շ�ID���������Ҫ�����������ֽ⵽����շ���¼��
    '�ֽ��ԭ���ǰ���Ŵ�����ȷ��䣨�Ѱ���Ž�������
    Dim n As Integer
    Dim dbl׼������ As Double
    Dim dblʣ������ As Double
    Dim int�շ���� As Integer
    Dim lng����ID As Long
    Dim str����ʱ�� As String
    
    With mrsBatch
        If mrsBatch.RecordCount <> 0 Then .MoveFirst
        For n = 1 To .RecordCount
            dbl׼������ = !׼������
            
            If lng����ID = !����ID And str����ʱ�� = !����ʱ�� Then

            Else
                dblʣ������ = !��������
            End If
            
            If dblʣ������ >= dbl׼������ Then
                dblʣ������ = dblʣ������ - dbl׼������
                !�������� = dbl׼������
            Else
                !�������� = dblʣ������
                dblʣ������ = 0
            End If
            
            lng����ID = !����ID
            str����ʱ�� = !����ʱ��
            
            .Update
            .MoveNext
        Next
    End With
End Sub
Private Sub LoadDetailList(ByVal lng����ID As Long)
    '-----------------------------------------------------------------------------------------------------------
    '����:������ϸ����
    '���:lng����ID-����ID
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-03 22:25:04
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    With vsDetail
        .Clear 1
        .Rows = 2
        mrsDetail.Filter = "����ID=" & lng����ID
        If mrsDetail.RecordCount = 0 Then Exit Sub
        .Rows = mrsDetail.RecordCount + 1
        lngRow = 0
        Do While Not mrsDetail.EOF
            lngRow = lngRow + 1
            .RowData(lngRow) = Val(NVL(mrsDetail!����ID))
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsDetail!����)
            .TextMatrix(lngRow, .ColIndex("NO")) = NVL(mrsDetail!NO)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(mrsDetail!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(lngRow, .ColIndex("����(סԺ)��")) = NVL(mrsDetail!��ʶ��)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsDetail!����)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsDetail!����)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(Val(NVL(mrsDetail!��������)) / mrsDetail!����ϵ��, mFMT.FM_����)
            .TextMatrix(lngRow, .ColIndex("���ʽ��")) = Format(Val(NVL(mrsDetail!���ʽ��)), mFMT.FM_���)
            .TextMatrix(lngRow, .ColIndex("��λ")) = NVL(mrsDetail!��λ)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsDetail!��������)
            
            .Cell(flexcpData, lngRow, .ColIndex("��������")) = Val(NVL(mrsDetail!��������))
            .Cell(flexcpData, lngRow, .ColIndex("NO")) = Val(NVL(mrsDetail!����ϵ��))
            mrsDetail.MoveNext
        Loop
        .Cell(flexcpForeColor, 1, .ColIndex("��������"), .Rows - 1, .ColIndex("��������")) = vbBlue
    End With
End Sub
Private Sub LoadBatchList(ByVal int���� As Integer, _
                ByVal strNo As String, ByVal lng����ID As Long, _
                ByVal str����ʱ�� As String, ByVal bln���±�־ As Boolean)
                
    '-----------------------------------------------------------------------------------------------------------
    '����:����������Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-03 22:42:02
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    With vsBatch
        mrsBatch.Filter = "����=" & int���� & _
                " And No='" & strNo & "' " & _
                " And ����ID=" & lng����ID & _
                " And ����ʱ��='" & str����ʱ�� & "' "
        mrsBatch.Sort = "�շ���� Desc"
        .Clear 1
        .Rows = 2
        If mrsBatch.RecordCount = 0 Then Exit Sub
        
        If mrsBatch.RecordCount = 0 Then
            vsDetail.Height = Me.ScaleHeight - vsDetail.Top
            picBatHsc.Visible = False
        Else
            picBatHsc.Top = Me.ScaleHeight - 1935
            vsDetail.Height = picBatHsc.Top - vsDetail.Top
            vsBatch.Top = picBatHsc.Top + picBatHsc.Height
            vsBatch.Height = Me.ScaleHeight - Me.vsBatch.Top
            picBatHsc.Visible = True
        End If
            
        .Rows = mrsBatch.RecordCount + 1
        mrsBatch.MoveFirst
        lngRow = 0
        Do While Not mrsBatch.EOF
                lngRow = lngRow + 1
                .RowData(lngRow) = Val(NVL(mrsBatch!����ID))
                vsBatch.TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsBatch!����)
                vsBatch.TextMatrix(lngRow, .ColIndex("NO")) = NVL(mrsBatch!NO)
                vsBatch.TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(mrsBatch!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                vsBatch.TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsBatch!����)
                vsBatch.TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsBatch!����)
                vsBatch.TextMatrix(lngRow, .ColIndex("Ч��")) = Format(mrsBatch!Ч��, "yyyy-mm-dd")
                vsBatch.TextMatrix(lngRow, .ColIndex("׼������")) = Format(Val(NVL(mrsBatch!׼������)) / mrsBatch!����ϵ��, mFMT.FM_����)
                vsBatch.TextMatrix(lngRow, .ColIndex("��������")) = Format(Val(NVL(mrsBatch!��������)) / mrsBatch!����ϵ��, mFMT.FM_����)
                vsBatch.TextMatrix(lngRow, .ColIndex("���ʽ��")) = Format(Val(NVL(mrsBatch!���ʽ��)), mFMT.FM_���)
                vsBatch.TextMatrix(lngRow, .ColIndex("��λ")) = NVL(mrsBatch!��λ)
                
                vsBatch.Cell(flexcpData, lngRow, .ColIndex("׼������")) = Val(NVL(mrsBatch!׼������))
                vsBatch.Cell(flexcpData, lngRow, .ColIndex("��������")) = Val(NVL(mrsBatch!��������))
                vsBatch.Cell(flexcpData, lngRow, .ColIndex("NO")) = Val(NVL(mrsBatch!����ϵ��))
                If bln���±�־ Then
                    mrsBatch!��˱�־ = mint��˱�־
                    mrsBatch.Update
                End If
            mrsBatch.MoveNext
        Loop
        .Cell(flexcpForeColor, 1, .ColIndex("��������"), lngRow, .ColIndex("��������")) = vbBlue
    End With
End Sub

Private Sub Form_Resize()
    
    With vsHeadGrid
        .Left = ScaleLeft
        .Width = ScaleWidth - .Left
        .Height = IIf(picHsc.Top - Top < 500, 500, picHsc.Top - Top)
        picHsc.Top = .Top + .Height
    End With
    With vsDetail
        .Top = picHsc.Top + picHsc.Height
        .Width = vsHeadGrid.Width
        If picBatHsc.Visible Then
            .Height = IIf(picBatHsc.Top - .Top < 500, 500, picBatHsc.Top - .Top)
            picBatHsc.Top = .Top + .Height
        Else
            .Height = Me.ScaleHeight - .Top
        End If
    End With
    
    If picBatHsc.Visible = True Then
        With vsBatch
            .Top = picBatHsc.Top + picBatHsc.Height
            .Width = vsHeadGrid.Width
            .Height = IIf(Me.ScaleHeight - .Top < 0, 0, Me.ScaleHeight - .Top)
            .Left = ScaleLeft
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsHeadGrid, Me.Caption, "����δ��_Head"
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "����δ��_Detail"
    zl_vsGrid_Para_Save mlngModule, vsBatch, Me.Caption, "����δ��_Batch"
End Sub

Private Sub vsBatch_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBatch
        Select Case .Col
        Case .ColIndex("��������")
            If .TextMatrix(Row, .ColIndex("����")) = "" Then Cancel = True
        Case Else
            Cancel = True
        End Select
    End With
End Sub

 

Private Sub vsDetail_EnterCell()
    Dim lng����ID As Long
    With vsDetail
        If .Row > 0 Then
            lng����ID = IIf(IsNull(.RowData(.Row)), 0, .RowData(.Row))
            '��ȡ������ϸ����
            Call LoadBatchList(Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), lng����ID, .TextMatrix(.Row, .ColIndex("����ʱ��")), False)
        End If
    End With
End Sub

 
Private Sub vsHeadGrid_Click()
    Dim bln���±�־ As Boolean
    
    With vsHeadGrid
        If .Row > 0 Then
            If .Cell(flexcpData, .Row, .ColIndex("��������")) = 0 Then Exit Sub
        End If
        
        
        If .Row > 0 And .Col = .ColIndex("���") Then
            If .TextMatrix(.Row, .Col) = "��" Then
                .TextMatrix(.Row, .Col) = "��"
                mint��˱�־ = 2
            ElseIf .TextMatrix(.Row, .Col) = "��" Then
                .TextMatrix(.Row, .Col) = ""
                mint��˱�־ = 0
            Else
                .TextMatrix(.Row, .Col) = "��"
                mint��˱�־ = 1
            End If
            bln���±�־ = True
        End If
        
        If .Row > 0 Then
            '��ȡ��ϸ����
            Call LoadDetailList(Val(.Cell(flexcpData, .Row, .ColIndex("��������"))))
        End If
    End With
    
    '��ȡ������ϸ����
    With vsDetail
        Call LoadBatchList(Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), .RowData(.Row), .TextMatrix(.Row, .ColIndex("����ʱ��")), bln���±�־)
    End With
End Sub


Private Sub vsBatch_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    'ֻ����������
    With vsBatch
        If Col = .ColIndex("��������") Then
            If InStr("1234567890" + Chr(46) + Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0
            End If
        End If
    End With
End Sub

Private Sub vsBatch_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim dblKey As Double
    With vsBatch
        dblKey = Val(.EditText)

        If dblKey > Val(.TextMatrix(Row, .ColIndex("׼������"))) Or dblKey < 0 Then
            dblKey = Val(.TextMatrix(Row, .ColIndex("׼������")))
        End If
        .EditText = Format(dblKey, mFMT.FM_����)
        .TextMatrix(Row, .ColIndex("��������")) = Format(dblKey, mFMT.FM_����)

        mrsBatch.Filter = "����=" & Val(.TextMatrix(Row, .ColIndex("����"))) & _
                        " And No='" & .TextMatrix(Row, .ColIndex("NO")) & "' " & _
                        " And ����ID=" & .RowData(Row) & _
                        " And �շ����=" & Val(.TextMatrix(Row, .ColIndex("�շ����"))) & _
                        " And ����ʱ��='" & Val(.TextMatrix(Row, .ColIndex("����ʱ��"))) & "' "
        If mrsBatch.EOF Then Exit Sub
        mrsBatch!�������� = Val(.TextMatrix(Row, .ColIndex("��������"))) * mrsBatch!����ϵ��
        mrsBatch.Update
    End With
End Sub
Public Function zlRefreshData(ByVal frmMain As Form, ByVal strPrivs As String, ByVal lngModule As Long, ByVal intUnit As Integer, _
    ByVal arrFilter As Variant) As Boolean
     '-----------------------------------------------------------------------------------------------------------
    '����:����ˢ������
    '���:frmMain-������
    '     strPrivs-Ȩ�޴�
    '     lngModule-ģ���
    '     intUnit-��ʾ��λ(0-ɢװ��λ,1-��װ��λ)
    '     arrFilter-��������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-22 14:25:18
    '-----------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mstrPrivs = strPrivs: mlngModule = lngModule:
    Set mArrFilter = arrFilter
    mintUnit = intUnit
    
    '��ʼ��ֵ
    Call Form_Load
    zlRefreshData = RefreshData
End Function
Public Function zlVerifyData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-04 00:23:54
    '-----------------------------------------------------------------------------------------------------------
    Screen.MousePointer = 11
    zlVerifyData = SaveData()
    Screen.MousePointer = 0
End Function
Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-04 00:08:26
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strCurDate As String
    
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim int��˱�־ As Integer
    Dim bln�Ƿ������� As Boolean
    Dim str������� As String
    Dim cllPro As Collection
    Dim strAudit As String  '��¼�ѽ������ʵķ��ü�¼�������ظ�ִ��
    Dim strReturnInfo As String
    Dim strReserve As String
    
    If vsHeadGrid.Rows = 1 Then Exit Function
    If Val(vsHeadGrid.Cell(flexcpData, 1, vsHeadGrid.ColIndex("��������"))) = 0 Then Exit Function
    strCurDate = Format(Sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    Set cllPro = New Collection
    
    With mrsBatch
        .Filter = 0
        If .State = 0 Then Exit Function
        If .RecordCount = 0 Then Exit Function
        Do While Not .EOF
            If !��˱�־ <> 0 And InStr("," & strAudit & ",", "," & !����ID & !����ʱ�� & ",") = 0 Then
                strAudit = IIf(strAudit = "", !����ID & !����ʱ��, strAudit & "," & !����ID & !����ʱ��)
                
                'Zl_���˷�������_Audit
                gstrSQL = "Zl_���˷�������_Audit("
                '  Id_In       ���˷�������.����id%Type,
                gstrSQL = gstrSQL & "" & Val(NVL(!����ID)) & ","
                '  ����ʱ��_In ���˷�������.����ʱ��%Type,
                gstrSQL = gstrSQL & "To_Date('" & !����ʱ�� & "','YYYY-MM-DD HH24:MI:SS'),"
                '  �����_In   ���˷�������.�����%Type,
                gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                '  ���ʱ��_In ���˷�������.���ʱ��%Type,
                gstrSQL = gstrSQL & "To_Date('" & strCurDate & "','YYYY-MM-DD HH24:MI:SS'),"
                '  ״̬_In     ���˷�������.״̬%Type,
                gstrSQL = gstrSQL & "" & Val(NVL(!��˱�־)) & ","
                '  int�Զ����� Integer:=1
                gstrSQL = gstrSQL & "0)"
                AddArray cllPro, gstrSQL
            End If
            
            '���ϴ���
            If !��˱�־ = 1 And !�������� <> 0 Then
                    'Zl_�����շ���¼_��������
                    gstrSQL = "Zl_�����շ���¼_��������("
                    '    �շ�id_In   In ҩƷ�շ���¼.ID%Type,
                    gstrSQL = gstrSQL & "" & NVL(!�շ�ID) & ","
                    '    �����_In   In ҩƷ�շ���¼.�����%Type,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '    �������_In In ҩƷ�շ���¼.�������%Type,
                    gstrSQL = gstrSQL & "to_date('" & strCurDate & "','yyyy-mm-dd HH24:mi:ss'),"
                    '    ����_In     In ҩƷ���.�ϴ�����%Type := Null,
                    gstrSQL = gstrSQL & "'" & NVL(!����) & "',"
                    '    Ч��_In     In ҩƷ���.Ч��%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(IsNull(!Ч��), "NULL", IIf(NVL(!Ч��) = "", "NULL", "To_Date('" & Format(!Ч��, "yyyy-MM-dd") & "','yyyy-MM-dd')")) & ","
                    '    ����_In     In ҩƷ���.�ϴβ���%Type := Null,
                    gstrSQL = gstrSQL & "'" & NVL(!����) & "',"
                    '    ��������_In In ҩƷ�շ���¼.ʵ������%Type := Null,
                    gstrSQL = gstrSQL & "" & NVL(!��������) & ","
                    '    �Զ�����_In Integer := 0,
                    gstrSQL = gstrSQL & "" & 0 & ","
                    '    ������_In   In ҩƷ�շ���¼.������%Type := Null
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    
                    '    �Ƿ�����_In Integer := 1,
                    gstrSQL = gstrSQL & "" & 0 & ")"
                    AddArray cllPro, gstrSQL
                    bln�Ƿ������� = True
                    '���ʴ���
                    str������� = !������� & ":" & !��������
                    '--��ţ���ʽ��"1,3,5,7,8",��"1:2,3:2,5:2,7:2,8:2",ð��ǰ������ֱ�ʾ�к�,��������ֱ�ʾ�˵�����,Ŀǰ�����������ʱ��ҩƷ�Ŵ���
                    '--      Ϊ�ձ�ʾ�������пɳ�����

                    If !��ҳid = 0 Then
                        gstrSQL = "Zl_������ʼ�¼_Delete('" & !NO & "','" & !������� & "','" & gstrUserCode & "','" & gstrUserName & "')"
                    Else
                        gstrSQL = "ZL_סԺ���ʼ�¼_Delete('" & !NO & "','" & str������� & "','" & gstrUserCode & "','" & gstrUserName & "'," & !��¼���� & ",1)"
                    End If
                    AddArray cllPro, gstrSQL
                    
                    'ҽ������
                    If Not IsNull(!����) And InStr(1, strMCNO, !NO) = 0 Then
                        MCPAR.���������ϴ� = gclsInsure.GetCapability(support���������ϴ�, , Val(!����))
                        MCPAR.������ɺ��ϴ� = gclsInsure.GetCapability(support������ɺ��ϴ�, , Val(!����))
                        strMCNO = strMCNO & IIf(strMCNO = "", "", "|") & !NO & "," & !���� & _
                                "," & IIf(MCPAR.���������ϴ�, "1", "0") & "," & IIf(MCPAR.������ɺ��ϴ�, "1", "0")
                    End If
                    
                    strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & NVL(!�շ�ID) & "," & NVL(!��������)
            End If
            .MoveNext
        Loop
    End With
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllPro, Me.Caption, True
    'ҽ�������������ϴ�������ʱ�ϴ�
    If strMCNO <> "" Then
        arrMCRec = Split(strMCNO, "|")
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 0 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    gcnOracle.RollbackTrans:  Exit Function
                End If
            End If
        Next
    End If
                            
    gcnOracle.CommitTrans
    
    'ҽ�������������ϴ�����ɺ��ϴ�
    If strMCNO <> "" Then
        For i = 0 To UBound(arrMCRec)
            arrMCPar = Split(arrMCRec(i), ",")
            If arrMCPar(2) = 1 And arrMCPar(3) = 1 Then
                If Not gclsInsure.TranChargeDetail(2, CStr(arrMCPar(0)), 2, 2, "", , Val(arrMCPar(1))) Then
                    MsgBox "����""" & CStr(arrMCPar(0)) & """������������ҽ������ʧ�ܣ��õ��������ʡ�", vbInformation, gstrSysName
                End If
            End If
        Next
    End If
    Screen.MousePointer = 0
    err = 0: On Error GoTo ErrHandRpt:
    If bln�Ƿ������� = True Then
      If zlStr.IsHavePrivs(mstrPrivs, "����֪ͨ��") Then
            If MsgBox("����Ҫ��ӡ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "����ʱ��=" & strCurDate, "��λ=" & mintUnit + 1, 2)
            End If
     End If
    End If
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing And bln�Ƿ������� Then
        mobjPlugIn.DrugReturnByID Val(mArrFilter("���ϲ���id")), strReturnInfo, CDate(strCurDate), strReserve
    End If
    
    SaveData = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Function
ErrHandRpt:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    SaveData = True
End Function
  

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        If vsHeadGrid.Height + Y <= 500 Or vsDetail.Height - Y <= 500 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        
        vsHeadGrid.Height = vsHeadGrid.Height + Y
        If picBatHsc.Visible Then
            vsDetail.Height = vsDetail.Height - Y
            vsDetail.Top = vsDetail.Top + Y
        Else
            vsDetail.Top = vsDetail.Top + Y
            vsDetail.Height = Me.ScaleHeight - vsDetail.Top
        End If
        Me.Refresh
    End If
End Sub

Private Sub picBatHsc_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        If vsDetail.Height + Y <= 500 Then Exit Sub
        picBatHsc.Top = picBatHsc.Top + Y
        If Me.ScaleHeight - picBatHsc.Top < 500 Then picBatHsc.Top = Me.ScaleHeight - 500
        vsDetail.Height = picBatHsc.Top - vsDetail.Top
        vsBatch.Top = picBatHsc.Top + picBatHsc.Height
        vsBatch.Height = Me.ScaleHeight - vsBatch.Top
            
        Me.Refresh
    End If
End Sub

