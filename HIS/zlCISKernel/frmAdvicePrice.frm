VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdvicePrice 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   1290
   ClientLeft      =   -30
   ClientTop       =   -315
   ClientWidth     =   5670
   ControlBox      =   0   'False
   Icon            =   "frmAdvicePrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "frmAdvicePrice"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   75
      ScaleHeight     =   210
      ScaleWidth      =   5520
      TabIndex        =   1
      Top             =   75
      Width           =   5520
      Begin VB.Label lblClose 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5310
         TabIndex        =   3
         Top             =   15
         Width           =   210
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ƼƼ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Left            =   75
         TabIndex        =   2
         Top             =   15
         Width           =   780
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPrice 
      Height          =   900
      Left            =   75
      TabIndex        =   0
      Top             =   330
      Width           =   5520
      _cx             =   9737
      _cy             =   1587
      Appearance      =   0
      BorderStyle     =   0
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
      BackColor       =   15659506
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   13811126
      ForeColorSel    =   0
      BackColorBkg    =   15659506
      BackColorAlternate=   15659506
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   15659506
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdvicePrice.frx":000C
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.Shape Bdr 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   1230
      Left            =   45
      Top             =   45
      Width           =   5595
   End
End
Attribute VB_Name = "frmAdvicePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event PanelHide()
Private COL_��� As Long
Private COL_���ID As Long
Private COL_ҽ��״̬ As Long
Private COL_������� As Long
Private COL_������ĿID As Long
Private COL_�շ�ϸĿID As Long
Private COL_�걾��λ As Long
Private COL_��鷽�� As Long
Private COL_ִ�б�� As Long
Private COL_�Ƽ����� As Long
Private COL_ִ������ As Long
Private COL_ִ�п���ID As Long

Private mfrmParent As Object
Private vsAdvice As VSFlexGrid
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mint���� As Integer
Private mint���� As Integer '1=����,2=סԺ
Private mlng�������� As Long
Private mstrҩƷ�۸�ȼ� As String '���˵�ҩƷ�۸�ȼ�
Private mstr���ļ۸�ȼ� As String '���˵����ļ۸�ȼ�
Private mstr��ͨ��Ŀ�۸�ȼ� As String '���˵���ͨ��Ŀ�۸�ȼ�

Public Sub HideMe()
    If mlng����ID <> 0 Then Me.Hide
End Sub

Public Sub ShowMe(frmParent As Object, vsEdit As Object, ByVal lng����ID As Long, lng��ҳID As Long, ByVal lng����id As Long, _
    ByVal lng�������� As Long, ByVal int���� As Integer, ByVal strCol As String)
'������lng��ҳID=�������ʱ����0
'      lng��������:0-��ͨסԺ����,1-�������۲��˻����ﲡ��,2-סԺ���۲���
    Dim arrCol As Variant
    
    Set mfrmParent = frmParent
    Set vsAdvice = vsEdit
    
    arrCol = Split(strCol, ",")
    COL_��� = arrCol(0)
    COL_���ID = arrCol(1)
    COL_ҽ��״̬ = arrCol(2)
    COL_������� = arrCol(3)
    COL_������ĿID = arrCol(4)
    COL_�շ�ϸĿID = arrCol(5)
    COL_�걾��λ = arrCol(6)
    COL_��鷽�� = arrCol(7)
    COL_ִ�б�� = arrCol(8)
    COL_�Ƽ����� = arrCol(9)
    COL_ִ������ = arrCol(10)
    COL_ִ�п���ID = arrCol(11)
    
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlng����ID = lng����id
    mint���� = int����
    mlng�������� = lng��������
    mint���� = IIF(mlng��ҳID = 0, 1, 2)
    If lng�������� = 1 Then mint���� = 1
    If InitObjPublicExpense Then Call gobjPublicExpense.zlGetPriceGrade(gstrNodeNo, mlng����ID, mlng��ҳID, "", mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
    
    Call ShowPrice
    Me.Show , frmParent
    
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Function ShowPrice() As Boolean
'���ܣ���ȡָ��ҽ���ļƼ�,�����ݵ�ǰ�������շ� ��ϵ���и���
    Dim rs�շ�ϸĿ As New ADODB.Recordset
    Dim rsTmp As New ADODB.Recordset
    Dim str�շ�ϸĿIDs As String
    Dim StrSQL As String, i As Long, j As Long
    Dim bln�䷽�� As Boolean, bln������ As Boolean, blnLoad As Boolean
    Dim lng���˿���ID As Long, lngִ�п���ID As Long
    Dim dblPrice As Double, lngRow As Long, lngW As Long
    
    Dim strAdvice As String, lngBegin As Long, lngEnd As Long
    Dim str�����շ� As String, str������Ŀ As String, strTmp As String
    Dim strPriceType As String
    
    On Error GoTo errH
        
    With vsAdvice
        lngRow = .Row
        
        '���ɲ���ҽ����¼��ʱ��
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_������ĿID)) <> 0 Then
                strAdvice = strAdvice & " Union ALL " & _
                    "Select " & .RowData(i) & " as ID," & Val(.TextMatrix(i, COL_���)) & " as ���," & ZVal(.TextMatrix(i, COL_���ID)) & " as ���ID," & _
                    Val(.TextMatrix(i, COL_ҽ��״̬)) & " as ҽ��״̬,'" & .TextMatrix(i, COL_�������) & "' as �������," & _
                    Val(.TextMatrix(i, COL_������ĿID)) & " as ������ĿID," & ZVal(.TextMatrix(i, COL_�շ�ϸĿID)) & " as �շ�ϸĿID," & _
                    "'" & .TextMatrix(i, COL_�걾��λ) & "' as �걾��λ,'" & .TextMatrix(i, COL_��鷽��) & "' as ��鷽��," & _
                    Val(.TextMatrix(i, COL_ִ�б��)) & " as ִ�б��," & Val(.TextMatrix(i, COL_�Ƽ�����)) & " as �Ƽ�����," & _
                    Val(.TextMatrix(i, COL_ִ������)) & " as ִ������," & ZVal(.TextMatrix(i, COL_ִ�п���ID), True) & " as ִ�п���ID From Dual"
                
                strTmp = Val(.TextMatrix(i, COL_������ĿID)) & ":" & Val(.TextMatrix(i, COL_ִ�п���ID))
                If InStr("," & str������Ŀ & ",", "," & strTmp & ",") = 0 Then str������Ŀ = str������Ŀ & "," & strTmp
                
            End If
        Next
        strAdvice = Mid(strAdvice, 12)
        str������Ŀ = Mid(str������Ŀ, 2)
    End With
    
    With vsPrice
        .Redraw = False
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
    
        If vsAdvice.RowData(lngRow) = 0 Then
            .Redraw = True: ShowPrice = True: Exit Function
        End If
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "E" Then
            bln�䷽�� = RowIn�䷽��(lngRow)
            bln������ = RowIn������(lngRow)
        End If
                                    
        blnLoad = True
        
        'ҩƷ�����ĵļƼ�
        If vsAdvice.TextMatrix(lngRow, COL_�������) = "4" Then
            '���ģ��̶�������´�
            StrSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,NULL as ��鷽��,0 as ִ�б��,0 as ��������," & _
                " A.�շ�ϸĿID,1 as ҩ����װ,C.���㵥λ,1 as ����,Decode(Nvl(C.�Ƿ���,0),1,Nvl(B.����,D.ȱʡ�۸�),D.�ּ�) as ����,A.ִ�п���ID,0 as ����" & _
                " From (" & strAdvice & ") A,����ҽ���Ƽ� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where Rownum=1 And A.ID=[1] And A.ID=B.ҽ��ID(+) And Nvl(A.ִ������,0)<>5" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "6", "7", "8") & _
                " And A.�շ�ϸĿID=C.ID And C.������� IN([4],3) And D.�շ�ϸĿID=C.ID" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
            blnLoad = False
        ElseIf InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
            '��,����ҩ:���ܰ������ҽ��,����1��ҩ����װ�ĵ���
            StrSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,NULL as ��鷽��,0 as ִ�б��,0 as ��������," & _
                " C.ID as �շ�ϸĿID,Decode([4],1,B.�����װ,B.סԺ��װ) as ҩ����װ,Decode([4],1,B.���ﵥλ,B.סԺ��λ) as ���㵥λ," & _
                " 1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*Decode([4],1,B.�����װ,B.סԺ��װ) as ����," & _
                " A.ִ�п���ID,0 as ����" & _
                " From (" & strAdvice & ") A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where Rownum=1 And A.ID=[1]" & _
                " And A.������ĿID=B.ҩ��ID And B.ҩƷID=C.ID And Nvl(A.ִ������,0)<>5" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "6", "7", "8") & _
                " And (A.�շ�ϸĿID is NULL Or A.�շ�ϸĿID=B.ҩƷID)" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And C.������� IN([4],3) And D.�շ�ϸĿID=C.ID" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
                
                '��һ����ҩ(�����)�ĵ�һ��ҩ�в���ʾ��ҩ;���ļƼ�
                blnLoad = Val(vsAdvice.TextMatrix(lngRow - 1, COL_���ID)) <> Val(vsAdvice.TextMatrix(lngRow, COL_���ID))
        ElseIf bln�䷽�� Then
            '�в�ҩ:һ����Ӧ�й���¼����д���շ�ϸĿID
            StrSQL = "Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,NULL as �걾��λ,NULL as ��鷽��,0 as ִ�б��,0 as ��������," & _
                " C.ID as �շ�ϸĿID,Decode([4],1,B.�����װ,B.סԺ��װ) as ҩ����װ,Decode([4],1,B.���ﵥλ,B.סԺ��λ) as ���㵥λ," & _
                " 1 as ����,Decode(Nvl(C.�Ƿ���,0),1,-NULL,D.�ּ�)*Decode([4],1,B.�����װ,B.סԺ��װ) as ����," & _
                " A.ִ�п���ID,0 as ����" & _
                " From (" & strAdvice & ") A,ҩƷ��� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                " Where A.�������='7' And A.���ID=[1]" & _
                " And A.�շ�ϸĿID=B.ҩƷID And A.�շ�ϸĿID=C.ID And C.������� IN([4],3)" & _
                GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "6", "7", "8") & _
                " And D.�շ�ϸĿID=C.ID And Nvl(A.ִ������,0)<>5" & _
                " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))"
        End If
        
        '��ȡ���мƼ�(ȡ���¼۸�)����ҩƷ������ҽ����ļƼ�,�������ҽ���Ƽ�
        '���Ƽ�,�ֹ��Ƽ۵�ҽ������ȡ
        '��Union��ʽ������������
        If blnLoad Then
            '�����¿���ҽ�������ݲ���ҽ���Ƽ���ȡ
            If InStr(",1,2,-1,", vsAdvice.TextMatrix(lngRow, COL_ҽ��״̬)) = 0 Then
                StrSQL = StrSQL & IIF(StrSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0) as ��������," & _
                    " B.�շ�ϸĿID,1 as ҩ����װ,C.���㵥λ,B.����,Decode(C.�Ƿ���,1,B.����,Sum(D.�ּ�)) as ����," & _
                    " Nvl(B.ִ�п���ID,A.ִ�п���ID) as ִ�п���ID,Nvl(B.����,0) as ����" & _
                    " From (" & strAdvice & ") A,����ҽ���Ƽ� B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                    " Where A.������� Not IN('4','5','6','7') And A.ID=B.ҽ��ID" & _
                    " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0)<>5 And B.�շ�ϸĿID=C.ID And B.�շ�ϸĿID=D.�շ�ϸĿID" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "6", "7", "8") & _
                    " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                    " And (A.ID=[1] Or A.ID=[2] Or A.���ID=[1])" & _
                    " Group by A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0)," & _
                    " B.�շ�ϸĿID,C.���㵥λ,B.����,C.�Ƿ���,B.����,Nvl(B.ִ�п���ID,A.ִ�п���ID),Nvl(B.����,0)"
                        
            Else
                '�¿�δ����ģ�ҽ��״̬��-1
                '�����ڱ����ҵ��շѶ�������,����û�мӲ�λ������������Ҫ��Distinct
                str�����շ� = "Select * From (" & _
                    "Select Distinct C.������ĿID,C.�շ���ĿID,C.��鲿λ,C.��鷽��,C.��������,C.�շ�����,C.���ж���,C.������Ŀ,C.�շѷ�ʽ,C.���ÿ���id" & _
                    " ,Max(Nvl(c.���ÿ���id, 0)) Over(Partition By c.������Ŀid, c.��鲿λ, c.��鷽��, c.��������) As Top" & _
                    " From �����շѹ�ϵ C,Table(f_Num2list2([5])) D Where C.������ĿID=D.c1" & _
                    "      And (C.���ÿ���ID is Null or C.���ÿ���ID = D.c2 And C.������Դ = " & IIF(mint���� = 1, 1, 2) & ")" & _
                    " ) Where Nvl(���ÿ���id, 0) = Top"
                
                '�¿���ҽ��(�ѱ����)�����������շ� ��ϵ��ȡ(��ҩ�����ʾΪȱʡ�۸�)
                StrSQL = StrSQL & IIF(StrSQL = "", "", " Union ALL") & _
                    " Select A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0) as ��������," & _
                    " B.�շ���ĿID as �շ�ϸĿID,1 as ҩ����װ,C.���㵥λ,B.�շ����� as ����,Decode(C.�Ƿ���,1,Sum(D.ȱʡ�۸�),Sum(D.�ּ�)) as ����," & _
                    " A.ִ�п���ID,Nvl(B.������Ŀ,0) as ����" & _
                    " From (" & strAdvice & ") A,(" & str�����շ� & ") B,�շ���ĿĿ¼ C,�շѼ�Ŀ D" & _
                    " Where A.������� Not IN('4','5','6','7') And A.ҽ��״̬ IN(-1,1,2) And A.������ĿID=B.������ĿID" & _
                    GetPriceGradeSQL(mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�, "C", "D", "6", "7", "8") & _
                    " And (A.���ID is Null And A.ִ�б�� IN(1,2) And B.��������=1" & _
                    "       Or A.�걾��λ=B.��鲿λ And A.��鷽��=B.��鷽�� And Nvl(B.��������,0)=0" & _
                    "       Or A.��鷽�� is Null And Nvl(B.��������,0)=0 And B.��鲿λ is Null And B.��鷽�� is Null)" & _
                    " And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5) And B.�շ���ĿID=C.ID And B.�շ���ĿID=D.�շ�ϸĿID" & _
                    " And ((Sysdate Between D.ִ������ and D.��ֹ����) or (Sysdate>=D.ִ������ And D.��ֹ���� is NULL))" & _
                    " And (C.����ʱ�� is NULL Or C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD')) And C.������� IN([4],3)" & _
                    " And (C.վ��='" & gstrNodeNo & "' Or C.վ�� is Null)" & _
                    " And (A.ID=[1] Or A.ID=[2] Or A.���ID=[1])" & _
                    " Group by A.ID,A.���ID,A.���,A.�������,A.������ĿID,A.�걾��λ,A.��鷽��,A.ִ�б��,Nvl(B.��������,0)," & _
                    " B.�շ���ĿID,C.���㵥λ,B.�շ�����,C.�Ƿ���,A.ִ�п���ID,Nvl(B.������Ŀ,0)"
            End If
        End If
        
        '��ȡ������Ŀ��Ϣ
        StrSQL = "Select /*+ RULE */ A.*,B.���� as ������Ŀ,C.���� as �����������" & _
            " From (" & StrSQL & ") A,������ĿĿ¼ B,������Ŀ��� C" & _
            " Where A.������ĿID=B.ID And B.���=C.����"
        StrSQL = StrSQL & " Order by ���,��������,����"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Name, Val(vsAdvice.RowData(lngRow)), Val(vsAdvice.TextMatrix(lngRow, COL_���ID)), 0, mint����, str������Ŀ, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
        
        '��ʾ�Ƽ�����
        If Not rsTmp.EOF Then
            'ȷ����ʾ����
            .Rows = .FixedRows + rsTmp.RecordCount
            
            '��ȡ������Ŀ,�շ�ϸĿ��Ϣ
            For i = 1 To rsTmp.RecordCount
                If InStr("," & str�շ�ϸĿIDs & ",", "," & rsTmp!�շ�ϸĿID & ",") = 0 Then str�շ�ϸĿIDs = str�շ�ϸĿIDs & "," & rsTmp!�շ�ϸĿID
                rsTmp.MoveNext
            Next
            str�շ�ϸĿIDs = Mid(str�շ�ϸĿIDs, 2)
                        
            StrSQL = "Select A.ID,A.���,B.���� as �������,A.����," & _
                " A.����,A.���,A.����,A.��������,A.�Ƿ���" & _
                " From �շ���ĿĿ¼ A,�շ���Ŀ��� B,Table(f_Num2list([1])) D" & _
                " Where A.���=B.���� And A.ID = D.Column_Value"
            StrSQL = "Select/*+ Rule*/ A.ID,A.���,A.�������,A.����,Nvl(B.����,A.����) as ����," & _
                " A.���,A.����,A.��������,N.���� as ҽ������,A.�Ƿ���,C.��������" & _
                " From (" & StrSQL & ") A,�շ���Ŀ���� B,�������� C,����֧����Ŀ M,����֧������ N" & _
                " Where A.ID=C.����ID(+) And A.ID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIF(gbytҩƷ������ʾ = 0, 1, 3) & _
                " And A.ID=M.�շ�ϸĿID(+) And M.����ID=N.ID(+) And M.����(+)=[2]"
            Set rs�շ�ϸĿ = zlDatabase.OpenSQLRecord(StrSQL, Me.Name, str�շ�ϸĿIDs, mint����)
            
            '��ʾÿ������
            rsTmp.MoveFirst
            For i = 1 To rsTmp.RecordCount
                rs�շ�ϸĿ.Filter = "ID=" & rsTmp!�շ�ϸĿID
                
                '�Ƽ�ҽ��
                If rsTmp!������� = "4" Then
                    .TextMatrix(i, 0) = "����"
                ElseIf InStr(",5,6,7,", rsTmp!�������) > 0 Then
                    .TextMatrix(i, 0) = "ҩƷ"
                ElseIf rsTmp!������� = "E" And InStr(",5,6,", vsAdvice.TextMatrix(lngRow, COL_�������)) > 0 Then
                    .TextMatrix(i, 0) = "��ҩ"
                ElseIf rsTmp!������� = "E" And vsAdvice.TextMatrix(lngRow, COL_�������) = "K" Then
                    .TextMatrix(i, 0) = "��Ѫ"
                ElseIf rsTmp!������� = "E" And (bln�䷽�� Or bln������) Then
                    If bln������ Then
                        .TextMatrix(i, 0) = "�ɼ�"
                    ElseIf Not IsNull(rsTmp!���ID) Then
                        .TextMatrix(i, 0) = "�巨"
                    Else
                        .TextMatrix(i, 0) = "�÷�"
                    End If
                ElseIf Not IsNull(rsTmp!���ID) Then
                    If rsTmp!������� = "C" Then
                        .TextMatrix(i, 0) = "����"
                    ElseIf rsTmp!������� = "D" Then
                        '��λ������
                        .TextMatrix(i, 0) = Nvl(rsTmp!�걾��λ) & "(" & Nvl(rsTmp!��鷽��) & ")"
                        '.TextMatrix(i, 0) = "��λ"
                    ElseIf rsTmp!������� = "F" Then
                        .TextMatrix(i, 0) = "����"
                    ElseIf rsTmp!������� = "G" Then
                        .TextMatrix(i, 0) = "����"
                    End If
                Else
                    If Nvl(rsTmp!��������, 0) = 1 Then
                        '���Ի����м��շ���
                        .TextMatrix(i, 0) = Decode(Nvl(rsTmp!ִ�б��, 0), 1, "(����)", 2, "(����)", "(����)")
                    Else
                        .TextMatrix(i, 0) = rsTmp!�����������
                    End If
                End If
                
                '���
                .TextMatrix(i, 1) = rs�շ�ϸĿ!�������
                '�շ���Ŀ:���/����
                .TextMatrix(i, 2) = rs�շ�ϸĿ!����
                If Not IsNull(rs�շ�ϸĿ!���) Then
                    .TextMatrix(i, 2) = .TextMatrix(i, 2) & " " & rs�շ�ϸĿ!���
                End If
                
                '�Ƽ�����:ҩ��ҩƷΪ1,��ҩ��ҩƷΪ��Ӧ�ۼ���
                '���㵥λ:ҩ��ҩƷΪҩ����λ,��ҩ��ҩƷΪ�ۼ۵�λ
                .TextMatrix(i, 3) = FormatEx(rsTmp!����, 5) & Nvl(rsTmp!���㵥λ)
                
                'ִ�п���
                lngִ�п���ID = Nvl(rsTmp!ִ�п���ID, 0)
                If rs�շ�ϸĿ!��� = "4" And Nvl(rs�շ�ϸĿ!��������, 0) = 1 _
                    Or InStr(",5,6,7,", rs�շ�ϸĿ!���) > 0 And InStr(",5,6,7,", rsTmp!�������) = 0 Then
                    lng���˿���ID = mlng����ID
                    lngִ�п���ID = Get�շ�ִ�п���ID(mlng����ID, mlng��ҳID, rs�շ�ϸĿ!���, rs�շ�ϸĿ!ID, 4, lng���˿���ID, 0, mint����, lngִ�п���ID, , , 2)
                End If
                
                '���۴���
                If InStr(",5,6,7,", rs�շ�ϸĿ!���) > 0 Then
                    If Nvl(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                        '��ҩƷʱ��
                        If InStr(",5,6,7,", rsTmp!�������) > 0 Then
                            'ҩ��ҩƷ����һ��ҩ����װ��ʱ��
                            .TextMatrix(i, 4) = CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, Nvl(rsTmp!ҩ����װ, 1), , ,, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�)
                            .TextMatrix(i, 4) = FormatEx(Val(.TextMatrix(i, 4)) * Nvl(rsTmp!ҩ����װ, 0), gbytDecPrice)
                        Else
                            '��ҩ��ҩƷ��������ۼ��������ۼ�ʵ��
                            .TextMatrix(i, 4) = FormatEx(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, Nvl(rsTmp!����, 0), ,, , mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gbytDecPrice)
                        End If
                    Else
                        'ҩ��ҩƷΪҩ������,��ҩҩƷΪ�ۼ�
                        .TextMatrix(i, 4) = FormatEx(Nvl(rsTmp!����), gbytDecPrice)
                    End If
                ElseIf rs�շ�ϸĿ!��� = "4" And Nvl(rs�շ�ϸĿ!��������, 0) = 1 And Nvl(rs�շ�ϸĿ!�Ƿ���, 0) = 1 Then
                    'ʱ�����ĵĵ��ۺ�ҩƷһ������
                    .TextMatrix(i, 4) = FormatEx(CalcDrugPrice(rs�շ�ϸĿ!ID, lngִ�п���ID, Nvl(rsTmp!����, 0), , ,, mstrҩƷ�۸�ȼ�, mstr���ļ۸�ȼ�, mstr��ͨ��Ŀ�۸�ȼ�), gbytDecPrice)
                Else
                    .TextMatrix(i, 4) = FormatEx(Nvl(rsTmp!����), gbytDecPrice)
                End If
                
                '��ʾҽ����������
                If Val(rsTmp!�շ�ϸĿID & "") <> 0 Then
                    strPriceType = GetPriceType(Val(mlng����ID), Val(rsTmp!�շ�ϸĿID & ""), Val(mint����), mlng�������� = 1)
                End If
                '��������
                If strPriceType = "" Then
                    .TextMatrix(i, 5) = Nvl(rs�շ�ϸĿ!��������)
                Else
                    .TextMatrix(i, 5) = strPriceType
                End If
                         
                .TextMatrix(i, 6) = Nvl(rs�շ�ϸĿ!ҽ������)
                
                dblPrice = dblPrice + FormatEx(Nvl(rsTmp!����, 0) * Val(.TextMatrix(i, 4)), 5)
                
                rsTmp.MoveNext
            Next
        End If
        
        '������ߴ�
        With vsPrice
            If .Rows < 3 Then .Rows = 3
            Call .AutoSize(0, .Cols - 1)
            For i = 0 To .Cols - 1
                If .ColWidth(i) > 1500 Then
                    .ColWidth(i) = 1500
                Else
                    .ColWidth(i) = .ColWidth(i) - 90
                End If
                lngW = lngW + .ColWidth(i)
            Next
            .Width = lngW + IIF(.Rows > 6, 225, 0)
            .Height = .RowHeight(1) * IIF(.Rows > 6, 6, .Rows)
        End With
        
        .CellBorderRange 0, 0, 0, .Cols - 1, &H80000008, 0, 0, 0, 1, 0, 0
        
        .Row = 1: .Col = 0
        .Redraw = True
    End With
    Call SetFormSize
    ShowPrice = True
    Exit Function
errH:
    vsPrice.Redraw = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    Dim lngS��ID As Long, lngO��ID As Long, i As Long
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS��ID = IIF(Val(.TextMatrix(lngRow, COL_���ID)) = 0, .RowData(lngRow), Val(.TextMatrix(lngRow, COL_���ID)))
        For i = lngRow - 1 To .FixedRows Step -1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_���ID)))
            If lngO��ID = lngS��ID Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow + 1 To .Rows - 1
            lngO��ID = IIF(Val(.TextMatrix(i, COL_���ID)) = 0, .RowData(i), Val(.TextMatrix(i, COL_���ID)))
            If lngO��ID = lngS��ID Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Function RowIn������(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ����ڼ�������е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "E" And Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
            '�ɼ�������
            If .TextMatrix(lngRow - 1, COL_�������) = "C" _
                And Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) Then
                RowIn������ = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, COL_�������) = "C" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '������Ŀ��
            RowIn������ = True: Exit Function
        End If
    End With
End Function

Private Function RowIn�䷽��(ByVal lngRow As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ�������ҩ�䷽�е�һ��
'˵���������е�ǰ�Ƿ�����
    If lngRow = -1 Then Exit Function
    If vsAdvice.RowData(lngRow) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, COL_�������) = "E" Then
            If Val(.TextMatrix(lngRow, COL_���ID)) = 0 Then
                '�÷���
                If Val(.TextMatrix(lngRow - 1, COL_���ID)) = .RowData(lngRow) _
                    And .TextMatrix(lngRow - 1, COL_�������) = "E" Then
                    RowIn�䷽�� = True: Exit Function
                End If
            Else
                '�巨��
                If .TextMatrix(lngRow - 1, COL_�������) = "7" _
                    And Val(.TextMatrix(lngRow - 1, COL_���ID)) = Val(.TextMatrix(lngRow, COL_���ID)) Then
                    RowIn�䷽�� = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, COL_�������) = "7" And Val(.TextMatrix(lngRow, COL_���ID)) <> 0 Then
            '��ҩ��
            RowIn�䷽�� = True: Exit Function
        End If
    End With
End Function

Private Sub Form_Load()
    Dim strPos As String
    
    Call zlControl.FormSetCaption(Me, False, False)
    If mint���� = 0 Then
        vsPrice.ColHidden(6) = True
        vsPrice.Width = vsPrice.Width - vsPrice.ColWidth(6)
    End If

    strPos = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & mfrmParent.Name, "PricePanePostion", "1600,5500")
    Me.Top = mfrmParent.Top + Val(Split(strPos, ",")(0))
    Me.Left = mfrmParent.Left + Val(Split(strPos, ",")(1))
End Sub

Private Sub SetFormSize()
    zlControl.FormLock Me.Hwnd
    Me.Width = vsPrice.Width + (Bdr.BorderWidth * 15 + 30) * 2
    Me.Height = vsPrice.Height + picTitle.Height + (Bdr.BorderWidth * 15 + 30) * 2 - 15
    
    Bdr.Left = 15
    Bdr.Top = 15
    Bdr.Width = Me.Width - 15
    Bdr.Height = Me.Height - 15
    
    picTitle.Left = Bdr.Left + Bdr.BorderWidth * 15 + 15
    picTitle.Top = Bdr.Top + Bdr.BorderWidth * 15 + 15
    picTitle.Width = Me.Width - picTitle.Left * 2
    
    vsPrice.Left = picTitle.Left
    vsPrice.Top = picTitle.Top + picTitle.Height
    
    Call SetCloseButton(0, True)
    zlControl.FormLock 0
End Sub

Private Sub SetCloseButton(ByVal intState As Integer, Optional ByVal blnSize As Boolean)
'������intState=0-����,1-����,2-����
    If intState = 0 Then
        lblClose.BackColor = picTitle.BackColor
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 0
    ElseIf intState = 1 Then
        lblClose.BackColor = vsPrice.BackColorSel
        lblClose.ForeColor = vbBlack
        lblClose.BorderStyle = 1
    ElseIf intState = 2 Then
        lblClose.BackColor = 11899525
        lblClose.ForeColor = vbWhite
        lblClose.BorderStyle = 1
    End If
    
    If blnSize Then
        lblClose.Width = 210
        lblClose.Height = 195
        lblClose.Left = picTitle.Width - lblClose.Width - 15
        lblClose.Top = (picTitle.Height - lblClose.Height) / 2
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
    End If
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetCloseButton(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngTop As Long, lngLeft As Long
    
    '������������������Ͻǵ�λ��
    If mfrmParent.WindowState = 0 Then
        lngTop = Me.Top - mfrmParent.Top
        lngLeft = Me.Left - mfrmParent.Left
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & mfrmParent.Name, "PricePanePostion", lngTop & "," & lngLeft
    End If
    
    mlng����ID = 0
    mlng��ҳID = 0
    mlng����ID = 0
    Set mfrmParent = Nothing
    Set vsAdvice = Nothing
End Sub

Private Sub lblClose_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call SetCloseButton(2)
    End If
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Sub lblClose_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And Y >= 0 And x <= lblClose.Width And Y <= lblClose.Height Then
        If Button = 1 Then
            Call SetCloseButton(2)
        Else
            Call SetCloseButton(1)
        End If
    Else
        Call SetCloseButton(1)
    End If
End Sub

Private Sub lblClose_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And Y >= 0 And x <= lblClose.Width And Y <= lblClose.Height Then
        Me.Hide
        RaiseEvent PanelHide
        If mfrmParent.Visible Then mfrmParent.SetFocus
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
    End If
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        Call MoveObj(Me.Hwnd)
        If mfrmParent.Visible Then mfrmParent.SetFocus
    End If
End Sub

Private Sub picTitle_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetCloseButton(0)
End Sub

Private Sub vsPrice_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mfrmParent.Visible Then mfrmParent.SetFocus
End Sub

Private Sub vsPrice_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call SetCloseButton(0)
    With vsPrice
        If .MouseCol = 2 And Between(.MouseRow, .FixedRows, .Rows - 1) Then
            .ToolTipText = .TextMatrix(.MouseRow, .MouseCol)
        Else
            .ToolTipText = ""
        End If
    End With
End Sub
