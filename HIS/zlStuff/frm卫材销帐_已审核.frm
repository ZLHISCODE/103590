VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frm��������_����� 
   BorderStyle     =   0  'None
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10005
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   12615
   End
   Begin VB.PictureBox picBatHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   12735
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4155
      Width           =   12735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsHeadGrid 
      Height          =   2055
      Left            =   75
      TabIndex        =   0
      Tag             =   "������"
      Top             =   105
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
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm��������_�����.frx":0000
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
      TabIndex        =   1
      Tag             =   "��ϸ"
      Top             =   2205
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm��������_�����.frx":00B8
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
      TabIndex        =   2
      Tag             =   "��ϸ"
      Top             =   4245
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm��������_�����.frx":01F7
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
Attribute VB_Name = "frm��������_�����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '��������
Private mrsVerifyBatch As ADODB.Recordset
Private mrsVerifyDetail As ADODB.Recordset      '�������ϸ��¼���ݼ�

Private mint��˱�־ As Integer
Private mintUnit As Integer '0-ɢװ��λ,1-��װ��λ
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
 
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
       ' .Editable = flexEDKbdMouse
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
    
    On Error GoTo ErrHandle
    Call InitRsStruct
    mint��˱�־ = 1
    ''''1����ȡ��������
    '��λ����װ����
    Select Case mintUnit
    Case 0
        strFields = "X.���㵥λ ��λ,1 ����ϵ��,A.���� As �������� "
    Case Else
        strFields = "D.��װ��λ ��λ,d.����ϵ�� ,A.���� As �������� "
    End Select
    
    strWere = strWere & " And A.����� Is Not Null And A.״̬ <> 0 And A.���ʱ�� Between [3] And [4] "
      
    '����/ҽ������
    If Val(mArrFilter("�������ID")) > 0 Then strWere = strWere & " And A.���벿��id = [2] "
    '������
    If Trim(mArrFilter("������")) <> "" Then strWere = strWere & " And A.������=[7] "
    '��������
    If Trim(mArrFilter("��������")) <> "" Then strWere = strWere & " And nvl(F.����,B.����)=[8] "
    strWere = strWere & IIf(Val(mArrFilter("סԺ��")) = 0, "", "             AND b.��ʶ��=[9] and b.�����־=2 ")
    strWere = strWere & IIf(Val(mArrFilter("����ID")) = 0, "", "             AND b.����iD=[10]  ")
    strWere = strWere & IIf(Trim(mArrFilter("����")) = "", "", "             AND b.����iD=[10]  ")

    
    gstrSQL = "" & _
    "   Select Distinct A.״̬,A.�շ�ϸĿid,'['||X.����||']'||X.���� as ��������, X.���, " & strFields & _
    "   From (  Select A.״̬,A.�շ�ϸĿid, Sum(A.����) As ���� " & _
    "           From סԺ���ü�¼ B, ������ҳ F, ���˷������� A " & _
    "           Where A.�������=1 And A.����id = B.ID And A.��˲���id = [1] And B.����id = F.����id(+)  " & _
    "                 And B.��ҳid = F.��ҳid(+)  And F.��Ժ����(+)  Is Null And F.״̬(+)<> 3 " & _
    "                 " & vbCrLf & strWere & _
    "                 And Exists (Select 1 From ҩƷ�շ���¼ C  Where C.����id = A.����id And C.����� Is Not Null And (C.��¼״̬ = 1 Or Mod(C.��¼״̬, 3) = 0))" & _
    "           Group By A.�շ�ϸĿid,A.״̬) A,�������� D, �շ���Ŀ���� E, �շ���ĿĿ¼ X " & _
    " Where A.�շ�ϸĿid = D.����id And A.�շ�ϸĿid = X.ID And X.ID = E.�շ�ϸĿid(+) And E.����(+) = 3 " & _
    " Order By ��������"
     
    '[5]��[6]����Ϊ����ʱ��,��ȡ��
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", _
        Val(mArrFilter("���ϲ���id")), Val(mArrFilter("�������id")), _
        CDate(mArrFilter("�������")(0)), CDate(mArrFilter("�������")(1)), _
        CDate(mArrFilter("�������")(0)), CDate(mArrFilter("�������")(1)), _
        Trim(mArrFilter("������")), Trim(mArrFilter("��������")), _
        Val(mArrFilter("סԺ��")), Val(mArrFilter("����ID")), Trim(mArrFilter("����")))
    
    With vsHeadGrid
        .Clear 1
        .Rows = 2
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 0
        Do While Not rsTemp.EOF
            lngRow = lngRow + 1
            .TextMatrix(lngRow, .ColIndex("���")) = IIf(Val(NVL(rsTemp!״̬)) = 1, "��", "��")
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(rsTemp!��������)
            .TextMatrix(lngRow, .ColIndex("���")) = NVL(rsTemp!���)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(Val(NVL(rsTemp!��������)) / rsTemp!����ϵ��, mFMT.FM_����)
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
        "   Select ����, NO, ҩƷID as ����ID, ����ʱ��, ��ʶ��, ����, ����, ��λ, ����ϵ��,��������, Sum(����) As �������� " & _
        "   From (  Select Distinct C.����, C.NO, C.ҩƷID, A.����ʱ��, B.��ʶ��, nvl(F.����,B.����) ����, B.����,P.���� ��������, " & strFields & " " & _
        "           From ���˷������� A, סԺ���ü�¼ B,ҩƷ�շ���¼ C, �������� D, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E " & _
        "           Where   A.�������=1 And A.����id = B.ID And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.����id And B.�շ�ϸĿid = X.ID  " & _
        "                   And  B.����id = F.����id(+) And B.��ҳid = F.��ҳid(+) And F.��Ժ����(+) Is Null And F.״̬(+) <> 3 " & _
        "                   And A.���벿��id = E.ID And B.ִ�в���id = [1]  " & _
        "                   And C.����� Is Not Null And C.���� In (24, 25, 26) And (C.��¼״̬ = 1 Or Mod(C.��¼״̬, 3) = 0)   " & strWere & ")" & _
        "           Group By ����, NO, ҩƷID, ����ʱ��, ��ʶ��, ����, ����, ��λ, ����ϵ��,�������� "
        
           
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", _
          Val(mArrFilter("���ϲ���id")), Val(mArrFilter("�������id")), _
        CDate(mArrFilter("�������")(0)), CDate(mArrFilter("�������")(1)), _
        CDate(mArrFilter("�������")(0)), CDate(mArrFilter("�������")(1)), _
          Trim(mArrFilter("������")), Trim(mArrFilter("��������")), _
          Val(mArrFilter("סԺ��")), Val(mArrFilter("����ID")), Trim(mArrFilter("����")))
      
    Do While Not rsTemp.EOF
        With mrsVerifyDetail
            .AddNew
    
            !���� = rsTemp!����
            !NO = rsTemp!NO
            !����ID = rsTemp!����ID
            !����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            !��ʶ�� = rsTemp!��ʶ��
            !���� = rsTemp!����
            !���� = rsTemp!����
            !�������� = rsTemp!��������
            !����ϵ�� = rsTemp!����ϵ��
            !��λ = rsTemp!��λ
            !�������� = rsTemp!��������
            .Update
            rsTemp.MoveNext
        End With
    Loop
     
    ''''3����ȡ������ϸ����
    '��λ����װ����
    Select Case mintUnit
    Case 0
        strFields = "X.���㵥λ ��λ,1 ����ϵ��,C.ʵ������ As ׼������,c.ʵ������ * c.���ϵ�� As ��������"
    Case Else
        strFields = "D.��װ��λ ��λ,D.����ϵ��  ,C.ʵ������ As ׼������,c.ʵ������ * c.���ϵ�� As ��������"
    End Select
    
'    gstrSQL = "" & _
'        "   Select C.ID As �շ�ID, C.ҩƷID, C.����, C.NO, C.��� As �շ����, C.����, C.����, C.Ч��, C.�������� As ����ʱ��, " & _
'        "           F.����, P.���� As ��������, A.����id, B.��� As �������, B.��¼����, B.��ҳID, A.����ʱ��, " & strFields & " " & _
'        "   From ���˷������� A, סԺ���ü�¼ B,ҩƷ�շ���¼ C, �������� D, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E " & _
'        "   Where A.�������=1 And A.����id = B.ID And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.����id And B.�շ�ϸĿid = X.ID  " & _
'        "       And B.����id = F.����id(+) And B.��ҳid = F.��ҳid(+) And F.��Ժ����(+) Is Null And F.״̬(+) <> 3 And A.���벿��id = E.ID " & _
'        "       And B.ִ�в���id = [1]  " & _
'        "       And C.����� Is Not Null And C.���� In (24, 25, 26) And (C.��¼״̬ = 1 Or Mod(C.��¼״̬, 3) = 0) " & strWere & _
'        "   Order By A.���ʱ��, C.����, C.NO, C.���"

        gstrSQL = "Select C.ID As �շ�ID, C.ҩƷID, C.����, C.NO, C.��� As �շ����, C.����, C.����, C.Ч��, F.����, P.���� As ��������,C.����, " & _
            " A.����id, B.��� As �������, B.��¼����, B.��ҳID, A.����ʱ��, A.���ʱ��, C.���ۼ� As ����, " & strFields & " " & _
            " From ���˷������� A, סԺ���ü�¼ B,ҩƷ�շ���¼ C, �������� D, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E " & _
            " Where A.�������=1 And A.����id = B.ID And B.No = C.No And B.ID = C.����id And B.��������id = P.ID And B.�շ�ϸĿid = D.����id And B.�շ�ϸĿid = X.ID And B.����id = F.����id(+) And B.��ҳid = F.��ҳid(+) And A.���벿��id = E.ID " & _
            " And B.ִ�в���id = [1]  " & strWere & _
            " And C.������� Is Not Null " & _
            " And ((A.״̬ = 1 And Mod(C.��¼״̬, 3) = 2 And A.���ʱ�� = C.�������) Or (A.״̬ = 2 And (C.��¼״̬ = 1 Or Mod(C.��¼״̬, 3) = 0))) "
        
        gstrSQL = gstrSQL & " Order By A.���ʱ��, C.����, C.NO, C.���"

     
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", _
          Val(mArrFilter("���ϲ���id")), Val(mArrFilter("�������id")), _
          CDate(mArrFilter("�������")(0)), CDate(mArrFilter("�������")(1)), _
          CDate(mArrFilter("�������")(0)), CDate(mArrFilter("�������")(1)), _
          Trim(mArrFilter("������")), Trim(mArrFilter("��������")), _
          Val(mArrFilter("סԺ��")), Val(mArrFilter("����ID")), Trim(mArrFilter("����")))
    Do While Not rsTemp.EOF
        With mrsVerifyBatch
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
    Set mrsVerifyDetail = New ADODB.Recordset
    With mrsVerifyDetail
        If .State = 1 Then .Close
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��ʶ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ϵ��", adDouble, 18, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 50, adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set mrsVerifyBatch = New ADODB.Recordset
    With mrsVerifyBatch
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

Private Sub Form_Load()
    zl_vsGrid_Para_Restore mlngModule, vsHeadGrid, Me.Caption, "��������_Head"
    zl_vsGrid_Para_Restore mlngModule, vsDetail, Me.Caption, "��������_Detail"
    zl_vsGrid_Para_Restore mlngModule, vsBatch, Me.Caption, "��������_Batch"
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
    
    With mrsVerifyBatch
        .Sort = "�շ���� desc"
        If mrsVerifyBatch.RecordCount <> 0 Then .MoveFirst
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
        mrsVerifyDetail.Filter = "����ID=" & lng����ID
        If mrsVerifyDetail.RecordCount = 0 Then Exit Sub
        .Rows = mrsVerifyDetail.RecordCount + 1
        lngRow = 0
        Do While Not mrsVerifyDetail.EOF
            lngRow = lngRow + 1
            .RowData(lngRow) = Val(NVL(mrsVerifyDetail!����ID))
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsVerifyDetail!����)
            .TextMatrix(lngRow, .ColIndex("NO")) = NVL(mrsVerifyDetail!NO)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(mrsVerifyDetail!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(lngRow, .ColIndex("����(סԺ)��")) = NVL(mrsVerifyDetail!��ʶ��)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsVerifyDetail!����)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsVerifyDetail!����)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(Val(NVL(mrsVerifyDetail!��������)) / mrsVerifyDetail!����ϵ��, mFMT.FM_����)
            .TextMatrix(lngRow, .ColIndex("��λ")) = NVL(mrsVerifyDetail!��λ)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsVerifyDetail!��������)
            
            .Cell(flexcpData, lngRow, .ColIndex("��������")) = Val(NVL(mrsVerifyDetail!��������))
            .Cell(flexcpData, lngRow, .ColIndex("NO")) = Val(NVL(mrsVerifyDetail!����ϵ��))
            mrsVerifyDetail.MoveNext
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
        .Clear 1
        .Rows = 2
        If mrsVerifyBatch Is Nothing Then
            Exit Sub
        End If
        
        mrsVerifyBatch.Filter = "����=" & int���� & _
                " And No='" & strNo & "' " & _
                " And ����ID=" & lng����ID & _
                " And ����ʱ��='" & str����ʱ�� & "' "
        mrsVerifyBatch.Sort = "�շ���� Desc"
        If mrsVerifyBatch.RecordCount = 0 Then Exit Sub
        
        If mrsVerifyBatch.RecordCount = 0 Then
            vsDetail.Height = Me.ScaleHeight - vsDetail.Top
            picBatHsc.Visible = False
        Else
            picBatHsc.Top = Me.ScaleHeight - 1935
            vsDetail.Height = picBatHsc.Top - vsDetail.Top
            vsBatch.Top = picBatHsc.Top + picBatHsc.Height
            vsBatch.Height = Me.ScaleHeight - Me.vsBatch.Top
            picBatHsc.Visible = True
        End If
        
        .Rows = mrsVerifyBatch.RecordCount + 1
        mrsVerifyBatch.MoveFirst
        lngRow = 0
        Do While Not mrsVerifyBatch.EOF
            lngRow = lngRow + 1
            .RowData(lngRow) = Val(NVL(mrsVerifyBatch!����ID))
            vsBatch.TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsVerifyBatch!����)
            vsBatch.TextMatrix(lngRow, .ColIndex("NO")) = NVL(mrsVerifyBatch!NO)
            vsBatch.TextMatrix(lngRow, .ColIndex("����ʱ��")) = Format(mrsVerifyBatch!����ʱ��, "yyyy-mm-dd hh:mm:ss")
            vsBatch.TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsVerifyBatch!����)
            vsBatch.TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsVerifyBatch!����)
            vsBatch.TextMatrix(lngRow, .ColIndex("Ч��")) = Format(mrsVerifyBatch!Ч��, "yyyy-mm-dd")
            vsBatch.TextMatrix(lngRow, .ColIndex("׼������")) = Format(Val(NVL(mrsVerifyBatch!׼������)) / mrsVerifyBatch!����ϵ��, mFMT.FM_����)
            vsBatch.TextMatrix(lngRow, .ColIndex("��������")) = Abs(Format(Val(NVL(mrsVerifyBatch!��������)) / mrsVerifyBatch!����ϵ��, mFMT.FM_����))
            vsBatch.TextMatrix(lngRow, .ColIndex("��λ")) = NVL(mrsVerifyBatch!��λ)
            
            vsBatch.Cell(flexcpData, lngRow, .ColIndex("׼������")) = Val(NVL(mrsVerifyBatch!׼������))
            vsBatch.Cell(flexcpData, lngRow, .ColIndex("��������")) = Abs(Val(NVL(mrsVerifyBatch!��������)))
            vsBatch.Cell(flexcpData, lngRow, .ColIndex("NO")) = Val(NVL(mrsVerifyBatch!����ϵ��))
            mrsVerifyBatch.MoveNext
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
    zl_vsGrid_Para_Save mlngModule, vsHeadGrid, Me.Caption, "��������_Head"
    zl_vsGrid_Para_Save mlngModule, vsDetail, Me.Caption, "��������_Detail"
    zl_vsGrid_Para_Save mlngModule, vsBatch, Me.Caption, "��������_Batch"
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
 
    With vsHeadGrid
        If .Row > 0 Then
            If .Cell(flexcpData, .Row, .ColIndex("��������")) = 0 Then Exit Sub
        End If
        If .Row > 0 Then
            '��ȡ��ϸ����
            Call LoadDetailList(Val(.Cell(flexcpData, .Row, .ColIndex("��������"))))
        End If
    End With

    '��ȡ������ϸ����
    With vsDetail
        Call LoadBatchList(Val(.TextMatrix(.Row, .ColIndex("����"))), .TextMatrix(.Row, .ColIndex("NO")), .RowData(.Row), .TextMatrix(.Row, .ColIndex("����ʱ��")), False)
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
Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Private Sub picBatHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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



