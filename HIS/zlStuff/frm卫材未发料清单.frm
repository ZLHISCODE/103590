VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm����δ�����嵥 
   BorderStyle     =   0  'None
   Caption         =   "δ�����嵥"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdSel 
      Caption         =   "��"
      Height          =   285
      Left            =   2955
      TabIndex        =   5
      Top             =   4935
      Width           =   285
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   660
      MaxLength       =   20
      TabIndex        =   4
      Top             =   4920
      Width           =   2595
   End
   Begin VB.ComboBox cboEdit 
      Height          =   300
      Index           =   0
      Left            =   4620
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4560
      Width           =   2625
   End
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   4125
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   7320
      _cx             =   12912
      _cy             =   7276
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   12632256
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   21
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm����δ�����嵥.frx":0000
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
      ExplorerBar     =   7
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
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ϴ�ӡ���ݸ�ʽ"
      Height          =   180
      Index           =   1
      Left            =   3165
      TabIndex        =   2
      Top             =   4650
      Width           =   1440
   End
   Begin VB.Label lblEdit 
      Caption         =   "������"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   4965
      Width           =   615
   End
End
Attribute VB_Name = "frm����δ�����嵥"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsNotPayStuff As ADODB.Recordset       '�ڲ���¼��:δ���ϲ���
Private mrsChargeOff As New ADODB.Recordset                   '������ʾ���������¼
Private mbln����ǰ�շѻ���� As Boolean
Private mbln����δ�շѵ����ﻮ�۴������� As Boolean
Private mbln����δ��˵ļ��˴������� As Boolean
Private mbln����ʱ�������� As Boolean
Private mstrPrivs As String
Private mlngModule As Long
Private mArrFilter As Variant   '��������
Private mfrmMain As Form        '������
Private mlngȱ�ϼ�� As Long
Private mbln�����ݷ��� As Boolean
Private mbln��ʾ�������� As Boolean     '����
Private mintUnit As Integer     '��ʾ��λ
Private mbln������ǩ�� As Boolean   '����
Private mrsMatStock As ADODB.Recordset      '�洢�ⷿ
Private mstrNo As String                    '��ǰѡ���NO
Private Const mstrAllType As String = "�ٴ�,����,���,����,����,����,Ӫ��"
Private mfrmFilter As New frm���ķ��Ź���
'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Enum mcboIdx
    idx_���ݸ�ʽ = 0
End Enum
Private Enum mtxtIdx
    idx_������ = 0
End Enum
Private Enum mlblIdx
    idx_lbl������ = 0
    idx_lbl���ݸ�ʽ = 1
End Enum
Private mblnHave���� As Boolean     '�Ƿ���ڷ�����Ŀ
Private mblnHave�ܷ� As Boolean         '�Ƿ���ھܷ�����Ŀ
Private mstrĬ�ϵ��ݸ�ʽ As String '
Private mstr���ܱ�ʶ�� As String
Private mbln������ʱ����� As Boolean

Public Event zlRefreshDataRecordSet(ByVal rsNotStuffStuff As ADODB.Recordset, ByVal rsChargeOff As ADODB.Recordset)

Private gclsInsure As New clsInsure
Private Type TYPE_MedicarePAR
    �������� As Boolean
    �����ϴ� As Boolean
    ������ɺ��ϴ� As Boolean
    ���������ϴ� As Boolean
End Type
Private MCPAR As TYPE_MedicarePAR

Private mobjPlugIn As Object             '��ҽӿڶ���

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
Private Function CheckIsStockUp(ByVal lng���ϲ���ID As Long, ByVal lng����ID As Long, ByVal lng���� As Long, _
        ByVal lng����ID As Long, ByVal dblʵ������ As Double) As Boolean
    '1����鵱ǰ��¼�Ƿ��Ǳ������ĵķ��ϼ�¼��ͨ�������Ƿ��ж�Ӧ��δ��˵�����ⷿ�������ⵥ�����ж�
    '2�����������ⷿ��ʵ�������Ƿ��㹻
    Dim rsData As ADODB.Recordset
    Dim lng����ⷿid As Long
    
    On Error GoTo ErrHand
    gstrSQL = "Select a.�ⷿid From ҩƷ�շ���¼ A, ����ⷿ���� B " & _
        " Where a.�ⷿid + 0 = b.����ⷿid And a.���� = 21 And a.������� Is Null And b.����id = [1] And a.ҩƷid + 0 = [2] " & _
        " And Nvl(a.����, 0) = [3] And a.����id = [4] And Rownum = 1 "
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIsStockUp", lng���ϲ���ID, lng����ID, lng����, lng����ID)
    
    If rsData.EOF Then
        CheckIsStockUp = False
        Exit Function
    Else
        lng����ⷿid = rsData!�ⷿID
    End If
    
    gstrSQL = "Select nvl(ʵ������,0) As ʵ������ From ҩƷ��� Where ����=1 And �ⷿID=[1] And ҩƷID=[2] And nvl(����,0)=[3]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckIsStockUp", lng����ⷿid, lng����ID, lng����)
    
    If rsData.EOF Then
        CheckIsStockUp = False
        Exit Function
    ElseIf Val(rsData!ʵ������) < dblʵ������ Then
         CheckIsStockUp = False
         Exit Function
    End If
    
    CheckIsStockUp = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RefreshChargeOffStation(ByVal rsStuffSendData As ADODB.Recordset, ByRef rsChargeOff As ADODB.Recordset)
    '���ݷ��ϼ�¼��ִ��״̬�������˼�¼��ִ��״̬
    
    rsStuffSendData.Filter = 0
    If rsStuffSendData.RecordCount = 0 Then Exit Sub
    
    rsChargeOff.Filter = 0
    If rsChargeOff.RecordCount = 0 Then Exit Sub
    
    rsChargeOff.MoveFirst
    Do While Not rsChargeOff.EOF
        '�Ƚ�ִ��״̬��Ϊ0
        rsChargeOff!ִ��״̬ = 0
        rsChargeOff!��˱�־ = 0
        
        rsStuffSendData.MoveFirst
        Do While Not rsStuffSendData.EOF
            'ֻҪ��һ����Ӧ�ķ��Ͽ��ң�����ID��ִ��״̬=1������¶�Ӧ�����˼�¼ִ��״̬=1
            If rsChargeOff!���ϲ���id = rsStuffSendData!����id And rsChargeOff!����ID = rsStuffSendData!����ID And rsStuffSendData!ִ��״̬ = 1 Then
                rsChargeOff!ִ��״̬ = 1
                rsChargeOff!��˱�־ = 1
                Exit Do
            End If
            
            rsStuffSendData.MoveNext
        Loop
        
        rsChargeOff.Update
        
        rsChargeOff.MoveNext
    Loop
End Sub

Private Sub GetChargeOffRecord(ByVal rsStuffSendData As ADODB.Recordset)
    '1.ͳ�Ʒ��ϼ�¼��������Щ���ϲ���(����)
    '2.��ѯ���ʼ�¼��������˲����Ǹÿⷿ��Ӧ�������������ݣ������벿��id��״̬=0Ϊ��������ȡ����¼���벿��ID���շ�ϸĿID������ʱ�䣬����ID�����������ȹؼ���Ϣ
    '3.ѭ���������ݼ����жϴ�1���ҵ����շ�ϸĿID�����벿��ID������ID�ڷ������ݼ����Ƿ���ڣ�������ڱ�ʾ����ͬʱ���Ϻ����ʵ����
    '4.��3���ҵ������벿��ID���շ�ϸĿid������ʱ���ٸ��ݷ���ID�����շ���¼�ȱ�����֯������������
    '5.��4�п���һ������ID��Ӧ����շ�ID����ͬ���Σ����������������жϸ��Ե�׼�������Ƿ��㹻�����������ֽ⵽��ͬ���շ�ID�����Σ���
    
    Dim rsTmp As ADODB.Recordset
    Dim rsChargeOffTmp As ADODB.Recordset
    Dim strDeptIDs As String
    Dim lngDeptId As Long
    Dim lngStuffID As Long
    Dim lngChargeID As Long
    Dim str��װ��λ As String
    Dim dblʣ���������� As Double
    Dim str����ʱ�� As String
    
    On Error GoTo ErrHandle
    
    If mbln����ʱ�������� = False Then Exit Sub
    
    rsStuffSendData.Filter = "ִ��״̬=1"
    If rsStuffSendData.RecordCount = 0 Then Exit Sub
    
    Set rsChargeOffTmp = New ADODB.Recordset
    With rsChargeOffTmp
        If .State = 1 Then .Close
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���벿��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "�շ�ϸĿID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    '1.���ϲ��Ż���
    rsStuffSendData.Sort = "����id"
    Do While Not rsStuffSendData.EOF
        If lngDeptId <> rsStuffSendData!����id Then
            lngDeptId = rsStuffSendData!����id
            strDeptIDs = IIf(strDeptIDs = "", "", strDeptIDs & ",") & lngDeptId
        End If
        rsStuffSendData.MoveNext
    Loop
    
    '2.��ѯ���ʼ�¼
    If InStr(strDeptIDs, ",") = 0 Then
        gstrSQL = "Select ����id, ���벿��id, �շ�ϸĿid, ����, ����ʱ�� " & _
            " From ���˷������� " & _
            " Where ��˲���id = [1] And ������� = 1 And ״̬ = 0 And ���벿��id = [2] " & _
            " Order By �շ�ϸĿid, ���벿��id "
    Else
        gstrSQL = "Select a.����id, a.���벿��id, a.�շ�ϸĿid, a.����, a.����ʱ�� " & _
            " From ���˷������� A, Table(f_Str2list([2])) T " & _
            " Where A.��˲���id = [1] And A.������� = 1 And A.״̬ = 0 And A.���벿��id = t.Column_Value " & _
            " Order By a.�շ�ϸĿid, a.���벿��id "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "GetChargeOffRecord", Val(mArrFilter("���ϲ���ID")), strDeptIDs)
    
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    '3.����ƥ���ͬʱ���˷�����Ŀ������ҵ��򱣴浽��ʱ���˼�¼��
    rsStuffSendData.Sort = "����id,����id"
    
    lngChargeID = 0
    lngStuffID = 0
    lngDeptId = 0
    
    Do While Not rsTmp.EOF
        lngStuffID = rsTmp!�շ�ϸĿid
        lngDeptId = rsTmp!���벿��id
        
        rsStuffSendData.MoveFirst
        Do While Not rsStuffSendData.EOF
            If lngStuffID = rsStuffSendData!����ID And lngDeptId = rsStuffSendData!����id Then
                rsChargeOffTmp.AddNew
                rsChargeOffTmp!����ID = rsTmp!����ID
                rsChargeOffTmp!���벿��id = rsTmp!���벿��id
                rsChargeOffTmp!�շ�ϸĿid = rsTmp!�շ�ϸĿid
                rsChargeOffTmp!���� = rsTmp!����
                rsChargeOffTmp!����ʱ�� = Format(rsTmp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                rsChargeOffTmp.Update
            End If
            rsStuffSendData.MoveNext
        Loop
        
        rsTmp.MoveNext
    Loop
    
    If rsChargeOffTmp.RecordCount = 0 Then Exit Sub
    
    '4.��֯��������
    If mintUnit = 0 Then
        str��װ��λ = ",x.���㵥λ as ��λ,1 as ��װ "
    Else
        str��װ��λ = ",d.��װ��λ as ��λ,d.����ϵ�� as ��װ "
    End If
    
    gstrSQL = "Select Distinct '[' || x.���� || ']' || x.���� As ��������, c.Id As �շ�id, c.ҩƷid as ����id, c.����, c.No, c.��� As �շ����, c.����, c.����, c.Ч��," & vbNewLine & _
        "              f.����, p.���� As ��������, e.���� As ���ϲ���, e.Id As ���ϲ���id, a.����id, b.��� As �������, b.��¼����, b.��ҳid, b.����id, a.����ʱ��," & vbNewLine & _
        "              c.ʵ������ As ׼������, a.���� As ��������" & str��װ��λ & vbNewLine & _
        " From ���˷������� A, סԺ���ü�¼ B," & vbNewLine & _
        "     (Select a.Id, a.����, a.No, a.���, a.ҩƷid, a.����, a.����, a.Ч��, a.����id, b.ʵ������" & vbNewLine & _
        "       From ҩƷ�շ���¼ A," & vbNewLine & _
        "            (Select c.����, c.No, c.���, c.ҩƷid, Sum(Nvl(c.����, 1) * c.ʵ������) As ʵ������" & vbNewLine & _
        "              From ҩƷ�շ���¼ C, ���˷������� A, סԺ���ü�¼ B" & vbNewLine & _
        "              Where a.������� = 1 And a.״̬ = 0 And a.����id = b.Id And b.No = c.No And b.Id = c.����id And c.���� In (24, 25) And" & vbNewLine & _
        "                    c.������� Is Not Null And c.�ⷿid = a.��˲���id And c.�ⷿid = [1] And c.����id = [2] And a.���벿��id = [3] And a.����ʱ�� = [4] " & vbNewLine & _
        "              Group By c.����, c.No, c.���, c.ҩƷid" & vbNewLine & _
        "              Having Sum(Nvl(c.����, 1) * c.ʵ������) > 0) B" & vbNewLine & _
        "       Where a.No = b.No And a.���� = b.���� And a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And a.����� Is Not Null And" & vbNewLine & _
        "             (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)) C, �������� D, �շ���ĿĿ¼ X, ���ű� P, ������ҳ F, ���ű� E" & vbNewLine & _
        " Where a.������� = 1 And a.״̬ = 0 And a.����id = b.Id And b.No = c.No And b.Id = c.����id And b.��������id = p.Id And" & vbNewLine & _
        "      b.�շ�ϸĿid = d.����id And b.�շ�ϸĿid = x.Id And b.����id = f.����id And b.��ҳid = f.��ҳid And a.���벿��id = e.Id And" & vbNewLine & _
        "      f.��Ժ���� Is Null And b.ִ�в���id = [1] And a.����id = [2] And a.���벿��id = [3] And a.����ʱ�� = [4] " & vbNewLine & _
        "Order By a.����ʱ��, c.����, c.No, c.��� Desc"
    
    rsChargeOffTmp.MoveFirst
    Do While Not rsChargeOffTmp.EOF
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "GetChargeOffRecord", Val(mArrFilter("���ϲ���ID")), _
            Val(rsChargeOffTmp!����ID), Val(rsChargeOffTmp!���벿��id), CDate(rsChargeOffTmp!����ʱ��))
        
        Do While Not rsTmp.EOF
            With mrsChargeOff
                .AddNew
                !�������� = rsTmp!��������
                !���ϲ��� = rsTmp!���ϲ���
                !���ϲ���id = rsTmp!���ϲ���id
                !���� = rsTmp!����
                !NO = rsTmp!NO
                !����ID = rsTmp!����ID
                !����ʱ�� = Format(rsTmp!����ʱ��, "yyyy-mm-dd hh:mm:ss")
                !����ID = rsTmp!����ID
                !�շ���� = rsTmp!�շ����
                !���� = rsTmp!����
                !���� = rsTmp!����
                !Ч�� = rsTmp!Ч��
'                !׼������ = Format(rsTmp!׼������ / rsTmp!��װ, mFMT.FM_����)
'                !�������� = Format(rsTmp!�������� / rsTmp!��װ, mFMT.FM_����)
                !׼������ = rsTmp!׼������
                !�������� = rsTmp!��������
                !��װ = rsTmp!��װ
                !��λ = rsTmp!��λ
                !�շ�ID = rsTmp!�շ�ID
                !��ҳid = IIf(IsNull(rsTmp!��ҳid), 0, rsTmp!��ҳid)
                !������� = rsTmp!�������
                !���� = rsTmp!����
                !����ID = rsTmp!����ID
                !��¼���� = rsTmp!��¼����
                !��˱�־ = 0
                !ִ��״̬ = 0
    
                .Update
            End With
            
            rsTmp.MoveNext
        Loop
        
        rsChargeOffTmp.MoveNext
    Loop
    
    If mrsChargeOff.RecordCount = 0 Then Exit Sub
    
    '5.ͬһ����ID�ж���շ�ID�İ�׼�������������������з���
    lngChargeID = 0
    dblʣ���������� = 0
    str����ʱ�� = ""
    mrsChargeOff.Sort = "����id,����ʱ��,�շ���� Desc"
    mrsChargeOff.MoveFirst
    
    Do While Not mrsChargeOff.EOF
        If lngChargeID = mrsChargeOff!����ID And str����ʱ�� = mrsChargeOff!����ʱ�� Then
            '��ʾ�Ƕ�����Σ����ϸ�����ʣ������������������
            If dblʣ���������� > 0 Then
                If dblʣ���������� - mrsChargeOff!׼������ > 0 Then
                    '����ʣ�࣬����ֻ�ܰ�׼����������
                    mrsChargeOff!�������� = mrsChargeOff!׼������
                    dblʣ���������� = dblʣ���������� - mrsChargeOff!׼������
                Else
                    'û��ʣ�࣬��ʣ����������
                    mrsChargeOff!�������� = dblʣ����������
                End If
            Else
                '��ʾ�ϴη������������û��ʣ�࣬ʣ�µ���������������Ϊ0
                mrsChargeOff!�������� = 0
            End If
            mrsChargeOff.Update
        Else
            '������ID������ʱ�俪ʼ�µ�������������
            lngChargeID = mrsChargeOff!����ID
            str����ʱ�� = mrsChargeOff!����ʱ��
            
            dblʣ���������� = mrsChargeOff!�������� - mrsChargeOff!׼������
            If dblʣ���������� > 0 Then
                '��ʾ��ʣ�࣬����ֻ�ܰ�׼����������
                mrsChargeOff!�������� = mrsChargeOff!׼������
                mrsChargeOff.Update
            End If
        End If
         
        mrsChargeOff.MoveNext
    Loop
    
    '6.����ִ��״̬
    Call RefreshChargeOffStation(rsStuffSendData, mrsChargeOff)
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetMatStock(ByVal lng�ⷿID As Long)
    On Error GoTo ErrHandle
    gstrSQL = "Select �շ�ϸĿid From �շ�ִ�п��� Where ִ�п���id = [1] "
    Set mrsMatStock = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�洢�ⷿ", lng�ⷿID)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub InitVsGrid()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����ؼ�
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-12 10:27:06
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        '0-��ѡ,1-��ѡ,-1-����
        .ColData(.ColIndex("״̬")) = 1
        .ColData(.ColIndex("��������")) = 1
        .ColData(.ColIndex("���ݺ�")) = 1
        .ColData(.ColIndex("����")) = 1
        .ColData(.ColIndex("����")) = 1
    End With
End Sub

Private Function SaveChargeOffData(ByVal strDate As String) As Boolean
    '�������+����
    Dim i As Integer
    Dim strMCNO As String, arrMCRec As Variant, arrMCPar As Variant
    Dim bln�Ƿ������� As Boolean
    Dim str������� As String
    Dim cllPro As Collection
    Dim strAudit As String
    Dim strReturnInfo As String
    Dim strReserve As String
    
    If mbln����ʱ�������� = False Then
        SaveChargeOffData = True
        Exit Function
    End If

    Set cllPro = New Collection
    
    With mrsChargeOff
        .Filter = "ִ��״̬=1 And ��������>0 "
        If .RecordCount = 0 Then
            SaveChargeOffData = True
            Exit Function
        End If
        
        Do While Not .EOF
            '�ų��ظ������˼�¼
            If InStr("," & strAudit & ",", "," & !����ID & !����ʱ�� & ",") = 0 Then
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
                gstrSQL = gstrSQL & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS'),"
                '  ״̬_In     ���˷�������.״̬%Type,
                gstrSQL = gstrSQL & "1,"
                '  int�Զ����� Integer:=1
                gstrSQL = gstrSQL & "0)"
                AddArray cllPro, gstrSQL
            End If
            
            '���ϴ���
            'Zl_�����շ���¼_��������
            gstrSQL = "Zl_�����շ���¼_��������("
            '    �շ�id_In   In ҩƷ�շ���¼.ID%Type,
            gstrSQL = gstrSQL & "" & NVL(!�շ�ID) & ","
            '    �����_In   In ҩƷ�շ���¼.�����%Type,
            gstrSQL = gstrSQL & "'" & gstrUserName & "',"
            '    �������_In In ҩƷ�շ���¼.�������%Type,
            gstrSQL = gstrSQL & "to_date('" & strDate & "','yyyy-mm-dd HH24:mi:ss'),"
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
            gstrSQL = gstrSQL & "'" & gstrUserName & "')"
            
            AddArray cllPro, gstrSQL
            
            bln�Ƿ������� = True
            
            strReturnInfo = IIf(strReturnInfo = "", "", strReturnInfo & "|") & NVL(!�շ�ID) & "," & NVL(!��������)
            
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

    err = 0: On Error GoTo ErrHandRpt:
    If bln�Ƿ������� = True Then
      If zlStr.IsHavePrivs(mstrPrivs, "����֪ͨ��") Then
            If MsgBox("����Ҫ��ӡ�����嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_2", Me, "����ʱ��=" & strDate, "��λ=" & mintUnit + 1, 2)
            End If
     End If
    End If
    
    '������ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing And bln�Ƿ������� Then
        mobjPlugIn.DrugReturnByID Val(mArrFilter("���ϲ���id")), strReturnInfo, CDate(strDate), strReserve
    End If
    
    SaveChargeOffData = True
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    Exit Function
ErrHandRpt:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    SaveChargeOffData = True
End Function

Public Function zlFullData(ByVal frmMain As Form, ByVal strPrivs As String, ByVal lngModule As Long, ByVal intUnit As Integer, _
    ByVal arrFilter As Variant) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�����ص�δ���ϵ�����Ϣ
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
    Set mfrmMain = frmMain: mstrPrivs = strPrivs: mlngModule = lngModule
    Set mArrFilter = arrFilter
    mintUnit = intUnit
    mlngȱ�ϼ�� = Val(zlDatabase.GetPara("ȱ�ϼ��", glngSys, lngModule))
    mbln�����ݷ��� = (Val(zlDatabase.GetPara("�����ݺŷ���", glngSys, lngModule, "0")) = 1)
    mbln������ǩ�� = (Val(zlDatabase.GetPara("������ǩ��", glngSys, mlngModule, 0)) = 1)
    mbln����ʱ�������� = (Val(zlDatabase.GetPara("����ʱ�����������ʼ�¼", glngSys, mlngModule, 0)) = 1)
    mbln������ʱ����� = Val(zlDatabase.GetPara("����ҽ��������ʱ�����", glngSys, 1723, 0))
    
    '��ʼ�ؼ�����
    Call InitData
    '������ݸ�δ��������
    If RefreshData() = False Then Exit Function
    zlFullData = False
End Function
Public Function zlPayStuff() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:������������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-22 14:25:18
    '-----------------------------------------------------------------------------------------------------------
    Dim strDate As String
          
    mstrNo = ""
    If vsGrid.Row > 0 Then
        If vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("���ݺ�")) <> "" And vsGrid.IsSubtotal(vsGrid.Row) = False Then
            mstrNo = vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("���ݺ�"))
        End If
    End If
    
    If ISValied() = False Then Exit Function
    
    strDate = Format(sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    
    If SaveData(strDate) = False Then Exit Function
    
    If SaveChargeOffData(strDate) = False Then Exit Function
    
    zlPayStuff = True
End Function


Private Function GetNext���ܱ�ʶ��() As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ���ܱ�ʶ��
    '���:
    '����:
    '����:�ɹ�,���ر�ʶ��
    '����:���˺�
    '����:2008-04-23 14:20:49
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    GetNext���ܱ�ʶ�� = sys.GetNextNo(20)
    Exit Function
End Function
Private Function CheckBillStruct() As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ��������
    '���:
    '����:
    '����:�ɹ�,���ؿռ�¼���ṹ
    '����:���˺�
    '����:2008-04-23 14:41:41
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    
    With rsTemp
        If .State = 1 Then .Close
        .Fields.Append "���ݱ�ʶ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "�����־", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set CheckBillStruct = rsTemp
End Function

Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��鷢���Ƿ�Ϸ�
    '���:
    '����:
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 14:25:36
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Integer, lng����ID As Long, rsCheck As ADODB.Recordset
    Dim str��� As String
    Dim intCardCount As Integer '���η�����Ҫˢ������
    Dim intִ��״̬ As Integer
    
    ISValied = False
    
    '���������
    If Trim(txtEdit(mtxtIdx.idx_������).Tag) = "" Then
        MsgBox "������δ��������벻��ȷ�������������ˣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�ȳ�ʼ���
    Set rsCheck = CheckBillStruct
    
    '���ִ�пⷿ
    With mrsNotPayStuff
        .Filter = ""
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        
        If mrsMatStock Is Nothing Then
            GetMatStock Val(mArrFilter("���ϲ���ID"))
            
            If mrsMatStock.RecordCount = 0 Then
                MsgBox "δ���ô洢�ⷿ�����ܷ���,���飡", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        If mbln�����ݷ��� = True Then
            .Filter = "NO='" & mstrNo & "'"
        End If
        
        .Sort = "����ID Asc"
        Do While Not .EOF
            If lng����ID <> !����ID Then
                If !ִ��״̬ = 1 Then
                    mrsMatStock.Filter = "�շ�ϸĿid=" & Val(!����ID)
                    If mrsMatStock.EOF Then
                        MsgBox !�������� & "δ���ô洢�ⷿ�����ܷ���,���飡", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    lng����ID = !����ID
                Else
                    lng����ID = 0
                End If
            End If
            
            If !ִ��״̬ = 1 Then
                '��Ҫ���ṩ����ٶȣ��ȴ����ڲ����ݼ�
                rsCheck.Filter = "���ݱ�ʶ='" & NVL(!NO) & "|" & NVL(!����) & "'"
                If rsCheck.RecordCount <> 0 Then
                    rsCheck.Find "����ID=" & Val(NVL(!����ID))
                    If rsCheck.EOF Then rsCheck.AddNew
                Else
                    rsCheck.AddNew
                End If
                
                rsCheck!���ݱ�ʶ = NVL(!NO) & "|" & NVL(!����)
                rsCheck!����ID = Val(NVL(!����ID))
                rsCheck!��¼���� = Val(NVL(!��¼����))
                rsCheck!�����־ = Val(NVL(!�����־))
                str��� = NVL(rsCheck!���)
                If InStr(1, "," & str��� & ",", "," & Val(NVL(!���)) & ",") = 0 Then
                    If str��� = "" Then
                        str��� = Val(NVL(!���))
                    Else
                        str��� = str��� & "," & Val(NVL(!���))
                    End If
                    rsCheck!��� = str���
                End If
                rsCheck.Update
                rsCheck.Filter = 0
            End If
            
            .MoveNext
        Loop
    End With
    Dim strNo As String, lng���� As Long, lng����id As Long
    '��鵥��,��Ҫ�Ǽ�鴦���Ƿ��Ѿ�����,�����Ƿ��Ѿ���Ժ�����Ȩ�޽�����صļ��
    With rsCheck
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = !���ݱ�ʶ & "|"
            lng���� = Split(strNo, "|")(1)
            strNo = Split(strNo, "|")(0)
            lng����id = !����ID
            str��� = NVL(!���)
            
            '�����ʴ����Ƿ��ܷ���
            If Check���ʴ���(mstrPrivs, lng����, strNo, str���, Val(!��¼����), Val(!�����־)) = False Then Exit Function
            If Check��Ժ����(mstrPrivs, lng����, strNo, Val(!��¼����), Val(!�����־), lng����id) = False Then Exit Function
            .MoveNext
        Loop
    End With
    
    strNo = ""
    lng���� = 0
    lng����id = 0
    
    'һ��ͨ���Ѽ��
    If mbln����ǰ�շѻ���� = True Then
        With mrsNotPayStuff
            '���1�����������Ҫˢ������ôһ�η���ֻ��һ������ˢ������
            .Filter = "ִ��״̬=1 And ���շ�=0 And ����ID>0"
            .Sort = "����ID"
            Do While Not .EOF
                If lng����id = 0 Then
                    lng����id = !����ID
                End If
                If lng����id <> !����ID Then
                    MsgBox "��֧�ֶ�����˷���ʱ����ˢ�����ѡ����β��ܷ��ϣ����飡", vbInformation, gstrSysName
                    Exit Function
                End If
                .MoveNext
            Loop
            
            '���2������ʱ�����Ҫˢ���������������������Ƿ���״̬
            If lng����id > 0 Then
                .Filter = "���շ�=0 And ����ID=" & lng����id
                .Sort = "����,NO,ִ��״̬"
                Do While Not .EOF
                    If lng���� <> !���� And strNo <> !NO Then
                        lng���� = !����
                        strNo = !NO
                        intִ��״̬ = !ִ��״̬
                    ElseIf intִ��״̬ <> !ִ��״̬ Then
                        MsgBox "����ʱˢ�����ѱ���������������һ���ϡ����β��ܷ��ϣ����飡", vbInformation, gstrSysName
                        Exit Function
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    ISValied = True
End Function

Private Function CardConfirm(ByVal rsData As ADODB.Recordset) As Boolean
    '���ѿ�����ȷ�Ͻӿ�
    '������������ϣ����Ұ���������ˣ������˶�ε���ˢ�����ѽӿ�
    'ʵ����֮ǰ�ѽ���У�飬����������������Ҫˢ�����ѣ����ֹ���ϣ���������Ӧ�ò������������ˢ������
    '��ʱ�������ִ���ʽ�������Ժ��䶯
    Dim lngCard����ID As Long
    Dim strCardNo As String
        
    On Error GoTo ErrHand
    
    If mbln����ǰ�շѻ���� = False Then
        CardConfirm = True
        Exit Function
    End If
        
    'ע�⴫��ļ�¼���Ǵ�����ϸ
    '�շѵ���
     rsData.Filter = "ִ��״̬=1 And ��¼����=1 And ���շ�=0"
     rsData.Sort = "����ID,NO"
     Do While Not rsData.EOF
         If lngCard����ID <> rsData!����ID Then
             If strCardNo <> "" Then
                 'ˢ������
                If gobjSquareCard.zlSquareAffirm(Me, mlngModule, mstrPrivs, lngCard����ID, mfrmFilter.PatiCardID, False, 1, strCardNo) = False Then
                    Exit Function
                End If
             End If
             
             lngCard����ID = rsData!����ID
             strCardNo = rsData!NO
         Else
             If strCardNo = "" Then
                 strCardNo = rsData!NO
             ElseIf InStr(1, strCardNo, rsData!NO) = 0 Then
                 strCardNo = strCardNo & "," & rsData!NO
             End If
         End If
         rsData.MoveNext
     Loop
     
     If strCardNo <> "" Then
        'ˢ������
        If gobjSquareCard.zlSquareAffirm(Me, mlngModule, mstrPrivs, lngCard����ID, mfrmFilter.PatiCardID, False, 1, strCardNo) = False Then
            Exit Function
        End If
     End If
    
    lngCard����ID = 0
    strCardNo = ""
    
    '���˵��ݣ�ֻ�����ﲡ�˽��д���
    rsData.Filter = "ִ��״̬=1 And ��¼����=2 And ���շ�=0"
    rsData.Sort = "����ID,NO"
    Do While Not rsData.EOF
        If rsData!�����־ = 1 Or rsData!�����־ = 4 Then
            If lngCard����ID <> rsData!����ID Then
                If strCardNo <> "" Then
                    'ˢ������
                    If gobjSquareCard.zlSquareAffirm(Me, mlngModule, mstrPrivs, lngCard����ID, mfrmFilter.PatiCardID, False, 2, strCardNo) = False Then
                        Exit Function
                    End If
                    strCardNo = ""
                End If
                
                lngCard����ID = rsData!����ID
                strCardNo = rsData!NO
            Else
                If strCardNo = "" Then
                    strCardNo = rsData!NO
                ElseIf InStr(1, strCardNo, rsData!NO) = 0 Then
                    strCardNo = strCardNo & "," & rsData!NO
                End If
            End If
        End If
        rsData.MoveNext
    Loop
    If strCardNo <> "" Then
        'ˢ������
        If gobjSquareCard.zlSquareAffirm(Me, mlngModule, mstrPrivs, lngCard����ID, mfrmFilter.PatiCardID, False, 2, strCardNo) = False Then
            Exit Function
        End If
    End If
    
    CardConfirm = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    CardConfirm = False
End Function
Private Function SaveData(ByVal strDate As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ָ���ķ�����Ŀ���з��ϴ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 11:48:06
    '-----------------------------------------------------------------------------------------------------------
    Dim str������ As String, lng����id As Long, strID���� As String
    Dim cllPro As Collection
    Dim strReserve As String
        
    SaveData = False
    err = 0: On Error GoTo ErrHand:
    mstr���ܱ�ʶ�� = GetNext���ܱ�ʶ��()
   
    Set cllPro = New Collection
    With mrsNotPayStuff
        .Filter = ""
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        If .EOF Then Exit Function
        
        If mbln�����ݷ��� = True Then
            If mstrNo = "" Then Exit Function
            If MsgBox("������ȷ��Ҫ�Ե���[" & mstrNo & "]���з��ϲ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        Else
            If MsgBox("������ȷ��Ҫ���з��ϲ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        End If
        
        '�µ����ѿ�ˢ�����ѽӿ�
        If Not CardConfirm(mrsNotPayStuff) Then Exit Function
        
        '��ҩ��ǩ��
        str������ = ""
        If mbln������ǩ�� Then
            str������ = zlDatabase.UserIdentify(Me, "������ǩ��", glngSys, mlngModule, "")
            If str������ = "" Then
                Exit Function
            End If
        End If
        
        If mbln�����ݷ��� = True Then
            .Filter = "NO='" & mstrNo & "'"
        Else
            .Filter = ""
        End If
        
        '������ID������ID����
        .Sort = "����ID Asc ,����ID Asc"
        
        Do While Not .EOF
            If !ִ��״̬ = 1 Then
                            
                If lng����id = 0 Then
                    lng����id = !����ID
                End If
                '����ID��ͬʱ��
                If lng����id = !����ID Then
                    '���������ַ�������3950ʱ���ύ��������ַ���Ϊ4000��
                    If zlCommFun.ActualLen(strID����) > 3990 Then
                        'Zl_ҩƷ�շ���¼_��������
                        gstrSQL = "Zl_ҩƷ�շ���¼_��������("
                        '    �շ�id_In     In Varchar2, --��ʽ:"id1,����1|id2,����2|....."
                        gstrSQL = gstrSQL & "'" & strID���� & "',"
                        '    �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                        gstrSQL = gstrSQL & "" & Val(mArrFilter("���ϲ���id")) & ","
                        '    �����_In     In ҩƷ�շ���¼.�����%Type,
                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                        '    �������_In   In ҩƷ�շ���¼.�������%Type,
                        gstrSQL = gstrSQL & "To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'),"
                        '    ���Ϸ�ʽ_In   In ҩƷ�շ���¼.��ҩ��ʽ%Type := 3, --1-��������;2-��������;3-���ŷ���;-1 ֹͣ����
                        gstrSQL = gstrSQL & "3,"
                        '    ������_In     In ҩƷ�շ���¼.������%Type := Null,
                        gstrSQL = gstrSQL & "'" & str������ & "',"
                        '    ���ϱ�ʶ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null,
                        gstrSQL = gstrSQL & "" & Val(mstr���ܱ�ʶ��) & ","
                        '    ������_In     In ҩƷ�շ���¼.��ҩ��%Type := Null
                        gstrSQL = gstrSQL & "'" & txtEdit(mtxtIdx.idx_������).Text & "',"
                        '    ����Ա����
                        gstrSQL = gstrSQL & "'" & UserInfo.��� & "')"
                        Call AddArray(cllPro, gstrSQL)
                        lng����id = 0
                        strID���� = !Id & "," & NVL(!����, 0)
                    Else
                        strID���� = IIf(strID���� = "", !Id & "," & NVL(!����, 0), strID���� & "|" & !Id & "," & NVL(!����, 0))
                    End If
                Else
                    '�������ID��ͬ���ύ����
                    'Zl_ҩƷ�շ���¼_��������
                    gstrSQL = "Zl_ҩƷ�շ���¼_��������("
                    '    �շ�id_In     In Varchar2, --��ʽ:"id1,����1|id2,����2|....."
                    gstrSQL = gstrSQL & "'" & strID���� & "',"
                    '    �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                    gstrSQL = gstrSQL & "" & Val(mArrFilter("���ϲ���id")) & ","
                    '    �����_In     In ҩƷ�շ���¼.�����%Type,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '    �������_In   In ҩƷ�շ���¼.�������%Type,
                    gstrSQL = gstrSQL & "To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'),"
                    '    ���Ϸ�ʽ_In   In ҩƷ�շ���¼.��ҩ��ʽ%Type := 3, --1-��������;2-��������;3-���ŷ���;-1 ֹͣ����
                    gstrSQL = gstrSQL & "3,"
                    '    ������_In     In ҩƷ�շ���¼.������%Type := Null,
                    gstrSQL = gstrSQL & "'" & str������ & "',"
                    '    ���ϱ�ʶ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null,
                    gstrSQL = gstrSQL & "" & mstr���ܱ�ʶ�� & ","
                    '    ������_In     In ҩƷ�շ���¼.��ҩ��%Type := Null
                    gstrSQL = gstrSQL & "'" & txtEdit(mtxtIdx.idx_������).Text & "',"
                    '    ����Ա����
                    gstrSQL = gstrSQL & "'" & UserInfo.��� & "')"
                    Call AddArray(cllPro, gstrSQL)
                    lng����id = !����ID
                    strID���� = !Id & "," & NVL(!����, 0)
                End If
            End If
            .MoveNext
            
            '�������û�м�¼���Ҵ����ַ�����Ϊ�գ����ύ����
            If .EOF And strID���� <> "" Then
                    'Zl_ҩƷ�շ���¼_��������
                    gstrSQL = "Zl_ҩƷ�շ���¼_��������("
                    '    �շ�id_In     In Varchar2, --��ʽ:"id1,����1|id2,����2|....."
                    gstrSQL = gstrSQL & "'" & strID���� & "',"
                    '    �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
                    gstrSQL = gstrSQL & "" & Val(mArrFilter("���ϲ���id")) & ","
                    '    �����_In     In ҩƷ�շ���¼.�����%Type,
                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                    '    �������_In   In ҩƷ�շ���¼.�������%Type,
                    gstrSQL = gstrSQL & "To_Date('" & strDate & "','yyyy-MM-dd hh24:mi:ss'),"
                    '    ���Ϸ�ʽ_In   In ҩƷ�շ���¼.��ҩ��ʽ%Type := 3, --1-��������;2-��������;3-���ŷ���;-1 ֹͣ����
                    gstrSQL = gstrSQL & "3,"
                    '    ������_In     In ҩƷ�շ���¼.������%Type := Null,
                    gstrSQL = gstrSQL & "'" & str������ & "',"
                    '    ���ϱ�ʶ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null,
                    gstrSQL = gstrSQL & "" & mstr���ܱ�ʶ�� & ","
                    '    ������_In     In ҩƷ�շ���¼.��ҩ��%Type := Null
                    gstrSQL = gstrSQL & "'" & txtEdit(mtxtIdx.idx_������).Text & "',"
                    '    ����Ա����
                    gstrSQL = gstrSQL & "'" & UserInfo.��� & "')"
                    Call AddArray(cllPro, gstrSQL)
            End If
        Loop
    End With
        
    On Error GoTo ErrExcute:
    Call ExecuteProcedureArrAy(cllPro, Me.Caption)
    SaveData = True
    err = 0: On Error GoTo ErrHand:
    Call BillListPrint(strDate)
    mstrNo = ""
    
    '���÷�ҩ�����ҽӿ�
    If Not mobjPlugIn Is Nothing Then
        mobjPlugIn.StuffSendBySumID Val(mArrFilter("���ϲ���id")), mstr���ܱ�ʶ��, strReserve
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Function
ErrExcute:
      gcnOracle.RollbackTrans
      If ErrCenter = 1 Then Resume
      Call SaveErrLog
End Function

Private Sub SetExecuteStaut(ByVal lngRow As Long)
    '-----------------------------------------------------------------------------------------------------------
    '����:����ִ��״̬
    '���:lngRow-ָ������
    '����:
    '����:
    '����:���˺�
    '����:2008-04-23 11:31:04
    '-----------------------------------------------------------------------------------------------------------
    Dim str״̬ As String, int״̬ As Integer, lngλ�� As Long
    With vsGrid
        str״̬ = Trim(.TextMatrix(lngRow, .ColIndex("״̬")))
        int״̬ = Decode(str״̬, "ȱ��", 0, "����", 1, "�ܷ�", 2, "������", 3, 4)
        lngλ�� = Val(.Cell(flexcpData, lngRow, .ColIndex("���ݺ�")))
    End With
    
    With mrsNotPayStuff
         .Filter = 0
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        .Find "λ��=" & lngλ��
        If .EOF = False Then
            !ִ��״̬ = int״̬:
            !״̬ = str״̬
            .Update
            Call CheckStock(Val(NVL(!����ID)))
        End If
        
        '���ܿ�������Ҫ����ǰ��״̬
        .MoveFirst
        .Find "λ��=" & lngλ��
        If .EOF = False Then
             vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("״̬")) = Decode(NVL(!ִ��״̬), 0, "ȱ��", 1, "����", 2, "�ܷ�", "������")
        End If
        .MoveFirst
        .Find "ִ��״̬=1"
        mblnHave���� = (.EOF = False)
        .MoveFirst
        .Find "ִ��״̬=2"
        mblnHave�ܷ� = (.EOF = False)
    End With
    
    Call RefreshChargeOffStation(mrsNotPayStuff, mrsChargeOff)
End Sub
  
Private Sub cmdSel_Click()
   Call SelectItem(txtEdit(mtxtIdx.idx_������), "")
End Sub

Private Sub Form_Load()
    '���˺�:����С����ʽ����
     
    zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "δ����"
    
    mstrĬ�ϵ��ݸ�ʽ = Trim(zlDatabase.GetPara("���ϵ��ݴ�ӡ��ʽ", glngSys, mlngModule, , Array(cboEdit(mcboIdx.idx_���ݸ�ʽ)), zlStr.IsHavePrivs(mstrPrivs, "��������")))

    mbln����ǰ�շѻ���� = zlDatabase.GetPara("��Ŀִ��ǰ�������շѻ��ȼ������", glngSys)
    mbln����δ�շѵ����ﻮ�۴������� = Val(zlDatabase.GetPara("����δ�շѵ����ﻮ�۴�������", glngSys))
    mbln����δ��˵ļ��˴������� = Val(zlDatabase.GetPara("����δ��˵ļ��˴�������", glngSys))
            
    Call InitVsGrid
    vsGrid.RowHeightMin = 300
    
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
    
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    With txtEdit(mtxtIdx.idx_������)
        .Left = lblEdit(mlblIdx.idx_lbl������).Left + lblEdit(mlblIdx.idx_lbl������).Width
        cmdSel.Left = .Left + .Width - cmdSel.Width - 10
    End With
    With cboEdit(mcboIdx.idx_���ݸ�ʽ)
        .Top = ScaleHeight - .Height - 50
        .Left = ScaleWidth - .Width - 50
        lblEdit(mlblIdx.idx_lbl���ݸ�ʽ).Left = .Left - lblEdit(mlblIdx.idx_lbl���ݸ�ʽ).Width - 10
        lblEdit(mlblIdx.idx_lbl���ݸ�ʽ).Top = .Top + (.Height - lblEdit(mlblIdx.idx_lbl���ݸ�ʽ).Height) \ 2
        lblEdit(mlblIdx.idx_lbl������).Top = lblEdit(mlblIdx.idx_lbl���ݸ�ʽ).Top
        txtEdit(mtxtIdx.idx_������).Top = .Top
        cmdSel.Top = .Top + (txtEdit(mtxtIdx.idx_������).Height - cmdSel.Height) \ 2
    End With
    
    With vsGrid
        .Top = ScaleTop
        .Width = ScaleWidth
        .Left = ScaleLeft
        .Height = cboEdit(mcboIdx.idx_���ݸ�ʽ).Top - .Top - 50
    End With
End Sub

Private Function InitRsStruct() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ���ڲ���¼��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 09:54:46
    '-----------------------------------------------------------------------------------------------------------
    Set mrsNotPayStuff = New ADODB.Recordset
    With mrsNotPayStuff
        If .State = 1 Then .Close
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ҽ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ҽ������", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "״̬", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 8, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 10, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "סԺ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ϵ��", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���÷���", adDouble, 2, adFldIsNullable
        .Fields.Append "��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����Ա", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "������λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "Ƶ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�÷�", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "���շ�", adDouble, 2, adFldIsNullable
        .Fields.Append "�Ƿ���", adDouble, 2, adFldIsNullable
        .Fields.Append "λ��", adDouble, 18, adFldIsNullable
        .Fields.Append "ִ��״̬", adDouble, 1, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable            '�жϿ����
        .Fields.Append "��λ", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "˵��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�����", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "ҽ��id", adDouble, 18, adFldIsNullable
        .Fields.Append "������", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "�ⷿ��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "�����־", adDouble, 18, adFldIsNullable
    
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set mrsChargeOff = New ADODB.Recordset
    With mrsChargeOff
        If .State = 1 Then .Close
        .Fields.Append "���ϲ���", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "���ϲ���ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "NO", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "����ʱ��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "�շ����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "Ч��", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "׼������", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adDouble, 18, adFldIsNullable
        .Fields.Append "��װ", adDouble, 18, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "�շ�ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��ҳID", adDouble, 18, adFldIsNullable
        .Fields.Append "�������", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��¼����", adDouble, 18, adFldIsNullable
        .Fields.Append "��˱�־", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ִ��״̬", adDouble, 2, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
        
    InitRsStruct = True
End Function
Private Function WhiteDataToRecord(ByVal rsSource As ADODB.Recordset) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ص�����д���ڲ���¼��(δ���ϲ���)
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 10:03:41
    '-----------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:

    With rsSource
        If rsSource.RecordCount <> 0 Then rsSource.MoveFirst
        Do While Not .EOF
            mrsNotPayStuff.AddNew
            mrsNotPayStuff!Id = !Id
            mrsNotPayStuff!״̬ = "����"    'ȫ��Ĭ��Ϊ����
            mrsNotPayStuff!���� = !����
            mrsNotPayStuff!����ҽ�� = !����ҽ��
            mrsNotPayStuff!���� = Decode(NVL(!����), 24, "�շѵ�", 25, "���ʵ�", 26, "���ʱ�", "��֪") & IIf(!���շ� = 0, "(δ)", "")
            mrsNotPayStuff!����ID = !����ID
            mrsNotPayStuff!λ�� = .AbsolutePosition
            mrsNotPayStuff!NO = !NO
            mrsNotPayStuff!���� = !����
            mrsNotPayStuff!����ID = Val(NVL(!����ID))
            mrsNotPayStuff!��� = !���
            mrsNotPayStuff!���� = !����
            mrsNotPayStuff!���� = NVL(!����)
            mrsNotPayStuff!סԺ�� = IIf(Val(NVL(!�����־)) = 2, NVL(!��ʶ��), "")
            mrsNotPayStuff!�������� = NVL(!��������)
            mrsNotPayStuff!��� = NVL(!���)
            mrsNotPayStuff!���� = NVL(!����)
            mrsNotPayStuff!���� = Val(NVL(!����))
            mrsNotPayStuff!���� = NVL(!����)
            mrsNotPayStuff!����ϵ�� = Val(NVL(!����ϵ��))
            mrsNotPayStuff!���÷��� = Val(NVL(!����))
            mrsNotPayStuff!�Ƿ��� = Val(NVL(!�Ƿ���))
            mrsNotPayStuff!�� = IIf(Val(NVL(!��)) = 0, 1, Val(NVL(!��)))
            mrsNotPayStuff!ʵ������ = IIf(Val(NVL(!����)) = 0, 1, Val(NVL(!����)))
            mrsNotPayStuff!��λ = !��λ
            mrsNotPayStuff!���� = Format(IIf(Val(NVL(!����)) = 0, 1, Val(NVL(!����))) / !����ϵ��, mFMT.FM_����) & !��λ
            mrsNotPayStuff!���� = !����
            mrsNotPayStuff!��� = !���
            mrsNotPayStuff!����Ա = NVL(!����Ա����)
            mrsNotPayStuff!���� = IIf(IsNull(!����), "", zlStr.FormatEx(!����, 5) & NVL(!���㵥λ))
            mrsNotPayStuff!������λ = NVL(!���㵥λ)
            mrsNotPayStuff!Ƶ�� = IIf(IsNull(!Ƶ��), "", !Ƶ��)
            mrsNotPayStuff!�÷� = IIf(IsNull(!�÷�), "", !�÷�)
            mrsNotPayStuff!˵�� = IIf(IsNull(!˵��), "", !˵��)
            mrsNotPayStuff!����ID = Val(NVL(!����ID))
            mrsNotPayStuff!��¼���� = Val(NVL(!��¼����))
            mrsNotPayStuff!�����־ = Val(NVL(!�����־))
            mrsNotPayStuff!ҽ������ = IIf(IsNull(!ҽ������), "", !ҽ������)
            If IsNull(!�Ǽ�ʱ��) Then
                mrsNotPayStuff!����ʱ�� = ""
            Else
                mrsNotPayStuff!����ʱ�� = Format(!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss")
            End If
            
            mrsNotPayStuff!������ = NVL(!������)
            mrsNotPayStuff!����� = NVL(!�����)
            mrsNotPayStuff!���շ� = !���շ�                          'δ�շѻ���ʴ�������������
            mrsNotPayStuff!ҽ��id = !ҽ��id
            If mbln��ʾ�������� = True Then
                mrsNotPayStuff!������ = !������         '��Ҫ�����һ�ε�������
            Else
                mrsNotPayStuff!������ = ""
            End If
            
            mrsNotPayStuff!�ⷿ��λ = IIf(IsNull(!�ⷿ��λ), "", !�ⷿ��λ)
            mrsNotPayStuff!����id = IIf(IsNull(!����id), 0, !����id)
            mrsNotPayStuff!������� = !�������
            
            '����Ƿ������� :0-ȱ��,1-����,2-�ܷ�,3-������
            mrsNotPayStuff!ִ��״̬ = 1     'ȱʡΪ����
            If mrsNotPayStuff!���շ� = 0 And mbln����ǰ�շѻ���� = False And mbln����δ�շѵ����ﻮ�۴������� = False And mrsNotPayStuff!���� = 24 Then
                mrsNotPayStuff!ִ��״̬ = 3   'δ�շѵģ�������
                mrsNotPayStuff!״̬ = "������"    'ȫ��Ĭ��Ϊ����
            End If
            
            '���˵���Ǿܷ���������ò����Ѿܷ���ͬʱ������ִ��״̬
            If NVL(!˵��) = "�ܷ�" Then mrsNotPayStuff!ִ��״̬ = 2
'            If mbln����δ��˴������� = False Then
                If NVL(mrsNotPayStuff!�����) = "" And mbln����ǰ�շѻ���� = False And mbln����δ��˵ļ��˴������� = False And mrsNotPayStuff!���� = 25 Then
                    mrsNotPayStuff!ִ��״̬ = 3  'δ��ģ���ʾΪ������
                    mrsNotPayStuff!״̬ = "������"    'ȫ��Ĭ��Ϊ����
'                End If
            Else
                'ȫ��Ϊ����
                mrsNotPayStuff!ִ��״̬ = 1
            End If
            mrsNotPayStuff.Update
            If err <> 0 Then GoTo ErrHand
            .MoveNext
        Loop
    End With
   
    WhiteDataToRecord = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function FullDataToVsGrid() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ص�������䵽ָ��������ؼ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-23 11:06:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    FullDataToVsGrid = False
    
    err = 0: On Error GoTo ErrHand:

    '������ݵ��ؼ���
    mrsNotPayStuff.Filter = 0
    If mrsNotPayStuff.RecordCount <> 0 Then mrsNotPayStuff.MoveFirst
    
    mblnHave���� = False
    mblnHave�ܷ� = False
    
    With vsGrid
        .Clear (1)
        If mrsNotPayStuff.EOF Then '
            .Rows = 2
            FullDataToVsGrid = True
            Exit Function
        End If
        .Subtotal flexSTClear

        .Rows = mrsNotPayStuff.RecordCount + .FixedRows
        lngRow = .FixedRows
        Do While Not mrsNotPayStuff.EOF
            .RowData(lngRow) = Val(mrsNotPayStuff!Id)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
            .Cell(flexcpData, lngRow, .ColIndex("���ݺ�")) = Val(NVL(mrsNotPayStuff!λ��))
            .TextMatrix(lngRow, .ColIndex("����ҽ��")) = NVL(mrsNotPayStuff!����ҽ��)
            .TextMatrix(lngRow, .ColIndex("ҽ������")) = NVL(mrsNotPayStuff!ҽ������)
            .TextMatrix(lngRow, .ColIndex("״̬")) = IIf(zlStr.IsHavePrivs(mstrPrivs, "�������Ϸ���") Or zlStr.IsHavePrivs(mstrPrivs, "�������Ͼܷ�"), NVL(mrsNotPayStuff!״̬), "")
            '24-�շѴ������ϣ�25-���ʵ��������ϣ�26-���ʱ������ϣ�
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!����)
            .TextMatrix(lngRow, .ColIndex("���ݺ�")) = NVL(mrsNotPayStuff!NO)
            .TextMatrix(lngRow, .ColIndex("����Ա")) = NVL(mrsNotPayStuff!����Ա)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!����)
            .TextMatrix(lngRow, .ColIndex("סԺ��")) = NVL(mrsNotPayStuff!סԺ��)
            .TextMatrix(lngRow, .ColIndex("��������")) = NVL(mrsNotPayStuff!��������)
            .TextMatrix(lngRow, .ColIndex("���")) = NVL(mrsNotPayStuff!���)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(NVL(mrsNotPayStuff!��)), "###")
            .TextMatrix(lngRow, .ColIndex("����")) = NVL(mrsNotPayStuff!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(NVL(mrsNotPayStuff!����)) * mrsNotPayStuff!����ϵ��, mFMT.FM_���ۼ�)
            .TextMatrix(lngRow, .ColIndex("���")) = Format(Val(NVL(mrsNotPayStuff!���)), mFMT.FM_���)
            .TextMatrix(lngRow, .ColIndex("˵��")) = NVL(mrsNotPayStuff!˵��)
            .TextMatrix(lngRow, .ColIndex("����ʱ��")) = NVL(mrsNotPayStuff!����ʱ��)
            If mrsNotPayStuff!ִ��״̬ = 1 Then mblnHave���� = True
            If mrsNotPayStuff!ִ��״̬ = 2 Then mblnHave�ܷ� = True
            
            lngRow = lngRow + 1
            mrsNotPayStuff.MoveNext
         Loop
    End With
    
    '�����ݽ��л���
    If SetTotalRowData = False Then Exit Function
    FullDataToVsGrid = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function RefreshData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ˢ��δ���ϵ�������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-22 09:59:36
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, strWhere As String, strFields As String
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    Dim str�������� As String
    Dim strסԺ As String
    Dim strSqlTmp As String
    
    On Error GoTo ErrHandle

    str�������� = zlDatabase.GetPara("�������Ϸ�ʽ", glngSys, mlngModule, "�ٴ�,����,���,����,����,����,Ӫ��")
    
    If mintUnit = 0 Then
        strFields = "x.���㵥λ as ��λ,1 as ����ϵ��,"
    Else
        strFields = "d.��װ��λ as ��λ,d.����ϵ��,"
    End If
    
    gstrSQL = "" & _
        "      Select Distinct s.Id, s.ҩƷid AS ����ID, Nvl(n.���շ�, 0) ���շ�, p.���� ����, s.��ҩ�� AS ������ ,S.����ID, c.������ ����ҽ��, " & _
        "          c.����Ա���� �����, s.����, Nvl(s.����, 0) ����, s.No, s.���, nvl(c.����id,0) as ����ID, '' ����, c.����, " & _
        "          c.��ʶ��, c.����Ա����, '[' || x.���� || ']' || x.���� ��������, s.���� ��, s.ʵ������ ����, " & _
        "          Nvl(d.���÷���, 0) ����, x.���, c.�Ǽ�ʱ��," & strFields & _
        "          s.���ۼ� ����, s.���۽�� ���, s.����, s.Ƶ��, s.�÷�, s.ժҪ ˵��, " & _
        "          Decode(s.����, Null, '', s.����) || Decode(s.����, Null, '', 0, '', '(' || s.���� || ')') ����, " & _
        "          Nvl(s.����, 0) ����, c.ҽ�����, i.���㵥λ, Nvl(s.����, Nvl(x.����, '')) ����, " & _
        "          Nvl(m.�����, -1) �����, Nvl(c.ҽ�����, -1) ҽ��id, '' �ⷿ��λ,x.�Ƿ���, m.���id,m.ҽ������, " & _
        "          s.�Է�����id As ����id, c.��� �������, C.��¼����,C.�����־,0 �������, z.���� As ������ " & _
        "       From δ��ҩƷ��¼ n,ҩƷ�շ���¼ s, ������ü�¼ c,������Ϣ c1, ����ҽ����¼ m,   " & _
        "          ���ű� p, �������� d, �շ���ĿĿ¼ x, �շ���Ŀ���� e,������ĿĿ¼ i, ������Ŀ���� z " & _
        "       Where n.���� = s.���� And  n.No = s.No AND nvl(n.�ⷿid,[1])+0=nvl(s.�ⷿid,[1])  " & _
        "             And s.����id = c.Id AND s.�Է�����id + 0 = p.Id  " & _
        "             And s.ҩƷid = d.����id And S.ҩƷid = x.Id  " & _
        "             And s.ҩƷid = e.�շ�ϸĿid(+)  And e.����(+) = 3 " & _
        "             And Nvl(Ltrim(Rtrim(s.ժҪ)), 'NOT�ܷ�') <> '�ܷ�'  AND s.����� Is Null And Nvl(s.��ҩ��ʽ, 0) <> -1 " & _
        "             And Mod(s.��¼״̬, 3) = 1 And instr([4],','||s.����||',')>0 " & _
        "             AND d.����ID=i.id  and C.����ID=c1.����ID(+) and C.����ID=c1.����ID(+) " & _
        "             AND D.����id = z.������Ŀid(+) And z.����(+) = 2    " & _
        "             AND c.ҽ����� = m.Id(+)  And Nvl(c.����״̬,0)<>1 " & _
        "             And Nvl(n.�ⷿid, [1]) + 0 = [1]  " & _
        "             And n.�������� Between [2] And  [3]" & _
        "             "
    
    '�ų���δ��ҩƷ�����ʼ�¼
    gstrSQL = gstrSQL & " And Not Exists (Select 1 From ���˷������� X " & _
        " Where X.������� = 0 And X.״̬+0 = 0 And X.�շ�ϸĿid+0 = S.ҩƷid And X.����id = S.����id) "
    
    '�շѴ�����ʾ��ʽ
    If Val(mArrFilter("�շѴ���")) = 1 Then
        gstrSQL = gstrSQL & " And n.���շ�=1 "
    ElseIf Val(mArrFilter("�շѴ���")) = 2 Then
        gstrSQL = gstrSQL & " And n.���շ�=0 "
    End If
        
    If Trim(mArrFilter("��������ID")) <> "" Then
        Select Case Val(mArrFilter("��������"))
        Case 0  '�ٴ�
            gstrSQL = gstrSQL & " And Instr([5], ',' || C.��������id || ',') > 0 And C.���˿���id=C.��������id"
        Case 1 'ҽ��
            gstrSQL = gstrSQL & " And Instr([5], ',' || C.��������id || ',') > 0 And C.���˿���id<>C.��������id"
        Case Else
            '����
            If str�������� = "" Then
                gstrSQL = gstrSQL & " And Instr([5], ',' || C.���˲���ID || ',') > 0 And C.���˿���id=C.��������id"
            Else
                gstrSQL = gstrSQL & " And Instr([5], ',' || C.���˲���ID || ',') > 0 "
                If str�������� <> mstrAllType Then
                    gstrSQL = gstrSQL & " And C.��������id Not In (Select Distinct ����id From ��������˵�� " & _
                        " Where Instr([13],',' || �������� || ',') > 0) "
                End If
            End If
        End Select
    End If
    
    strWhere = ""
    If (Trim(mArrFilter("���ݺ�")(0)) <> "" And Trim(mArrFilter("���ݺ�")(1)) = "") Then
        strWhere = strWhere & "            AND s.NO =[6]  "
    ElseIf (Trim(mArrFilter("���ݺ�")(1)) <> "" And Trim(mArrFilter("���ݺ�")(0)) = "") Then
        strWhere = strWhere & "            AND s.NO =[7]  "
    ElseIf Trim(mArrFilter("���ݺ�")(0)) <> "" And Trim(mArrFilter("���ݺ�")(1)) <> "" Then
        strWhere = strWhere & "            AND ( s.NO between [6] and [7] )"
    End If
    
    gstrSQL = gstrSQL & strWhere
    
    gstrSQL = gstrSQL & IIf(Val(mArrFilter("����ID")) = 0 And Val(mArrFilter("IC����")) = 0, "", "       AND c.����iD=[8]  ")
    gstrSQL = gstrSQL & IIf(Val(mArrFilter("סԺ��")) = 0, "", "       AND c.��ʶ��=[9] and c.�����־=2 ")
    gstrSQL = gstrSQL & IIf(Trim(mArrFilter("����")) = "", "", "       AND C.���� like [10] ")
    gstrSQL = gstrSQL & IIf(Val(mArrFilter("�����")) = 0, "", "       AND c.��ʶ��=[11] and c.�����־=1 ")
    gstrSQL = gstrSQL & IIf(Trim(mArrFilter("���￨��")) = "", "", "   AND c1.���￨�� =[12] ")
    
    If mbln������ʱ����� = True Then
        strSqlTmp = Replace(gstrSQL, "n.��������", "c.����ʱ��") & " And C.ҽ����� Is Not Null"
        gstrSQL = gstrSQL & " And C.ҽ����� Is Null"
        gstrSQL = gstrSQL & " Union All " & strSqlTmp
    End If
    
    If mbln��ʾ�������� Then
        gstrSQL = " Select a.*, b.��ҩ�� as ������ " & _
                 "  From ( " & gstrSQL & _
                 "          Order By s.No, s.���� " & _
                 "       ) a, " & _
                 "      (Select a.����, a.No, a.���, a.������ ��ҩ�� " & _
                 "       From ҩƷ�շ���¼ a, " & _
                 "          (   Select s.����, s.No, s.���, Max(s.��¼״̬) ��¼״̬ " & _
                 "              From ҩƷ�շ���¼ s, δ��ҩƷ��¼ n " & _
                 "              Where s.No = n.No And s.���� = n.���� And Nvl(s.��ҩ��ʽ, 0) <> -1 " & _
                 "                     And Nvl(s.�ⷿid, [1]) + 0 = Nvl(n.�ⷿid, [1])  " & _
                 "                     And Nvl(s.�ⷿid, [1]) + 0 = [1]  " & _
                 "                     And n.�������� Between [2] And [3]  " & _
                 "                     And Mod(s.��¼״̬, 3) = 2 And instr([4],','||s.����||',')>0 " & strWhere & _
                 "              Group By s.����, s.No, s.��� " & _
                 "            ) b " & _
                 "       Where a.���� = b.���� And a.No = b.No And a.��� = b.��� And a.��¼״̬ = b.��¼״̬) b " & _
                 "  Where a.���� = b.����(+) And a.No = b.No(+) And a.��� = b.���(+) "
    End If
    
    If Val(mArrFilter("��������")) = 0 Then
        '����
        str���� = Replace(gstrSQL, "C.���˲���ID", "C.��������id")
        strסԺ = Replace(gstrSQL, "'' ����", "c.����")
        strסԺ = Replace(strסԺ, "C.����", "nvl(r.����,C.����)")
        strסԺ = Replace(strסԺ, "c.����", "nvl(r.����,c.����) ����")
        strסԺ = Replace(strסԺ, "������ü�¼ c", "סԺ���ü�¼ c,������ҳ r")
        strסԺ = Replace(strסԺ, "And Nvl(c.����״̬,0)<>1", " and r.����id=c.����id and r.��ҳid=c.��ҳid " & IIf(Trim(mArrFilter("����")) = "", "", "   AND c.���� =[14] "))
        If Trim(mArrFilter("����")) <> "" Then str���� = str���� & " and 1=0"
        gstrSQL = str���� & " Union All " & strסԺ
    ElseIf Val(mArrFilter("��������")) = 1 Then
        gstrSQL = Replace(gstrSQL, "C.���˲���ID", "C.��������id")
    ElseIf Val(mArrFilter("��������")) = 2 Then
        'סԺ���ʵ�
        gstrSQL = Replace(gstrSQL, "'' ����", "c.����")
        gstrSQL = Replace(gstrSQL, "C.����", "nvl(r.����,C.����)")
        gstrSQL = Replace(gstrSQL, "c.����", "nvl(r.����,c.����) ����")
        gstrSQL = Replace(gstrSQL, "������ü�¼ c", "סԺ���ü�¼ c,������ҳ r")
        gstrSQL = Replace(gstrSQL, "And Nvl(c.����״̬,0)<>1", " and r.����id=c.����id and r.��ҳid=c.��ҳid " & IIf(Trim(mArrFilter("����")) = "", "", "   AND c.���� =[14] "))
    End If
    
    gstrSQL = gstrSQL & "  Order By No, �������"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
        Val(mArrFilter("���ϲ���ID")), _
        CDate(mArrFilter("���ڷ�Χ")(0)), _
        CDate(mArrFilter("���ڷ�Χ")(1)), _
        CStr("," & mArrFilter("����") & ","), _
        "," & Trim(mArrFilter("��������ID")) & ",", _
        CStr(mArrFilter("���ݺ�")(0)), _
        CStr(mArrFilter("���ݺ�")(1)), _
        Val(mArrFilter("����ID")), _
        Val(mArrFilter("סԺ��")), _
        CStr(mArrFilter("����")), _
        Val(mArrFilter("�����")), _
        CStr(mArrFilter("���￨��")), _
        "," & str�������� & ",", Val(mArrFilter("����")))
    
    '��ʼ�����ݽṹ
    Call InitRsStruct
    '�����ص����ݵ��ڲ���¼��
    Call WhiteDataToRecord(rsTemp)
    '������Ƿ����
    Call CheckStock
    '��֯�������ݼ�
    Call GetChargeOffRecord(mrsNotPayStuff)
    
    RaiseEvent zlRefreshDataRecordSet(mrsNotPayStuff, mrsChargeOff)
    
    Call FullDataToVsGrid
    RefreshData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetTotalRowData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�����еĻ�������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-22 10:22:21
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, lngCol As Long, strUpper As String
    Dim lngOldRow As Long, lngOldTopRow As Long
    With vsGrid
        .Redraw = flexRDNone
         .OutlineCol = 0
        .SubtotalPosition = flexSTAbove
        .Subtotal flexSTSum, -1, .ColIndex("���"), "###.00", , vbBlue, True, "�ϼ�"
        .Subtotal flexSTSum, .ColIndex("��������"), .ColIndex("���"), "###.00", , vbBlue, True
        .Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("���"), "###.00", , vbBlue, True
        '.Subtotal flexSTSum, .ColIndex("���ݺ�"), .ColIndex("����"), "###.00", , , True, "����С��"
        .Editable = flexEDKbdMouse
        If .Rows > 2 Then
            .Cell(flexcpBackColor, 1, .ColIndex("״̬"), .Rows - 1, .ColIndex("״̬")) = &HE7CFBA
            .Cell(flexcpBackColor, 1, .ColIndex("����"), .Rows - 1, .ColIndex("����")) = &HE7CFBA
        End If
        
        For lngRow = 1 To .Rows - 1
            If .IsSubtotal(lngRow) = True Then
                If InStr(1, .TextMatrix(lngRow, .ColIndex("��������")), "Total") > 0 Then
                    .TextMatrix(lngRow, .ColIndex("��������")) = Replace(.TextMatrix(lngRow, .ColIndex("��������")), "Total", "")
                    If Trim(.TextMatrix(lngRow, .ColIndex("��������"))) <> "" Then
                        .TextMatrix(lngRow, .ColIndex("��������")) = Trim(.TextMatrix(lngRow, .ColIndex("��������"))) & "(С��)"
                    End If
                End If
                
                .TextMatrix(lngRow, .ColIndex("���ݺ�")) = Replace(.TextMatrix(lngRow, .ColIndex("���ݺ�")), "Total", "")
                If Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) <> "" Then
                    .TextMatrix(lngRow, .ColIndex("���ݺ�")) = Trim(.TextMatrix(lngRow, .ColIndex("���ݺ�"))) & "(С��)"
                End If

            End If
        Next
        
        '���е��ݺϲ�
'        .MergeCells = flexMergeRestrictRows
'        For lngRow = 1 To .Rows - 1
'            .MergeRow(lngRow) = False
'            If .IsSubtotal(lngRow) = True Then
'                .MergeRow(lngRow) = True
'                strUpper = Trim(.TextMatrix(lngRow, .ColIndex("���"))) & " (��д:" & zlCommFun.UppeMoney(Val(.TextMatrix(lngRow, .ColIndex("���")))) & ")"
'                For lngCol = .ColIndex("���ݺ�") To .Cols - 1
'                    .TextMatrix(lngRow, lngCol) = strUpper
'                Next
'            End If
'        Next
        .OutlineBar = 1

        .Redraw = flexRDBuffered
    End With
End Function

Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "δ����"
    Call zlDatabase.SetPara("���ϵ��ݴ�ӡ��ʽ", cboEdit(mcboIdx.idx_���ݸ�ʽ), glngSys, mlngModule)
End Sub
Private Sub txtEDIT_Change(Index As Integer)
    txtEdit(Index).Tag = ""
    
End Sub

Private Sub txtEDIT_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    OS.OpenIme True
End Sub

Private Sub txtEDIT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Index <> mtxtIdx.idx_������ Then Exit Sub
    If Trim(txtEdit(Index).Text) = "" Then Exit Sub
    If txtEdit(Index).Tag <> "" Then OS.PressKey vbKeyTab
    Call SelectItem(txtEdit(Index), Trim(txtEdit(Index)))
End Sub

Private Sub txtEDIT_LostFocus(Index As Integer)
    OS.OpenIme False
End Sub

Private Sub vsGrid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngRow As Long, str״̬ As String
    With vsGrid
        Select Case Col
        Case .ColIndex("״̬")
            Call ChangeSelStaut(Row)
        End Select
    End With
End Sub
Private Sub ChangeSelStaut(ByVal Row As Long)
    '-----------------------------------------------------------------------------------------------------------
    '����:�ı�ָ���е�״̬
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-05-01 23:58:23
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, str״̬ As String
    Dim lng���� As Long
    
    With vsGrid
            If .IsSubtotal(Row) Then
               str״̬ = Trim(.TextMatrix(Row, .ColIndex("״̬")))
               lng���� = .RowOutlineLevel(Row)
               For lngRow = Row + 1 To .Rows - 1
                   If .RowOutlineLevel(lngRow) <> lng���� Then
                       If .TextMatrix(lngRow, .ColIndex("״̬")) <> "ȱ��" Then
                            .TextMatrix(lngRow, .ColIndex("״̬")) = str״̬
                            '������ص�ִ��״̬
                            Call SetExecuteStaut(lngRow)
                            
                       End If
                   Else
                        Exit For
                   End If
               Next
            Else
                '������ص�ִ��״̬
                Call SetExecuteStaut(Row)
            End If
    End With
End Sub
 

Private Sub vsGrid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        Select Case .Col
        Case .ColIndex("״̬")
                  If zlStr.IsHavePrivs(mstrPrivs, "�������Ϸ���") = False And zlStr.IsHavePrivs(mstrPrivs, "�������Ͼܷ�") = False Then
                        Cancel = True
                        Exit Sub
                  End If

                If .TextMatrix(Row, Col) = "ȱ��" Then
                    Cancel = True
                End If
'                If Row = 1 And .IsSubtotal(Row) = True Then
'                    Cancel = True
'                End If
'                If .IsSubtotal(Row) = True And InStr(1, Trim(.TextMatrix(Row, .ColIndex("��������"))), "С��") > 0 Then
'                    Cancel = True
'                End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub
Private Function InitCheckStock() As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ�������ļ�¼��
    '���:
    '����:
    '����:���ؼ�¼��
    '����:���˺�
    '����:2008-04-22 20:47:33
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Set rsTemp = New ADODB.Recordset
    With rsTemp
        If .State = 1 Then .Close
        .Fields.Append "����ID", adDouble, 18
        .Fields.Append "����", adDouble, 18
        .Fields.Append "���", adDouble, 18
        .Fields.Append "����", adDouble, 18
        .Fields.Append "���", adDouble, 5
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    Set InitCheckStock = rsTemp
End Function
Private Function CheckStock(Optional lng����ID As Long = 0) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�����ÿ��,ȷ���Ƿ����ȱ��
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-04-22 20:38:29
    '-----------------------------------------------------------------------------------------------------------
    Dim rsStock As ADODB.Recordset
    Dim lngRow As Long, lng��� As Long
    Dim arrtemp As Variant
    
    Set rsStock = InitCheckStock
    With mrsNotPayStuff
        '�����:
        '   1.���ڿ��ܴ�����ͬ�����κͲ��ϣ���ˣ���Ҫ�𲽼���ÿ�ʿ�棬��������ȷ�������ĵĿ�������Ƿ�����
        '   2.
        If .RecordCount <> 0 Then .MoveFirst
        If .RecordCount = 0 Then Exit Function
        
        If lng����ID <> 0 Then
            '���ܶ�ĳ�ֲ��Ͻ��м��
            .Filter = "����ID=" & lng����ID
            If .RecordCount = 0 Then .Filter = 0: Exit Function
        End If
        Do While Not .EOF
            If !ִ��״̬ <= 1 Then
                'ֻ���ȱ�Ϻͷ����������
                If LocaleStockData(rsStock, Val(mArrFilter("���ϲ���ID")), Val(NVL(!����ID)), Val(NVL(!����)), lng���) = True Then
                   '�ҵ���ָ���Ŀ��:��Ҫ��������Ƿ���㣬��������Ҫ�����жϣ������Ƿ��ϲ��ŵĿ������
                        If Val(NVL(rsStock!����)) - Val(NVL(mrsNotPayStuff!ʵ������)) < 0 Then
                            '��������,��ȷ��Ϊȱ��:�������λ��۵�,����Ϊ����棩
                            If mrsNotPayStuff!���շ� = 0 And mbln����ǰ�շѻ���� = False And mbln����δ�շѵ����ﻮ�۴������� = False And mrsNotPayStuff!���� = 24 Then
                                !ִ��״̬ = 3
                            ElseIf NVL(mrsNotPayStuff!�����) = "" And mbln����ǰ�շѻ���� = False And mbln����δ��˵ļ��˴������� = False And mrsNotPayStuff!���� = 25 Then
                                !ִ��״̬ = 3
                            Else
                                !ִ��״̬ = IIf(mlngȱ�ϼ�� = 1 Or rsStock!���� <> 0 Or rsStock!��� = 1, 0, !ִ��״̬)
                            End If
                            
                            .Update
                        Else
                            '��������:
                            !ִ��״̬ = 1   'ȱʡΪ����
                            If mrsNotPayStuff!���շ� = 0 And mbln����ǰ�շѻ���� = False And mbln����δ�շѵ����ﻮ�۴������� = False And mrsNotPayStuff!���� = 24 Then !ִ��״̬ = 3      'δ�շѵģ�ǿ��Ϊ������
                            If NVL(mrsNotPayStuff!�����) = "" And mbln����ǰ�շѻ���� = False And mbln����δ��˵ļ��˴������� = False And mrsNotPayStuff!���� = 25 Then !ִ��״̬ = 3 'δ��˵ĵ���,Ĭ��Ϊ������
                            .Update
                        End If
                        If !ִ��״̬ = 1 Then
                            '�������,��Ҫ���Ŀɿ����
                            With rsStock
                                !���� = Val(NVL(!����)) - Val(NVL(mrsNotPayStuff!ʵ������))
                                .Update
                            End With
                        End If
                End If
                
                If !ִ��״̬ = 0 Then
                    If CheckIsStockUp(Val(mArrFilter("���ϲ���ID")), Val(NVL(!����ID)), Val(NVL(!����)), Val(NVL(!����ID)), Val(NVL(mrsNotPayStuff!ʵ������))) = True Then
                        !ִ��״̬ = 1
                    End If
                End If
                
                !״̬ = Decode(!ִ��״̬, 0, "ȱ��", 1, "����", 2, "�ܷ�", "������")
                .Update
            End If
            .MoveNext
        Loop
     End With
     mrsNotPayStuff.Filter = 0
End Function

Private Function LocaleStockData(ByRef rsStock As ADODB.Recordset, _
    ByVal lng���ϲ���ID As Long, ByVal lng����ID As Long, ByVal lng���� As Long, Optional ByRef lng��� As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ָ������ָ�����Ŀ���Ƿ����
    '���:rsStock-ָ�����Ŀ������(����Ϊ�ռ�¼),�����Զ���չ
    '     lng���ϲ���ID-���ϲ���id
    '     lng����id-����id
    '     lng����-����
    '
    '����:lng���-���ؿ������
    '����:�ɹ�,��ʾ�ҵ�,�����ʾδ�ҵ�
    '����:���˺�
    '����:2008-04-22 21:07:47
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim dbl��� As Double
    LocaleStockData = False
    
    err = 0: On Error GoTo ErrHand:
    With rsStock
        .Filter = "����ID=" & lng����ID & " and ����=" & lng����
        If .RecordCount = 0 Then
            .Filter = 0: lng��� = .RecordCount + 1
            
            gstrSQL = "" & _
            " Select nvl(F.�Ƿ���,0) ���,nvl(A.ʵ������,0) ����" & _
            " From �������� B,�շ���ĿĿ¼ F," & _
            "      (Select A.ҩƷid as ����ID,a.ʵ������ From ҩƷ��� A Where ����=1 And �ⷿID=[1] And ҩƷID=[2] And nvl(����,0)=[3]) A" & _
            " Where B.����ID=F.ID And B.����ID=A.����ID(+) And B.����ID=[2] "
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng���ϲ���ID, lng����ID, lng����)
            
            dbl��� = Val(NVL(rsTemp!����))
            .AddNew
            !����ID = lng����ID
            !���� = lng����
            !��� = rsTemp!���
            !���� = dbl���
            !��� = lng���
            .Update
        Else
            lng��� = Val(NVL(!���))
            .Filter = 0
        End If
        .MoveFirst
        .Find "���=" & lng���
        LocaleStockData = True
    End With
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub InitData()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ���ؼ�����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-04-23 16:29:05
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, intDefault As Integer
    
    On Error GoTo ErrHandle
    If mbln�����ݷ��� = True Then
        vsGrid.ColHidden(vsGrid.ColIndex("״̬")) = True
    Else
        vsGrid.ColHidden(vsGrid.ColIndex("״̬")) = False
    End If
    
    intDefault = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\zl9Report\LocalSet\ZL1_BILL_1723_1", "Format", 1))
    txtEdit(mtxtIdx.idx_������).Text = gstrUserName: txtEdit(mtxtIdx.idx_������).Tag = gstrUserName
    '������ش�ӡ����
    If cboEdit(mcboIdx.idx_���ݸ�ʽ).ListCount <> 0 Then Exit Sub
    mstrĬ�ϵ��ݸ�ʽ = Trim(zlDatabase.GetPara("���ϵ��ݴ�ӡ��ʽ", glngSys, mlngModule, , Array(cboEdit(mcboIdx.idx_���ݸ�ʽ)), zlStr.IsHavePrivs(mstrPrivs, "��������")))
    gstrSQL = "Select  ���,˵�� From zltools.zlRPTFMTs Where ����id = (Select ID From zltools.zlReports Where ��� = 'ZL1_BILL_1723_1') Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���ݴ�ӡ��ʽ")
    With cboEdit(mcboIdx.idx_���ݸ�ʽ)
        Do While Not rsTemp.EOF
            .AddItem rsTemp!˵��
            If mstrĬ�ϵ��ݸ�ʽ <> "" Then
                If NVL(rsTemp!˵��) = mstrĬ�ϵ��ݸ�ʽ Then
                    .ListIndex = .NewIndex
                End If
            ElseIf Val(NVL(rsTemp!���)) = intDefault Then
                    .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If .ListCount <> 0 And .ListIndex < 0 Then .ListIndex = 0
        If rsTemp.RecordCount = 1 Then .Enabled = False
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Function SelectItem(ByVal objCtl As Control, ByVal strKey As String) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, strTittle As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    strTittle = "������ѡ��"
    
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    gstrSQL = "" & _
        "   Select distinct a.��� as ����,A.���� As ����,����" & _
        "   From ��Ա�� A,������Ա B,��������˵�� C,��Ա����˵�� D " & _
        "   Where A.Id=B.��Աid And B.����id=C.����Id And D.��Աid=A.Id " & _
        "       And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) AND B.����id in (Select ����ID From ������Ա where ��Աid=[2] ) "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "    And  ((A.����) like [1] or  A.���  like [1] or  ����  like  upper([1]))  " & _
        "    "
    End If
    gstrSQL = "Select rownum as ID,a.* from (" & gstrSQL & ") A" & _
        "   ORDER BY ���� "
     
     strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = zlControl.GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey, UserInfo.Id)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgBox "û���ҵ���������������,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
            .Cell(flexcpData, .Row, .Col) = NVL(rsTemp!����)
        End With
    Else
        If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = NVL(rsTemp!����)
        objCtl.Tag = NVL(rsTemp!����)
        OS.PressKey vbKeyTab
    End If
    SelectItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
 Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function

Private Sub vsGrid_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    With vsGrid
        If Position <= .ColIndex("���ݺ�") Then
            ShowMsgBox "���ܽ����ƶ������ݺ���ǰ����!"
            Position = Col
        End If
    End With
End Sub

Private Sub vsGrid_DblClick()
    Dim str״̬ As String
    If zlStr.IsHavePrivs(mstrPrivs, "�������Ϸ���") = False And zlStr.IsHavePrivs(mstrPrivs, "�������Ͼܷ�") = False Then Exit Sub
    With vsGrid
'         If .Row = 1 And .IsSubtotal(.Row) = True Then
'            Exit Sub
'        End If
         
         If mbln�����ݷ��� = True Then Exit Sub
         
         str״̬ = Trim(.TextMatrix(.Row, .ColIndex("״̬")))
         
        .TextMatrix(.Row, .ColIndex("״̬")) = Decode(str״̬, "����", "�ܷ�", "�ܷ�", "������", "ȱ��", "ȱ��", "����")
        
        If .IsSubtotal(.Row) = True Then
            Call ChangeSelStaut(.Row)
            Exit Sub
        End If
     
        Call ChangeSelStaut(.Row)
    End With
End Sub
Public Property Get zlHaveSel����() As Boolean
    zlHaveSel���� = mblnHave����
End Property
Public Property Get zlHaveSel�ܷ�() As Boolean
    zlHaveSel�ܷ� = mblnHave�ܷ�
End Property
Public Property Get zlHaveData() As Boolean
    zlHaveData = mrsNotPayStuff.RecordCount <> 0
End Property

Private Sub BillListPrint(Optional strDate As String = "", Optional IntStyle As Integer = 0)
    '-----------------------------------------------------------------------------------------------------------
    '����:���ݻ�����ӡ
    '���:
    '     intStyle:0-�����Ϸ�ʽ��ӡ,1-���ݴ�ӡ,2-���ϵ���
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-05 10:36:44
    '-----------------------------------------------------------------------------------------------------------
    Dim strReg As String
    Dim bln���ϵ� As Boolean
    Dim bln�ѷ����嵥 As Boolean
    Dim bln���ݴ�ӡ As Boolean
    Dim intMsg As Integer   '0-��ʾ��ӡ,1-�Զ���ӡ,2-����ӡ
    
    
    intMsg = Val(zlDatabase.GetPara("���ϴ�ӡ���ѷ�ʽ", glngSys, mlngModule, "0"))
    
    bln���ϵ� = zlStr.IsHavePrivs(mstrPrivs, "����֪ͨ��")
    bln�ѷ����嵥 = zlStr.IsHavePrivs(mstrPrivs, "��ӡ�ѷ����嵥")
    bln���ݴ�ӡ = zlStr.IsHavePrivs(mstrPrivs, "���ݴ�ӡ")
    If intMsg = 0 Then
        '��ʾ��ӡ
        If bln�ѷ����嵥 = False Then Exit Sub
        If MsgBox("����Ҫ��ӡ��ص�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    ElseIf intMsg = 1 Then
        '�Զ���ӡ
    Else
        Exit Sub
    End If
        '���ŷ���
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1723_1", Me, _
       "�ⷿ=" & Val(mArrFilter("���ϲ���ID")), _
       "���Ϸ�ʽ=���ŷ���|3", _
       "��������=" & Val(mArrFilter("��������")), _
       "���տ���=" & ��ȡ���ղ�������(strDate), _
       "��λ=" & IIf(mintUnit = 0, 0, 1), _
       "���Ϻ�=" & strDate, _
       "���ܷ��Ϻ�=" & Val(mstr���ܱ�ʶ��), _
       "ReportFormat=" & IIf(cboEdit(mcboIdx.idx_���ݸ�ʽ).ListIndex = -1, 1, cboEdit(mcboIdx.idx_���ݸ�ʽ).ListIndex + 1), "PrintEmpty=0", 2)
End Sub
Private Function ��ȡ���ղ�������(ByVal strDate As String) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ���ղ��ŵĴ�ӡ����
    '���:
    '����:
    '����:�ɹ�,���� ��ʾ|IN(����ID,..) ,���򷵻�""
    '����:���˺�
    '����:2008-05-05 13:31:28
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, str��ʾ As String, strIDIn As String
    
    On Error GoTo ErrHandle
    If mArrFilter("��������id") = "" Then
        'û������,���Ը���ѡ�������ȡ��ʾ����
        gstrSQL = "Select distinct D.ID,D.����,D.���� as ���� " & _
                 " From ҩƷ�շ���¼ S,������ü�¼ C,���ű� d " & _
                 " Where S.����ID=C.ID And Mod(S.��¼״̬,3) In (0,1) And S.����� Is Not Null " & _
                 "      And C.ִ��״̬=1 And S.�ⷿID=[1] And S.��ҩ��ʽ=3 And S.�������=[2] " & _
                 "      And S.���� In (24,25,26) "
        Select Case Val(mArrFilter("��������"))
            Case 0  '
                gstrSQL = gstrSQL & " and C.���˿���id=d.id(+) "
            Case 1 'ҽ��
                gstrSQL = gstrSQL & "  and C.��������id =d.id(+)"
            Case Else '����
                gstrSQL = gstrSQL & "  and C.���˲���ID =d.id(+)"
        End Select
        
        Select Case Val(mArrFilter("��������"))
            Case 0, 1
                gstrSQL = gstrSQL & " Union All " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
            Case Else '����
                gstrSQL = Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
        End Select
        
        gstrSQL = gstrSQL & "order by ����"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mArrFilter("���ϲ���id")), CDate(strDate))
        With rsTemp
            Do While Not .EOF
                str��ʾ = str��ʾ & "," & !����
                strIDIn = strIDIn & "," & !Id
                rsTemp.MoveNext
            Loop
        End With
        strIDIn = "0" & strIDIn
        str��ʾ = str��ʾ & "|" & " IN (" & strIDIn & ")"
        ��ȡ���ղ������� = str��ʾ
        Exit Function
    End If
    gstrSQL = "Select ID, ���� From ���ű� A, Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)) J Where ID = J.Column_Value order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(mArrFilter("��������id")))
    With rsTemp
        Do While Not .EOF
            str��ʾ = str��ʾ & "," & !����
            rsTemp.MoveNext
        Loop
    End With
    str��ʾ = str��ʾ & "|" & " IN (0" & CStr(mArrFilter("��������id")) & ")"
    ��ȡ���ղ������� = str��ʾ
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub zlSetFontSize(ByVal curFontSize As Currency)
    '-----------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-05-06 17:00:44
    '-----------------------------------------------------------------------------------------------------------
    With vsGrid
        .Font.Size = curFontSize
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .Rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("��") + 120
        .RowHeightMax = TextHeight("��") + 120
        .Refresh
    End With
    lblEdit(mlblIdx.idx_lbl������).Font.Size = curFontSize
    lblEdit(mlblIdx.idx_lbl������).AutoSize = True
    lblEdit(mlblIdx.idx_lbl���ݸ�ʽ).FontSize = curFontSize
    lblEdit(mlblIdx.idx_lbl���ݸ�ʽ).AutoSize = True
    cboEdit(mcboIdx.idx_���ݸ�ʽ).Font.Size = curFontSize
    txtEdit(mtxtIdx.idx_������).Font.Size = curFontSize
    Call Form_Resize
End Sub
Public Property Get zl_�ϴλ��ܷ��Ϻ�() As String
    '�����ϴλ��ܷ��Ϻ�
    zl_�ϴλ��ܷ��Ϻ� = mstr���ܱ�ʶ��
End Property

 
Private Sub vsGrid_EnterCell()
    Dim lngRow As Long
    Dim strNo As String
    Dim lng����id As Long
    
    If mbln�����ݷ��� = True Then
        mblnHave���� = False
        If vsGrid.Row > 0 Then
            If vsGrid.IsSubtotal(vsGrid.Row) = False Then
                If vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("״̬")) <> "����" Then Exit Sub
                strNo = vsGrid.TextMatrix(vsGrid.Row, vsGrid.ColIndex("���ݺ�"))
                For lngRow = 1 To vsGrid.Rows - 1
                    If vsGrid.IsSubtotal(lngRow) = False Then
                        If strNo = vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("���ݺ�")) Then
                            If vsGrid.TextMatrix(lngRow, vsGrid.ColIndex("״̬")) <> "����" Then
                                mblnHave���� = False
                                Exit For
                            End If
                        End If
                    End If
                Next
                
                mblnHave���� = True
                
                If mbln����ǰ�շѻ���� = True Then
                    With mrsNotPayStuff
                        .Filter = ""
                        If Not .EOF Then
                            'Ĭ��ȫ��ִ��״̬Ϊ"������"
                            Do While Not .EOF
                                !ִ��״̬ = 0
                                .MoveNext
                            Loop
                            
                            'Ѱ�ҵ�ǰѡ��Ĳ���
                            .Filter = "NO = '" & strNo & "'"
                            lng����id = !����ID
                            
                            .Filter = "����id =" & lng����id
                            Do While Not .EOF
                                !ִ��״̬ = 1
                                .MoveNext
                            Loop
                            
                            .Filter = ""
                        End If
                    End With
                End If
            End If
        End If
    End If
End Sub


