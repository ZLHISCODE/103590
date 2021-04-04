VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.UserControl usrInSpecResult 
   BackColor       =   &H80000005&
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   LockControls    =   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   8130
   Begin VB.PictureBox PicItem 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   8130
      TabIndex        =   7
      Top             =   0
      Width           =   8130
      Begin VB.CommandButton cmdP1 
         Caption         =   "&P"
         Height          =   300
         Left            =   4230
         TabIndex        =   6
         ToolTipText     =   "ѡ��걾"
         Top             =   -15
         Width           =   315
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   1290
         TabIndex        =   5
         Top             =   0
         Width           =   2955
      End
      Begin VB.CommandButton cmdP 
         Caption         =   "��"
         Height          =   285
         Left            =   6825
         TabIndex        =   10
         ToolTipText     =   "ѡ��걾"
         Top             =   0
         Width           =   300
      End
      Begin VB.Label lblBBCaption 
         BackStyle       =   0  'Transparent
         Caption         =   "�걾(&B)"
         Height          =   180
         Left            =   4725
         TabIndex        =   8
         Top             =   45
         Width           =   630
      End
      Begin VB.Label lblBB 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�걾����"
         Height          =   180
         Left            =   5445
         TabIndex        =   9
         Top             =   45
         Width           =   720
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ(&C)"
         Height          =   180
         Left            =   270
         TabIndex        =   4
         Top             =   45
         Width           =   990
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   8130
         Y1              =   315
         Y2              =   315
      End
   End
   Begin VB.ListBox listCell 
      Height          =   1110
      Left            =   765
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   570
      Width           =   3075
   End
   Begin VB.ComboBox CmbCell 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   300
      Width           =   810
   End
   Begin VB.TextBox txtCell 
      Height          =   330
      Left            =   960
      TabIndex        =   1
      Top             =   345
      Width           =   810
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfMain 
      Height          =   1530
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   2699
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   11
      FixedCols       =   0
      BackColorSel    =   -2147483639
      BackColorBkg    =   -2147483628
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483631
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      BorderStyle     =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
End
Attribute VB_Name = "usrInSpecResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const LAWLChar = "';`|,"""
Private i As Long, j As Long
Private strSQL As String
Private rsTmp As New ADODB.Recordset

Private Enum EnmCTLType
    CTLTxt = 0  '-�ı�
    CTLUpDown = 1 '-����,
    CTLDownList = 2 '-����,
    CTLCheck = 3 '-��ѡ,
    CTLOption = 4   '-��ѡ
End Enum
Private Enum EnmValType
    ValNumber = 0   '��ֵ��
    ValText = 1     '�ı���
    ValDate = 2     '������
End Enum

Private Enum EnmGridCol
    ItemID = 0
    Item���� = 1
    Item��ʾ�� = 2
    Item��ֵ�� = 3
    Item��ʼֵ = 4
    Item���� = 5
    ItemС���� = 6
    Item�к� = 7
    Itemָ���� = 8
    ItemӢ�� = 9
    Item����ֵ = 10
    Item�������� = 11
    Item��λ = 12
    Item���� = 13
End Enum

Private mDispMode As Boolean
Private mReturnErrnumber As Long
Private mReturnErrDescription As String
Private mID������Ŀ As Long
Private mblnCancel As Boolean
Private mlng����id As Long
Private mShowItem As Boolean

Private mblnLawless As Boolean
Private mblnFirst As Boolean

Private mItemIndex As Long

Private Function zlGetSymbol(StrInput As String, Optional bytIsWB As Byte) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & StrInput & "') from dual"
    Else
        strSQL = "select zlSpellcode('" & StrInput & "') from dual"
    End If
    On Error GoTo ErrHand
    With rsTmp
        If .State = adStateOpen Then .Close
        Call SQLTest(App.ProductName, "mdlCISBase", strSQL)
        rsTmp.Open strSQL, gcnOracle, adOpenKeyset
        Call SQLTest
        zlGetSymbol = IIf(IsNull(.Fields(0).Value), "", .Fields(0).Value)
    End With
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Property Get ShowItem() As Boolean
'��ʾ������Ŀ,�����û������Լ��ı������Ŀ
    ShowItem = mShowItem
End Property

Public Property Let ShowItem(ByVal New_ShowItem As Boolean)
    mShowItem = New_ShowItem
    If mShowItem = True Then
        PicItem.Height = 510
        PicItem.Visible = True
    Else
        PicItem.Height = 0
        PicItem.Visible = False
    End If
    UserControl_Resize
    PropertyChanged "ShowItem"
End Property

Public Property Get ID������Ŀ() As Long
    ID������Ŀ = mID������Ŀ
End Property

Public Property Let ID������Ŀ(ByVal New_ID������Ŀ As Long)
'����������Ŀ
On Error GoTo ErrHandle
Dim lngWidth As Long
Dim lngWidth��λ As Long
Dim lngWidth�к� As Long

Dim rs������Ŀ As New ADODB.Recordset

    txtCell.Visible = False
    CmbCell.Visible = False
    listCell.Visible = False
    lngWidth = 800
    
    '��ʼ���ؼ�
    mblnCancel = True
    InitMe
    mblnCancel = False
    mID������Ŀ = New_ID������Ŀ
    PropertyChanged "ID������Ŀ"
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Property
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Property
        
    strSQL = _
        "SELECT C.ID," & vbCrLf & _
        "       A.������� ���," & vbCrLf & _
        "       A.����걾," & vbCrLf & _
        "       C.����," & vbCrLf & _
        "       C.��ʾ��," & vbCrLf & _
        "       C.��ֵ��," & vbCrLf & _
        "       C.��ʼֵ," & vbCrLf & _
        "       C.����," & vbCrLf & _
        "       C.С��," & vbCrLf & _
        "       C.������ ָ����," & vbCrLf & _
        "       C.Ӣ���� Ӣ����," & vbCrLf & _
        "       C.��λ" & vbCrLf & _
        "  FROM ������ĿĿ¼ B, ���鱨����Ŀ A,����������Ŀ C" & vbCrLf & _
        " WHERE B.ID IN (SELECT DISTINCT ������ĿID FROM ���鱨����Ŀ) AND  A.������Ŀid=C.Id AND " & vbCrLf & _
        "      B.�걾��λ = A.����걾 AND B.ID = A.������ĿID  AND A.������ĿID =" & New_ID������Ŀ & vbCrLf & _
        " ORDER BY A.�������"
    If rsTmp.State = adStateOpen Then rsTmp.Close
    Set rsTmp = Nothing
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "����������")
    '����������Ŀ�ļ���ָ��
    If rsTmp.RecordCount > 0 Then
        mblnCancel = True
        rsTmp.MoveFirst
        '����
        msfMain.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            msfMain.TextMatrix(i, EnmGridCol.ItemID) = zlCommFun.Nvl(rsTmp!ID, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item����) = zlCommFun.Nvl(rsTmp!����, 1)
            msfMain.TextMatrix(i, EnmGridCol.Item��ʾ��) = zlCommFun.Nvl(rsTmp!��ʾ��, 0)
            If zlCommFun.Nvl(rsTmp!����, 1) = 0 Then
                If Trim(zlCommFun.Nvl(rsTmp!��ֵ��)) <> ";" Then
                    msfMain.TextMatrix(i, EnmGridCol.Item����ֵ) = Replace(Trim(zlCommFun.Nvl(rsTmp!��ֵ��)), ";", " �� ")
                End If
            End If
            msfMain.TextMatrix(i, EnmGridCol.Item��ֵ��) = Trim(zlCommFun.Nvl(rsTmp!��ֵ��))
            msfMain.TextMatrix(i, EnmGridCol.Item��ʼֵ) = Trim(zlCommFun.Nvl(rsTmp!��ʼֵ))
            msfMain.TextMatrix(i, EnmGridCol.Item����) = zlCommFun.Nvl(rsTmp!����, 0)
            msfMain.TextMatrix(i, EnmGridCol.ItemС����) = zlCommFun.Nvl(rsTmp!С��, 0)
            msfMain.TextMatrix(i, EnmGridCol.Itemָ����) = Trim(zlCommFun.Nvl(rsTmp!ָ����)) & IIf(Trim(zlCommFun.Nvl(rsTmp!Ӣ����)) = "", "", "��" & Trim(zlCommFun.Nvl(rsTmp!Ӣ����)) & "��")
            msfMain.TextMatrix(i, EnmGridCol.ItemӢ��) = Trim(zlCommFun.Nvl(rsTmp!Ӣ����))
            msfMain.TextMatrix(i, EnmGridCol.Item��������) = Trim(zlCommFun.Nvl(rsTmp!��ʼֵ))
            msfMain.TextMatrix(i, EnmGridCol.Item��λ) = Trim(zlCommFun.Nvl(rsTmp!��λ))
            If i = 1 Then lblBB.Caption = zlCommFun.Nvl(rsTmp!����걾)
            rsTmp.MoveNext
        Next
        strSQL = "select * from ������ĿĿ¼  where id=" & mID������Ŀ
        Call zlDatabase.OpenRecordset(rs������Ŀ, strSQL, "����������")
        If rs������Ŀ.RecordCount > 0 Then
            txtItem.Text = zlCommFun.Nvl(rs������Ŀ!����)
            txtItem.SelStart = Len(txtItem.Text)
            txtItem.Tag = zlCommFun.Nvl(rs������Ŀ!����)
            cmdP1.Tag = rs������Ŀ!ID
        Else
            txtItem.Text = ""
            txtItem.Tag = ""
            cmdP1.Tag = 0
        End If
        ReSetRowCode msfMain
        PicItem_Resize
        If rs������Ŀ.RecordCount > 0 Then
            On Error Resume Next
            If msfMain.Enabled And msfMain.Visible Then
                msfMain.SetFocus
            End If
        End If
        mblnCancel = False
    Else
        lblBB.Caption = ""
        txtItem.Text = ""
        txtItem.Tag = ""
    End If
    UserControl_Resize
    mblnCancel = False
    Exit Property
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Property

Private Sub ReadData(lng����ID As Long)
'����:����ָ��ID��ȡָ��
On Error GoTo ErrHandle
Dim rs������Ŀ As New ADODB.Recordset
Dim rs������Ŀ1 As New ADODB.Recordset
Dim strItemName As String   '��������������Ŀ����
Dim lngWidth As Long

    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Sub
    
    txtCell.Visible = False
    CmbCell.Visible = False
    listCell.Visible = False
    '�ȸ��ݲ���ID�����������û�������ٸ�����ĿID����ָ��.
    strSQL = _
        "SELECT A.������ID ID,A.�ؼ��� ���, " & vbCrLf & _
        "           B.����, " & vbCrLf & _
        "           B.��ʾ��, " & vbCrLf & _
        "           B.��ֵ��, " & vbCrLf & _
        "           B.��ʼֵ, " & vbCrLf & _
        "           B.����, " & vbCrLf & _
        "           B.С��, " & vbCrLf & _
        "           B.������ ָ����," & vbCrLf & _
        "           B.Ӣ���� Ӣ����," & vbCrLf & _
        "           A.��������,A.�ϲ���," & vbCrLf & _
        "           A.������λ ��λ " & vbCrLf & _
        " FROM ���˲��������� a,����������Ŀ b,���鱨����Ŀ c   " & vbCrLf & _
        " WHERE a.������ID(+)=B.ID AND c.������Ŀid=b.id " & vbCrLf & _
        "   AND nvl(a.�ϲ���,0)=c.������Ŀid " & vbCrLf & _
        "   AND a.����ID=" & lng����ID & vbCrLf & _
        " ORDER BY c.�������"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "����������")
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        '��������ݾͶ�������,
        msfMain.Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            msfMain.TextMatrix(i, EnmGridCol.ItemID) = zlCommFun.Nvl(rsTmp!ID, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item����) = zlCommFun.Nvl(rsTmp!����, 1)
            msfMain.TextMatrix(i, EnmGridCol.Item��ʾ��) = zlCommFun.Nvl(rsTmp!��ʾ��, 0)
            If zlCommFun.Nvl(rsTmp!����, 1) = 0 Then
                If Trim(zlCommFun.Nvl(rsTmp!��ֵ��)) <> ";" Then
                    msfMain.TextMatrix(i, EnmGridCol.Item����ֵ) = Replace(Trim(zlCommFun.Nvl(rsTmp!��ֵ��)), ";", " �� ")
                End If
            End If
            msfMain.TextMatrix(i, EnmGridCol.Item��ֵ��) = Trim(zlCommFun.Nvl(rsTmp!��ֵ��))
            msfMain.TextMatrix(i, EnmGridCol.Item��ʼֵ) = Trim(zlCommFun.Nvl(rsTmp!��ʼֵ))
            msfMain.TextMatrix(i, EnmGridCol.Item����) = zlCommFun.Nvl(rsTmp!����, 0)
            msfMain.TextMatrix(i, EnmGridCol.ItemС����) = zlCommFun.Nvl(rsTmp!С��, 0)
            msfMain.TextMatrix(i, EnmGridCol.Itemָ����) = Trim(zlCommFun.Nvl(rsTmp!ָ����)) & IIf(Trim(zlCommFun.Nvl(rsTmp!Ӣ����)) = "", "", "��" & Trim(zlCommFun.Nvl(rsTmp!Ӣ����)) & "��")
            msfMain.TextMatrix(i, EnmGridCol.ItemӢ��) = Trim(zlCommFun.Nvl(rsTmp!Ӣ����))
            msfMain.TextMatrix(i, EnmGridCol.Item��������) = zlCommFun.Nvl(rsTmp!��������)
            msfMain.TextMatrix(i, EnmGridCol.Item��λ) = Trim(zlCommFun.Nvl(rsTmp!��λ))
            If i = 1 Then
                mID������Ŀ = zlCommFun.Nvl(rsTmp!�ϲ���, 0)
                cmdP1.Tag = mID������Ŀ
            End If
            '���µ���ָ�����Ŀ��
            If UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Itemָ����)) > lngWidth Then
                lngWidth = UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Itemָ����))
            End If
            rsTmp.MoveNext
        Next
        '���µ���ָ�����Ŀ��
        If msfMain.ColWidth(EnmGridCol.Itemָ����) < lngWidth Then
            msfMain.ColWidth(EnmGridCol.Itemָ����) = lngWidth
        End If
        i = msfMain.Width - (msfMain.ColWidth(EnmGridCol.Item�к�) + msfMain.ColWidth(EnmGridCol.Itemָ����) + msfMain.ColWidth(EnmGridCol.Item��λ)) - Screen.TwipsPerPixelX * 6
        msfMain.ColWidth(EnmGridCol.Item��������) = IIf(i < 200, 200, i)
        
        '����������Ŀ��ǰ������
        strSQL = "select * from ���˲��������� where �ؼ��� in (-2,-1) and �ϲ���=" & mID������Ŀ & " and ����ID=" & lng����ID
        Call zlDatabase.OpenRecordset(rs������Ŀ, strSQL, "����������")
        rs������Ŀ.Filter = "�ؼ���=-2"
        If rs������Ŀ.RecordCount > 0 Then
            txtItem.Tag = zlCommFun.Nvl(rs������Ŀ!����)
            '�ٴμ���Ƿ���ڸ���Ŀ����������ʾ
            strSQL = "select * from ������ĿĿ¼  where id=" & mID������Ŀ
            If rs������Ŀ1.State = adStateOpen Then rs������Ŀ1.Close
            Set rs������Ŀ1 = Nothing
            Call zlDatabase.OpenRecordset(rs������Ŀ1, strSQL, "����������")
            If rs������Ŀ1.RecordCount > 0 Then
                txtItem.Text = zlCommFun.Nvl(rs������Ŀ1!����)
            Else
                '���û�оͳ�ʼ��
                InitMe
                txtItem.Text = ""
                txtItem.Tag = ""
            End If
        Else
            '��û�в����е���Ŀ�ͼ����û���Ǹ���Ŀ
            strSQL = "select * from ������ĿĿ¼  where id=" & mID������Ŀ
            Call zlDatabase.OpenRecordset(rs������Ŀ1, strSQL, "����������")
            ID������Ŀ = mID������Ŀ
            If rs������Ŀ1.RecordCount < 1 Then
                '���û�о��˳�
                Exit Sub
            End If
        End If
        '�õ��걾
        rs������Ŀ.Filter = "�ؼ���=-1"
        If rs������Ŀ.RecordCount > 0 Then
            lblBB.Caption = zlCommFun.Nvl(rs������Ŀ!����)
        Else
            lblBB.Caption = ""
        End If
        
        mblnCancel = True
        ReSetRowCode msfMain
        PicItem_Resize
        mblnCancel = False
        ReSetRowCode msfMain
    Else
        '����û�����ݾͳ�ʼ����������Ŀ�ı��
        ID������Ŀ = mID������Ŀ
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function SaveData(lng����ID As Long, lng��ҳID As Long, lng����ID As Long, strReturnSQL As String, strError As String) As Boolean
'�����������Ŀ����
Dim strTmp As String, strName As String
Dim lngRow As Long
Dim lngCol As Long

On Error GoTo ErrHandle

    If msfMain.Rows < 3 And Trim(msfMain.TextMatrix(1, EnmGridCol.ItemID)) = "" Then
        strReturnSQL = ""
        strError = "����������Ϊ��"
        Exit Function
    End If
    If mID������Ŀ < 1 Then
        strReturnSQL = ""
        strError = "������Ŀ��ȷ����������ѡ�������Ŀ"
        Exit Function
    End If
    '�õ���Ŀ����
    strName = IIf(txtItem.Tag = "", Trim(txtItem.Text), txtItem.Tag)
    '������Ŀ
    strTmp = mID������Ŀ & "''"
    strTmp = strTmp & strName & "''"
    strTmp = strTmp & " ''"
    strTmp = strTmp & "-2''-2'' ''"
    '����걾
    strName = Trim(lblBB.Caption)
    strTmp = strTmp & mID������Ŀ & "''"
    strTmp = strTmp & strName & "''"
    strTmp = strTmp & " ''"
    strTmp = strTmp & "-1''-1'' ''"
    For lngRow = 1 To msfMain.Rows - 1
        '�����Ӣ������ȡӢ����,�����ȡ��������ͬʱ�滻�����������ݿ�ķָ���
        strName = IIf(Replace(Trim(msfMain.TextMatrix(lngRow, EnmGridCol.ItemӢ��)), "'", "��") = "", Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Itemָ����)), Trim(msfMain.TextMatrix(lngRow, EnmGridCol.ItemӢ��)))
        '�洢���̲�����ʽ:������ĿID'�����ı�'��λ'������ĿID'�ؼ���'��������'������ĿID1'�����ı�1'��λ1'������ĿID1'�ؼ���1'��������1'
        strTmp = strTmp & mID������Ŀ & "''"
        strTmp = strTmp & strName & "''"
        strTmp = strTmp & Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Item��λ)) & "''"
        strTmp = strTmp & Trim(msfMain.TextMatrix(lngRow, EnmGridCol.ItemID)) & "''"
        
        For lngCol = 1 To Len("'`|,""")
            If InStr(Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Item��������)), Mid("'`|,""", lngCol, 1)) > 0 Then
                msfMain.Row = lngRow: msfMain.Col = EnmGridCol.Item��������
                msfMain_EnterCell
                strError = "��" & lngRow & "�д��ڷǷ��ַ���"
                SetErr 0, "��" & lngRow & "�д��ڷǷ��ַ���"
                Exit Function
            End If
        Next
        strTmp = strTmp & CStr(lngRow - 1) & "''" & Trim(msfMain.TextMatrix(lngRow, EnmGridCol.Item��������)) & "''"
    Next
    strReturnSQL = "ZL_��������¼_INSERT(" & lng����ID & ",'" & strTmp & "')"
    SaveData = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitMe()
Dim rsNewTmp As New ADODB.Recordset
    '��ʼ���ؼ�
    msfMain.Clear
    msfMain.Rows = 2
    msfMain.FixedRows = 1
    msfMain.Cols = 13
    '��һ������Ϊ��
    msfMain.RowHeight(0) = 0
    '��ͷ
    msfMain.TextMatrix(0, EnmGridCol.ItemID) = "ID"
    msfMain.TextMatrix(0, EnmGridCol.Item����) = "����"
    msfMain.TextMatrix(0, EnmGridCol.Item��ʾ��) = "��ʾ��"
    msfMain.TextMatrix(0, EnmGridCol.Item��ֵ��) = "��ֵ��"
    msfMain.TextMatrix(0, EnmGridCol.Item��ʼֵ) = "��ʼֵ"
    msfMain.TextMatrix(0, EnmGridCol.Item����) = "����"
    msfMain.TextMatrix(0, EnmGridCol.ItemС����) = "С����"
    msfMain.TextMatrix(0, EnmGridCol.Item�к�) = "�к�"
    msfMain.TextMatrix(0, EnmGridCol.Itemָ����) = "ָ����"
    msfMain.TextMatrix(0, EnmGridCol.ItemӢ��) = "Ӣ����"
    msfMain.TextMatrix(0, EnmGridCol.Item����ֵ) = "����ֵ"
    msfMain.TextMatrix(0, EnmGridCol.Item��������) = "��������"
    msfMain.TextMatrix(0, EnmGridCol.Item��λ) = "��λ"
    '���ø��еĿ�
    msfMain.ColWidth(EnmGridCol.ItemID) = 0
    msfMain.ColWidth(EnmGridCol.Item����) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item��ʾ��) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item��ֵ��) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item��ʼֵ) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item����) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.ItemС����) = msfMain.ColWidth(EnmGridCol.ItemID)
    msfMain.ColWidth(EnmGridCol.Item�к�) = 300
    msfMain.ColWidth(EnmGridCol.Itemָ����) = 2400
    msfMain.ColWidth(EnmGridCol.Item����ֵ) = 2400
    msfMain.ColWidth(EnmGridCol.ItemӢ��) = 0
    
    '��λ�еĿ�
    msfMain.ColWidth(EnmGridCol.Item��λ) = 1000
    '�Զ�������������ݵĿ�
    i = msfMain.Width - (msfMain.ColWidth(EnmGridCol.Item�к�) + msfMain.ColWidth(EnmGridCol.Itemָ����) + msfMain.ColWidth(EnmGridCol.Item����ֵ) + msfMain.ColWidth(EnmGridCol.Item��λ)) - Screen.TwipsPerPixelX * 6
    msfMain.ColWidth(EnmGridCol.Item��������) = IIf(i < 200, 200, i)
    '�����ж���
    msfMain.ColAlignment(EnmGridCol.Item�к�) = AlignmentSettings.flexAlignLeftCenter
    msfMain.ColAlignment(EnmGridCol.Itemָ����) = AlignmentSettings.flexAlignLeftCenter
    msfMain.ColAlignment(EnmGridCol.Item����ֵ) = AlignmentSettings.flexAlignCenterCenter
    msfMain.ColAlignment(EnmGridCol.ItemӢ��) = AlignmentSettings.flexAlignLeftCenter
    msfMain.ColAlignment(EnmGridCol.Item��������) = msfMain.ColAlignment(EnmGridCol.Itemָ����)
    msfMain.ColAlignment(EnmGridCol.Item��λ) = msfMain.ColAlignment(EnmGridCol.Itemָ����)
    UserControl_Resize
    txtItem.Tag = ""
    txtItem.Text = ""
End Sub

Private Sub ReSetRowCode(objMSH As MSHFlexGrid)
'���кŽ�����������
Dim lngWidth�к� As Long

    For i = 1 To objMSH.Rows - 1
'        objMSH.RowHeight(i) = CmbCell.Height
        objMSH.TextMatrix(i, EnmGridCol.Item�к�) = CStr(i) & "��"
        If UserControl.TextWidth(objMSH.TextMatrix(i, EnmGridCol.Item�к�)) > lngWidth�к� Then lngWidth�к� = UserControl.TextWidth(objMSH.TextMatrix(i, EnmGridCol.Item�к�))
    Next
    objMSH.ColWidth(EnmGridCol.Item�к�) = lngWidth�к�
End Sub

Private Function InDesign() As Boolean
'���ܣ��жϵ�ǰ���г����Ƿ���VB�Ĺ��̻�����
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Private Sub SetErr(lngErrNum As Long, strErr As String)
'���ô��������������
'���lngErrNum=-1 ��ʾ �ؼ��Լ�����Ĵ���
mReturnErrnumber = lngErrNum
mReturnErrDescription = strErr
End Sub

Public Property Get ID���˲���() As Long
'���ز��˲���ID
    ID���˲��� = mlng����id
End Property

Public Property Let ID���˲���(ByVal New_ID���˲��� As Long)
'���ò��˲���ID,�����ò����ǲ��Ǵ���
    mlng����id = New_ID���˲���
    ReadData mlng����id
End Property

Public Property Get ReturnErrNumber() As Long
'�������һ�εĴ����
    ReturnErrNumber = mReturnErrnumber
End Property

Public Property Get ReturnErrDescription() As String
'�������һ�δ��������ַ���
    ReturnErrDescription = mReturnErrDescription
End Property

Public Property Get DispMode() As Boolean
'�Ƿ�Ϊ��ʾģʽ
    DispMode = mDispMode
End Property

Public Property Let DispMode(ByVal New_DispMode As Boolean)
    mDispMode = New_DispMode
    msfMain_EnterCell
    PropertyChanged "DispMode"
End Property

Public Property Get BorderStyle() As BorderStyleSettings
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleSettings)
    UserControl.BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub CmbCell_Click()
If mblnCancel = False Then
    msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������) = CmbCell.Text
Else
    mblnCancel = False
End If
End Sub

Private Sub CmbCell_DblClick()
If mblnCancel = False Then
    msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������) = CmbCell.Text
End If
End Sub

Private Sub CmbCell_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or ((KeyCode = vbKeyDown Or KeyCode = vbKeyPageDown) And Shift = 2) Then
        KeyCode = 0
        If msfMain.Row < msfMain.Rows - 1 Then
            mblnCancel = True
            msfMain.Row = msfMain.Row + 1
            msfMain_EnterCell
            mblnCancel = True
            Exit Sub
        End If
    ElseIf (KeyCode = vbKeyUp Or KeyCode = vbKeyPageUp) And Shift = 2 Then
        KeyCode = 0
        If msfMain.Row > 1 Then
            mblnCancel = True
            msfMain.Row = msfMain.Row - 1
            msfMain_EnterCell
            mblnCancel = True
            Exit Sub
        End If
    End If
    '
    If mblnCancel = False Then
        msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������) = CmbCell.Text
    End If
End Sub

Private Sub cmdP_Click()
On Error GoTo ErrHandle
Dim strSQL As String
Dim strReturn As String
Dim CurPoint As POINTAPI
Dim rsNewTmp As New ADODB.Recordset
    If gcnOracle Is Nothing Then SetErr -1, "���Ӷ���û�г�ʼ��": Exit Sub
    If gcnOracle.State <> adStateOpen Then SetErr -1, "���Ӷ���û������": Exit Sub
    strSQL = "SELECT DISTINCT ����걾 �걾���� FROM ���鱨����Ŀ where ������ĿID=" & mID������Ŀ
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "����������")
    If rsTmp.RecordCount = 1 Then
        lblBB.Caption = zlCommFun.Nvl(rsTmp!�걾����)
        PicItem_Resize
    ElseIf rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        '��λѡ����
        CurPoint.X = (cmdP.Left) / Screen.TwipsPerPixelX
        CurPoint.Y = (cmdP.Top + cmdP.Height + Screen.TwipsPerPixelY * 2) / Screen.TwipsPerPixelY
        ClientToScreen PicItem.hwnd, CurPoint
        CurPoint.X = CurPoint.X * Screen.TwipsPerPixelX
        CurPoint.Y = CurPoint.Y * Screen.TwipsPerPixelY
        If CurPoint.X + 3000 > Screen.Width Then CurPoint.X = Screen.Width - 3200
        If CurPoint.X < 0 Then CurPoint.X = 0
        If CurPoint.Y + 2400 > Screen.Height Then CurPoint.Y = Screen.Height - 2800
        If CurPoint.Y < 0 Then CurPoint.Y = 0
        strReturn = frmSelectChild.ShowSelectChild(Me, CurPoint.X, CurPoint.Y, 3200, 2400, rsTmp, "2800")
        If Trim(strReturn) = "" Or Trim(strReturn) = ";" Then Exit Sub
        lblBB.Caption = Split(strReturn, ";")(0)
        PicItem_Resize
    End If
    '����ָ���걾��ָ����Ŀ
    strSQL = _
        "SELECT C.ID," & vbCrLf & _
        "       A.������� ���," & vbCrLf & _
        "       A.����걾," & vbCrLf & _
        "       C.����," & vbCrLf & _
        "       C.��ʾ��," & vbCrLf & _
        "       C.��ֵ��," & vbCrLf & _
        "       C.��ʼֵ," & vbCrLf & _
        "       C.����," & vbCrLf & _
        "       C.С��," & vbCrLf & _
        "       C.������ ָ����," & vbCrLf & _
        "       C.Ӣ���� Ӣ����," & vbCrLf & _
        "       C.��λ" & vbCrLf & _
        "  FROM ������ĿĿ¼ B, ���鱨����Ŀ A,����������Ŀ C" & vbCrLf & _
        " WHERE B.ID IN (SELECT DISTINCT ������ĿID FROM ���鱨����Ŀ) AND  A.������Ŀid=C.Id AND " & vbCrLf & _
        "      A.����걾='" & lblBB.Caption & "' AND B.ID = A.������ĿID  AND A.������ĿID =" & ID������Ŀ & vbCrLf & _
        " ORDER BY A.�������"
    Call zlDatabase.OpenRecordset(rsNewTmp, strSQL, "����������")
    '����������Ŀ�ļ���ָ��
    If rsNewTmp.RecordCount > 0 Then
        mblnCancel = True
        rsNewTmp.MoveFirst
        '����
        msfMain.Rows = rsNewTmp.RecordCount + 1
        For i = 1 To rsNewTmp.RecordCount
            msfMain.TextMatrix(i, EnmGridCol.ItemID) = zlCommFun.Nvl(rsNewTmp!ID, 0)
            msfMain.TextMatrix(i, EnmGridCol.Item����) = zlCommFun.Nvl(rsNewTmp!����, 1)
            msfMain.TextMatrix(i, EnmGridCol.Item��ʾ��) = zlCommFun.Nvl(rsNewTmp!��ʾ��, 0)
            If zlCommFun.Nvl(rsNewTmp!����, 1) = 0 Then
                If Trim(zlCommFun.Nvl(rsNewTmp!��ֵ��)) <> ";" Then
                    msfMain.TextMatrix(i, EnmGridCol.Item����ֵ) = Replace(Trim(zlCommFun.Nvl(rsNewTmp!��ֵ��)), ";", " �� ")
                End If
            End If
            msfMain.TextMatrix(i, EnmGridCol.Item��ֵ��) = Trim(zlCommFun.Nvl(rsNewTmp!��ֵ��))
            msfMain.TextMatrix(i, EnmGridCol.Item��ʼֵ) = Trim(zlCommFun.Nvl(rsNewTmp!��ʼֵ))
            msfMain.TextMatrix(i, EnmGridCol.Item����) = zlCommFun.Nvl(rsNewTmp!����, 0)
            msfMain.TextMatrix(i, EnmGridCol.ItemС����) = zlCommFun.Nvl(rsNewTmp!С��, 0)
            msfMain.TextMatrix(i, EnmGridCol.Itemָ����) = Trim(zlCommFun.Nvl(rsNewTmp!ָ����)) & IIf(Trim(zlCommFun.Nvl(rsNewTmp!Ӣ����)) = "", "", "��" & Trim(zlCommFun.Nvl(rsNewTmp!Ӣ����)) & "��")
            msfMain.TextMatrix(i, EnmGridCol.ItemӢ��) = Trim(zlCommFun.Nvl(rsNewTmp!Ӣ����))
            msfMain.TextMatrix(i, EnmGridCol.Item��������) = Trim(zlCommFun.Nvl(rsNewTmp!��ʼֵ))
            msfMain.TextMatrix(i, EnmGridCol.Item��λ) = Trim(zlCommFun.Nvl(rsNewTmp!��λ))
            rsNewTmp.MoveNext
        Next
        ReSetRowCode msfMain
        PicItem_Resize
    End If
    UserControl_Resize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub listCell_ItemCheck(Item As Integer)
Dim strTmp As String
    If mblnCancel = True Then Exit Sub
    For i = 0 To listCell.ListCount - 1
        If listCell.Selected(i) = True Then
            strTmp = strTmp & listCell.List(i) & ";"
        End If
    Next
    If Right(strTmp, 1) = ";" Then strTmp = Left(strTmp, Len(strTmp) - 1)
    msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������) = strTmp
End Sub

Private Sub listCell_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If msfMain.Row < msfMain.Rows - 1 Then
            listCell.Visible = False
            msfMain.Row = msfMain.Row + 1
            msfMain_EnterCell
        End If
    End If
End Sub

Private Sub SetSelColor(objMsf As MSHFlexGrid, ByVal lngRow As Long, Optional ByVal oleForeColor As OLE_COLOR = 0, Optional ByVal oleBackColor As OLE_COLOR = &HFFFFFF)
'����ѡ���е���ɫ
Dim lngSelCol As Long, lngSelRow As Long

    objMsf.Redraw = False
    lngSelCol = objMsf.Col
    lngSelRow = objMsf.Row
    
    For i = 1 To objMsf.Rows - 1
        objMsf.Row = i
        If i = lngRow Then
            For j = 0 To objMsf.Cols - 1
                objMsf.Col = j
                objMsf.CellFontBold = True
                objMsf.CellForeColor = oleForeColor
                objMsf.CellBackColor = oleBackColor
            Next
        Else
            For j = 0 To objMsf.Cols - 1
                objMsf.Col = j
                objMsf.CellFontBold = False
                objMsf.CellForeColor = 0
                objMsf.CellBackColor = RGB(255, 255, 255)
            Next
        End If
    Next
    objMsf.Col = lngSelCol
    objMsf.Row = lngSelRow
    objMsf.Refresh
    objMsf.Redraw = True
End Sub

Private Sub msfMain_EnterCell()
On Error Resume Next
mblnCancel = True
txtCell.Visible = False
listCell.Visible = False
CmbCell.Visible = False

SetSelColor msfMain, msfMain.Row
If msfMain.Row > 0 Then
    '�����еĸ����¸�ֵ
    msfMain.RowHeight(msfMain.Row) = 255
    txtCell.Height = 255
    If Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ʾ��)) = CStr(EnmCTLType.CTLTxt) Or Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ʾ��)) = CStr(EnmCTLType.CTLUpDown) Then
    '���Ϊ�ı�������ʱ
        '�����
        txtCell.Left = msfMain.ColPos(EnmGridCol.Item��������) + Screen.TwipsPerPixelX * 2
        '���
        i = msfMain.ColWidth(EnmGridCol.Item��������) - Screen.TwipsPerPixelX * 4
        txtCell.Width = IIf(i < Screen.TwipsPerPixelX * 4, Screen.TwipsPerPixelX * 4, i)
        '��
        txtCell.Top = msfMain.Top + msfMain.RowPos(msfMain.Row) + Screen.TwipsPerPixelY * 2
        '���
        i = msfMain.CellHeight - Screen.TwipsPerPixelY * 4
        txtCell.Height = IIf(i < Screen.TwipsPerPixelY * 4, Screen.TwipsPerPixelY * 4, i)
        '�õ���ǰ����
        mblnCancel = True
        txtCell.Text = msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������)
        mblnCancel = False
        If txtCell.Enabled And UserControl.Enabled And mDispMode = False Then
            txtCell.Visible = True
            txtCell.ZOrder
            txtCell.SetFocus
        End If
    ElseIf Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ʾ��)) = CStr(EnmCTLType.CTLDownList) Or Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ʾ��)) = CStr(EnmCTLType.CTLOption) Then
        '���Ϊ�����͵�ѡʱ
        '�����
        CmbCell.Left = msfMain.ColPos(EnmGridCol.Item��������) + Screen.TwipsPerPixelX * 2
        '���
        i = msfMain.ColWidth(EnmGridCol.Item��������) - Screen.TwipsPerPixelX * 4
        CmbCell.Width = IIf(i < Screen.TwipsPerPixelX * 4, Screen.TwipsPerPixelX * 4, i)
        '��
        CmbCell.Top = msfMain.Top + msfMain.RowPos(msfMain.Row) + Screen.TwipsPerPixelY * 2
        msfMain.RowHeight(msfMain.Row) = CmbCell.Height
        '�õ���ǰ����
        mblnCancel = True
        CmbCell.Clear
        If InStr(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ֵ��)), ";") > 0 Then
            'ѡ��ʼ��
            For i = 0 To UBound(Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ֵ��)), ";"))
                CmbCell.AddItem Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ֵ��)), ";")(i)
            Next
            CmbCell.ListIndex = 0
            '����ֵ
            If Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������)) = "" Then
                msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������) = CmbCell.Text
            Else
                For i = 0 To CmbCell.ListCount - 1
                    If CmbCell.List(i) = msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������) Then
                        CmbCell.ListIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
        If CmbCell.Enabled And UserControl.Enabled And mDispMode = False Then
            CmbCell.Visible = True
            CmbCell.ZOrder
            CmbCell.SetFocus
        End If
        mblnCancel = False
    ElseIf Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ʾ��)) = CStr(EnmCTLType.CTLCheck) Then
        '���Ϊ��ѡ��ʱ
        '�����
        listCell.Left = msfMain.ColPos(EnmGridCol.Item��������) + Screen.TwipsPerPixelX * 2
        '���
        i = msfMain.ColWidth(EnmGridCol.Item��������) - Screen.TwipsPerPixelX * 4
        listCell.Width = IIf(i < Screen.TwipsPerPixelX * 4, Screen.TwipsPerPixelX * 4, i)
        '��
        listCell.Top = msfMain.Top + msfMain.RowPos(msfMain.Row) + msfMain.CellHeight + Screen.TwipsPerPixelY * 2
        '�������ø�
        listCell.Height = 1200
        If listCell.Top + listCell.Height > UserControl.Height Then
            listCell.Top = listCell.Top - msfMain.CellHeight - Screen.TwipsPerPixelY * 2 - listCell.Height
        End If
        listCell.Clear
        mblnCancel = True
        If InStr(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ֵ��)), ";") > 0 Then
            'ѡ��ʼ��
            For i = 0 To UBound(Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ֵ��)), ";"))
                listCell.AddItem Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��ֵ��)), ";")(i)
            Next
            '������ֵ
            For i = 0 To listCell.ListCount - 1
                For j = 0 To UBound(Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������)), ";"))
                    If listCell.List(i) = Split(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������)), ";")(j) Then
                        listCell.Selected(i) = True
                    End If
                Next
            Next
        End If
        If listCell.Enabled And UserControl.Enabled And mDispMode = False Then
            listCell.Visible = True
            listCell.ZOrder
            listCell.SetFocus
        End If
        mblnCancel = False
    End If
End If
End Sub

Private Sub msfMain_Scroll()
    txtCell.Visible = False
    listCell.Visible = False
    CmbCell.Visible = False
End Sub

Private Sub PicItem_Resize()
    Line1.X2 = PicItem.ScaleWidth - Line1.X1
    cmdP.Left = Line1.X2 - cmdP.Width
    lblBB.Left = cmdP.Left - lblBB.Width - Screen.TwipsPerPixelX * 10
    lblBBCaption.Left = lblBB.Left - lblBBCaption.Width - Screen.TwipsPerPixelX * 4
End Sub

Private Sub txtCell_Change()
If mblnCancel = False Then
    msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������) = txtCell.Text
End If
End Sub

Private Sub txtCell_GotFocus()
    If IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) Then
        If Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) = 0 And IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) Then
            If IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) Then
                If Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) > 0 Then
                    txtCell.MaxLength = Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����))
                End If
            End If
        End If
    End If
    zlControl.TxtSelAll txtCell
    zlCommFun.OpenIme True
End Sub

Private Sub txtCell_KeyDown(KeyCode As Integer, Shift As Integer)
Dim blnCancel As Boolean
'�ȼ���ǲ��������˷Ƿ��ַ�
    If InStr(LAWLChar, Chr(KeyCode)) > 0 Then
        KeyCode = 0
        Exit Sub
    End If
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown, vbKeyPageDown
            KeyCode = 0
            If msfMain.Row < msfMain.Rows - 1 Then
                txtCell_Validate blnCancel
                If mblnLawless = True Then mblnLawless = False:   Exit Sub
                msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������) = txtCell.Text
                txtCell.Visible = False
                msfMain.Row = msfMain.Row + 1
                msfMain_EnterCell
                Exit Sub
            End If
        Case vbKeyUp, vbKeyPageUp
            KeyCode = 0
            If msfMain.Row > 1 Then
                txtCell_Validate blnCancel
                If mblnLawless = True Then mblnLawless = False:   Exit Sub
                txtCell.Visible = False
                msfMain.Row = msfMain.Row - 1
                msfMain_EnterCell
                Exit Sub
            End If
    End Select
End Sub

Private Sub txtCell_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub
    
    If Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) = "0" Then
        If IsNumeric(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.ItemС����))) Then
            If Format(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.ItemС����))) > 0 Then
                'ΪС��ʱ
                If InStr("0123456789.", Chr(KeyAscii)) < 1 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            Else
                'Ϊ����
                If InStr("0123456789", Chr(KeyAscii)) < 1 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        Else
            'Ϊ����
            If InStr("0123456789", Chr(KeyAscii)) < 1 Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub txtCell_LostFocus()
    zlCommFun.OpenIme
End Sub

Private Sub txtCell_Validate(Cancel As Boolean)
'���ȼ��
If IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) Then
    If zlCommFun.ActualLen(txtCell.Text) > 1000 And Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) > 0 And _
        zlCommFun.ActualLen(txtCell.Text) > Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) Then
        MsgBox "���볬��,���������룡", vbInformation, gstrSysName
        mblnLawless = True
        Cancel = True
        Exit Sub
    End If
    If Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item����)) = "0" Then
        If IsNumeric(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������)) = False And Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������)) <> "" Then
            MsgBox "ֻ��������ֵ,���������룡", vbInformation, gstrSysName
            mblnLawless = True
            Cancel = True
            Exit Sub
        End If
        If Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.ItemС����)) > 0 And InStr(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������)), ".") > 0 Then
            i = InStr(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������)), ".")
            i = Len(Trim(msfMain.TextMatrix(msfMain.Row, EnmGridCol.Item��������))) - i
            If i > Format(msfMain.TextMatrix(msfMain.Row, EnmGridCol.ItemС����)) Then
               MsgBox "����С�����ֳ���,���������룡", vbInformation, gstrSysName
               mblnLawless = True
               Cancel = True
               Exit Sub
            End If
        End If
    Else
        If zlCommFun.ActualLen(txtCell.Text) > 1000 Then
            MsgBox "���볬��,���������룡", vbInformation, gstrSysName
            mblnLawless = True
            Cancel = True
            Exit Sub
        End If
    End If
Else
    If zlCommFun.ActualLen(txtCell.Text) > 1000 Then
        MsgBox "���볬��,���������룡", vbInformation, gstrSysName
        mblnLawless = True
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub txtItem_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txtItem
End Sub

Public Property Get Cur��ǰ�걾() As String
    '
    Cur��ǰ�걾 = lblBB.Caption
End Property

Public Property Let Cur��ǰ�걾(ByVal New_Cur As String)
    lblBB.Caption = New_Cur
End Property

Private Sub cmdP1_Click()
On Error GoTo ErrHandle
Dim strWidth As String
Dim CurPoint As POINTAPI

    strSQL = _
        "SELECT a.* FROM (SELECT DISTINCT A.ID, A.����, A.����, B.���� ����, A.���㵥λ" & vbCrLf & _
        "  FROM ������ĿĿ¼ A, ������Ŀ���� B" & vbCrLf & _
        " WHERE B.������ĿID = A.ID(+) AND" & vbCrLf & _
        "      A.ID IN (SELECT DISTINCT ������ĿID FROM ���鱨����Ŀ)) A" & vbCrLf & _
        " order by a.���� "
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "����������")
    If rsTmp.RecordCount = 1 Then
        ID������Ŀ = rsTmp!ID
    ElseIf rsTmp.RecordCount > 0 Then
        '��λѡ����
        CurPoint.X = (txtItem.Left) / Screen.TwipsPerPixelX
        CurPoint.Y = (txtItem.Top + txtItem.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
        ClientToScreen UserControl.hwnd, CurPoint
        CurPoint.X = CurPoint.X * Screen.TwipsPerPixelX
        CurPoint.Y = CurPoint.Y * Screen.TwipsPerPixelY
        If CurPoint.X + 4800 > Screen.Width Then CurPoint.X = Screen.Width - 5180
        If CurPoint.X < 0 Then CurPoint.X = 0
        If CurPoint.Y + Screen.TwipsPerPixelY * 200 > Screen.Height Then CurPoint.Y = CurPoint.Y - txtItem.Height - Screen.TwipsPerPixelY * 200 - Screen.TwipsPerPixelY * 2
        If CurPoint.Y < 0 Then CurPoint.Y = 0
        
        '��ʼѡ����
        strWidth = "0;800;1500;1500;1000"
        strWidth = frmSelectChild.ShowSelectChild(Nothing, CurPoint.X, CurPoint.Y, 4800 + 380, Screen.TwipsPerPixelY * 200, rsTmp, strWidth)
        If Trim(strWidth) = "" Or Trim(strWidth) = ";;" Then
            Exit Sub
        End If
        ID������Ŀ = CLng(Split(strWidth, ";")(0))
    Else
        txtItem.Text = ""
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtItem_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrHandle
Dim blnMatching As Boolean
Dim strWidth As String
Dim CurPoint As POINTAPI

    If KeyCode = vbKeyReturn Then
        blnMatching = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", "0") = "0", True, False)
        KeyCode = 0
        strSQL = _
            "SELECT a.* FROM (SELECT DISTINCT A.ID, A.����,A.����, B.���� ����, A.���㵥λ" & vbCrLf & _
            "  FROM ������ĿĿ¼ A, ������Ŀ���� B" & vbCrLf & _
            " WHERE B.������ĿID = A.ID(+) AND " & vbCrLf & _
            "      (Upper(Nvl(a.����,'')) like '" & UCase(txtItem.Text) & "%' Or  " & vbCrLf & _
            "       Upper(Nvl(a.����,'')) like '" & IIf(blnMatching = True, "%", "") & UCase(txtItem.Text) & "%' Or  " & vbCrLf & _
            "       Upper(Nvl(b.����,'')) like '" & IIf(blnMatching = True, "%", "") & UCase(txtItem.Text) & "%' Or  " & vbCrLf & _
            "       Upper(Nvl(b.����,'')) like '" & IIf(blnMatching = True, "%", "") & UCase(txtItem.Text) & "%' )  and  " & vbCrLf & _
            "      A.ID IN (SELECT DISTINCT ������ĿID FROM ���鱨����Ŀ)) A " & vbCrLf & _
            " order by a.���� "

        Call zlDatabase.OpenRecordset(rsTmp, strSQL, "����������")
        If rsTmp.RecordCount = 1 Then
            ID������Ŀ = rsTmp!ID
        ElseIf rsTmp.RecordCount > 0 Then
            '��λѡ����
            CurPoint.X = (txtItem.Left) / Screen.TwipsPerPixelX
            CurPoint.Y = (txtItem.Top + txtItem.Height + Screen.TwipsPerPixelY) / Screen.TwipsPerPixelY
            ClientToScreen UserControl.hwnd, CurPoint
            CurPoint.X = CurPoint.X * Screen.TwipsPerPixelX
            CurPoint.Y = CurPoint.Y * Screen.TwipsPerPixelY
            If CurPoint.X + 4800 > Screen.Width Then CurPoint.X = Screen.Width - 5180
            If CurPoint.X < 0 Then CurPoint.X = 0
            If CurPoint.Y + Screen.TwipsPerPixelY * 200 > Screen.Height Then CurPoint.Y = CurPoint.Y - txtItem.Height - Screen.TwipsPerPixelY * 200 - Screen.TwipsPerPixelY * 2
            If CurPoint.Y < 0 Then CurPoint.Y = 0
            
            '��ʼѡ����
            strWidth = "0;800;1500;1500;1000"
            strWidth = frmSelectChild.ShowSelectChild(Nothing, CurPoint.X, CurPoint.Y, 4800 + 380, Screen.TwipsPerPixelY * 200, rsTmp, strWidth)
            If Trim(strWidth) = "" Or Trim(strWidth) = ";;" Then
                Exit Sub
            End If
            ID������Ŀ = CLng(Split(strWidth, ";")(0))
        Else
            ID������Ŀ = 0
            txtItem.Text = ""
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtItem_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub UserControl_GotFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub UserControl_Initialize()
    UserControl.Font.Name = "����"
    UserControl.Font.Size = 9
    UserControl.Font.Bold = True
    ID������Ŀ = mID������Ŀ
    mItemIndex = -1
    mblnCancel = False
End Sub

Private Sub UserControl_InitProperties()
    mDispMode = False
    mShowItem = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDispMode = PropBag.ReadProperty("DispMode", False)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    ShowItem = PropBag.ReadProperty("ShowItem", True)
    mblnCancel = True
    ID������Ŀ = PropBag.ReadProperty("ID������Ŀ", 0)
    mblnCancel = False
    txtCell.Visible = False
    CmbCell.Visible = False
    listCell.Visible = False
End Sub

Public Property Get Text() As String
'Ϊÿһ���ؼ������ı�ת������
Dim i As Long
Dim strTmp As String, strName As String

'ͨ���û���������ݵõ�ת���ı�
    If msfMain.Rows < 2 Then Exit Property
    If msfMain.Rows = 2 And msfMain.TextMatrix(1, EnmGridCol.Itemָ����) = "" Then Exit Property
    '�õ���Ŀ����
    strName = IIf(txtItem.Tag = "", Trim(txtItem.Text), txtItem.Tag)
    '������Ŀ
    strTmp = strName & "��" & lblBB.Caption & "����"
    For i = 1 To msfMain.Rows - 1
        strTmp = strTmp & IIf(Trim(msfMain.TextMatrix(i, EnmGridCol.ItemӢ��)) <> "", Trim(msfMain.TextMatrix(i, EnmGridCol.ItemӢ��)), msfMain.TextMatrix(i, EnmGridCol.Itemָ����)) & " " & msfMain.TextMatrix(i, EnmGridCol.Item��������) & msfMain.TextMatrix(i, EnmGridCol.Item��λ) & IIf(i = msfMain.Rows - 1, "", "��")
    Next
    Text = strTmp
End Property

Private Sub UserControl_Resize()
    Dim lngWidth As Long
    Dim lngWidth��λ As Long
    
    msfMain.Left = 0
    msfMain.Top = PicItem.Top + PicItem.Height
    msfMain.Width = ScaleWidth
    i = ScaleHeight - (PicItem.Top + PicItem.Height)
    msfMain.Height = IIf(i > Screen.TwipsPerPixelY, i, Screen.TwipsPerPixelY)
    msfMain_EnterCell
    
    msfMain.ColWidth(EnmGridCol.Itemָ����) = 2400
    msfMain.ColWidth(EnmGridCol.Item����ֵ) = 1200
    For i = 1 To msfMain.Rows - 1
        If UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Itemָ����)) > lngWidth Then
            lngWidth = UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Itemָ����)) / 2
        End If
        If UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Item��λ)) > lngWidth��λ Then
            lngWidth��λ = UserControl.TextWidth(msfMain.TextMatrix(i, EnmGridCol.Item��λ)) / 2
        End If
    Next
    If msfMain.ColWidth(EnmGridCol.Itemָ����) < lngWidth Then
        msfMain.ColWidth(EnmGridCol.Itemָ����) = lngWidth
    End If
    If msfMain.ColWidth(EnmGridCol.Item��λ) < lngWidth��λ Then
        msfMain.ColWidth(EnmGridCol.Item��λ) = lngWidth��λ
    End If
    '�Զ�������������ݵĿ�
    i = msfMain.Width - (msfMain.ColWidth(EnmGridCol.Item�к�) + msfMain.ColWidth(EnmGridCol.Itemָ����) + msfMain.ColWidth(EnmGridCol.Item����ֵ) + msfMain.ColWidth(EnmGridCol.ItemӢ��) + msfMain.ColWidth(EnmGridCol.Item��λ)) - Screen.TwipsPerPixelX * 20
    msfMain.ColWidth(EnmGridCol.Item��������) = IIf(i < 200, 200, i)
    msfMain_EnterCell
    mblnCancel = False
End Sub

Private Sub UserControl_Show()
    mblnCancel = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ShowItem", mShowItem, True)
    Call PropBag.WriteProperty("DispMode", mDispMode, False)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("ID������Ŀ", mID������Ŀ, 0)
End Sub
 
Private Sub UserControl_EnterFocus()
    On Error Resume Next
    UserControl.Parent.CallBack_GotFocus
End Sub

