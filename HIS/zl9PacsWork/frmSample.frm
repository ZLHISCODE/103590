VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSample 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10050
   Icon            =   "frmSample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   10050
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame framSample 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      Begin VB.Frame fram���� 
         Caption         =   "1���걾����"
         Height          =   3495
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   9615
         Begin VB.TextBox txtPathologyNo 
            Height          =   350
            Left            =   5040
            TabIndex        =   28
            Top             =   2400
            Width           =   1815
         End
         Begin VB.ComboBox cboCheckType 
            Height          =   300
            Left            =   2520
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   2425
            Width           =   1575
         End
         Begin VB.TextBox txt����ԭ�� 
            Height          =   350
            Left            =   2520
            TabIndex        =   21
            Top             =   2880
            Width           =   6855
         End
         Begin VB.ComboBox cbo���ռ�ʦ 
            Height          =   300
            Left            =   8160
            TabIndex        =   20
            Text            =   "Combo1"
            Top             =   2422
            Width           =   1260
         End
         Begin VB.CommandButton cmdRefuse 
            Caption         =   "����"
            Height          =   350
            Left            =   120
            TabIndex        =   19
            Top             =   2880
            Width           =   1100
         End
         Begin VB.CommandButton cmdCheckIn 
            Caption         =   "����"
            Height          =   350
            Left            =   120
            TabIndex        =   18
            Top             =   2400
            Width           =   1100
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgList 
            Height          =   1935
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   9375
            _cx             =   16536
            _cy             =   3413
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
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
         Begin VB.Label Label11 
            Caption         =   "�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   255
            Left            =   4320
            TabIndex        =   27
            Top             =   2445
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "�������"
            Height          =   255
            Left            =   1560
            TabIndex        =   25
            Top             =   2448
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "����ԭ��"
            Height          =   255
            Left            =   1560
            TabIndex        =   23
            Top             =   2928
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "���ռ�ʦ"
            Height          =   255
            Left            =   7200
            TabIndex        =   22
            Top             =   2445
            Width           =   855
         End
      End
      Begin VB.Frame framȡ�� 
         Caption         =   "2���޼�"
         Height          =   3855
         Left            =   120
         TabIndex        =   1
         Top             =   4200
         Width           =   9615
         Begin VB.CommandButton cmdSave 
            Caption         =   "����"
            Height          =   350
            Left            =   120
            TabIndex        =   24
            Top             =   3360
            Width           =   1100
         End
         Begin VB.TextBox txt�޼���� 
            Height          =   1695
            Left            =   600
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   240
            Width           =   8895
         End
         Begin VB.TextBox txt���� 
            Height          =   350
            Left            =   600
            TabIndex        =   7
            Top             =   2040
            Width           =   4575
         End
         Begin VB.TextBox txt��ע 
            Height          =   350
            Left            =   600
            TabIndex        =   6
            Top             =   2475
            Width           =   8895
         End
         Begin VB.TextBox txtʣ��걾λ�� 
            Height          =   350
            Left            =   6480
            TabIndex        =   5
            Top             =   2040
            Width           =   3015
         End
         Begin VB.ComboBox cboȡ�ļ�ʦ 
            Height          =   300
            Left            =   1080
            TabIndex        =   4
            Text            =   "Combo2"
            Top             =   2932
            Width           =   1500
         End
         Begin VB.ComboBox cbo�޼�ҽʦ 
            Height          =   300
            Left            =   4560
            TabIndex        =   3
            Text            =   "Combo3"
            Top             =   2910
            Width           =   1500
         End
         Begin VB.ComboBox cbo��Ƭ��ʦ 
            Height          =   300
            Left            =   7800
            TabIndex        =   2
            Text            =   "Combo4"
            Top             =   2910
            Width           =   1500
         End
         Begin VB.Label Label4 
            Caption         =   "�޼�"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label5 
            Caption         =   "����"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2085
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "��ע"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "ʣ��걾λ��"
            Height          =   255
            Left            =   5280
            TabIndex        =   12
            Top             =   2085
            Width           =   1215
         End
         Begin VB.Label Label8 
            Caption         =   "ȡ�ļ�ʦ"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2955
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "�޼�ҽʦ"
            Height          =   255
            Left            =   3480
            TabIndex        =   10
            Top             =   2955
            Width           =   855
         End
         Begin VB.Label Label10 
            Caption         =   "��Ƭ��ʦ"
            Height          =   255
            Left            =   6840
            TabIndex        =   9
            Top             =   2955
            Width           =   735
         End
      End
      Begin VB.Label lblInfo 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   9495
      End
   End
End
Attribute VB_Name = "frmSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mlngҽ��ID As Long
Private mlng���ͺ� As Long
Private blnInit As Boolean
Private blnChangedDept As Boolean
Private mblnMoved As Boolean
Private mstrPrivs As String

Private Enum ColList
        ColID = 0
        Col���1
        Col�걾��λ������1
        Col����1
        Col���2
        Col�걾��λ������2
        Col����2
        Col���3
        Col�걾��λ������3
        Col����3
End Enum

'״̬�ı��¼�lngState = 0 δ������lngState=1 �걾���գ�lngState = 2 �걾����
Public Event StateChanged(lngState As Long, str����� As String, str��������� As String)

'�������ú���
Public Sub zlRefresh(lng����ID As Long, lngҽ��ID As Long, lng���ͺ� As Long, strPrivs As String, ByVal blnReadOnly As Boolean, ByVal blnMoved As Boolean)
    If mlng����ID <> lng����ID Or lng����ID = 0 Then
        blnChangedDept = True
        mlng����ID = lng����ID
    Else
        blnChangedDept = False
    End If
    
    mlngҽ��ID = lngҽ��ID
    mlng���ͺ� = lng���ͺ�
    mblnMoved = blnMoved
    mstrPrivs = strPrivs
    
    Call InitBillSamples    '��ʼ���걾��λ����
    If blnInit = True Then  '����װ�أ���ʼ����������
        Call FillCheckType      '��ʼ������������
    End If
    If blnChangedDept = True Then Call InitDoctors
    '������Ԫ��
    Call FillInterface
    If lngҽ��ID = 0 Or lng���ͺ� = 0 Or blnReadOnly = True Or mblnMoved = True Then
        framSample.Enabled = False
    Else
        framSample.Enabled = True
    End If
End Sub

Private Sub FillCheckType()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    blnInit = False
    cboCheckType.Clear
    strSQL = "Select ���� From Ӱ�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡӰ�������")
    
    If rsTemp.EOF = True Then
        MsgBoxD Me, "���Ҳ������������͵���Ϣ���������ֵ��Ӱ������������á�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    While Not rsTemp.EOF
        cboCheckType.AddItem (Nvl(rsTemp!����))
        rsTemp.MoveNext
    Wend
    If cboCheckType.ListCount > 0 And cboCheckType.ListIndex = -1 Then cboCheckType.ListIndex = 0
End Sub

Private Sub FillInterface()
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    
    '��յ�ǰ����
    lblInfo.Caption = ""
    txt����ԭ��.Text = ""
    txt�޼����.Text = ""
    txt����.Text = ""
    txtʣ��걾λ��.Text = ""
    txt��ע.Text = ""
    txtPathologyNo.Text = ""
    
    On Error GoTo errHandle
    
    '��д�޼������
    strSQL = "Select ҽ��ID,���ͺ�,�޼�����,����,ʣ��걾λ��,��ע,�޼�ҽʦ,ȡ�ļ�ʦ,��Ƭ��ʦ,���ռ�ʦ," & _
             "����ԭ��,�������,�����,��������� From Ӱ��걾����ȡ�� Where ҽ��ID=[1]"
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "Ӱ��걾����ȡ��", "HӰ��걾����ȡ��")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����걾����ȡ��", mlngҽ��ID)
    If Not rsTemp.EOF Then
        If Nvl(rsTemp!�������, 0) = 2 Then '����
            lblInfo.Caption = "�걾������"
            lblInfo.ForeColor = vbRed
            cbo���ռ�ʦ.Text = Nvl(rsTemp!���ռ�ʦ)
            txt����ԭ��.Text = Nvl(rsTemp!����ԭ��)
        Else
            lblInfo.Caption = "�걾�Ѻ���"
            lblInfo.ForeColor = vbBlue
            txt�޼����.Text = Nvl(rsTemp!�޼�����)
            txt����.Text = Nvl(rsTemp!����)
            txtʣ��걾λ��.Text = Nvl(rsTemp!ʣ��걾λ��)
            txt��ע.Text = Nvl(rsTemp!��ע)
            cboȡ�ļ�ʦ.Text = Nvl(rsTemp!ȡ�ļ�ʦ)
            cbo�޼�ҽʦ.Text = Nvl(rsTemp!�޼�ҽʦ)
            cbo��Ƭ��ʦ.Text = Nvl(rsTemp!��Ƭ��ʦ)
            Call SetCheckType(Nvl(rsTemp!���������))
            txtPathologyNo.Text = Nvl(rsTemp!�����)
        End If
    End If
    
    If txtPathologyNo.Text = "" Then    'ˢ����ȡ������
        Call cboCheckType_Click
    End If
    '�������ݿ���д������
    Call FillBILL
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetCheckType(str��������� As String)
    Dim i As Integer
    
    For i = 0 To cboCheckType.ListCount - 1
        If cboCheckType.List(i) = str��������� Then Exit For
    Next i
    If i < cboCheckType.ListCount Then
        cboCheckType.ListIndex = i
    Else
        cboCheckType.ListIndex = -1
    End If
End Sub

Private Sub InitBillSamples()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    With vfgList
        .Clear
        .FixedRows = 1
        .Rows = 6
        .Cols = 10
        .ColWidth(ColID) = 0    'ID
        .ColWidth(Col���1) = 500
        .ColWidth(Col�걾��λ������1) = 1500
        .ColWidth(Col����1) = 1000
        .ColWidth(Col���2) = 500
        .ColWidth(Col�걾��λ������2) = 1500
        .ColWidth(Col����2) = 1000
        .ColWidth(Col���3) = 500
        .ColWidth(Col�걾��λ������3) = 1500
        .ColWidth(Col����3) = 1000
        
        .TextMatrix(0, ColID) = "ID"
        .TextMatrix(0, Col���1) = "���"
        .TextMatrix(0, Col�걾��λ������1) = "�걾��λ������"
        .TextMatrix(0, Col����1) = "����"
        .TextMatrix(0, Col���2) = "���"
        .TextMatrix(0, Col�걾��λ������2) = "�걾��λ������"
        .TextMatrix(0, Col����2) = "����"
        .TextMatrix(0, Col���3) = "���"
        .TextMatrix(0, Col�걾��λ������3) = "�걾��λ������"
        .TextMatrix(0, Col����3) = "����"
        
        .ColAlignment(ColID) = flexAlignCenterCenter
        .ColAlignment(Col���1) = flexAlignCenterCenter
        .ColAlignment(Col�걾��λ������1) = flexAlignCenterCenter
        .ColAlignment(Col����1) = flexAlignCenterCenter
        .ColAlignment(Col���2) = flexAlignCenterCenter
        .ColAlignment(Col�걾��λ������2) = flexAlignCenterCenter
        .ColAlignment(Col����2) = flexAlignCenterCenter
        .ColAlignment(Col���3) = flexAlignCenterCenter
        .ColAlignment(Col�걾��λ������3) = flexAlignCenterCenter
        .ColAlignment(Col����3) = flexAlignCenterCenter
        
        '����걾��λ�����������б������
        strSQL = "Select ���� From Ӱ����걾��λ Order By ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����걾��λ")
        strSQL = ""
        While Not rsTemp.EOF
            strSQL = strSQL & "|" & zlCommFun.SpellCode(Nvl(rsTemp!����)) & "-" & Nvl(rsTemp!����)
            
            rsTemp.MoveNext
        Wend
        strSQL = Mid(strSQL, 2)
        If strSQL = "" Then strSQL = " "
        .ColComboList(Col�걾��λ������1) = "|" & strSQL
        .ColComboList(Col�걾��λ������2) = "|" & strSQL
        .ColComboList(Col�걾��λ������3) = "|" & strSQL
        '����̶������ֱ��
        .TextMatrix(1, Col���1) = "1"
        .TextMatrix(2, Col���1) = "2"
        .TextMatrix(3, Col���1) = "3"
        .TextMatrix(4, Col���1) = "4"
        .TextMatrix(5, Col���1) = "5"
        .TextMatrix(1, Col���2) = "6"
        .TextMatrix(2, Col���2) = "7"
        .TextMatrix(3, Col���2) = "8"
        .TextMatrix(4, Col���2) = "9"
        .TextMatrix(5, Col���2) = "10"
        .TextMatrix(1, Col���3) = "11"
        .TextMatrix(2, Col���3) = "12"
        .TextMatrix(3, Col���3) = "13"
        .TextMatrix(4, Col���3) = "14"
        .TextMatrix(5, Col���3) = "15"
    End With
End Sub

Private Sub FillBILL()
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim lng��� As Long
    Dim i As Integer
    Dim int�������� As Integer
    
    If mlngҽ��ID = 0 And mlng���ͺ� = 0 Then Exit Sub
    On Error GoTo errHandle
    
    strSQL = "Select a.���,a.ҽ��ID,a.���ͺ�,a.�걾��λ,a.���� From Ӱ����걾 a " & _
             "Where a.ҽ��ID = [1] order by a.���"
    
    If mblnMoved = True Then
        strSQL = Replace(strSQL, "Ӱ����걾", "HӰ����걾")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����걾", mlngҽ��ID)
    
    '���ԭ������
    For i = 1 To 5
        vfgList.TextMatrix(i, Col�걾��λ������1) = ""
        vfgList.TextMatrix(i, Col�걾��λ������2) = ""
        vfgList.TextMatrix(i, Col�걾��λ������3) = ""
        vfgList.TextMatrix(i, Col����1) = ""
        vfgList.TextMatrix(i, Col����2) = ""
        vfgList.TextMatrix(i, Col����3) = ""
    Next i
    
    int�������� = 0
    While Not rsTemp.EOF
        lng��� = rsTemp!���
        With vfgList
            If lng��� <= 5 Then
                .TextMatrix(lng���, Col�걾��λ������1) = rsTemp!�걾��λ
                .TextMatrix(lng���, Col����1) = rsTemp!����
            ElseIf lng��� <= 10 Then
                .TextMatrix(lng��� - 5, Col�걾��λ������2) = rsTemp!�걾��λ
                .TextMatrix(lng��� - 5, Col����2) = rsTemp!����
            Else
                .TextMatrix(lng��� - 10, Col�걾��λ������3) = rsTemp!�걾��λ
                .TextMatrix(lng��� - 10, Col����3) = rsTemp!����
            End If
            int�������� = int�������� + Nvl(rsTemp!����, 0)
        End With
        rsTemp.MoveNext
    Wend
    If lblInfo.Caption <> "" And int�������� <> 0 Then
        lblInfo.Caption = lblInfo.Caption & " ���� " & int��������
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitDoctors()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "Select /*+RULE*/" & vbNewLine & _
                "Distinct b.id,b.����, Upper(b.����) As ����" & vbNewLine & _
                " From ������Ա a, ��Ա�� b, ��Ա����˵�� c" & vbNewLine & _
                " Where a.����id = [1] And a.��Աid = b.Id And b.Id = c.��Աid And c.��Ա���� = 'ҽ��' And" & vbNewLine & _
                "      (b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.����ʱ�� Is Null)" & vbNewLine & _
                " Order By ���� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡҽ��", mlng����ID)
    cbo���ռ�ʦ.Clear
    cbo�޼�ҽʦ.Clear
    cbo��Ƭ��ʦ.Clear
    cboȡ�ļ�ʦ.Clear
    While Not rsTemp.EOF
        cbo���ռ�ʦ.AddItem rsTemp!���� & "-" & rsTemp!����
        If rsTemp!ID = UserInfo.ID Then cbo���ռ�ʦ.ListIndex = cbo���ռ�ʦ.NewIndex
        cbo�޼�ҽʦ.AddItem rsTemp!���� & "-" & rsTemp!����
        If rsTemp!ID = UserInfo.ID Then cbo�޼�ҽʦ.ListIndex = cbo�޼�ҽʦ.NewIndex
        cbo��Ƭ��ʦ.AddItem rsTemp!���� & "-" & rsTemp!����
        If rsTemp!ID = UserInfo.ID Then cbo��Ƭ��ʦ.ListIndex = cbo��Ƭ��ʦ.NewIndex
        cboȡ�ļ�ʦ.AddItem rsTemp!���� & "-" & rsTemp!����
        If rsTemp!ID = UserInfo.ID Then cboȡ�ļ�ʦ.ListIndex = cboȡ�ļ�ʦ.NewIndex
        rsTemp.MoveNext
    Wend
    If cbo���ռ�ʦ.ListCount > 0 And cbo���ռ�ʦ.ListIndex = -1 Then cbo���ռ�ʦ.ListIndex = 0
    If cbo�޼�ҽʦ.ListCount > 0 And cbo�޼�ҽʦ.ListIndex = -1 Then cbo�޼�ҽʦ.ListIndex = 0
    If cbo��Ƭ��ʦ.ListCount > 0 And cbo��Ƭ��ʦ.ListIndex = -1 Then cbo��Ƭ��ʦ.ListIndex = 0
    If cboȡ�ļ�ʦ.ListCount > 0 And cboȡ�ļ�ʦ.ListIndex = -1 Then cboȡ�ļ�ʦ.ListIndex = 0
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboCheckType_Click()
    '��ȡ������
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim lngBigNumber As Long
    
    strSQL = "Select ����,������,ǰ����� From Ӱ������� where ���� = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������", CStr(cboCheckType.Text))
    
    If Not rsTemp.EOF Then
        '�ж��Ƿ����ǰ����Ǻ�������
        If IsNull(rsTemp!ǰ�����) Then
            MsgBoxD Me, "�������ֵ��Ӱ������������ò���ŵ�ǰ����ǡ�", vbInformation, gstrSysName
            Exit Sub
        End If
        lngBigNumber = Nvl(rsTemp!������, 0)
        txtPathologyNo.Text = Nvl(rsTemp!ǰ�����) & lngBigNumber
    Else
        txtPathologyNo.Text = ""
    End If
End Sub

Private Sub cmdCheckIn_Click()
    Call CheckInSamples
End Sub

Private Sub CheckInSamples()
    '���ձ걾
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim str����걾�� As String
    Dim i As Integer
    Dim j As Integer
    Dim iSampleCount As Integer
    Dim str����� As String
    Dim str��������� As String
    
    str����� = txtPathologyNo.Text
    str��������� = cboCheckType.Text
    If str����� = "" Or str��������� = "" Then
        MsgBoxD Me, "����Ż��߲�����������벻��ȷ�����顣 ", vbInformation, gstrSysName
        Exit Sub
    Else
        '��鲡��ţ��Ƿ���һ��ǰ���ַ�
        If Not IsNumeric(Mid(str�����, 2)) Or IsNumeric(Left(str�����, 1)) Then
            MsgBoxD Me, "����Ų����ϡ����ǰ���ַ�+���ֱ�š��Ĺ������顣", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    iSampleCount = 0
    For i = 1 To 5
        If vfgList.TextMatrix(i, Col�걾��λ������1) <> "" And vfgList.TextMatrix(i, Col����1) <> "" _
            And Val(vfgList.TextMatrix(i, Col����1)) <> 0 Then
            str����걾�� = str����걾�� & "<" & vfgList.TextMatrix(i, Col���1) & "-" & _
                            vfgList.TextMatrix(i, Col�걾��λ������1) & "-" & _
                            Val(vfgList.TextMatrix(i, Col����1)) & ">"
            iSampleCount = iSampleCount + Val(vfgList.TextMatrix(i, Col����1))
        Else
            Exit For
        End If
    Next i
    
    For i = 1 To 5
        If vfgList.TextMatrix(i, Col�걾��λ������2) <> "" And vfgList.TextMatrix(i, Col����2) <> "" _
            And Val(vfgList.TextMatrix(i, Col����2)) <> 0 Then
            str����걾�� = str����걾�� & "<" & vfgList.TextMatrix(i, Col���2) & "-" & _
                            vfgList.TextMatrix(i, Col�걾��λ������2) & "-" & _
                            Val(vfgList.TextMatrix(i, Col����2)) & ">"
            iSampleCount = iSampleCount + Val(vfgList.TextMatrix(i, Col����2))
        Else
            Exit For
        End If
    Next i
    
    For i = 1 To 5
        If vfgList.TextMatrix(i, Col�걾��λ������3) <> "" And vfgList.TextMatrix(i, Col����3) <> "" _
            And Val(vfgList.TextMatrix(i, Col����3)) <> 0 Then
            str����걾�� = str����걾�� & "<" & vfgList.TextMatrix(i, Col���3) & "-" & _
                            vfgList.TextMatrix(i, Col�걾��λ������3) & "-" & _
                            Val(vfgList.TextMatrix(i, Col����3)) & ">"
            iSampleCount = iSampleCount + Val(vfgList.TextMatrix(i, Col����3))
        Else
            Exit For
        End If
    Next i

    On Error GoTo errHandle
    
    strSQL = "ZL_Ӱ��걾����(" & mlngҽ��ID & "," & mlng���ͺ� & ",'" & NeedName(cbo���ռ�ʦ.Text) & "',sysdate,'" _
                & str��������� & "','" & str����� & "'," & Mid(str�����, 2) & ",'" & str����걾�� & "')"
    zlDatabase.ExecuteProcedure strSQL, "Ӱ����걾����"
    lblInfo.Caption = "�걾�Ѻ���" & " ���� " & iSampleCount
    lblInfo.ForeColor = vbBlue
    
    '��鲡����Ƿ����
    strSQL = "Select ����� From Ӱ��걾����ȡ�� where ҽ��ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡӰ�����", mlngҽ��ID)
    If rsTemp!����� <> str����� Then
        MsgBoxD Me, "ԭ����ţ� " & str����� & " �Ѿ���ʹ�ã��Զ���������޸�Ϊ�� " & rsTemp!�����, vbInformation, gstrSysName
        txtPathologyNo.Text = rsTemp!�����
    End If
    '�����¼�
    RaiseEvent StateChanged(1, txtPathologyNo.Text, str���������)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdRefuse_Click()
    '���ձ걾
    Dim strSQL As String
    
    strSQL = "ZL_Ӱ��걾����(" & mlngҽ��ID & "," & mlng���ͺ� & ",'" & NeedName(cbo���ռ�ʦ.Text) & "',sysdate,'" & txt����ԭ��.Text & "')"
    zlDatabase.ExecuteProcedure strSQL, "Ӱ����걾����"
    lblInfo.Caption = "�걾������"
    lblInfo.ForeColor = vbRed
    RaiseEvent StateChanged(2, "", "")
End Sub

Private Sub CmdSave_Click()
    '����޼��ȡ��
    Dim strSQL As String
    
    '�Ⱥ��ձ걾
    Call CheckInSamples
    
    strSQL = "ZL_Ӱ����޼�ȡ��(" & mlngҽ��ID & "," & mlng���ͺ� & ",'" & txt�޼����.Text & "','" & _
            txt����.Text & "','" & txtʣ��걾λ��.Text & "','" & txt��ע.Text & "','" & NeedName(cbo�޼�ҽʦ.Text) & _
            "',sysdate,'" & NeedName(cboȡ�ļ�ʦ.Text) & "','" & NeedName(cbo��Ƭ��ʦ.Text) & "')"
    zlDatabase.ExecuteProcedure strSQL, "Ӱ����޼�ȡ��"
    
End Sub

Private Sub Form_Load()
    blnInit = True
    vfgList.SelectionMode = flexSelectionFree
    vfgList.Editable = flexEDKbdMouse
End Sub

Private Sub vfgList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '����Ǳ걾��λ��ɾ�����걾��λ����ǰ���ƴ������ĸ
    Dim strTemp As String
    Dim strSQL As String
    
    On Error Resume Next
    
    If Col = Col�걾��λ������1 Or Col = Col�걾��λ������2 Or Col = Col�걾��λ������3 Then
        strTemp = vfgList.TextMatrix(Row, Col)
        If InStr(strTemp, "-") <> 0 Then
            strTemp = Mid(strTemp, InStr(strTemp, "-") + 1)
            vfgList.TextMatrix(Row, Col) = strTemp
        Else
            '�ж��û��Ƿ�����Ӳ���걾��Ȩ�ޣ�����У�����걾�Ƿ���ڣ������������
            If InStr(mstrPrivs, "�걾����") > 0 And Trim(strTemp) <> "" Then
                strSQL = "ZL_Ӱ����걾��λ_Insert('" & strTemp & "','" & zlCommFun.SpellCode(strTemp) & "' )"
                zlDatabase.ExecuteProcedure strSQL, "���²���걾"
            End If
        End If
    End If
End Sub

Private Sub vfgList_DblClick()
    With vfgList
        If .Col = Col���1 Or .Col = Col���2 Or .Col = Col���3 Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vfgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = Col�걾��λ������1 Or Col = Col�걾��λ������2 Or Col = Col�걾��λ������3 Then
'        If KeyAscii <> vbKeyReturn Then KeyAscii = 0
    ElseIf Col = Col���1 Or Col = Col���2 Or Col = Col���3 Then
        KeyAscii = 0
    ElseIf Col = Col����1 Or Col = Col����2 Or Col = Col����3 Then
        If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        End If
    End If
    
    If KeyAscii = vbKeyReturn Then
        If Col = 2 Or Col = 5 Or Col = 8 Then
            vfgList.Select Row, Col + 1
            vfgList.EditCell
        ElseIf Col = 3 Or Col = 6 Or Col = 9 Then
            If Row < 5 Then
                vfgList.Select Row + 1, Col - 1
                vfgList.EditCell
            ElseIf Col <> 9 Then
                vfgList.Select 1, Col + 2
                vfgList.EditCell
            End If
        End If
    End If
End Sub
