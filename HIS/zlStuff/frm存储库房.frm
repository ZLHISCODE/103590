VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frm�洢�ⷿ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�洢�ⷿ����"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   Icon            =   "frm�洢�ⷿ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdStuff 
      Caption         =   "��"
      Height          =   285
      Left            =   8520
      TabIndex        =   23
      TabStop         =   0   'False
      Tag             =   "����"
      ToolTipText     =   "��*��ѡ����"
      Top             =   623
      Width           =   285
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   5280
      MaxLength       =   10
      TabIndex        =   20
      Tag             =   "������"
      Top             =   990
      Width           =   1440
   End
   Begin VB.CommandButton cmdChoose 
      Caption         =   "ȫѡ(&A)"
      Height          =   350
      Left            =   1200
      Picture         =   "frm�洢�ⷿ.frx":000C
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5970
      Width           =   957
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Ӧ���ڱ���(&L)"
      Height          =   350
      Left            =   3400
      Picture         =   "frm�洢�ⷿ.frx":0156
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "���ڷ������ѡ���˿��Һ󣬽������������ѡ���ֵӦ��ǰ�湴ѡ�˵����У�"
      Top             =   5970
      Width           =   1395
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "�����������(&Z)"
      Height          =   350
      Left            =   5760
      Picture         =   "frm�洢�ⷿ.frx":02A0
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "�洢�ⷿ�ͷ�����ҵ����ò������һ�ν�������!"
      Top             =   5970
      Width           =   1572
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4635
      Left            =   -6135
      TabIndex        =   18
      Top             =   6000
      Visible         =   0   'False
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Frame frame 
      Caption         =   "Ӧ����ͬԺ��(&B)"
      Height          =   1050
      Left            =   135
      TabIndex        =   4
      Top             =   4800
      Width           =   9180
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ���ڱ�Ʒ�����й��(&2)"
         Height          =   255
         Index           =   4
         Left            =   2985
         TabIndex        =   6
         Top             =   285
         Width           =   2730
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ���ڴ˷����µ�������������(&4)"
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   8
         Top             =   630
         Width           =   4320
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ���ڱ���������������(&3)"
         Height          =   255
         Index           =   1
         Left            =   5880
         TabIndex        =   7
         Top             =   285
         Width           =   3045
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "��Ӧ���ڱ���������(&1)"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   5
         Top             =   285
         Value           =   -1  'True
         Width           =   4065
      End
      Begin VB.OptionButton optӦ���� 
         Caption         =   "Ӧ�������С��������ϡ�(&5)"
         Height          =   255
         Index           =   3
         Left            =   5880
         TabIndex        =   9
         Top             =   630
         Width           =   3045
      End
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "�ָ�(&R)"
      Height          =   350
      Left            =   4800
      Picture         =   "frm�洢�ⷿ.frx":03EA
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5970
      Width           =   957
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ȫ��(&C)"
      Height          =   350
      Left            =   2250
      Picture         =   "frm�洢�ⷿ.frx":0534
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5970
      Width           =   957
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����(&S)"
      Height          =   350
      Left            =   7350
      TabIndex        =   10
      Top             =   5970
      Width           =   957
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      Picture         =   "frm�洢�ⷿ.frx":067E
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5970
      Width           =   885
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8325
      TabIndex        =   11
      Top             =   5970
      Width           =   957
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2985
      Top             =   5610
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�洢�ⷿ.frx":07C8
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�洢�ⷿ.frx":0D62
            Key             =   "ItemStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm�洢�ⷿ.frx":12FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit mshBillEdit 
      Height          =   3180
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   5609
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.TextBox txtStuff 
      Height          =   300
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   2
      Top             =   615
      Width           =   6855
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frm�洢�ⷿ.frx":3006
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblFindComment 
      Caption         =   "������룬���ƣ����룬�����س����в��ң��������Ұ�F3��"
      Height          =   375
      Left            =   6840
      TabIndex        =   22
      Top             =   980
      Width           =   2535
   End
   Begin VB.Label lblFind 
      AutoSize        =   -1  'True
      Caption         =   "���ҿⷿ"
      Height          =   180
      Left            =   4440
      TabIndex        =   21
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "ָ�������ķ��ࣺ"
      Height          =   180
      Left            =   240
      TabIndex        =   19
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "ָ����������(&M)"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   675
      Width           =   1350
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    ��ѡ���������Ϻ����ø��������ϵĴ洢�ⷿ�ķ�����ҡ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   432
      TabIndex        =   0
      Top             =   240
      Width           =   5244
   End
End
Attribute VB_Name = "frm�洢�ⷿ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean
Private mbln�༭ As Boolean
Private mlng����ID As Long                  '������ĿID
Private Const mlngModule = 1711
Private mstrPreStuffName As String          '�ϴ�ѡ����ı�����
Private mlngLastFindRows As Long            '�ϴ��ҵ����к�
Private mstrFind As String
Private mrs���� As ADODB.Recordset
Private mblnFind As Boolean             '�Ƿ��ǵ�һ�β�ѯ
Private mstr��¼ֵ As String            '��¼���е�ֵ
Private mblnSave As Boolean             '�Ƿ񱣴�ɹ�   true ����ɹ������Ѿ����� false δ����ɹ���δ����
Private mblnKeyMethod As Boolean        '�Ƿ���ͨ���༭�س���ʽ

'Private Sub FindRow(ByVal strInput As String, ByVal lngStartRows As Long)
'    Dim intTargetRow As Integer
'    Dim lngRows As Long
'    Dim blnFind As Boolean
'    Dim strMatch As String
'
'    If strInput = "" Then Exit Sub
'
'    '������ƥ�䣬���ƺͼ���˫��ƥ��
'    If IsNumeric(strInput) Then
'        intTargetRow = 5
'        strMatch = strInput
'    ElseIf zlStr.IsCharAlpha (strInput) Then
'        intTargetRow = 6
'        strMatch = "*" & strInput & "*"
'    Else
'        intTargetRow = 2
'        strMatch = "*" & strInput & "*"
'    End If
'
'    With mshBillEdit
'        If .Rows = 1 Then Exit Sub
'
'        If lngStartRows > .Rows - 1 Then
'            MsgBox "�Ѳ�ѯ�����", vbInformation, gstrSysName
'            lngRows = 1
'        Else
'            lngRows = lngStartRows
'        End If
'
'        .SetFocus
'
'        '��ָ���Ŀ�ʼλ�ÿ�ʼ����
'        For lngRows = lngRows To .Rows - 1
'            If .TextMatrix(lngRows, intTargetRow) Like strMatch Then
'                .MsfObj.TopRow = lngRows
'                .Row = lngRows
'                .Col = 1
'                mlngLastFindRows = lngRows + 1
'                blnFind = True
'                Exit For
'            End If
'        Next
'
'        'û�ҵ��ʹ�ͷ����һ��
'        If Not blnFind And lngStartRows > 1 Then
'            For lngRows = 1 To .Rows - 1
'                If .TextMatrix(lngRows, intTargetRow) Like strInput & "*" Then
'                    .MsfObj.TopRow = lngRows
'                    .Row = lngRows
'                    .Col = 1
'                    mlngLastFindRows = lngRows + 1
'                    Exit For
'                End If
'            Next
'        End If
'    End With
'End Sub

Private Function SelectStuff(ByVal strSeach As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ������������
    '����:strKey-��ѡ�������
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strWhere As String
    Dim objCtl As Object: Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
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
    
    On Error GoTo ErrHand
    Set objCtl = txtStuff
    vRect = zlControl.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    strKey = GetMatchingSting(strSeach)
      
    strTittle = "���в���ѡ��"
    If strSeach = "" Then

        gstrSQL = " " & _
        "   Select Id,�ϼ�id,����,����,'' As ���㵥λ,0 as ĩ�� ,'' վ��,'' ����ʱ��" & _
        "   From ���Ʒ���Ŀ¼  " & _
        "   Where ����='7' " & _
        "   Start With �ϼ�id Is Null Connect By Prior Id =�ϼ�id " & _
        "   Union All  " & _
        "   Select I.ID,B.����ID As �ϼ�ID, I.����, I.���� || LPad(' ', 3, ' ') || I.��� || LPad(' ', 3, ' ') || I.���� As ����" & _
        "       ,I.���㵥λ,1 as  ĩ��,I.վ��,to_char(I.����ʱ��,'yyyy-mm-dd') as ����ʱ�� " & _
        "   From �շ���ĿĿ¼ I, �������� T,������ĿĿ¼ B " & _
        "   Where I.ID = T.����id And   T.����id=b.Id     And I.��� = '4' And (I.����ʱ�� Is Null Or I.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 2, strTittle, True, "", "", False, False, False, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False)
    Else
        
        strWhere = " And (I.���� Like [1] OR N.���� Like [1] OR ( N.���� LIKE Upper([1]) and N.����=[2]))"
        If IsNumeric(strSeach) Then                         '���������,��ֻȡ����
            If Mid(gSystem_Para.Para_���뷽ʽ, 1, 1) = "1" Then strWhere = " And (I.���� Like [1] And N.����=[2])"
        ElseIf zlStr.IsCharAlpha(strSeach) Then          '����ȫ����ĸʱֻƥ�����
            If Mid(gSystem_Para.Para_���뷽ʽ, 2, 1) = "1" Then strWhere = " And (N.���� Like Upper([1]) And N.����=[2]) "
        ElseIf zlStr.IsCharChinese(strSeach) Then
            strWhere = " And (N.���� Like [1] And N.����=[2]) "
        End If
    
        gstrSQL = "" & _
        "   Select distinct I.ID,I.����, I.���� ||LPAD(' ',3,' ')||I.���||LPAD(' ',3,' ')||I.���� as ����,I.���㵥λ,I.վ��,to_char(I.����ʱ��,'yyyy-mm-dd') as ����ʱ��" & _
        "   From �շ���ĿĿ¼ I,�������� T,�շ���Ŀ���� N" & _
        "   Where I.ID=T.����ID and I.ID=N.�շ�ϸĿID and I.���='4'" & _
        "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
        "        " & strWhere
        
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, gSystem_Para.int���뷽ʽ + 1)
    End If
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    '���ز���
    If rsTemp Is Nothing Then
        ShowMsgBox "û��������������������,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    With rsTemp
        mstrPreStuffName = "[" & !���� & "]" & !����:
        objCtl.Text = mstrPreStuffName
        objCtl.Tag = zlStr.Nvl(!Id)
        mlng����ID = !Id
    End With
    Call ShowData
    SelectStuff = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function



Private Sub cmdApply_Click()
    Dim lngRow As Long, lngRowCurr As Long
    Dim strObject As String, strObjectID As String
    
    With mshBillEdit
        lngRowCurr = .Row
        strObject = .TextMatrix(.Row, 3)
        strObjectID = .TextMatrix(.Row, 4)
        For lngRow = 1 To .Rows - 1
            If lngRow <> lngRowCurr And .TextMatrix(lngRow, 1) = "��" Then
                .TextMatrix(lngRow, 3) = strObject
                .TextMatrix(lngRow, 4) = strObjectID
            End If
        Next
    End With
End Sub

Private Sub cmdChoose_Click()
    Dim lngRow As Long, lngRows As Long
    With mshBillEdit
        lngRows = .Rows - 1
        For lngRow = 1 To lngRows
            .TextMatrix(lngRow, 1) = "��"
'            .TextMatrix(lngRow, 3) = ""
'            .TextMatrix(lngRow, 4) = ""
        Next
    End With
End Sub

Private Sub cmdClear_Click()
    Dim lngRow As Long, lngRows As Long
    With mshBillEdit
        lngRows = .Rows - 1
        For lngRow = 1 To lngRows
            .TextMatrix(lngRow, 1) = ""
            .TextMatrix(lngRow, 3) = ""
            .TextMatrix(lngRow, 4) = ""
        Next
    End With
End Sub

Private Sub cmdClose_Click()
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    
    If mblnSave = False Then
        With mshBillEdit
            For i = 1 To .Rows - 1
                For j = 1 To .Cols - 1
                    strTemp = strTemp & .TextMatrix(i, j) & "|"
                Next
            Next
        End With
        strTemp = txtStuff.Text & "|" & optӦ����(0).Value & "|" & optӦ����(4).Value & "|" & optӦ����(1).Value & "|" & _
                    optӦ����(2).Value & "|" & optӦ����(3).Value & "|" & strTemp
                    
        If strTemp <> mstr��¼ֵ Then
            If MsgBox("�����ݱ��޸��ˣ��Ƿ��˳���", vbYesNo, gstrSysName) = vbYes Then
                Unload Me
            End If
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdStuff_Click()
    If SelectStuff("") = False Then Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call ShowData
End Sub

Private Sub CmdSave_Click()
    Dim strPara As String
    Dim lngRow As Long, lngRows As Long
    Dim rsTemp As New ADODB.Recordset
    Dim arrInput As Variant
    Dim strTmp As String
    Dim intType As Integer
    Dim i As Integer
    
    arrInput = Array()
    
    On Error GoTo ErrHand
    
    If mlng����ID = 0 Then
        MsgBox "����ѡ���������ϣ�", vbInformation, gstrSysName
        txtStuff.SetFocus
        Exit Sub
    End If
    
    If mshBillEdit.Active = False Then
        MsgBox "û���ҵ��κοⷿ�����ڲ��Ź��������ã�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If optӦ����(0).Value = False Then
        For i = 0 To optӦ����.UBound
            If optӦ����(i).Value = True Then
                If MsgBox("�����Ĵ洢�ⷿӦ�÷�ΧΪ��" & optӦ����(i).Caption & "���Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    Exit For
                End If
            End If
        Next
    End If
    
    '�������봮������
    lngRows = mshBillEdit.Rows - 1
    For lngRow = 1 To lngRows
        If mshBillEdit.TextMatrix(lngRow, 1) <> "" Then
            strTmp = "1," & mshBillEdit.RowData(lngRow) & "|" & mshBillEdit.TextMatrix(lngRow, 4)
        Else
            strTmp = "0," & mshBillEdit.RowData(lngRow) & "|" & ""
        End If
        
        '�������ڷ�����ҹ��࣬��ɴ���Ĳ����ַ����ȹ��������������4K���ַ���
        If Len(IIf(strPara = "", "", strPara & "!!") & strTmp) > 4000 Then
            ReDim Preserve arrInput(UBound(arrInput) + 1)
            arrInput(UBound(arrInput)) = strPara
            
            strPara = strTmp
        Else
            strPara = IIf(strPara = "", "", strPara & "!!") & strTmp
        End If
    Next
    
    ReDim Preserve arrInput(UBound(arrInput) + 1)
    arrInput(UBound(arrInput)) = strPara
    
    '������
    '   ����ID_IN����_IN(���ñ�־,�ⷿID|����ID,����ID!!...),Ӧ��_IN(1-����������;2-����������������;3-�����µ�������������;4-���в���,5-��Ʒ�����й��)
    '���ñ�־:0-�����������ڸÿⷿ�Ĵ洢��1-Ҫ����
    If optӦ����(0).Value = True Then
        intType = 1
    ElseIf optӦ����(1).Value = True Then
        intType = 2
    ElseIf optӦ����(2).Value = True Then
        intType = 3
    ElseIf optӦ����(4).Value = True Then
        intType = 5
    Else
        intType = 4
    End If
    
    For i = 0 To UBound(arrInput)
        If arrInput(i) <> "" Then
            gstrSQL = "zl_�������ϴ洢�ⷿ_UPDATE(" & mlng����ID & ",'" & arrInput(i) & "'," & intType & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Next
    
    MsgBox "���������ϵĴ洢�ⷿ�ͷ�����ұ���ɹ���", vbInformation, gstrSysName
    mblnSave = True
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub



Private Sub cmd����_Click()
    Dim str���� As String
    Dim str����id As String
    Dim blnSel As Boolean
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "" & _
        "   Select  ��������id, ��������id, ִ�п���id ,K.���� " & _
        "   From �շ�ִ�п��� A," & _
        "       (Select �շ�ϸĿid From �շ�ִ�п��� Where Rowid = " & _
        "           (Select Max(a.Rowid) From �շ�ִ�п��� a,�������� c where a.�շ�ϸĿid=c.����id)" & _
        "       ) B ,���ű� K" & _
        "   Where A.�շ�ϸĿid = b.�շ�ϸĿid and a.��������id=K.id(+) "
    
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    With mshBillEdit
        For intRow = 1 To .Rows - 1
            
            str���� = "": str����id = ""
            rsTemp.Filter = "ִ�п���ID=" & .RowData(intRow)
            
            blnSel = False
            Do While Not rsTemp.EOF
                blnSel = True
                str���� = str���� & "," & zlStr.Nvl(rsTemp!����)
                str����id = str����id & "," & zlStr.Nvl(rsTemp!��������id, 0)
                rsTemp.MoveNext
            Loop
            If str���� <> "" Then
                str���� = Mid(str����, 2)
                str����id = Mid(str����id, 2)
                If str����id = "0" Then str����id = ""
            End If
            mshBillEdit.TextMatrix(intRow, 0) = intRow
            If blnSel Then
                mshBillEdit.TextMatrix(intRow, 1) = "��"
            Else
                mshBillEdit.TextMatrix(intRow, 1) = ""
            End If
            mshBillEdit.TextMatrix(intRow, 3) = str����
            mshBillEdit.TextMatrix(intRow, 4) = str����id
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If lvwItems.Visible Then
            lvwItems.Visible = False
            mshBillEdit.SetFocus
            Exit Sub
        Else
            Unload Me
            Exit Sub
        End If
    End If
    
    If KeyCode = vbKeyF3 Then
        Call txtfind_KeyPress(vbKeyReturn)
    End If
        
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim strTemp As String
    Dim j As Integer
    
    mlngLastFindRows = 1
    
    Call InitFace
    Call ShowData
    If mbln�༭ = False Then
        mshBillEdit.Active = False
        frame.Enabled = False
        cmdSave.Visible = False
        cmdRestore.Visible = False
        cmdClear.Visible = False
        cmdClose.Caption = "�˳�(&X)"
    End If
    Call CtlEnableSet
    
    With mshBillEdit
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
                strTemp = strTemp & .TextMatrix(i, j) & "|"
            Next
        Next
    End With
    mstr��¼ֵ = ""
    mstr��¼ֵ = txtStuff.Text & "|" & optӦ����(0).Value & "|" & optӦ����(4).Value & "|" & optӦ����(1).Value & "|" & _
                optӦ����(2).Value & "|" & optӦ����(3).Value & "|" & strTemp
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnFind = False
    mblnSave = False
End Sub

Private Sub mshBillEdit_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    mshBillEdit.TextMatrix(Row, 1) = ""
    mshBillEdit.TextMatrix(Row, 3) = ""
    mshBillEdit.TextMatrix(Row, 4) = ""
    Cancel = True
End Sub

Private Sub mshBillEdit_CommandClick()
    Dim str������� As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    cmdSave.Enabled = True
    str������� = ""
    gstrSQL = "select distinct ������� from ��������˵�� where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������", Val(mshBillEdit.RowData(mshBillEdit.Row)))
    
    Do While Not rsTemp.EOF
        str������� = str������� & "," & rsTemp!�������
        rsTemp.MoveNext
    Loop
    
    If str������� <> "" Then
        str������� = Mid(str�������, 2)
        If InStr(1, str�������, 3) <> 0 Then
            str������� = "0,1,2,3"
        ElseIf InStr(1, str�������, 1) <> 0 Or InStr(1, str�������, 2) <> 0 Then
            str������� = str������� & ",3"
        End If
    Else
        str������� = "0"
    End If
    
    '�ų������ֲ��ǿ������ҵĲ�������
    gstrSQL = " Select distinct ID,����,���� From ���ű� A,��������˵�� B, Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) C " & _
              " Where A.ID=B.����ID And B.�������=C.Column_Value and (a.����ʱ�� is null or to_char(a.����ʱ��,'yyyy-mm-dd')='3000-01-01')" & _
              " And B.�������� Not In ('����ⷿ', '��������', '��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��', '���Ŀ�', '���ϲ���') "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������", str�������)

    If rsTemp.RecordCount = 0 Then
        MsgBox "δ�����ٴ���ҽ���ȿ��ң�[���Ź���]", vbInformation, gstrSysName
        mshBillEdit.TextMatrix(mshBillEdit.Row, 3) = ""
        mshBillEdit.TextMatrix(mshBillEdit.Row, 4) = ""
    End If
    
    Call AddColumnHeader(False)
    Me.lvwItems.ListItems.Clear
    Me.lvwItems.Checkboxes = True
    
    Do While Not rsTemp.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTemp!Id, rsTemp!����, , 3)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = rsTemp!����
        If InStr(1, "," & mshBillEdit.TextMatrix(mshBillEdit.Row, 4) & ",", "," & rsTemp!Id & ",") > 0 Then
            objItem.Checked = True
        End If
        rsTemp.MoveNext
    Loop
    If lvwItems.ListItems.Count <> 0 Then lvwItems.ListItems(1).Selected = True
    With Me.lvwItems
        .Left = Me.mshBillEdit.Left + 2300
        .Top = Me.mshBillEdit.Top + Me.mshBillEdit.CellTop + Me.mshBillEdit.RowHeight(Me.mshBillEdit.Row)
        If .Top + .Height > Me.ScaleHeight Then
            .Top = Me.ScaleHeight - .Height
        End If
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBillEdit_EnterCell(Row As Long, Col As Long)
    With mshBillEdit
        If .Col = 3 Then
            .ColData(3) = 1
        End If
    End With
End Sub


Private Sub mshBillEdit_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim str������� As String
    Dim objItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    
    With mshBillEdit
'            mblnKeyMethod = False
        If KeyCode = vbKeyReturn And .Col = 3 Then
'            mblnKeyMethod = True
'            Call mshBillEdit_CommandClick
            If Trim(.Text) = "" Then
                If .Row <> .Rows - 1 Then
                    .Row = .Row + 1
                End If
                Exit Sub
            End If
            cmdSave.Enabled = True
            str������� = ""
            gstrSQL = "select distinct ������� from ��������˵�� where ����ID=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������", Val(mshBillEdit.RowData(mshBillEdit.Row)))
            
            Do While Not rsTemp.EOF
                str������� = str������� & "," & rsTemp!�������
                rsTemp.MoveNext
            Loop
            If str������� <> "" Then
                str������� = Mid(str�������, 2)
                If InStr(1, str�������, 3) <> 0 Then str������� = "0,1,2,3"
            Else
                str������� = "0"
            End If
            gstrSQL = " Select /*+ Rule*/ distinct ID,����,���� From ���ű� A,��������˵�� B, Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) C " & _
                      " Where A.ID=B.����ID And B.�������=C.Column_Value"
            gstrSQL = gstrSQL & " and (���� like [2] or ���� like [2] or ���� like [2])"
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������", str�������, UCase(.Text) & "%")

            If rsTemp.RecordCount = 0 Then
                MsgBox "δ���ø��ٴ����ң�[���Ź���]", vbInformation, gstrSysName
                mshBillEdit.TextMatrix(mshBillEdit.Row, 3) = ""
                mshBillEdit.TextMatrix(mshBillEdit.Row, 4) = ""
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            If rsTemp.RecordCount = 1 Then
                .TextMatrix(mshBillEdit.Row, 1) = "��"
                .TextMatrix(.Row, 4) = rsTemp!Id
                .Text = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(.Row, 3) = .Text
            Else
                Call AddColumnHeader(False)
                Me.lvwItems.ListItems.Clear
                Me.lvwItems.Checkboxes = True
                
                Do While Not rsTemp.EOF
                    Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsTemp!Id, rsTemp!����, , 3)
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = rsTemp!����
                    If InStr(1, "," & mshBillEdit.TextMatrix(mshBillEdit.Row, 4) & ",", "," & rsTemp!Id & ",") > 0 Then
                        objItem.Checked = True
                    End If
                    rsTemp.MoveNext
                Loop
                lvwItems.ListItems(1).Checked = True
                With Me.lvwItems
                    .Left = Me.mshBillEdit.Left + 2300
                    .Top = Me.mshBillEdit.Top + Me.mshBillEdit.CellTop + Me.mshBillEdit.RowHeight(Me.mshBillEdit.Row)
                    If .Top + .Height > Me.ScaleHeight Then
                        .Top = Me.ScaleHeight - .Height
                    End If
                    .ZOrder 0: .Visible = True
                    Cancel = True
                    .SetFocus
                End With
            End If
        End If
    End With
End Sub

Private Sub optӦ����_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optӦ����.UBound
        If i = Index Then
            optӦ����(i).FontBold = True
        Else
            optӦ����(i).FontBold = False
        End If
    Next
End Sub

Private Sub optӦ����_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call OS.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtFind_Change()
    mblnFind = False
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtfind_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeyReturn Then Exit Sub
'    If Trim(txtFind.Text) = "" Then Exit Sub
'
'    mlngLastFindRows = 1
'    FindRow Trim(txtFind.Text), 1
'    txtFind.SetFocus
'    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    
    On Error GoTo ErrHandle
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then
        If Trim(txtFind.Text) = "" Then Exit Sub
        If mstrFind <> Trim(txtFind.Text) And mblnFind = False Then
            gstrSQL = "Select a.ID,a.����, a.����, a.����" & _
                       " From ���ű� A, ��������˵�� B" & _
                       " Where a.Id = b.����id And b.�������� In ('���Ŀ�', '���ϲ���', '�Ƽ���', '����ⷿ') And (a.���� Like [1] Or a.���� Like [1] Or a.���� like [1])"

            Set mrs���� = zlDatabase.OpenSQLRecord(gstrSQL, "���Ҳ�ѯ", UCase(txtFind.Text) & "%")
            If mrs����.RecordCount > 0 Then
                mblnFind = True
                Call FindData(mrs����)
            Else
                MsgBox "û���ҵ�����Ҫ�����ݣ�", vbInformation, gstrSysName
                txtFind.SetFocus
                zlControl.TxtSelAll txtFind
            End If
        Else
            If Not mrs����.EOF Then
                mrs����.MoveNext
                If Not mrs����.EOF Then
                    Call FindData(mrs����)
                Else
                    MsgBox "�Ѳ�ѯ�����", vbInformation, gstrSysName
                End If
            Else
                mrs����.MoveFirst
                Call FindData(mrs����)
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txtStuff_Change()
    txtStuff.Tag = ""
End Sub

Private Sub txtStuff_GotFocus()
    zlControl.TxtSelAll txtStuff
    OS.OpenIme False
End Sub

Private Sub txtStuff_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtStuff.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If txtStuff.Text = "" Then OS.PressKey vbKeyTab: Exit Sub
    strKey = Trim(Me.txtStuff.Text)
    If SelectStuff(strKey) = False Then Exit Sub
End Sub
Private Sub txtStuff_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub
Private Sub txtStuff_LostFocus()
    Me.txtStuff.Text = mstrPreStuffName
End Sub

Private Sub ShowData()
    '��ȡ���ݲ���ʾ����
    Dim str�ⷿID As String, str���� As String, str����id As String
    Dim intRow As Integer, intRows As Integer
    Dim blnSel As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    
    Dim lng����id As Long, strWhere As String
    Dim strվ�� As String
    Dim lng������Ŀid As Long
    
    Call cmdClear_Click
    On Error GoTo ErrHandle
    
    gstrSQL = "Select ����ID,id From ������ĿĿ¼ where id =(Select Max(����ID) ID from �������� where ����id=[1]) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", mlng����ID)
    If rsTemp.EOF Then
        lng����id = 0
    Else
        lng����id = Val(zlStr.Nvl(rsTemp!����id))
        lng������Ŀid = rsTemp!Id
    End If
    
    gstrSQL = "Select A.����,A.���� From ���Ʒ���Ŀ¼ A where id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", lng����id)
   
    
    If rsTemp.EOF Then
        lbl����.Caption = "ָ�������ķ��ࣺ"
    Else
        lbl����.Caption = "ָ�������ķ��ࣺ" & "[" & rsTemp!���� & "]" & rsTemp!����
    End If
    
'    If mblnFirst Then
    '��ȡ������Ϣ
    strWhere = "": strվ�� = ""
    If mlng����ID <> 0 Then
        gstrSQL = " Select A.ID, A.����, A.���� ||LPAD(' ',3,' ')||A.���||LPAD(' ',3,' ')||A.���� as ����,A.վ��" & _
                  " From �շ���ĿĿ¼ A" & _
                  " Where A.ID=[1]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", mlng����ID)
        If rsTemp.EOF = False Then
            txtStuff.Text = "[" & rsTemp!���� & "]" & rsTemp!����
            mstrPreStuffName = txtStuff.Text

            txtStuff.Tag = zlStr.Nvl(rsTemp!Id)
            strվ�� = zlStr.Nvl(rsTemp!վ��)
        End If
        If strվ�� <> "" Then
            strWhere = " And (վ��=[1] or վ�� is null)" '
            lbl����.Caption = lbl����.Caption & "    վ�㣺" & strվ��
        End If
    End If
        
    '���ݲ��ϵ���;������ȡ������洢�Ŀⷿ
    gstrSQL = "" & _
        " Select ID,����,����,���� From ���ű� " & _
        " Where ID in (Select distinct ����id from ��������˵�� where �������� In('���Ŀ�','���ϲ���','�Ƽ���','����ⷿ')) and (����ʱ�� is null or to_char(����ʱ��,'yyyy-mm-dd')='3000-01-01')  " & _
        strWhere
    gstrSQL = gstrSQL & " Order By ���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������������ȡ������洢�Ŀⷿ", strվ��)
    
    With mshBillEdit
        .Rows = 2:
        If rsTemp.EOF Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
        End If
        Do While Not rsTemp.EOF
            .TextMatrix(.Rows - 1, 2) = rsTemp!����
            .TextMatrix(.Rows - 1, 5) = rsTemp!����
            .TextMatrix(.Rows - 1, 6) = rsTemp!����
            .RowData(.Rows - 1) = rsTemp!Id
            .Rows = .Rows + 1
            str�ⷿID = str�ⷿID & "," & rsTemp!Id
            rsTemp.MoveNext
        Loop
        If str�ⷿID <> "" Then
            str�ⷿID = Mid(str�ⷿID, 2)
            .Rows = .Rows - 1
            .Active = True
        Else
            .Active = False
        End If
    End With
    'End If
    
    'ȡ���пⷿ
    '    str�ⷿID = ""
    intRows = mshBillEdit.Rows - 1
    '    For intRow = 1 To intRows
    '        str�ⷿID = str�ⷿID & "," & mshBillEdit.RowData(intRow)
    '    Next
    '
    '    If str�ⷿID <> "" Then str�ⷿID = Mid(str�ⷿID, 2)
    
    
    '����Ӧ������֯��װ�뵥�ݿؼ�
    gstrSQL = "" & _
        "   Select A.�շ�ϸĿID,A.��������ID,A.ִ�п���ID,B.���� " & _
        "   From �շ�ִ�п��� A,���ű� B " & _
        "   Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=[1] And A.ִ�п���ID in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList)))" & _
        "   Order by A.ִ�п���ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����õ�����ִ�п�������", mlng����ID, str�ⷿID)
       
    If rsTemp.RecordCount = 0 And mlng����ID <> 0 Then
        '
        gstrSQL = "" & _
            " Select a.�շ�ϸĿid,a.��������id, a.ִ�п���id, b.����" & _
            " From �շ�ִ�п��� A, ���ű� B, (Select ����id From �������� Where ����id = [1] And Rownum < 2) C " & _
           " Where a.��������id = b.Id(+) And a.�շ�ϸĿid = c.����id And a.ִ�п���id in (select * from Table(Cast(f_Num2List([2]) As zlTools.t_NumList)))" & _
            " Order By a.ִ�п���id"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng������Ŀid, str�ⷿID)
            
        If rsTemp.RecordCount = 0 And mlng����ID <> 0 Then
            gstrSQL = "" & _
                "   Select A.�շ�ϸĿID,A.��������ID,A.ִ�п���ID,B.���� " & _
                "   From �շ�ִ�п��� A,���ű� B, " & _
                "       ( Select A.ID From �շ���ĿĿ¼ A,�������� B,�շ�ִ�п��� C " & _
                "         Where A.ID=B.����ID And A.���='4' And A.ID=C.�շ�ϸĿID And Rownum<2 ) C" & _
                "   Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=C.ID And A.ִ�п���ID in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList)))" & _
                "   Order by A.ִ�п���ID"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str�ⷿID)
            If rsTemp.RecordCount <> 0 Then
                MsgBox "��ǰ�������δ���ô洢�ⷿ����ȡ������ͬ�Ĺ�����ĵĴ洢�ⷿ��Ϊȱʡ���ݣ�", vbInformation, gstrSysName
            End If
        Else
            MsgBox "��ǰ�������δ���ô洢�ⷿ����ȡͬƷ���¹�����ĵĴ洢�ⷿ��Ϊȱʡ���ݣ�", vbInformation, gstrSysName
        End If
    End If
    With mshBillEdit
        For intRow = 1 To intRows
            str���� = "": str����id = ""
            rsTemp.Filter = "ִ�п���ID=" & .RowData(intRow)
            rsTemp.Sort = "��������ID"
            blnSel = False
            Do While Not rsTemp.EOF
                blnSel = True
                str���� = str���� & "," & zlStr.Nvl(rsTemp!����)
                str����id = str����id & "," & Nvl(rsTemp!��������id, 0)
                rsTemp.MoveNext
            Loop
            If str���� <> "" Then
                str���� = Mid(str����, 2)
                str����id = Mid(str����id, 2)
                If str����id = "0" Then str����id = ""
            End If
            .TextMatrix(intRow, 0) = intRow
            If blnSel Then .TextMatrix(intRow, 1) = "��"
            .TextMatrix(intRow, 3) = str����
            .TextMatrix(intRow, 4) = str����id
        Next
    End With
    mblnFirst = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitFace()
    '��ʼ���ؼ�
    With mshBillEdit
        .Rows = 2
        .Cols = 7
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "ѡ��"
        .TextMatrix(0, 2) = "�洢�ⷿ"
        .TextMatrix(0, 3) = "�������"
        .TextMatrix(0, 4) = "�������ID"
        .TextMatrix(0, 5) = "�ⷿ����"
        .TextMatrix(0, 6) = "�ⷿ����"
        .TextMatrix(1, 0) = "1"
        .ColData(0) = 5
        .ColData(1) = -1
        .ColData(2) = 5
        .ColData(3) = 1
        .ColData(4) = 5
        .ColData(5) = 5
        .ColData(6) = 5
        .ColWidth(0) = 300
        .ColWidth(1) = 500
        .ColWidth(2) = 1500
        .ColWidth(3) = 5000
        .ColWidth(4) = 0
        .ColWidth(5) = 0
        .ColWidth(6) = 0
        
        .PrimaryCol = 1
        .LocateCol = 1
        .AllowAddRow = False
        .Active = True
    End With
End Sub

Private Sub mshBillEdit_AfterAddRow(Row As Long)
    Dim lngCurRow As Long
    
    '�޸������
    With mshBillEdit
        For lngCurRow = Row To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub mshBillEdit_AfterDeleteRow()
    Dim lngCurRow As Long
    '�޸������
    With mshBillEdit
        For lngCurRow = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub
Private Sub AddColumnHeader(Optional ByVal bln���� As Boolean = True)
    If bln���� Then
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "����", "����", 3000
            .Add , "����", "����", 1000
            .Add , "���㵥λ", "���㵥λ", 800
        End With
        With Me.lvwItems
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 1
            .SortOrder = lvwAscending
        End With
        lvwItems.Tag = "1"
    Else
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "����", "����", 3000
            .Add , "����", "����", 1000
        End With
        With Me.lvwItems
            .ColumnHeaders("����").Position = 1
            .SortKey = .ColumnHeaders("����").Index - 1
            .SortOrder = lvwAscending
        End With
        lvwItems.Tag = "2"
    End If
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim blnCancel As Boolean
    Dim lngRow As Long, lngRows As Long
    Dim str���� As String, str����id As String
     
    If lvwItems.Tag = "1" Then
'        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
'        With Me.lvwItems
'            If mlng����ID <> Mid(.SelectedItem.Key, 2) Then
'                mlng����ID = Mid(.SelectedItem.Key, 2)
'                Me.txtStuff.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
'                Me.txtStuff.Text = Me.txtStuff.Tag
'                Call ShowData
'            End If
'            Me.txtStuff.SetFocus
'            Call OS.PressKey(vbKeyTab)
'        End With
    Else
        'ѭ����ȡ�û���ѡ��Ŀ���
        lngRows = lvwItems.ListItems.Count
        For lngRow = 1 To lngRows
            If lvwItems.ListItems(lngRow).Checked Then
                str���� = str���� & "," & lvwItems.ListItems(lngRow).Text
                str����id = str����id & "," & Mid(lvwItems.ListItems(lngRow).Key, 2)
            End If
        Next
        If str���� <> "" Then
            str���� = Mid(str����, 2)
            str����id = Mid(str����id, 2)
        End If
        
        If str���� <> "" Then mshBillEdit.TextMatrix(mshBillEdit.Row, 1) = "��"
        mshBillEdit.TextMatrix(mshBillEdit.Row, 3) = str����
        mshBillEdit.TextMatrix(mshBillEdit.Row, 4) = str����id
        mshBillEdit.SetFocus
        If mshBillEdit.Rows - 1 > mshBillEdit.Row Then mshBillEdit.Row = mshBillEdit.Row + 1
    End If
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Public Sub ShowMe(ByVal frmParent As Object, ByVal lng����ID As Long, ByVal bln�༭ As Boolean)
    On Error Resume Next
    mblnFirst = True
    mlng����ID = lng����ID
    mbln�༭ = bln�༭
    Me.Show 1, frmParent
End Sub
Private Sub CtlEnableSet()
    '����:������ؿؼ���Enable
    '---------------------------------------------------------------------------------------------------------------------
    Dim strReg As String
    err = 0: On Error GoTo ErrHand:
    '��ʽ:3λ�ַ�����,1,��������0��������,��111.���е�һλ��������,�ڶ�λ����������,����λ�������������
    strReg = zlDatabase.GetPara("����Ӧ���ڵķ�Χ", glngSys, mlngModule)
    If Len(strReg) < 3 Then
        '����ȫѡ��
        strReg = "111"
    End If
    optӦ����(3).Enabled = IIf(Val(Mid(strReg, 1, 1)) = 1, True, False)
    optӦ����(1).Enabled = IIf(Val(Mid(strReg, 2, 1)) = 1, True, False)
    optӦ����(2).Enabled = IIf(Val(Mid(strReg, 3, 1)) = 1, True, False)
    If optӦ����(1).Enabled = False And optӦ����(1).Value = True Then
         optӦ����(1).Value = False
    End If
    If optӦ����(2).Enabled = False And optӦ����(2).Value = True Then
         optӦ����(2).Value = False
    End If
    If optӦ����(3).Enabled = False And optӦ����(3).Value = True Then
         optӦ����(3).Value = False
    End If
    optӦ����(0).Value = True
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub FindData(ByVal rsTemp As ADODB.Recordset)
    '��ѯ����
    Dim i As Integer
    
    With mshBillEdit
        For i = 1 To .Rows - 1
            If .RowData(i) = rsTemp!Id Then
                .SetFocus
                .Row = i
                .Col = 1
                .MsfObj.TopRow = i
            End If
        Next
    End With
End Sub
