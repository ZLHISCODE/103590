VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemsEdit 
   Caption         =   "�����Ŀ����"
   ClientHeight    =   6375
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   10650
   Icon            =   "frmItemsEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10650
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTitle 
      Height          =   600
      Left            =   75
      TabIndex        =   3
      Top             =   15
      Width           =   6870
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   4020
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   195
         Width           =   2745
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&1.���"
         Height          =   180
         Index           =   4
         Left            =   3435
         TabIndex        =   19
         Top             =   255
         Width           =   540
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ŀ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   195
         Width           =   1800
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6015
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmItemsEdit.frx":076A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13705
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin VB.Frame fra2 
      Height          =   3000
      Left            =   45
      TabIndex        =   1
      Top             =   525
      Width           =   8475
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   2
         Left            =   4620
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   14
         Top             =   180
         Width           =   1020
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   1
         Left            =   3075
         Locked          =   -1  'True
         MaxLength       =   16
         TabIndex        =   13
         Top             =   180
         Width           =   870
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   0
         Left            =   1095
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   930
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   6
         Left            =   7995
         Picture         =   "frmItemsEdit.frx":0FFE
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "��ݼ���F4"
         Top             =   165
         Width           =   345
      End
      Begin VB.CommandButton cmd 
         Height          =   345
         Index           =   5
         Left            =   7620
         Picture         =   "frmItemsEdit.frx":1E40
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "��ݼ���F3"
         Top             =   165
         Width           =   345
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   1650
         Left            =   90
         TabIndex        =   0
         Top             =   555
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   2910
      End
      Begin zl9Medical.VsfGrid vsfPrice 
         Height          =   1635
         Left            =   5070
         TabIndex        =   11
         Top             =   1005
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   2884
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ۿ�(Z)"
         Height          =   180
         Index           =   17
         Left            =   3975
         TabIndex        =   17
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���۸�(E)"
         Height          =   180
         Index           =   18
         Left            =   2070
         TabIndex        =   16
         Top             =   240
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����۸�(&B)"
         Height          =   180
         Index           =   19
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   9825
      Top             =   3645
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemsEdit.frx":23CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemsEdit.frx":7434
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemsEdit.frx":772E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemsEdit.frx":7CC8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraButton 
      Height          =   705
      Left            =   90
      TabIndex        =   5
      Top             =   3450
      Width           =   8460
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   5775
         TabIndex        =   8
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6975
         TabIndex        =   7
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   75
         TabIndex        =   6
         Top             =   210
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmItemsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngLoop As Long
Private mblnDataChange As Boolean
Private mrsItems As New ADODB.Recordset
Private mblnChanged As Boolean
Private mblnGroup As Boolean
Private mlngDept As Long
Private mstrGroup As String
Private mbytMode As Byte
Private mlngKey As Long
Private mlngArchiveKey As Long
Private mblnNo As Boolean
Private mstrSQL As String
Private mstr�Ա� As String
Private Enum mCol
    ��Ŀ = 1
    ִ�п���
    ��鲿λ
    �ɼ���ʽ
    �ɼ�����
    ����걾
    �����۸�
    ���۸�
    �ۿ�
    �������
    ���
    ���㷽ʽ
    ִ�п���id
    �ɼ���ʽid
    �ɼ�����id
    ��鲿λid
    �Ʒ���ϸ
    �¼�
    ǰ��ɫ
    ɾ��
    ����
    �嵥id
    
    p�Ƽ���Ŀ = 1
    p����
    p���㵥λ
    p����
    p��׼����
    p��쵥��
    p�ۿ�
    p��׼���
    p�����
    pִ�п���
    pִ�п���id
    p�շ���Ŀid
    p�Ƽ�����
    p���
    p���ÿ��
End Enum

'�������Զ�����̻���************************************************************************************************
Private Property Let DataChange(ByVal vData As Boolean)
        mblnDataChange = vData
End Property

Private Property Get DataChange() As Boolean
        DataChange = mblnDataChange
End Property

Private Function CreatePriceList(ByVal intRow As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim strKeys As String
    
    strKeys = CStr(Val(vsf.RowData(intRow))) & "'" & CStr(Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid))) & "'" & vsf.TextMatrix(intRow, mCol.��鲿λid)
    
    Dim str�Ƽ���Ŀ As String
    Dim str�Ƽ����� As String
    
    vsfPrice.Rows = 2
    str�Ƽ���Ŀ = vsfPrice.TextMatrix(1, mCol.p�Ƽ���Ŀ)
    str�Ƽ����� = vsfPrice.TextMatrix(1, mCol.p�Ƽ�����)
    
    vsfPrice.Body.Cell(flexcpText, 1, mCol.p�Ƽ���Ŀ + 1, 1, vsfPrice.Cols - 1) = ""
    vsfPrice.RowData(1) = 0

    vsfPrice.TextMatrix(1, mCol.p�Ƽ���Ŀ) = str�Ƽ���Ŀ
    vsfPrice.TextMatrix(1, mCol.p�Ƽ�����) = str�Ƽ�����
    
    If vsfPrice.ComboList(mCol.p�Ƽ���Ŀ) <> "" Then
        vsfPrice.TextMatrix(1, mCol.p�Ƽ���Ŀ) = Split(vsfPrice.ComboList(mCol.p�Ƽ���Ŀ), "|")(0)
    End If
    
    mstrSQL = GetPublicSQL(SQL.�����Ŀ�۱�, strKeys)
    If vsf.TextMatrix(intRow, mCol.��鲿λid) = "" Then
        '����򵥲�λ���
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(vsf.RowData(intRow)), Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid)))
    Else
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    End If
    
    If rs.BOF = False Then
        With vsfPrice
            Do While Not rs.EOF
                
                If Val(.TextMatrix(.Rows - 1, mCol.p�շ���Ŀid)) > 0 Then
                    .Rows = .Rows + 1
                End If
                
                If zlCommFun.NVL(rs("�Ƽ�����")) = 2 Then
                    .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "�ɼ���ʽ-" & vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ)
                ElseIf vsf.TextMatrix(vsf.Row, mCol.���) = "����" Then
                    .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "������Ŀ-" & vsf.TextMatrix(vsf.Row, mCol.��Ŀ)
                Else
                    .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "�����Ŀ-" & vsf.TextMatrix(vsf.Row, mCol.��Ŀ)
                End If
                
                .TextMatrix(.Rows - 1, mCol.p����) = zlCommFun.NVL(rs("����"))
                .TextMatrix(.Rows - 1, mCol.p���㵥λ) = zlCommFun.NVL(rs("���㵥λ"))
                .TextMatrix(.Rows - 1, mCol.p����) = zlCommFun.NVL(rs("�շ�����"))
                .TextMatrix(.Rows - 1, mCol.p��׼����) = zlCommFun.NVL(rs("�ּ�"))
                .TextMatrix(.Rows - 1, mCol.p��쵥��) = zlCommFun.NVL(rs("�ּ�"))
                .TextMatrix(.Rows - 1, mCol.p�ۿ�) = 10
                .TextMatrix(.Rows - 1, mCol.p��׼���) = zlCommFun.NVL(rs("�շ�����"), 0) * zlCommFun.NVL(rs("�ּ�"), 0)
                .TextMatrix(.Rows - 1, mCol.p�����) = zlCommFun.NVL(rs("�շ�����"), 0) * zlCommFun.NVL(rs("�ּ�"), 0)
                .TextMatrix(.Rows - 1, mCol.p�շ���Ŀid) = zlCommFun.NVL(rs("ID"))
                
                .TextMatrix(vsfPrice.Rows - 1, mCol.p�Ƽ�����) = zlCommFun.NVL(rs("�Ƽ�����"))
                
                .TextMatrix(vsfPrice.Rows - 1, mCol.p���) = zlCommFun.NVL(rs("���"))
                
                Call SetRowDefault(zlCommFun.NVL(rs("ID"), 0), vsfPrice.Rows - 1, "�շ�ִ�п���")
                
                If InStr("567", .TextMatrix(.Rows - 1, mCol.p���)) > 0 Then
                    .TextMatrix(.Rows - 1, mCol.p���ÿ��) = GetStorage(Val(.RowData(.Rows - 1)), Val(.TextMatrix(.Rows - 1, mCol.pִ�п���id)))
                    Call PromptStorageWarn(Val(.TextMatrix(.Rows - 1, mCol.p����)), Val(.TextMatrix(.Rows - 1, mCol.p���ÿ��)), .TextMatrix(.Rows - 1, mCol.p����), .TextMatrix(.Rows - 1, mCol.pִ�п���), .TextMatrix(.Rows - 1, mCol.p���㵥λ), 1)
                End If
                                
                rs.MoveNext
            Loop
        End With
    End If
    
    vsf.TextMatrix(intRow, mCol.�����۸�) = SumPrice(1)
    vsf.TextMatrix(intRow, mCol.���۸�) = SumPrice(2)
    
End Function


Public Function ShowEdit(ByVal frmMain As Object, _
                        ByVal lngKey As Long, _
                        ByRef rsItems As ADODB.Recordset, _
                        ByVal lngDept As Long, _
                        Optional blnGroup As Boolean = False, _
                        Optional ByVal bytMode As Byte = 1, _
                        Optional ByVal lngArchiveKey As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    Dim varGroup As Variant
    
    mblnStartUp = True
    mblnOK = False

    'bytMode    1��ʾ����ǰ,2��ʾ���ܺ�
    mblnNo = True
    
    Set mfrmMain = frmMain
        
    mlngKey = lngKey
    Call CopyRecord(rsItems, mrsItems)
    mblnGroup = blnGroup
    mlngDept = lngDept
    mbytMode = bytMode
    mlngArchiveKey = lngArchiveKey
    
    Call ClearData
    If InitData = False Then Exit Function
    If ReadData() = False Then Exit Function
            
    DataChange = False
    
    mblnNo = False
    
    Call cbo_Click
    
    vsf.Col = 2
    vsf.Col = 1
    
    Me.Show 1, frmMain
    
    rsItems.Filter = ""
    If mblnOK Then Call CopyRecord(mrsItems, rsItems)
    
    ShowEdit = mblnOK
    
End Function

Private Function ChangeTotal(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim db�ۿ� As Double
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    Dim dbTotal As Double

    If bytMode = 1 Then
        '�仯���
        
        If dbMoney = 0 Then Exit Function
        '1.�����ۿ�
        db�ۿ� = Format(10 * dbTmp / dbMoney, "0.0000")

    Else
        '�仯�ۿ�
        db�ۿ� = dbTmp

    End If
    
    txtSum(1).Text = Format(dbMoney * db�ۿ� / 10, "0.00")
    txtSum(2).Text = Format(db�ۿ�, "0.0000")
    dbTotal = 0
    
    For lngLoop = 1 To vsf.Rows - 1
    
        vsf.TextMatrix(lngLoop, mCol.�ۿ�) = db�ۿ�
        vsf.TextMatrix(lngLoop, mCol.���۸�) = Format(vsf.TextMatrix(lngLoop, mCol.�����۸�) * (db�ۿ� / 10), "0.00")
        
        dbTotal = dbTotal + Val(vsf.TextMatrix(lngLoop, mCol.���۸�))
                    
        varRow = Split(vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ), ";")
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                varCol(4) = Format(Val(varCol(3)) * (db�ۿ� / 10), "0.00000")
                varCol(10) = db�ۿ�
            End If
            varRow(lngRow) = Join(varCol, ":")
        Next
        vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ) = Join(varRow, ";")
    Next

    '����
    '------------------------------------------------------------------------------------------------------------------
    If dbTotal <> Val(txtSum(1).Text) Then

        For lngLoop = 1 To vsf.Rows - 1
            If Val(vsf.TextMatrix(lngLoop, mCol.���۸�)) <> 0 Then
            
                vsf.TextMatrix(lngLoop, mCol.���۸�) = Val(vsf.TextMatrix(lngLoop, mCol.���۸�)) + (Val(txtSum(1).Text) - dbTotal)
                
                If Val(vsf.TextMatrix(lngLoop, mCol.�����۸�)) <> 0 Then
                    vsf.TextMatrix(lngLoop, mCol.�ۿ�) = Format(10 * Val(vsf.TextMatrix(lngLoop, mCol.���۸�)) / Val(vsf.TextMatrix(lngLoop, mCol.�����۸�)), "0.0000")
                Else
                    vsf.TextMatrix(lngLoop, mCol.�ۿ�) = 0
                End If
                
                varRow = Split(vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ), ";")
                For lngRow = 0 To UBound(varRow)
                    If varRow(lngRow) <> "" Then
                        varCol = Split(varRow(lngRow), ":")
                        If Val(varCol(4)) <> 0 Then
                            varCol(4) = Val(varCol(4)) + (Val(txtSum(1).Text) - dbTotal)
                            If Val(varCol(3)) <> 0 Then
                                varCol(10) = Format(10 * Val(varCol(4)) / Val(varCol(3)), "0.0000")
                            Else
                                varCol(10) = 0
                            End If
                        End If
                    End If
                    varRow(lngRow) = Join(varCol, ":")
                Next
                vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ) = Join(varRow, ";")
                Exit For
            End If
        Next
    End If

    ChangeTotal = True
    
End Function

Private Function ChangeItem(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1, Optional ByVal blnUpdate As Boolean = True) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim dbSum As Double
    Dim db�ۿ� As Double
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    
    
    If blnUpdate Then
        If dbMoney = 0 Then Exit Function
        
        Call WritePrice(vsf.Row)
        
        If bytMode = 1 Then
            '�仯���
            
            '1.�����ۿ�
            db�ۿ� = Format(10 * dbTmp / dbMoney, "0.0000")
        Else
            '�仯�ۿ�
            db�ۿ� = dbTmp
            
        End If
        
        vsf.TextMatrix(vsf.Row, mCol.���۸�) = Format(dbMoney * db�ۿ� / 10, "0.00")
        vsf.TextMatrix(vsf.Row, mCol.�ۿ�) = Format(db�ۿ�, "0.0000")
    End If
    
    '��������
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.�����۸�))
    Next
    txtSum(0).Text = Format(dbSum, "0.00")
    
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.���۸�))
    Next
    txtSum(1).Text = Format(dbSum, "0.00")
    
    If Val(txtSum(0).Text) <> 0 Then
        txtSum(2).Text = Format(10 * Val(txtSum(1).Text) / Val(txtSum(0).Text), "0.0000")
    Else
        txtSum(2).Text = "0.0000"
    End If
    
    '���¼۸�
    '------------------------------------------------------------------------------------------------------------------
    If blnUpdate Then
        varRow = Split(vsf.TextMatrix(vsf.Row, mCol.�Ʒ���ϸ), ";")
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                varCol(4) = Format(Val(varCol(3)) * (db�ۿ� / 10), "0.00000")
                varCol(10) = db�ۿ�
            End If
            varRow(lngRow) = Join(varCol, ":")
        Next
        vsf.TextMatrix(vsf.Row, mCol.�Ʒ���ϸ) = Join(varRow, ";")
    End If
        
    ChangeItem = True
    
End Function

Private Function ChangePrice(ByVal dbMoney As Double, ByVal dbTmp As Double, Optional ByVal bytMode As Byte = 1) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    Dim dbSum As Double
    Dim db�ۿ� As Double
    
    
    If bytMode = 1 Then
        '�仯���
        If dbMoney = 0 Then Exit Function
        '1.�����ۿ�
        db�ۿ� = Format(10 * dbTmp / dbMoney, "0.0000")
    Else
        '�仯�ۿ�
        db�ۿ� = dbTmp
        
    End If
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p��쵥��) = Format(dbMoney * db�ۿ� / 10, "0.00000")
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p�ۿ�) = Format(db�ۿ�, "0.0000")
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p�����) = Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p����)) * Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.p��쵥��))
    
    '������Ŀ
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p��׼���))
    Next
    vsf.TextMatrix(vsf.Row, mCol.�����۸�) = dbSum
    
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p�����))
    Next
    vsf.TextMatrix(vsf.Row, mCol.���۸�) = dbSum
    
    If Val(vsf.TextMatrix(vsf.Row, mCol.�����۸�)) <> 0 Then
        vsf.TextMatrix(vsf.Row, mCol.�ۿ�) = Format(10 * Val(vsf.TextMatrix(vsf.Row, mCol.���۸�)) / Val(vsf.TextMatrix(vsf.Row, mCol.�����۸�)), "0.0000")
    Else
        vsf.TextMatrix(vsf.Row, mCol.�ۿ�) = "0.0000"
    End If
    
    '��������
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.�����۸�))
    Next
    txtSum(0).Text = Format(dbSum, "0.00")
    
    dbSum = 0
    For lngLoop = 1 To vsf.Rows - 1
       dbSum = dbSum + Val(vsf.TextMatrix(lngLoop, mCol.���۸�))
    Next
    txtSum(1).Text = Format(dbSum, "0.00")
    
    If Val(txtSum(0).Text) <> 0 Then
        txtSum(2).Text = Format(10 * Val(txtSum(1).Text) / Val(txtSum(0).Text), "0.0000")
    Else
        txtSum(2).Text = "0.0000"
    End If
        
    ChangePrice = True
    
End Function


Private Function SumPrice(ByVal bytMode As Byte) As Single
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim sglSum As Single
    
    For lngLoop = 1 To vsfPrice.Rows - 1
        If bytMode = 2 Then
            sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p�����))
        Else
            sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.p��׼����)) * Val(vsfPrice.TextMatrix(lngLoop, mCol.p����))
        End If
    Next
    SumPrice = sglSum
    
End Function

Private Function ClearData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------

    cbo.Clear
    Call ResetVsf(vsf)
    Call ResetVsf(vsfPrice)
    DataChange = False
    
        
End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------
    
    On Error GoTo errHand

    With vsf
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "����", 2100, 1, "...", 1
        .NewColumn "ִ�п���", 1080, 1, " ", 1
        
        .NewColumn "��鲿λ", 1800, 1, "...", 1
        .NewColumn "�ɼ���ʽ", 1200, 1, " ", 1
        .NewColumn "�ɼ�����", 1080, 1, " ", 1
        
        .NewColumn "����걾", 900, 1, " ", 1
        .NewColumn "�����۸�", 900, 7
        .NewColumn "���۸�", 900, 7
        .NewColumn "�ۿ�", 900, 7, , 1
        .NewColumn "�������", 0, 1
        .NewColumn "���", 0, 1
        .NewColumn "���㷽ʽ", 900, 1, "����|�շ�", 1
        .NewColumn "ִ�п���id", 0, 1
        .NewColumn "�ɼ���ʽid", 0, 1
        .NewColumn "�ɼ�����id", 0, 1
        .NewColumn "��鲿λid", 0, 1
        .NewColumn "�Ʒ���ϸ", 0, 1
        .NewColumn "�¼�", 0, 1
        .NewColumn "ǰ��ɫ", 0, 1
        .NewColumn "ɾ��", 0, 1
        .NewColumn "����", 0, 1
        .NewColumn "�嵥id", 0, 1
        .FixedCols = 1
        
        .SelectMode = True
        
        .Body.ColFormat(mCol.�����۸�) = "0.00"
        .Body.ColFormat(mCol.���۸�) = "0.00"
        .Body.ColFormat(mCol.�ۿ�) = "0.0000"
    End With
    
    With vsfPrice
        .Cols = 0
        .NewColumn "", 255, 4
        .NewColumn "�Ƽ���Ŀ", 2100, 1, " ", 1
        .NewColumn "�շ���Ŀ", 2700, 1, "...", 1
        .NewColumn "��λ", 600, 1
        .NewColumn "����", 540, 7, , 1
        .NewColumn "��׼����", 900, 7
        .NewColumn "��쵥��", 900, 7, , 1
        .NewColumn "�ۿ�", 900, 7, , 1
        .NewColumn "��׼�۸�", 900, 7
        .NewColumn "���۸�", 900, 7
        .NewColumn "ִ�п���", 1080, 1, " ", 1
        .NewColumn "ִ�п���id", 0
        .NewColumn "�շ���Ŀid", 0
        .NewColumn "�Ƽ�����", 0
        .NewColumn "���", 0
        .NewColumn "", 0
        .FixedCols = 1
        .Body.ColFormat(mCol.p��׼����) = "0.00000"
        .Body.ColFormat(mCol.p��쵥��) = "0.00000"
        .Body.ColFormat(mCol.p��׼���) = "0.00"
        .Body.ColFormat(mCol.p�����) = "0.00"
        .Body.ColFormat(mCol.p�ۿ�) = "0.0000"
        .SelectMode = True
    End With
    
    mstrGroup = ""
    cbo.AddItem "ȱʡ"
               
    If mblnGroup = False Then
        cbo.Visible = False
        lbl(4).Visible = False
'        cmd(5).Visible = False
'        cmd(6).Visible = False
        lblTitle.Caption = "��Ա��Ŀ"
    Else
'        cmd(5).Visible = True
'        cmd(6).Visible = True
        lblTitle.Caption = "�����Ŀ"
    End If
        
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadPrice(ByVal intRow As Integer) As Boolean
    '��ȡ��Ӧ�ļƷ���ϸ
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    Dim lngCol As Long
    
    Call ResetVsf(vsfPrice)
    
    If intRow = 0 Then Exit Function
    
    If vsf.TextMatrix(intRow, mCol.�Ʒ���ϸ) <> "" Then
        
        varRow = Split(vsf.TextMatrix(intRow, mCol.�Ʒ���ϸ), ";")
        
        vsfPrice.Rows = UBound(varRow) + 2
        
        For lngRow = 0 To UBound(varRow)
            If varRow(lngRow) <> "" Then
                varCol = Split(varRow(lngRow), ":")
                For lngCol = 0 To UBound(varCol)
                    
                    If Val(varCol(6)) = 2 Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p�Ƽ���Ŀ) = "�ɼ���ʽ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ))
                    ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "����" Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p�Ƽ���Ŀ) = "������Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                    Else
                        vsfPrice.TextMatrix(lngRow + 1, mCol.p�Ƽ���Ŀ) = "�����Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                    End If
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p����) = varCol(0)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p���㵥λ) = varCol(1)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p����) = varCol(2)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p��׼����) = varCol(3)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p��쵥��) = varCol(4)
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p��׼���) = Val(varCol(2)) * Val(varCol(3))
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p�����) = Val(varCol(2)) * Val(varCol(4))
                                        
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p�շ���Ŀid) = varCol(5)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p�Ƽ�����) = varCol(6)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.pִ�п���) = varCol(7)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.pִ�п���id) = varCol(8)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p���) = varCol(9)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p�ۿ�) = varCol(10)
                    
                    vsfPrice.RowData(lngRow + 1) = Val(varCol(5))
                    
                Next
            End If
        Next
        
    End If
    
    ReadPrice = True
End Function

Private Function WritePrice(ByVal intRow As Integer) As Boolean
    Dim strTmp As String
    Dim lngRow As Long
    Dim varCol As Variant
    
    On Error GoTo errHand
    
    If intRow <= 0 Then Exit Function
    
    For lngRow = 1 To vsfPrice.Rows - 1
        If Val(vsfPrice.TextMatrix(lngRow, mCol.p�շ���Ŀid)) > 0 Then
            
            varCol = Split(String(11, ":"), ":")
            
            varCol(0) = vsfPrice.TextMatrix(lngRow, mCol.p����)
            varCol(1) = vsfPrice.TextMatrix(lngRow, mCol.p���㵥λ)
            varCol(2) = vsfPrice.TextMatrix(lngRow, mCol.p����)
            varCol(3) = vsfPrice.TextMatrix(lngRow, mCol.p��׼����)
            varCol(4) = vsfPrice.TextMatrix(lngRow, mCol.p��쵥��)
            varCol(5) = vsfPrice.TextMatrix(lngRow, mCol.p�շ���Ŀid)
            varCol(6) = vsfPrice.TextMatrix(lngRow, mCol.p�Ƽ�����)
            varCol(7) = vsfPrice.TextMatrix(lngRow, mCol.pִ�п���)
            varCol(8) = vsfPrice.TextMatrix(lngRow, mCol.pִ�п���id)
            varCol(9) = vsfPrice.TextMatrix(lngRow, mCol.p���)
            varCol(10) = vsfPrice.TextMatrix(lngRow, mCol.p�ۿ�)
            
            If strTmp = "" Then
                strTmp = Join(varCol, ":")
            Else
                strTmp = strTmp & ";" & Join(varCol, ":")
            End If
        End If
    Next
    
    vsf.TextMatrix(intRow, mCol.�Ʒ���ϸ) = strTmp
    
    WritePrice = True
    
errHand:
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
            
    On Error GoTo errHand

    
    '��ȡ�����������Ŀ
    
    mblnNo = True
    
    cbo.Clear
    
    If mblnGroup = False Then
        
        gstrSQL = "Select �������,�Ա� From �����Ա���� Where ����id=[2] and �Ǽ�id=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey, mlngArchiveKey)
        If rs.BOF = False Then
            cbo.AddItem zlCommFun.NVL(rs("�������"))
            mstr�Ա� = zlCommFun.NVL(rs("�Ա�"), "δ֪")
        End If
        
        If cbo.ListCount = 0 Then cbo.AddItem "ȱʡ"
    Else
        gstrSQL = "SELECT A.�������, rownum AS ID FROM ������ A WHERE A.�Ǽ�id=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngKey)
        If rs.BOF = False Then
            Do While Not rs.EOF
                cbo.AddItem zlCommFun.NVL(rs("�������"))
                rs.MoveNext
            Loop
        Else
            cbo.AddItem "ȱʡ"
        End If
    End If
    
    If cbo.ListIndex = -1 And cbo.ListCount > 0 Then cbo.ListIndex = 0
    '��ȡ�����Ŀ
    mblnNo = False
    
    Call cbo_Click
    
    

    ReadData = True
    
    Exit Function
    
errHand:

    If ErrCenter = 1 Then
        Resume
    End If
    
End Function

Private Function CheckHave(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ����Ƿ����ظ�����Ŀ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) = lngKey And vsf.Row <> lngLoop Then
            CheckHave = True
            Exit Function
        End If
    Next
End Function

Private Function ShowOpenList(Optional strText As String, Optional ByVal lngCol As Long = 0) As Byte
    '------------------------------------------------------------------------------------------------------------------
    '����:  ���б�ʽ��ʾ����
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim strLvw As String
    Dim sglX As Single
    Dim sglY As Single
    Dim rs As New ADODB.Recordset
    Dim strClass As String
    Dim strPath As String
    Dim lngLoop As Long
    Dim strTmp As String
    On Error GoTo errHand
    
    ShowOpenList = 2

    Select Case lngCol
        Case mCol.��Ŀ
        
            strText = UCase(strText)
            
            strLvw = "����,1200,0,1;����,2700,0,0;��λ,900,0,0;�걾��λ,900,0,0;���,900,0,0"
            strPath = Me.Name & "\�����Ŀѡ��"
            
            gstrSQL = GetPublicSQL(SQL.�����Ŀ����ѡ��, strText)
            
            If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                strTmp = strText & "%"
            Else
                strTmp = "%" & strText & "%"
            End If
            
            Dim bytParam1 As Byte
            Dim bytParam2 As Byte
            
            Select Case mstr�Ա�
            Case "��"
                bytParam1 = 1
            Case "Ů"
                bytParam2 = 2
            End Select
            
            If Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp, bytParam1, bytParam2)
            ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "����" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "", strText & "%", strTmp, bytParam1, bytParam2)
            Else
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "D", "", strText & "%", strTmp, bytParam1, bytParam2)
            End If
            
'            If Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "" Then
'                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp)
'            ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "����" Then
'                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "", strText & "%", strTmp)
'            Else
'                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "D", "", strText & "%", strTmp)
'            End If

        Case mCol.ִ�п���
            
            strLvw = "����,1200,0,1;����,3300,0,0;����,1200,0,0"
            strPath = Me.Name & "\ִ�п���ѡ��"
            
            gstrSQL = GetPublicSQL(SQL.����ִ�п���)
            If gstrSQL = "" Then Exit Function
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), mlngDept, UserInfo.����ID, "%" & UCase(strText) & "%")
            
        Case mCol.��鲿λ
            
            strText = "'%" & UCase(strText) & "%'"
            
            strLvw = "����,3300,0,0"
            strPath = Me.Name & "\��鲿λѡ��"
            
            gstrSQL = "select B.�걾��λ AS ����,B.ID,0 AS ѡ�� from ������Ŀ��� A,������ĿĿ¼ B WHERE (B.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or B.����ʱ�� is NULL) AND A.������ĿID=B.ID AND A.�������ID=" & Val(vsf.RowData(vsf.Row)) & ""
            
            rs.CursorLocation = adUseClient
            If rs.State = adStateOpen Then rs.Close
            rs.Open gstrSQL, gcnOracle, adOpenStatic, adLockOptimistic
            
        Case mCol.�ɼ���ʽ
            
            strText = "%" & UCase(strText) & "%"
            
            strLvw = "����,1200,0,1;����,3300,0,0"
            strPath = Me.Name & "\�ɼ���ʽѡ��"
            
            gstrSQL = "SELECT A.ID,A.����,A.���� " & _
                "FROM ������ĿĿ¼ A,�����÷����� B " & _
                "WHERE (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) AND A.���='E' AND A.��������='6' AND A.ID=B.�÷�id AND B.��Ŀid=[1] "
            gstrSQL = gstrSQL & " AND (UPPER(A.����) Like [2] OR A.���� Like [2] OR A.ID IN (SELECT ������Ŀid FROM ������Ŀ���� WHERE (���� Like [2] OR UPPER(����) Like [2])))"
            
'            Call OpenRecord(rs, gstrSQL, Me.Caption)
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), strText)
            
            If rs.BOF Then
                gstrSQL = "SELECT A.ID,A.����,A.���� " & _
                    "FROM ������ĿĿ¼ A " & _
                    "WHERE (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) AND A.���='E' AND A.��������='6' "
                gstrSQL = gstrSQL & " AND (UPPER(A.����) Like [1] OR A.���� Like [1] OR A.ID IN (SELECT ������Ŀid FROM ������Ŀ���� WHERE (���� Like [1] OR UPPER(����) Like [1])))"
                    
            End If
'            Call OpenRecord(rs, gstrSQL, Me.Caption)
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText)
            
    End Select

    If rs.BOF Then
        ShowOpenList = 0
        Exit Function
    End If
    If rs.RecordCount = 1 And strText <> "'%%'" Then GoTo PointOver
    Call CalcPosition(sglX, sglY, vsf)
    
    If lngCol = mCol.��鲿λ Then
        If vsf.TextMatrix(vsf.Row, mCol.��鲿λid) <> "" Then
            Do While Not rs.EOF
                If InStr("," & vsf.TextMatrix(vsf.Row, mCol.��鲿λid) & ",", "," & rs("ID").Value & ",") > 0 Then rs("ѡ��").Value = 1
                rs.MoveNext
            Loop
        End If
        rs.MoveFirst
        
        If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "�������ѡ������Ŀ,Ȼ��س���˫���˳�", sglX + 60, sglY + 30, 9000, 4500, 300, , strPath, , False, True) Then GoTo PointOver
        
    Else
        If frmSelectDialog.ShowSelect(Me, 2, rs, strLvw, "�������ѡ��һ����Ŀ", sglX + 60, sglY + 30, 9000, 4500, 300, , strPath, , False) Then GoTo PointOver
    End If
        
    Exit Function
    
PointOver:
    Select Case lngCol
        Case mCol.��Ŀ
            If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                MsgBox "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���Ѿ���ѡ��", vbInformation, gstrSysName
                Exit Function
            End If
            
            vsf.Cell(flexcpText, vsf.Row, mCol.��Ŀ + 1, vsf.Row, vsf.Cols - 1) = ""
            
            vsf.EditText = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
            vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
            vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
            
        Case mCol.ִ�п���
        
            vsf.EditText = zlCommFun.NVL(rs("����").Value)
            vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Row, mCol.ִ�п���id) = zlCommFun.NVL(rs("ID").Value)
        
        Case mCol.��鲿λ
            
            vsf.TextMatrix(vsf.Row, vsf.Col) = ""
            vsf.TextMatrix(vsf.Row, mCol.��鲿λid) = ""
            
            rs.Filter = ""
            rs.Filter = "ѡ��=1"
            If rs.RecordCount > 0 Then
                rs.MoveFirst
                Do While Not rs.EOF
                    vsf.TextMatrix(vsf.Row, vsf.Col) = vsf.TextMatrix(vsf.Row, vsf.Col) & zlCommFun.NVL(rs("����").Value) & ","
                    vsf.TextMatrix(vsf.Row, mCol.��鲿λid) = vsf.TextMatrix(vsf.Row, mCol.��鲿λid) & zlCommFun.NVL(rs("ID").Value) & ","
                    rs.MoveNext
                Loop
                
                If vsf.TextMatrix(vsf.Row, mCol.��鲿λ) <> "" Then vsf.TextMatrix(vsf.Row, mCol.��鲿λ) = Mid(vsf.TextMatrix(vsf.Row, mCol.��鲿λ), 1, Len(vsf.TextMatrix(vsf.Row, mCol.��鲿λ)) - 1)
                If vsf.TextMatrix(vsf.Row, mCol.��鲿λid) <> "" Then vsf.TextMatrix(vsf.Row, mCol.��鲿λid) = Mid(vsf.TextMatrix(vsf.Row, mCol.��鲿λid), 1, Len(vsf.TextMatrix(vsf.Row, mCol.��鲿λid)) - 1)
                
            End If
        Case mCol.�ɼ���ʽ
        
            vsf.EditText = zlCommFun.NVL(rs("����").Value)
            vsf.Cell(flexcpData, vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
            vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid) = zlCommFun.NVL(rs("ID").Value)
    End Select
    
    ShowOpenList = 1
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ReadSample(ByVal lng������Ŀid As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ѡ�Ĳɼ���ʽ��������
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    gstrSQL = "SELECT 1 FROM ������ĿĿ¼ WHERE �����Ŀ=1 AND ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng������Ŀid)
    If rs.BOF = False Then
        '�������Ŀ
        
        gstrSQL = "SELECT DISTINCT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                "WHERE C.ID<>[1] AND nvl(C.�����Ŀ,0)=0 " & _
                    "AND B.������Ŀid=A.��Ŀid "
                    
        gstrSQL = gstrSQL & "AND B.������Ŀid IN (SELECT C.ID " & _
                     "FROM ���鱨����Ŀ A," & _
                          "(SELECT ������Ŀid FROM ���鱨����Ŀ WHERE ������Ŀid = [1]) B," & _
                          "������ĿĿ¼ C,����������Ŀ D,������Ŀ E,���鱨����Ŀ F " & _
                    "WHERE A.������Ŀid = B.������Ŀid AND A.������Ŀid <> [1] AND " & _
                          "nvl(C.�����Ŀ,0) = 0 AND A.������Ŀid = C.ID AND C.ID=F.������Ŀid AND F.������Ŀid=D.ID AND D.ID=E.������Ŀid)"
                                  
    Else
        gstrSQL = "SELECT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                "WHERE C.ID=[1] AND nvl(C.�����Ŀ,0)=0 AND B.������Ŀid=[1] and B.������Ŀid=A.��Ŀid"
    End If
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng������Ŀid)

    If rs.BOF = False Then
        Do While Not rs.EOF
            ReadSample = ReadSample & rs("����").Value & "|"
            rs.MoveNext
        Loop
    Else
        
        'û�ж�Ӧʱ����ȡ���б걾����
        gstrSQL = "SELECT ���� FROM ���Ƽ���걾 A "
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then
            Do While Not rs.EOF
                ReadSample = ReadSample & rs("����").Value & "|"
                rs.MoveNext
            Loop
        End If
        
    End If
    
    Exit Function
        
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function SetRowData(ByVal lngKey As Long, ByVal intRow As Integer, ParamArray arryMode() As Variant) As Boolean
'------------------------------------------------------------------------------------------------------------------
    '����:���������ݣ����в�ͬ����ͬ��
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCombList As String
    
    On Error Resume Next
    
    For lngLoop = 0 To UBound(arryMode)
        Select Case arryMode(lngLoop)
        Case "�շ�ִ�п���"
        
            If InStr("4,5,6,7", vsfPrice.TextMatrix(intRow, mCol.p���)) > 0 Then
                gstrSQL = GetPublicSQL(SQL.ҩƷִ�п���)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfPrice.TextMatrix(intRow, mCol.p���))
            Else
                gstrSQL = GetPublicSQL(SQL.�շ�ִ�п���, "1")
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.����ID, "%%")
            End If
            If rs.RecordCount > 1 Then
                vsfPrice.EditMode(mCol.pִ�п���) = 1
                vsfPrice.Body.ColComboList(mCol.pִ�п���) = vsfPrice.Body.BuildComboList(rs, "����", "ID")
            Else
                vsfPrice.EditMode(mCol.pִ�п���) = 0
                vsfPrice.Body.ColComboList(mCol.pִ�п���) = ""
            End If
        
        Case "�Ƽ���Ŀ"
            
            If Trim(vsf.TextMatrix(intRow, mCol.���)) = "���" Then
                strCombList = "�����Ŀ-" & Trim(vsf.TextMatrix(intRow, mCol.��Ŀ))
                vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 0
                vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = ""
                
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p�Ƽ���Ŀ) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p�Ƽ�����) = "1"
            Else
                strCombList = "������Ŀ-" & Trim(vsf.TextMatrix(intRow, mCol.��Ŀ))
                If Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid)) > 0 Then
                    strCombList = strCombList & "|�ɼ���ʽ-" & Trim(vsf.TextMatrix(intRow, mCol.�ɼ���ʽ))
                    vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 1
                    vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = strCombList
                Else
                    vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 0
                    vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = ""
                End If
            End If
            
        Case "����ִ�п���"
        
            gstrSQL = GetPublicSQL(SQL.����ִ�п���, "1")
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.����ID, "%%")
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.ִ�п���) = 1
                    vsf.Body.ColComboList(mCol.ִ�п���) = vsf.Body.BuildComboList(rs, "����", "ID")
                Else
                    vsf.EditMode(mCol.ִ�п���) = 0
                    vsf.Body.ColComboList(mCol.ִ�п���) = ""
                End If
            End If
        
        Case "�ɼ���ʽ"
        
            gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A,�����÷����� B WHERE A.ID=B.�÷�id AND A.���='E' AND A.��������='6' AND B.��ĿID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.RecordCount > 1 Then
                vsf.EditMode(mCol.�ɼ���ʽ) = 1
                vsf.Body.ColComboList(mCol.�ɼ���ʽ) = vsf.Body.BuildComboList(rs, "����", "ID")
            Else
                gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A WHERE A.���='E' AND A.��������='6'"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.�ɼ���ʽ) = 1
                    vsf.Body.ColComboList(mCol.�ɼ���ʽ) = vsf.Body.BuildComboList(rs, "����", "ID")
                Else
                    vsf.EditMode(mCol.�ɼ���ʽ) = 0
                    vsf.Body.ColComboList(mCol.�ɼ���ʽ) = ""
                End If
            End If
            
        Case "�ɼ�����"
        
            gstrSQL = GetPublicSQL(SQL.����ִ�п���)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid)), mlngDept, UserInfo.����ID, "%%")
                If rs.RecordCount > 1 Then
                    vsf.EditMode(mCol.�ɼ�����) = 1
                    vsf.Body.ColComboList(mCol.�ɼ�����) = vsf.Body.BuildComboList(rs, "*����", "ID")
                Else
                    vsf.EditMode(mCol.�ɼ�����) = 0
                    vsf.Body.ColComboList(mCol.�ɼ�����) = ""
                End If
            End If
        
        Case "����걾"
        
            gstrSQL = "SELECT 1 FROM ������ĿĿ¼ WHERE �����Ŀ=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                '�������Ŀ
                
                gstrSQL = "SELECT DISTINCT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                        "WHERE C.ID<>[1] AND nvl(C.�����Ŀ,0)=0 " & _
                            "AND B.������Ŀid=A.��Ŀid and rownum<2"
                            
                gstrSQL = gstrSQL & "AND B.������Ŀid IN (SELECT C.ID " & _
                             "FROM ���鱨����Ŀ A," & _
                                  "(SELECT ������Ŀid FROM ���鱨����Ŀ WHERE ������Ŀid = [1]) B," & _
                                  "������ĿĿ¼ C,����������Ŀ D,������Ŀ E,���鱨����Ŀ F " & _
                            "WHERE A.������Ŀid = B.������Ŀid AND A.������Ŀid <> [1] AND " & _
                                  "nvl(C.�����Ŀ,0) = 0 AND A.������Ŀid = C.ID AND C.ID=F.������Ŀid AND F.������Ŀid=D.ID AND D.ID=E.������Ŀid)  and rownum<2 "
                                          
            Else
                gstrSQL = "SELECT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                        "WHERE C.ID=[1] AND nvl(C.�����Ŀ,0)=0 AND B.������Ŀid=[1] and B.������Ŀid=A.��Ŀid  and rownum<2"
            End If
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.RecordCount > 1 Then
                
                vsf.EditMode(mCol.����걾) = 1
                vsf.Body.ColComboList(mCol.����걾) = vsf.Body.BuildComboList(rs, "����", "����")
                
            Else
                
                'û�ж�Ӧʱ����ȡ���б걾����
                gstrSQL = "SELECT ���� FROM ���Ƽ���걾 A"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.RecordCount > 1 Then
                
                    vsf.EditMode(mCol.����걾) = 1
                    vsf.Body.ColComboList(mCol.����걾) = vsf.Body.BuildComboList(rs, "����", "����")
                Else
                    vsf.EditMode(mCol.����걾) = 0
                    vsf.Body.ColComboList(mCol.����걾) = ""
                End If
                
            End If
        
        End Select
    Next
    
    SetRowData = True
    
End Function

Private Function SetRowDefault(ByVal lngKey As Long, ByVal intRow As Integer, ParamArray arryMode() As Variant) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡȱʡ
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim lngLoop As Long
    Dim strCombList As String
    
    On Error GoTo errHand
    
    For lngLoop = 0 To UBound(arryMode)
        
        Select Case arryMode(lngLoop)
        Case "���㷽ʽ"
            
            If mblnGroup Then
                vsf.TextMatrix(vsf.Row, mCol.���㷽ʽ) = "����"
            Else
                vsf.TextMatrix(vsf.Row, mCol.���㷽ʽ) = "�շ�"
            End If
            
        Case "ִ�п���"
'            lng��������id = mlngDept
            
            gstrSQL = GetPublicSQL(SQL.����ִ�п���)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.����ID, "%%")
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.ִ�п���) = zlCommFun.NVL(rs("����").Value)
                    vsf.TextMatrix(vsf.Row, mCol.ִ�п���id) = zlCommFun.NVL(rs("ID").Value)
                Else
                    vsf.TextMatrix(vsf.Row, mCol.ִ�п���) = gstrDeptName
                    vsf.TextMatrix(vsf.Row, mCol.ִ�п���id) = UserInfo.����ID
                End If
            End If
        
        Case "�ɼ���ʽ"
            
            gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A,�����÷����� B WHERE A.ID=B.�÷�id AND A.���='E' AND A.��������='6' AND B.��ĿID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid) = zlCommFun.NVL(rs("ID").Value)
            Else
                gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A WHERE A.���='E' AND A.��������='6'"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ) = zlCommFun.NVL(rs("����").Value)
                    vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid) = zlCommFun.NVL(rs("ID").Value)
                End If
            End If
            
        Case "�ɼ�����"
                    
            gstrSQL = GetPublicSQL(SQL.����ִ�п���)
            If gstrSQL <> "" Then
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid)), mlngDept, UserInfo.����ID, "%%")
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.�ɼ�����) = zlCommFun.NVL(rs("����").Value)
                    vsf.TextMatrix(vsf.Row, mCol.�ɼ�����id) = zlCommFun.NVL(rs("ID").Value)
                End If
            End If
        
        Case "����걾"
        
            gstrSQL = "SELECT 1 FROM ������ĿĿ¼ WHERE �����Ŀ=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                '�������Ŀ
                
                gstrSQL = "SELECT DISTINCT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                        "WHERE C.ID<>[1] AND nvl(C.�����Ŀ,0)=0 " & _
                            "AND B.������Ŀid=A.��Ŀid and rownum<2"
                            
                gstrSQL = gstrSQL & "AND B.������Ŀid IN (SELECT C.ID " & _
                             "FROM ���鱨����Ŀ A," & _
                                  "(SELECT ������Ŀid FROM ���鱨����Ŀ WHERE ������Ŀid = [1]) B," & _
                                  "������ĿĿ¼ C,����������Ŀ D,������Ŀ E,���鱨����Ŀ F " & _
                            "WHERE A.������Ŀid = B.������Ŀid AND A.������Ŀid <> [1] AND " & _
                                  "nvl(C.�����Ŀ,0) = 0 AND A.������Ŀid = C.ID AND C.ID=F.������Ŀid AND F.������Ŀid=D.ID AND D.ID=E.������Ŀid)  and rownum<2 "
                                          
            Else
                gstrSQL = "SELECT A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
                        "WHERE C.ID=[1] AND nvl(C.�����Ŀ,0)=0 AND B.������Ŀid=[1] and B.������Ŀid=A.��Ŀid  and rownum<2"
            End If
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                vsf.TextMatrix(vsf.Row, mCol.����걾) = rs("����").Value
            Else
                
                'û�ж�Ӧʱ����ȡ���б걾����
                gstrSQL = "SELECT ���� FROM ���Ƽ���걾 A where rownum<2"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
                If rs.BOF = False Then
                    vsf.TextMatrix(vsf.Row, mCol.����걾) = rs("����").Value
                End If
                
            End If
        
        Case "�շ�ִ�п���"
            If InStr("4,5,6,7", vsfPrice.TextMatrix(intRow, mCol.p���)) > 0 Then
                gstrSQL = GetPublicSQL(SQL.ҩƷִ�п���)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsfPrice.TextMatrix(intRow, mCol.p���))
            Else
                gstrSQL = GetPublicSQL(SQL.�շ�ִ�п���)
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, mlngDept, UserInfo.����ID, "%%")
            End If
            If rs.BOF = False Then
                vsfPrice.TextMatrix(intRow, mCol.pִ�п���) = zlCommFun.NVL(rs("����").Value)
                vsfPrice.TextMatrix(intRow, mCol.pִ�п���id) = zlCommFun.NVL(rs("ID").Value)
            Else
                vsfPrice.TextMatrix(intRow, mCol.pִ�п���) = vsf.TextMatrix(vsf.Row, mCol.ִ�п���)
                vsfPrice.TextMatrix(intRow, mCol.pִ�п���id) = vsf.TextMatrix(vsf.Row, mCol.ִ�п���id)
            End If
        Case "�Ƽ���Ŀ"
        
            If Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "���" Then
                strCombList = "�����Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 0
                vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p�Ƽ���Ŀ) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p�Ƽ�����) = "1"
            Else
                strCombList = "������Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                If Val(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid)) > 0 Then
                    strCombList = strCombList & "|�ɼ���ʽ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ))
                    vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 1
                    vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = strCombList
                Else
                    vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 0
                    vsfPrice.Body.ColComboList(mCol.p�Ƽ���Ŀ) = ""
                End If
            End If

        End Select
    Next
    
    SetRowDefault = True
    
    Exit Function
    
errHand:
    
End Function

Private Function SaveItems(ByVal strGroup As String) As Boolean
    
    Dim lngLoop As Long
    
    On Error GoTo errHand

    '������ѡ��ļ�����Ŀ
    mrsItems.Filter = ""
    mrsItems.Filter = "���='" & strGroup & "' AND ɾ��<>'1'"
    
    Call DeleteRecord(mrsItems)
    
    For lngLoop = 1 To vsf.Rows - 1
        
        If Val(vsf.RowData(lngLoop)) > 0 Then
            mrsItems.AddNew
            
            mrsItems("���").Value = strGroup
            mrsItems("ID").Value = vsf.RowData(lngLoop)
            mrsItems("���").Value = vsf.TextMatrix(lngLoop, mCol.���)
            mrsItems("����").Value = vsf.TextMatrix(lngLoop, mCol.��Ŀ)
            mrsItems("ִ�п���").Value = vsf.TextMatrix(lngLoop, mCol.ִ�п���)
            mrsItems("��鲿λ").Value = vsf.TextMatrix(lngLoop, mCol.��鲿λ)
            mrsItems("�ɼ���ʽ").Value = vsf.TextMatrix(lngLoop, mCol.�ɼ���ʽ)
            mrsItems("�ɼ�����").Value = vsf.TextMatrix(lngLoop, mCol.�ɼ�����)
            mrsItems("����걾").Value = vsf.TextMatrix(lngLoop, mCol.����걾)
            mrsItems("�������").Value = vsf.TextMatrix(lngLoop, mCol.�������)
            mrsItems("�����۸�").Value = vsf.TextMatrix(lngLoop, mCol.�����۸�)
            mrsItems("���۸�").Value = vsf.TextMatrix(lngLoop, mCol.���۸�)
            mrsItems("�ۿ�").Value = vsf.TextMatrix(lngLoop, mCol.�ۿ�)
            mrsItems("���㷽ʽ").Value = vsf.TextMatrix(lngLoop, mCol.���㷽ʽ)
            mrsItems("ִ�п���id").Value = vsf.TextMatrix(lngLoop, mCol.ִ�п���id)
            mrsItems("�ɼ���ʽid").Value = vsf.TextMatrix(lngLoop, mCol.�ɼ���ʽid)
            mrsItems("�ɼ�����id").Value = vsf.TextMatrix(lngLoop, mCol.�ɼ�����id)
            mrsItems("��鲿λid").Value = vsf.TextMatrix(lngLoop, mCol.��鲿λid)
            mrsItems("�Ʒ���ϸ").Value = vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ)
            mrsItems("�¼�").Value = vsf.TextMatrix(lngLoop, mCol.�¼�)
            mrsItems("ǰ��ɫ").Value = vsf.TextMatrix(lngLoop, mCol.ǰ��ɫ)
            mrsItems("ɾ��").Value = ""
            mrsItems("����").Value = vsf.TextMatrix(lngLoop, mCol.����)
            mrsItems("�嵥id").Value = vsf.TextMatrix(lngLoop, mCol.�嵥id)
            
        End If
    Next
    
    SaveItems = True
    
errHand:

End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
                
   
    If Val(vsf.TextMatrix(lngLoop, mCol.���۸�)) < 0 Then
            
            ShowSimpleMsg "���۸���Ϊ����"
            vsf.Row = lngLoop
            vsf.Col = mCol.���۸�
            vsf.ShowCell vsf.Row, vsf.Col
            vsf.SetFocus
            
            Exit Function
        End If
        
'        If Val(vsf.TextMatrix(lngLoop, mCol.���۸�)) > Val(vsf.TextMatrix(lngLoop, mCol.�����۸�)) Then
'
'            ShowSimpleMsg "���۸��ܴ��ڻ����۸�"
'            vsf.Row = lngLoop
'            vsf.Col = mCol.���۸�
'            vsf.ShowCell vsf.Row, vsf.Col
'            vsf.SetFocus
'
'            Exit Function
'        End If
    
    ValidEdit = True
    
End Function

Private Function ReadItems(ByVal strGroup As String) As Boolean
    
    mrsItems.Filter = ""
    mrsItems.Filter = "���='" & strGroup & "' AND ɾ��<>'1'"
    If mrsItems.RecordCount > 0 Then
        mrsItems.MoveFirst
        Call FillGrid(vsf, mrsItems)
    End If
    
    ReadItems = True
    
End Function

Private Function ReadTemplate(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rs As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset
    
    Dim strKeys As String
    Dim bytParam1 As Byte
    Dim bytParam2 As Byte
    
    bytParam1 = 1
    bytParam2 = 2
            
    Select Case mstr�Ա�
    Case "��"
        bytParam1 = 1
    Case "Ů"
        bytParam2 = 2
    End Select
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT DISTINCT A.ID,DECODE(A.���,'C','����','D','���') AS ���,A.����,A.����,C.���� AS �������,D.���� As �ɼ���ʽ,B.�ɼ���ʽid,B.����걾,B.��鲿λ,B.��鲿λid " & _
                "FROM ������ĿĿ¼ A,�������Ŀ¼ B,������� C,������ĿĿ¼ D " & _
                "WHERE A.ID=B.������ĿID AND C.���=B.��� AND D.ID(+)=B.�ɼ���ʽid AND B.���=[1] And Nvl(a.�����Ա�,0) In (0,[2],[3])"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey, bytParam1, bytParam2)
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            vsf.Row = vsf.Rows - 1
            If Val(vsf.RowData(vsf.Row)) > 0 Then
                vsf.Rows = vsf.Rows + 1
                vsf.Row = vsf.Rows - 1
            End If
            
            If CheckHave(rs("ID").Value) = False Then
            
                vsf.TextMatrix(vsf.Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
                vsf.TextMatrix(vsf.Row, mCol.��Ŀ) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(vsf.Row, mCol.�������) = zlCommFun.NVL(rs("�������").Value)
                
                vsf.TextMatrix(vsf.Row, mCol.����걾) = zlCommFun.NVL(rs("����걾").Value)
                vsf.TextMatrix(vsf.Row, mCol.��鲿λ) = zlCommFun.NVL(rs("��鲿λ").Value)
                vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ) = zlCommFun.NVL(rs("�ɼ���ʽ").Value)
                vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid) = zlCommFun.NVL(rs("�ɼ���ʽid").Value)
                vsf.TextMatrix(vsf.Row, mCol.��鲿λid) = zlCommFun.NVL(rs("��鲿λid").Value)
                
                vsf.RowData(vsf.Row) = zlCommFun.NVL(rs("ID").Value)
            End If
                        
            If vsf.TextMatrix(vsf.Row, mCol.���) = "����" Then
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "ִ�п���", "�ɼ���ʽ", "�ɼ�����", "����걾", "���㷽ʽ", "�Ƽ���Ŀ")
                
            ElseIf vsf.TextMatrix(vsf.Row, mCol.���) = "���" Then
                Call SetRowDefault(Val(vsf.RowData(vsf.Row)), vsf.Row, "ִ�п���", "���㷽ʽ", "�Ƽ���Ŀ")
            End If
            
            gstrSQL = "Select y.���,z.����,y.����,y.���㵥λ,x.�ּ�,y.id,Nvl(z.�Ƽ�����,1) As �Ƽ����� " & _
                        "From " & _
                            "( Select a.���,a.������Ŀid,a.�շ�ϸĿid,Sum(c.�ּ�) As �ּ� " & _
                              "From �շѼ�Ŀ c, " & _
                                   "������ͼƼ� a " & _
                              "Where a.�շ�ϸĿid = c.�շ�ϸĿid " & _
                                    "and c.ִ������<=SYSDATE and (c.��ֹ���� IS NULL OR c.��ֹ����>SYSDATE) " & _
                                    "and A.���=[2] " & _
                                    "and A.������Ŀid=[1] " & _
                              "Group by a.���,a.������Ŀid,a.�շ�ϸĿid " & _
                            ") x, " & _
                            "�շ���ĿĿ¼ y, " & _
                            "������ͼƼ� z " & _
                        "Where x.�շ�ϸĿid = y.ID " & _
                              "and z.���=x.��� " & _
                              "and z.������Ŀid=x.������Ŀid " & _
                              "and z.�շ�ϸĿid=x.�շ�ϸĿid "
                        
            Set rsPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(vsf.Row)), lngKey)
            If rsPrice.BOF = False Then
                With vsfPrice
                    Do While Not rsPrice.EOF
                        
                        If Val(.TextMatrix(.Rows - 1, mCol.p�շ���Ŀid)) > 0 Then
                            .Rows = .Rows + 1
                        End If
                        
                        .TextMatrix(.Rows - 1, mCol.p����) = zlCommFun.NVL(rsPrice("����"))
                        .TextMatrix(.Rows - 1, mCol.p���㵥λ) = zlCommFun.NVL(rsPrice("���㵥λ"))
                        .TextMatrix(.Rows - 1, mCol.p����) = zlCommFun.NVL(rsPrice("����"))
                        .TextMatrix(.Rows - 1, mCol.p��׼����) = zlCommFun.NVL(rsPrice("�ּ�"))
                        .TextMatrix(.Rows - 1, mCol.p��쵥��) = zlCommFun.NVL(rsPrice("�ּ�"))
                        .TextMatrix(.Rows - 1, mCol.p��׼���) = zlCommFun.NVL(rsPrice("����"), 0) * zlCommFun.NVL(rsPrice("�ּ�"), 0)
                        .TextMatrix(.Rows - 1, mCol.p�����) = zlCommFun.NVL(rsPrice("����"), 0) * zlCommFun.NVL(rsPrice("�ּ�"), 0)
                        .TextMatrix(.Rows - 1, mCol.p�շ���Ŀid) = zlCommFun.NVL(rsPrice("ID"))
                        .TextMatrix(.Rows - 1, mCol.p�Ƽ�����) = zlCommFun.NVL(rsPrice("�Ƽ�����"))
                        .TextMatrix(.Rows - 1, mCol.p���) = zlCommFun.NVL(rsPrice("���"))
                        .RowData(.Rows - 1) = zlCommFun.NVL(rsPrice("ID"), 0)
                        
                        If zlCommFun.NVL(rsPrice("�Ƽ�����"), 1) = 2 Then
                            .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "�ɼ���ʽ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ))
                        ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.���)) = "����" Then
                            .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "������Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                        Else
                            .TextMatrix(.Rows - 1, mCol.p�Ƽ���Ŀ) = "�����Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ))
                        End If
                        
                        Call SetRowDefault(Val(.RowData(.Rows - 1)), .Rows - 1, "�շ�ִ�п���")
                        
                        If InStr("567", .TextMatrix(.Rows - 1, mCol.p���)) > 0 Then
                            .TextMatrix(.Rows - 1, mCol.p���ÿ��) = GetStorage(Val(.RowData(.Rows - 1)), Val(.TextMatrix(.Rows - 1, mCol.pִ�п���id)))
                            Call PromptStorageWarn(Val(.TextMatrix(.Rows - 1, mCol.p����)), Val(.TextMatrix(.Rows - 1, mCol.p���ÿ��)), .TextMatrix(.Rows - 1, mCol.p����), .TextMatrix(.Rows - 1, mCol.pִ�п���), .TextMatrix(.Rows - 1, mCol.p���㵥λ), 1)
                        End If
                        rsPrice.MoveNext
                    Loop
                End With
                
                vsf.TextMatrix(vsf.Row, mCol.�����۸�) = SumPrice(1)
                vsf.TextMatrix(vsf.Row, mCol.���۸�) = SumPrice(2)
                
            End If
            
            Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
            Call vsfPrice_BeforeRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col, False)
            Call WritePrice(vsf.Row)
                                    
            rs.MoveNext
        Loop
    End If
    
    ReadTemplate = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub cbo_Click()
    If mblnNo Then Exit Sub
    
    If mstrGroup <> cbo.Text Then
        
        Call WritePrice(vsf.Row)
        Call SaveItems(mstrGroup)
        
        mstrGroup = cbo.Text
        
        Call ResetVsf(vsf)
        Call ResetVsf(vsfPrice)
        
        Call ReadItems(mstrGroup)
        Call ReadPrice(vsf.Row)
        
        Call ChangeItem(Val(vsf.TextMatrix(vsf.Row, mCol.�����۸�)), Val(vsf.TextMatrix(vsf.Row, mCol.���۸�)), 1, False)

    End If
       
End Sub

Private Sub cbo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then

        zlCommFun.PressKey vbKeyTab

    End If
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim rs As New ADODB.Recordset
    Dim lngKey As Long
    Dim strDate As String
    Dim objPoint As POINTAPI
    Dim strTmp As String
    Dim rsData As New ADODB.Recordset
    Dim rsPrice As New ADODB.Recordset

    Dim lngLoop As Long
    Dim objItem As ListItem
    Dim intRow As Long
    Dim strKeys As String

    On Error GoTo errHand
    
    Call ClientToScreen(cmd(Index).hWnd, objPoint)
    
    Select Case Index

    Case 5
            
        gstrSQL = GetPublicSQL(SQL.�����Ŀѡ��)

        Select Case mstr�Ա�
        Case "��"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 1)
        Case "Ů"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 2, 2)
        Case Else
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
        End Select
        
        If ShowTxtSelect(Me, cmd(Index), "����,1200,0,1;����,2700,0,0;��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 5100, True) Then

            rs.Filter = 0
            rs.Filter = "ѡ��=1"
            If rs.RecordCount > 0 Then

                rs.MoveFirst
                Do While Not rs.EOF
                    'ѡȡ��һ����Ŀ
                    vsf.Row = 0

                    If CheckHave(zlCommFun.NVL(rs("ID").Value)) = False Then

                        If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                            vsf.Rows = vsf.Rows + 1
                        End If
                        intRow = vsf.Rows - 1
                        vsf.Row = vsf.Rows - 1

                        vsf.Cell(flexcpText, intRow, mCol.��Ŀ + 1, intRow, vsf.Cols - 1) = ""

                        vsf.TextMatrix(intRow, mCol.���) = zlCommFun.NVL(rs("���").Value)
                        vsf.TextMatrix(intRow, mCol.��Ŀ) = zlCommFun.NVL(rs("����").Value)
                        vsf.RowData(intRow) = zlCommFun.NVL(rs("ID").Value)

'                        Call DefaultValue(Val(vsf.RowData(intRow)), 1)
'                        If vsf.TextMatrix(intRow, mCol.���) = "����" Then
'                            Call DefaultValue(Val(vsf.RowData(intRow)), 2)
'                            Call DefaultValue(Val(vsf.RowData(intRow)), 3)
'                        End If
                        
                        Call CreatePriceList(intRow)
                        Call WritePrice(intRow)

                        DataChange = True
                    End If

                    rs.MoveNext
                Loop
            End If

        End If

        EnterFocus vsf

    Case 6
    
        gstrSQL = GetPublicSQL(SQL.������ͷ���ѡ��)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(mblnGroup, 2, 1))
        If ShowTxtSelect(Me, cmd(Index), "����,1080,0,1;����,2400,0,0;����,900,0,0;˵��,1500,0,0", Me.Name & "\�������ѡ��", "����б���ѡ��һ��������͡�", rsData, rs, 8790, 5100, True) Then

            rs.Filter = 0
            rs.Filter = "ѡ��=1"
            If rs.RecordCount > 0 Then

                If Val(vsf.RowData(1)) > 0 Then
                    
                    If Not (mblnGroup = False And mbytMode = 2) Then
                        If MsgBox("�Ƿ�Ҫ�����ѡ��������Ŀ��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        
                            '��¼
                            If Trim(cbo.Text) <> "" Then
                                For lngLoop = 1 To vsf.Rows - 1
                                    mrsItems.Filter = ""
                                    mrsItems.Filter = "���='" & Trim(cbo.Text) & "' AND �嵥id=" & Val(vsf.TextMatrix(lngLoop, mCol.�嵥id))
                                    If mrsItems.RecordCount > 0 Then
                                        mrsItems.MoveFirst
                                        mrsItems("ɾ��").Value = "1"
                                    End If
                                Next
                            End If
                            
                            Call ResetVsf(vsf)
                            Call ResetVsf(vsfPrice)
                        End If
                    End If
                    
                End If

                rs.MoveFirst

                Do While Not rs.EOF

                    Call ReadTemplate(rs("ID").Value)
                    rs.MoveNext

                Loop

                DataChange = True
            End If

        End If

        EnterFocus vsf
ErrHandler:

    End Select
    
    Exit Sub
    
errHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()
        
    Dim lngKey As Long
    
    If Trim(cbo.Text) <> "" Then
        Call WritePrice(vsf.Row)
        Call SaveItems(Trim(cbo.Text))
    End If
        
    If ValidEdit = False Then Exit Sub
    
    mrsItems.Filter = ""
    
    mblnOK = True
    DataChange = False
    
    Unload Me
    
End Sub


Private Sub Form_Load()
    glngFormW = 10770
    glngFormH = 6780
    If Not InDesign Then
        glngOld = GetWindowLong(Me.hWnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf Custom_WndMessage)
    End If
    
    Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    With fraTitle
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth - .Left
    End With
    cbo.Move fraTitle.Width - cbo.Width - 45, cbo.Top
    lbl(4).Move cbo.Left - lbl(4).Width - 30
    
    With fra2
        .Left = 0
        .Top = fraTitle.Top + fraTitle.Height - 90
        .Width = fraTitle.Width
        .Height = Me.ScaleHeight - .Top - stbThis.Height - fraButton.Height + 90
    End With

    vsf.Move 45, vsf.Top, fra2.Width - vsf.Left - 45, fra2.Height - vsf.Top - 45 - vsfPrice.Height - 45
    
    With vsfPrice
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height + 45
        .Width = fra2.Width - .Left - 45
    End With
    
    cmd(6).Left = fra2.Width - cmd(6).Width - 60
    cmd(5).Left = cmd(6).Left - cmd(5).Width - 45
    
    With fraButton
        .Left = fra2.Left
        .Top = fra2.Top + fra2.Height - 90
        .Width = fra2.Width
    End With
    
    cmdCancel.Left = fraButton.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 45
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If DataChange Then
        Cancel = (MsgBox("���ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
    
    If Not InDesign Then
        Call SetWindowLong(Me.hWnd, GWL_WNDPROC, glngOld)
    End If
    
End Sub

Private Sub txtSum_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txtSum(Index)
        
End Sub

Private Sub txtSum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        zlCommFun.PressKey vbKeyTab

    End If
End Sub

Private Sub txtSum_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSum(Index).Locked Then
        glngTXTProc = GetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtSum_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtSum(Index).Locked Then
        Call SetWindowLong(txtSum(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub


Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    If vsf.Rows = 2 And Val(vsf.RowData(1)) = 0 Then
        Call ResetVsf(vsfPrice)
    Else
        Call ReadPrice(vsf.Row)
    End If
    
    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
    
    DataChange = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case mCol.ִ�п���
        
        vsf.TextMatrix(Row, mCol.ִ�п���id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.ִ�п���) = vsf.Cell(flexcpTextDisplay, Row, mCol.ִ�п���)
        
    Case mCol.�ɼ���ʽ
        
        vsf.TextMatrix(Row, mCol.�ɼ���ʽid) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.�ɼ���ʽ) = vsf.Cell(flexcpTextDisplay, Row, mCol.�ɼ���ʽ)
        
    Case mCol.�ɼ�����
    
        vsf.TextMatrix(Row, mCol.�ɼ�����id) = vsf.Body.ComboData
        vsf.TextMatrix(Row, mCol.�ɼ�����) = vsf.Cell(flexcpTextDisplay, Row, mCol.�ɼ�����)
        
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.���۸�
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
        Call ReadPrice(Row)

    '------------------------------------------------------------------------------------------------------------------
    Case mCol.�ۿ�
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.�ۿ�)), 2)
        Call ReadPrice(Row)
        
    End Select
    DataChange = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If NewRow = OldRow Then Exit Sub
    
    Call ReadPrice(NewRow)
    
    Call vsfPrice_BeforeRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col, False)

End Sub


Private Sub vsf_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case mbytMode
    Case 1
        If Val(vsf.TextMatrix(Row, mCol.����)) = 1 Then
            
            If mblnGroup = False Then
                If MsgBox("����ĿΪ���幫����Ŀ���Ƿ����Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If

    Case 2
'        If Val(vsf.TextMatrix(Row, mCol.����)) = 1 Then
            If MsgBox("����Ŀ�Ѿ���ʼ���Ƿ����Ҫȡ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
                Exit Sub
            End If
            
'        End If
    End Select
    
    '��¼
    If Trim(cbo.Text) <> "" Then
        mrsItems.Filter = ""
'        mrsItems.Filter = "���='" & Trim(cbo.Text) & "' AND ID=" & Val(vsf.RowData(Row))
        mrsItems.Filter = "���='" & Trim(cbo.Text) & "' AND �嵥id=" & Val(vsf.TextMatrix(Row, mCol.�嵥id))
        If mrsItems.RecordCount > 0 Then
            mrsItems.MoveFirst
            mrsItems("ɾ��").Value = "1"
        End If
    End If

End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf.RowData(Row)) <= 0)
    Cancel = (Val(vsf.TextMatrix(Row, mCol.ִ�п���id)) <= 0)
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    On Error GoTo errHand
    
    If NewRow = OldRow Then Exit Sub
    
    vsf.EditMode(mCol.��Ŀ) = 0
    vsf.EditMode(mCol.ִ�п���) = 0
    vsf.EditMode(mCol.��鲿λ) = 0
    vsf.EditMode(mCol.�ɼ���ʽ) = 0
    vsf.EditMode(mCol.�ɼ�����) = 0
    vsf.EditMode(mCol.����걾) = 0
    vsf.EditMode(mCol.���㷽ʽ) = 0
    vsf.EditMode(mCol.���۸�) = 0
    vsf.EditMode(mCol.�ۿ�) = 0
    
    vsf.ComboList(mCol.��Ŀ) = ""
    vsf.ComboList(mCol.ִ�п���) = ""
    vsf.ComboList(mCol.��鲿λ) = ""
    vsf.ComboList(mCol.�ɼ���ʽ) = ""
    vsf.ComboList(mCol.�ɼ�����) = ""
    vsf.ComboList(mCol.����걾) = ""
    vsf.ComboList(mCol.���㷽ʽ) = ""

'    '���ñ༭״̬

    If Val(vsf.TextMatrix(NewRow, mCol.�¼�)) = 0 Then
        vsf.EditMode(mCol.��Ŀ) = 1
        vsf.EditMode(mCol.ִ�п���) = 1
        vsf.EditMode(mCol.���㷽ʽ) = 1
        vsf.EditMode(mCol.���۸�) = 1
        vsf.EditMode(mCol.�ۿ�) = 1
        vsf.ComboList(mCol.��Ŀ) = "..."
        vsf.ComboList(mCol.ִ�п���) = " "
        vsf.ComboList(mCol.���㷽ʽ) = "����|�շ�"

        Select Case vsf.TextMatrix(NewRow, mCol.���)
        Case "���"
            vsf.EditMode(mCol.��鲿λ) = 1
             vsf.ComboList(mCol.��鲿λ) = "..."
        Case "����"
            vsf.EditMode(mCol.�ɼ���ʽ) = 1
            vsf.EditMode(mCol.�ɼ�����) = 1
            vsf.EditMode(mCol.����걾) = 1

            vsf.ComboList(mCol.�ɼ���ʽ) = " "
            vsf.ComboList(mCol.�ɼ�����) = " "
            vsf.ComboList(mCol.����걾) = " "

        End Select

        vsfPrice.ComboList(mCol.p����) = "..."
        vsfPrice.ComboList(mCol.pִ�п���) = " "
        vsfPrice.ComboList(mCol.p�Ƽ���Ŀ) = " "
        vsfPrice.EditMode(mCol.p����) = 1
        vsfPrice.EditMode(mCol.p����) = 1
        vsfPrice.EditMode(mCol.p��쵥��) = 1
        vsfPrice.EditMode(mCol.pִ�п���) = 1
        vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 1
        vsfPrice.EditMode(mCol.p�ۿ�) = 1
    Else
        vsfPrice.ComboList(mCol.p����) = ""
        vsfPrice.EditMode(mCol.p����) = 0
        vsfPrice.EditMode(mCol.p����) = 0
        vsfPrice.EditMode(mCol.p��쵥��) = 0
        vsfPrice.EditMode(mCol.pִ�п���) = 0
        vsfPrice.EditMode(mCol.p�Ƽ���Ŀ) = 0
        vsfPrice.EditMode(mCol.p�ۿ�) = 0
    End If
    
    
    If Val(vsf.TextMatrix(OldRow, mCol.�¼�)) = 0 And OldRow > 0 Then
        Call WritePrice(OldRow)
    End If
    
    If Val(vsf.TextMatrix(NewRow, mCol.�¼�)) = 0 Then
        If vsf.TextMatrix(NewRow, mCol.���) = "����" Then
            Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "�Ƽ���Ŀ", "����ִ�п���", "�ɼ���ʽ", "����걾")
            Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "�ɼ�����")
        ElseIf vsf.TextMatrix(NewRow, mCol.���) = "���" Then
            Call SetRowData(Val(vsf.RowData(NewRow)), NewRow, "�Ƽ���Ŀ", "����ִ�п���")
        End If
    End If
    
    Exit Sub
    
errHand:

    
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim bytResult As Byte
    Dim rsPrice As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strKeys As String
    Dim rsData As New ADODB.Recordset
    
    Select Case Col
        Case mCol.��Ŀ
            
            gstrSQL = GetPublicSQL(SQL.�����Ŀѡ��)
            
            Select Case mstr�Ա�
            Case "��"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 1)
            Case "Ů"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 2, 2)
            Case Else
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
            End Select
        
            If ShowGrdSelect(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;�걾��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 4500) Then
                'ѡȡ��һ����Ŀ
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
                
                vsf.Cell(flexcpText, Row, mCol.��Ŀ + 1, Row, vsf.Cols - 1) = ""
                
                vsf.EditText = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.���) = zlCommFun.NVL(rs("���").Value)
                vsf.TextMatrix(Row, mCol.��Ŀ) = zlCommFun.NVL(rs("����").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                
                If vsf.TextMatrix(Row, mCol.���) = "����" Then
                    Call SetRowDefault(Val(vsf.RowData(Row)), Row, "ִ�п���", "�ɼ���ʽ", "�ɼ�����", "����걾", "���㷽ʽ", "�Ƽ���Ŀ")
                    
                ElseIf vsf.TextMatrix(Row, mCol.���) = "���" Then
                    Call SetRowDefault(Val(vsf.RowData(Row)), Row, "ִ�п���", "���㷽ʽ", "�Ƽ���Ŀ")
                End If
                
                Call CreatePriceList(Row)
                Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                Call vsfPrice_BeforeRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col, False)
                
                Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
                
                DataChange = True
                
            End If
    End Select
    
    Select Case Col
            
        Case mCol.ִ�п���
            
            bytResult = ShowOpenList("", mCol.ִ�п���)
            If bytResult = 0 Then ShowSimpleMsg "û���ҵ���ƥ�����Ŀ��"
            If bytResult = 1 Then DataChange = True
                
        Case mCol.��鲿λ
            
            bytResult = ShowOpenList("", mCol.��鲿λ)
            If bytResult = 0 Then ShowSimpleMsg "û���ҵ���ƥ�����Ŀ��"
            If bytResult = 1 Then
                     
                Call CreatePriceList(Row)
                
                DataChange = True
            End If
            
        Case mCol.�ɼ���ʽ
            bytResult = ShowOpenList("", mCol.�ɼ���ʽ)
            If bytResult = 0 Then ShowSimpleMsg "û���ҵ���ƥ�����Ŀ��"
            If bytResult = 1 Then
                Call CreatePriceList(Row)
                DataChange = True
            End If
            
    End Select
End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim bytResult As Byte
    Dim rs As New ADODB.Recordset
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." And Col = mCol.��Ŀ Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
                        
            bytResult = ShowOpenList(UCase(vsf.EditText), Col)
            
            If bytResult = 0 Then
                'û��ƥ�����Ŀ
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
                MsgBox "û���ҵ���ƥ��������Ŀ��", vbInformation, gstrSysName
            End If
            
            If bytResult = 1 Then
                'ѡȡ��һ����Ŀ
                DataChange = True
                
                If Col = mCol.��Ŀ Then
                
                    If vsf.TextMatrix(Row, mCol.���) = "����" Then
                        Call SetRowDefault(Val(vsf.RowData(Row)), Row, "ִ�п���", "�ɼ���ʽ", "�ɼ�����", "����걾", "���㷽ʽ", "�Ƽ���Ŀ")
                        
                    ElseIf vsf.TextMatrix(Row, mCol.���) = "���" Then
                        Call SetRowDefault(Val(vsf.RowData(Row)), Row, "ִ�п���", "���㷽ʽ", "�Ƽ���Ŀ")
                    End If
                    
                    Call CreatePriceList(Row)
                    
                    Call vsf_BeforeRowColChange(0, 0, vsf.Row, vsf.Col, False)
                    Call vsfPrice_BeforeRowColChange(0, 0, vsfPrice.Row, vsfPrice.Col, False)
                    
                    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
                    
'                    Call DefaultValue(Val(vsf.RowData(Row)), 1)
'
'                    If vsf.TextMatrix(Row, mCol.���) = "����" Then
'                        Call DefaultValue(Val(vsf.RowData(Row)), 2)
'                        Call DefaultValue(Val(vsf.RowData(Row)), 3)
'                    End If
                    
'                    Call CreatePriceList(Row)
                    
                End If
                
            End If
            
            If bytResult = 2 Then
                'ȡ���˱���ѡ��
                KeyCode = 0
                Cancel = True
                
                vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                
            End If
            
        End If
    Else
        DataChange = True
    End If
End Sub

Private Sub vsf_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    
    If KeyAscii = vbKeyReturn Then
                
        If Col = 1 Then
            If Trim(vsf.TextMatrix(Row, Col)) = "" Then
                
                KeyAscii = 0
                
                cmdOK.SetFocus
                
                Cancel = True
                
            End If
        End If
    End If
    
End Sub
Private Sub vsfPrice_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    
    Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
    
    DataChange = True
End Sub

Private Sub vsfPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
    Case mCol.p�Ƽ���Ŀ
    
        If Left(vsfPrice.TextMatrix(Row, mCol.p�Ƽ���Ŀ), 4) = "�ɼ���ʽ" Then
            vsfPrice.TextMatrix(Row, mCol.p�Ƽ�����) = "2"
        Else
            vsfPrice.TextMatrix(Row, mCol.p�Ƽ�����) = "1"
        End If
        vsfPrice.TextMatrix(Row, mCol.p�Ƽ���Ŀ) = vsfPrice.Cell(flexcpTextDisplay, Row, mCol.p�Ƽ���Ŀ)
        
'    Case mCol.p����, mCol.p��쵥��
'
'        vsfPrice.TextMatrix(Row, mCol.p��׼���) = Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)) * Val(vsfPrice.TextMatrix(Row, mCol.p����))
'        vsfPrice.TextMatrix(Row, mCol.p�����) = Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)) * Val(vsfPrice.TextMatrix(Row, mCol.p����))
'
'        vsf.TextMatrix(vsf.Row, mCol.�����۸�) = SumPrice(1)
'        vsf.TextMatrix(vsf.Row, mCol.���۸�) = SumPrice(2)
        
    Case mCol.p����
        vsfPrice.TextMatrix(Row, mCol.p��׼���) = Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)) * Val(vsfPrice.TextMatrix(Row, mCol.p����))
        vsfPrice.TextMatrix(Row, mCol.p�����) = Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)) * Val(vsfPrice.TextMatrix(Row, mCol.p����))
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
                
        If InStr("567", vsfPrice.TextMatrix(Row, mCol.p���)) > 0 Then
            Call PromptStorageWarn(Val(vsfPrice.TextMatrix(Row, mCol.p����)), Val(vsfPrice.TextMatrix(Row, mCol.p���ÿ��)), vsfPrice.TextMatrix(Row, mCol.p����), vsfPrice.TextMatrix(Row, mCol.pִ�п���), vsfPrice.TextMatrix(Row, mCol.p���㵥λ), 1)
        End If
            
    Case mCol.p��쵥��
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
    
    Case mCol.p�ۿ�
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p�ۿ�)), 2)
        
    Case mCol.pִ�п���
        vsfPrice.TextMatrix(Row, mCol.pִ�п���id) = vsfPrice.Body.ComboData
        vsfPrice.TextMatrix(Row, mCol.pִ�п���) = vsfPrice.Cell(flexcpTextDisplay, Row, mCol.pִ�п���)
    End Select
    
    DataChange = True
    
End Sub

Private Sub vsfPrice_AfterNewRow(ByVal Row As Long, Col As Long)
    
    If Row > 1 Then
        vsfPrice.TextMatrix(Row, mCol.p�Ƽ���Ŀ) = vsfPrice.TextMatrix(Row - 1, mCol.p�Ƽ���Ŀ)
        If Left(vsfPrice.TextMatrix(Row, mCol.p�Ƽ���Ŀ), 4) = "�ɼ���ʽ" Then
            vsfPrice.TextMatrix(Row, mCol.p�Ƽ�����) = "2"
        Else
            vsfPrice.TextMatrix(Row, mCol.p�Ƽ�����) = "1"
        End If
    End If
    
End Sub

Private Sub vsfPrice_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str�Ƽ���Ŀ As String
    Dim str�Ƽ����� As String
    
    If vsfPrice.Rows = 2 Then
        
        str�Ƽ���Ŀ = vsfPrice.TextMatrix(1, mCol.p�Ƽ���Ŀ)
        str�Ƽ����� = vsfPrice.TextMatrix(1, mCol.p�Ƽ�����)
        
        vsfPrice.Body.Cell(flexcpText, 1, mCol.p�Ƽ���Ŀ + 1, 1, vsfPrice.Cols - 1) = ""
        vsfPrice.RowData(1) = 0

        vsfPrice.TextMatrix(1, mCol.p�Ƽ���Ŀ) = str�Ƽ���Ŀ
        vsfPrice.TextMatrix(1, mCol.p�Ƽ�����) = str�Ƽ�����
        Call vsfPrice_AfterDeleteRow(1, Col)
        
        Cancel = True
    End If
End Sub

Private Sub vsfPrice_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    If OldRow = NewRow Then Exit Sub
    
    Call SetRowData(Val(vsfPrice.RowData(NewRow)), NewRow, "�շ�ִ�п���")
    
End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    If Col = mCol.p���� Then
            
        gstrSQL = GetPublicSQL(SQL.�շ���Ŀѡ��)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If ShowGrdSelect(Me, vsfPrice, "����,1200,0,1;����,2700,0,0;��λ,600,0,0;���,1200,0,0;����,900,0,0;���,900,0,0", Me.Name & "\�շ���Ŀѡ��", "����б���ѡ��һ���շ���Ŀ��", rsData, rs, 8790, 5100) Then

            If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                Exit Sub
            End If
            
            With vsfPrice
                .EditText = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mCol.p����) = zlCommFun.NVL(rs("����").Value)
                .TextMatrix(Row, mCol.p���㵥λ) = zlCommFun.NVL(rs("��λ").Value)
                
                .TextMatrix(Row, mCol.p��׼����) = zlCommFun.NVL(rs("����").Value, 0)
                .TextMatrix(Row, mCol.p��쵥��) = .TextMatrix(Row, mCol.p��׼����)
                
                .TextMatrix(Row, mCol.p�շ���Ŀid) = zlCommFun.NVL(rs("ID").Value, 0)
                If Val(.TextMatrix(Row, mCol.p����)) < 1 Then .TextMatrix(Row, mCol.p����) = 1
                
                .TextMatrix(Row, mCol.p��׼���) = Val(.TextMatrix(Row, mCol.p��׼����)) * Val(.TextMatrix(Row, mCol.p����))
                .TextMatrix(Row, mCol.p�����) = .TextMatrix(Row, mCol.p��׼���)
                .TextMatrix(Row, mCol.p���) = zlCommFun.NVL(rs("���").Value)
                
                .RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                
                Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
                
                Call SetRowDefault(Val(.RowData(Row)), Row, "�շ�ִ�п���")
                Call SetRowData(Val(.RowData(Row)), Row, "�շ�ִ�п���")
            
                If InStr("567", .TextMatrix(Row, mCol.p���)) > 0 Then
                    .TextMatrix(Row, mCol.p���ÿ��) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.pִ�п���id)))
                    Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p����)), Val(.TextMatrix(Row, mCol.p���ÿ��)), .TextMatrix(Row, mCol.p����), .TextMatrix(Row, mCol.pִ�п���), .TextMatrix(Row, mCol.p���㵥λ), 1)
                End If
                
                
            End With
            
            DataChange = True

        End If
        
    End If
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strTmp As String
    Dim strText As String
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsfPrice.EditText, "'") > 0 Then
                KeyCode = 0
                vsfPrice.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.p����
                    
                    strText = UCase(vsfPrice.EditText)
                    gstrSQL = GetPublicSQL(SQL.�շ���Ŀ����, strText)
                    
                    If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If
                    
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText & "%", strTmp)
                    
                    If ShowGrdFilter(Me, vsfPrice, "����,1200,0,1;����,2700,0,0;��λ,600,0,0;���,1200,0,0;����,900,0,0;���,900,0,0", Me.Name & "\�շ���Ŀ����", "����б���ѡ��һ���շ���Ŀ��", rsData, rs, 8790, 5100) Then
                        
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
                        With vsfPrice
                            .EditText = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, mCol.p����) = zlCommFun.NVL(rs("����").Value)
                            .TextMatrix(Row, mCol.p���㵥λ) = zlCommFun.NVL(rs("��λ").Value)
                            
                            .TextMatrix(Row, mCol.p��׼����) = zlCommFun.NVL(rs("����").Value, 0)
                            .TextMatrix(Row, mCol.p��쵥��) = .TextMatrix(Row, mCol.p��׼����)
                            
                            .TextMatrix(Row, mCol.p�շ���Ŀid) = zlCommFun.NVL(rs("ID").Value, 0)
                            If Val(.TextMatrix(Row, mCol.p����)) < 1 Then .TextMatrix(Row, mCol.p����) = 1
                            
                            .TextMatrix(Row, mCol.p��׼���) = Val(.TextMatrix(Row, mCol.p��׼����)) * Val(.TextMatrix(Row, mCol.p����))
                            .TextMatrix(Row, mCol.p�����) = .TextMatrix(Row, mCol.p��׼���)
                            .TextMatrix(Row, mCol.p���) = zlCommFun.NVL(rs("���").Value)
                            
                            .RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                            
                            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.p��׼����)), Val(vsfPrice.TextMatrix(Row, mCol.p��쵥��)), 1)
                            
                            Call SetRowDefault(Val(.RowData(Row)), Row, "�շ�ִ�п���")
                            Call SetRowData(Val(.RowData(Row)), Row, "�շ�ִ�п���")
                            
                            If InStr("567", .TextMatrix(Row, mCol.p���)) > 0 Then
                                .TextMatrix(Row, mCol.p���ÿ��) = GetStorage(Val(.RowData(Row)), Val(.TextMatrix(Row, mCol.pִ�п���id)))
                                Call PromptStorageWarn(Val(.TextMatrix(Row, mCol.p����)), Val(.TextMatrix(Row, mCol.p���ÿ��)), .TextMatrix(Row, mCol.p����), .TextMatrix(Row, mCol.pִ�п���), .TextMatrix(Row, mCol.p���㵥λ), 1)
                            End If
                        End With
                        
                        DataChange = True
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsfPrice.Cell(flexcpData, Row, Col) = vsfPrice.Cell(flexcpData, Row, Col)
                        vsfPrice.EditText = vsfPrice.Cell(flexcpData, Row, Col)
                        vsfPrice.TextMatrix(Row, Col) = vsfPrice.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        DataChange = True
    End If
End Sub









