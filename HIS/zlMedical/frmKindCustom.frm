VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKindCustom 
   Caption         =   "�����Ŀ����"
   ClientHeight    =   7035
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   10815
   Icon            =   "frmKindCustom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10815
   Begin VB.PictureBox picBack 
      Height          =   3870
      Left            =   0
      ScaleHeight     =   3810
      ScaleWidth      =   10440
      TabIndex        =   4
      Top             =   645
      Width           =   10500
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   2
         Left            =   8760
         MaxLength       =   16
         TabIndex        =   16
         Top             =   60
         Width           =   1020
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   1
         Left            =   7215
         MaxLength       =   16
         TabIndex        =   15
         Top             =   60
         Width           =   870
      End
      Begin VB.TextBox txtSum 
         Height          =   300
         Index           =   0
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   60
         Width           =   930
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&P"
         Height          =   285
         Left            =   3675
         TabIndex        =   6
         Top             =   75
         Width           =   300
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1110
         TabIndex        =   5
         Top             =   60
         Width           =   2550
      End
      Begin zl9Medical.VsfGrid vsf 
         Height          =   1500
         Left            =   60
         TabIndex        =   8
         Top             =   405
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   2646
      End
      Begin zl9Medical.VsfGrid vsfPrice 
         Height          =   1635
         Left            =   60
         TabIndex        =   9
         Top             =   2130
         Width           =   6705
         _ExtentX        =   11827
         _ExtentY        =   2884
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ۿ�(Z)"
         Height          =   180
         Index           =   3
         Left            =   8115
         TabIndex        =   17
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���۸�(E)"
         Height          =   180
         Index           =   2
         Left            =   6210
         TabIndex        =   13
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����۸�(&B)"
         Height          =   180
         Index           =   0
         Left            =   4230
         TabIndex        =   12
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������(&T)"
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   120
         Width           =   990
      End
   End
   Begin VB.Frame fraTitle 
      Height          =   600
      Left            =   0
      TabIndex        =   10
      Top             =   -90
      Width           =   6870
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
         TabIndex        =   11
         Top             =   195
         Width           =   1800
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6675
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmKindCustom.frx":058A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13996
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
   Begin VB.Frame fraButton 
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   4455
      Width           =   6870
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4230
         TabIndex        =   3
         Top             =   225
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5400
         TabIndex        =   2
         Top             =   225
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmKindCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mlngKey As Long
Private mblnChanged As Boolean
Private mstrName As String
Private Enum mCol
    ��Ŀ��� = 1
    ��Ŀ����
    �ɼ���ʽ
    
    ����걾
    
    ��鲿λ
    
    �����۸�
    ���۸�
    �ۿ�
    �Ʒ���ϸ
    �ɼ���ʽid
    ��鲿λid
    
    �Ƽ���Ŀ = 1
    �շ���Ŀ
    ���㵥λ
    �շ�����
    �շѵ���
    ��쵥��
    p�ۿ�
    �շѽ��
    �����
    �շ���Ŀid
    �Ƽ�����
End Enum

Private mstrSQL As String

'�������Զ�����̻���************************************************************************************************

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
    str�Ƽ���Ŀ = vsfPrice.TextMatrix(1, mCol.�Ƽ���Ŀ)
    str�Ƽ����� = vsfPrice.TextMatrix(1, mCol.�Ƽ�����)
    
    vsfPrice.Body.Cell(flexcpText, 1, mCol.�Ƽ���Ŀ + 1, 1, vsfPrice.Cols - 1) = ""
    vsfPrice.RowData(1) = 0

    vsfPrice.TextMatrix(1, mCol.�Ƽ���Ŀ) = str�Ƽ���Ŀ
    vsfPrice.TextMatrix(1, mCol.�Ƽ�����) = str�Ƽ�����

    
    mstrSQL = GetPublicSQL(SQL.�����Ŀ�۱�, strKeys)
    
    If vsf.TextMatrix(intRow, mCol.��鲿λid) = "" Then
        '����򵥲�λ���
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, Val(vsf.RowData(intRow)), Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid)))
    Else
        Set rs = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption)
    End If
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If Val(vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�շ���Ŀid)) > 0 Then
                vsfPrice.Rows = vsfPrice.Rows + 1
            End If
            
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�շ���Ŀ) = zlCommFun.NVL(rs("����"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.���㵥λ) = zlCommFun.NVL(rs("���㵥λ"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�շ�����) = zlCommFun.NVL(rs("�շ�����"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�շѵ���) = zlCommFun.NVL(rs("�ּ�"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.��쵥��) = zlCommFun.NVL(rs("�ּ�"))
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.p�ۿ�) = 10
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�շѽ��) = zlCommFun.NVL(rs("�շ�����"), 0) * zlCommFun.NVL(rs("�ּ�"), 0)
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�����) = zlCommFun.NVL(rs("�շ�����"), 0) * zlCommFun.NVL(rs("�ּ�"), 0)
            vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�շ���Ŀid) = zlCommFun.NVL(rs("ID"))
            
            rs.MoveNext
        Loop
    End If
    
    vsf.TextMatrix(intRow, mCol.�����۸�) = SumPrice(mCol.�շѽ��)
    vsf.TextMatrix(intRow, mCol.���۸�) = SumPrice(mCol.�����)

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
    
    If dbMoney = 0 Then Exit Function
    
    If bytMode = 1 Then
        '�仯���
        
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
                varCol(6) = Format(Val(varCol(3)) * (db�ۿ� / 10), "0.00000")
                varCol(7) = db�ۿ�
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
                        If Val(varCol(6)) <> 0 Then
                            varCol(6) = Val(varCol(6)) + (Val(txtSum(1).Text) - dbTotal)
                            If Val(varCol(3)) <> 0 Then
                                varCol(7) = Format(10 * Val(varCol(6)) / Val(varCol(3)), "0.0000")
                            Else
                                varCol(7) = 0
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
        If bytMode = 1 Then
            '�仯���
            If dbMoney = 0 Then Exit Function
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
                varCol(6) = Format(Val(varCol(3)) * (db�ۿ� / 10), "0.00000")
                varCol(7) = db�ۿ�
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
    
    If dbMoney = 0 Then Exit Function
    
    If bytMode = 1 Then
        '�仯���
        
        '1.�����ۿ�
        db�ۿ� = Format(10 * dbTmp / dbMoney, "0.0000")
    Else
        '�仯�ۿ�
        db�ۿ� = dbTmp
        
    End If
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.��쵥��) = Format(dbMoney * db�ۿ� / 10, "0.00000")
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.p�ۿ�) = Format(db�ۿ�, "0.0000")
    
    vsfPrice.TextMatrix(vsfPrice.Row, mCol.�����) = Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.�շ�����)) * Val(vsfPrice.TextMatrix(vsfPrice.Row, mCol.��쵥��))
    
    '������Ŀ
    '------------------------------------------------------------------------------------------------------------------
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.�շѽ��))
    Next
    vsf.TextMatrix(vsf.Row, mCol.�����۸�) = dbSum
    
    dbSum = 0
    For lngLoop = 1 To vsfPrice.Rows - 1
       dbSum = dbSum + Val(vsfPrice.TextMatrix(lngLoop, mCol.�����))
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

Private Function SumPrice(ByVal intCol As Integer, Optional ByVal bytMode As Byte = 1) As Single
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim sglSum As Single
    
    If bytMode = 2 Then
        '������͵ļ۸�
        For lngLoop = 1 To vsf.Rows - 1
           sglSum = sglSum + Val(vsf.TextMatrix(lngLoop, intCol))
        Next
    Else
        For lngLoop = 1 To vsfPrice.Rows - 1
           sglSum = sglSum + Val(vsfPrice.TextMatrix(lngLoop, intCol))
        Next
    End If
    SumPrice = sglSum
    
End Function

Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
            
    txt.Locked = False
    cmd.Enabled = True
    
    If vData = False Then
        cmdOK.Tag = ""
    Else
        cmdOK.Tag = "Changed"
        txt.Locked = True
        cmd.Enabled = False
    End If
End Property

Private Property Get EditChanged() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
            
    EditChanged = (cmdOK.Tag = "Changed")
    
End Property

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
        
        Case "�Ƽ���Ŀ"
            
            If Trim(vsf.TextMatrix(intRow, mCol.��Ŀ���)) = "���" Then
                strCombList = "�����Ŀ-" & Trim(vsf.TextMatrix(intRow, mCol.��Ŀ����))
                vsfPrice.EditMode(mCol.�Ƽ���Ŀ) = 0
                vsfPrice.Body.ColComboList(mCol.�Ƽ���Ŀ) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�Ƽ���Ŀ) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�Ƽ�����) = "1"
            Else
                strCombList = "������Ŀ-" & Trim(vsf.TextMatrix(intRow, mCol.��Ŀ����))
                If Val(vsf.TextMatrix(intRow, mCol.�ɼ���ʽid)) > 0 Then
                    strCombList = strCombList & "|�ɼ���ʽ-" & Trim(vsf.TextMatrix(intRow, mCol.�ɼ���ʽ))
                    vsfPrice.EditMode(mCol.�Ƽ���Ŀ) = 1
                    vsfPrice.Body.ColComboList(mCol.�Ƽ���Ŀ) = strCombList
                Else
                    vsfPrice.EditMode(mCol.�Ƽ���Ŀ) = 0
                    vsfPrice.Body.ColComboList(mCol.�Ƽ���Ŀ) = ""
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
        Case "�Ƽ���Ŀ"
        
            If Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ���)) = "���" Then
                strCombList = "�����Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ����))
                vsfPrice.EditMode(mCol.�Ƽ���Ŀ) = 0
                vsfPrice.Body.ColComboList(mCol.�Ƽ���Ŀ) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�Ƽ���Ŀ) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�Ƽ�����) = "1"
            Else
                strCombList = "������Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ����))
                If Val(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid)) > 0 Then
                    strCombList = strCombList & "|�ɼ���ʽ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ))
                    vsfPrice.EditMode(mCol.�Ƽ���Ŀ) = 1
                    vsfPrice.Body.ColComboList(mCol.�Ƽ���Ŀ) = strCombList
                Else
                    vsfPrice.EditMode(mCol.�Ƽ���Ŀ) = 0
                    vsfPrice.Body.ColComboList(mCol.�Ƽ���Ŀ) = ""
                End If
            End If

        End Select
    Next
    
    SetRowDefault = True
    
    Exit Function
    
errHand:
    
End Function

Private Function SetDefault(ByVal lng������Ŀid As Long, ParamArray arryMode() As Variant) As Boolean
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
            
        Case "�ɼ���ʽ"
        
            gstrSQL = "SELECT A.���� AS ����,A.ID FROM ������ĿĿ¼ A,�����÷����� B WHERE A.ID=B.�÷�id AND A.���='E' AND A.��������='6' AND B.��ĿID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng������Ŀid)
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
        
        Case "�Ƽ���Ŀ"
        
            If Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ���)) = "���" Then
            
                strCombList = "�����Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ����))
                vsfPrice.EditMode(mCol.�Ƽ���Ŀ) = 0
                vsfPrice.Body.ColComboList(mCol.�Ƽ���Ŀ) = ""
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�Ƽ���Ŀ) = strCombList
                vsfPrice.TextMatrix(vsfPrice.Rows - 1, mCol.�Ƽ�����) = "1"
                
            Else
            
                strCombList = "������Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ����))
                If Val(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽid)) > 0 Then
                    strCombList = strCombList & "|�ɼ���ʽ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ))
                    vsfPrice.EditMode(mCol.�Ƽ���Ŀ) = 1
                    vsfPrice.Body.ColComboList(mCol.�Ƽ���Ŀ) = strCombList
                Else
                    vsfPrice.EditMode(mCol.�Ƽ���Ŀ) = 0
                    vsfPrice.Body.ColComboList(mCol.�Ƽ���Ŀ) = ""
                End If
                
            End If
            
        Case "����걾"
        
            gstrSQL = "SELECT 1 FROM ������ĿĿ¼ WHERE �����Ŀ=1 AND ID=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng������Ŀid)
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
        
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng������Ŀid)
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
            
        End Select
    Next
    
        
    SetDefault = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadSample(ByVal lng������Ŀid As Long) As String
    '------------------------------------------------------------------------------------------------------------------
    '����:��ȡ��ѡ�Ĳɼ���ʽ��������
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand

    gstrSQL = "SELECT distinct A.�걾���� AS ���� FROM ������Ŀ�ο� A,���鱨����Ŀ B,������ĿĿ¼ C " & _
            "WHERE C.ID=[1] AND B.������Ŀid=[1] and B.������Ŀid=A.��Ŀid"

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


Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    vsf.Rows = 2
    vsf.RowData(1) = 0
    vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
    
    Call ResetVsf(vsfPrice)
        
End Function

Public Function ShowEdit(ByVal frmMain As Object, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
            
    mlngKey = lngKey
    
    
    Set mfrmMain = frmMain
    
    If InitData = False Then Exit Function
    Call InitSysPara
    
    If mlngKey > 0 Then
        Call ReadData(mlngKey)
        Call ReadPrice(vsf.Row)
    End If
            
    EditChanged = False
    
    Me.Show 1, frmMain
    
    ShowEdit = mblnOK
    
End Function

Private Function ReadData(ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
        
    On Error GoTo errHand
    
    gstrSQL = "SELECT * FROM ������� WHERE ���=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If rs.BOF = False Then txt.Text = rs("����").Value
    mstrName = txt.Text
    
    stbThis.Panels(2).Text = "�������:" & txt.Text & "  ����:" & rs("����").Value
    
    gstrSQL = GetPublicSQL(SQL.���������Ŀ, CStr(lngKey))
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    
    If rs.BOF = False Then
        Do While Not rs.EOF
            
            If Val(vsf.RowData(vsf.Rows - 1)) > 0 Then
                vsf.Rows = vsf.Rows + 1
            End If
            
            vsf.RowData(vsf.Rows - 1) = zlCommFun.NVL(rs("ID"), 0)
            vsf.TextMatrix(vsf.Rows - 1, mCol.��Ŀ����) = zlCommFun.NVL(rs("��Ŀ����"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.��Ŀ���) = zlCommFun.NVL(rs("��Ŀ���"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.��鲿λ) = zlCommFun.NVL(rs("��鲿λ"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.�ɼ���ʽ) = zlCommFun.NVL(rs("�ɼ���ʽ"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.����걾) = zlCommFun.NVL(rs("����걾"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.�����۸�) = zlCommFun.NVL(rs("�����۸�"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.���۸�) = zlCommFun.NVL(rs("���۸�"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.�ۿ�) = zlCommFun.NVL(rs("�ۿ�"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.�ɼ���ʽid) = zlCommFun.NVL(rs("�ɼ���ʽid"))
            vsf.TextMatrix(vsf.Rows - 1, mCol.��鲿λid) = zlCommFun.NVL(rs("��鲿λid"))
            
            vsf.TextMatrix(vsf.Rows - 1, mCol.�Ʒ���ϸ) = GetTypePriceList(lngKey, zlCommFun.NVL(rs("ID"), 0))
                        
            rs.MoveNext
        Loop
    End If
    
    Call ChangeItem(Val(vsf.TextMatrix(vsf.Row, mCol.�����۸�)), Val(vsf.TextMatrix(vsf.Row, mCol.���۸�)), 1, False)

                
    ReadData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    
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
        .NewColumn "��Ŀ���", 900, 1
        .NewColumn "��Ŀ����", 2700, 1, "...", 1
        .NewColumn "�ɼ���ʽ", 1200, 1
        .NewColumn "����걾", 900, 1
        .NewColumn "��鲿λ", 1500, 1
        .NewColumn "�����۸�", 900, 7
        .NewColumn "���۸�", 900, 7, , 1
        .NewColumn "�ۿ�", 900, 7, , 1
        .NewColumn "�Ʒ���ϸ", 0, 1
        .NewColumn "�ɼ���ʽid", 0, 1
        .NewColumn "��鲿λid", 0, 1
        .FixedCols = 1
        
        .Body.ColFormat(mCol.�����۸�) = "0.00"
        .Body.ColFormat(mCol.���۸�) = "0.00"
        .Body.ColFormat(mCol.�ۿ�) = "0.0000"
        .SelectMode = True
    End With
    
    With vsfPrice
        .Cols = 0
        
        .NewColumn "", 255, 4
        
        .NewColumn "�Ƽ���Ŀ", 3000, 1, " |", 1
        .NewColumn "�շ���Ŀ", 2100, 1, "...", 1
        .NewColumn "��λ", 900, 1
        .NewColumn "����", 600, 7, , 1
        .NewColumn "�շѵ���", 900, 7
        .NewColumn "��쵥��", 900, 7, , 1
        .NewColumn "�ۿ�", 900, 7, , 1
        .NewColumn "�շѽ��", 900, 7
        .NewColumn "�����", 900, 7
        .NewColumn "�շ���Ŀid", 0, 1
        .NewColumn "�Ƽ�����", 0, 1
        
        .Body.ColFormat(mCol.�շѵ���) = "0.00000"
        .Body.ColFormat(mCol.�շѽ��) = "0.00"
        .Body.ColFormat(mCol.p�ۿ�) = "0.0000"
        .Body.ColFormat(mCol.��쵥��) = "0.00000"
        .Body.ColFormat(mCol.�����) = "0.00"
        
        .FixedCols = 1
        .SelectMode = True
    End With
    
    InitData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
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
    
                    If Val(varCol(5)) = 2 Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.�Ƽ���Ŀ) = "�ɼ���ʽ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.�ɼ���ʽ))
                    ElseIf Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ���)) = "����" Then
                        vsfPrice.TextMatrix(lngRow + 1, mCol.�Ƽ���Ŀ) = "������Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ����))
                    Else
                        vsfPrice.TextMatrix(lngRow + 1, mCol.�Ƽ���Ŀ) = "�����Ŀ-" & Trim(vsf.TextMatrix(vsf.Row, mCol.��Ŀ����))
                    End If
                    
                    vsfPrice.TextMatrix(lngRow + 1, mCol.�շ���Ŀ) = varCol(0)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.���㵥λ) = varCol(1)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.�շ�����) = varCol(2)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.�շѵ���) = varCol(3)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.��쵥��) = varCol(6)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.�շѽ��) = Val(varCol(2)) * Val(varCol(3))
                    vsfPrice.TextMatrix(lngRow + 1, mCol.�����) = Val(varCol(2)) * Val(varCol(6))
                    vsfPrice.TextMatrix(lngRow + 1, mCol.�շ���Ŀid) = varCol(4)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.�Ƽ�����) = varCol(5)
                    vsfPrice.TextMatrix(lngRow + 1, mCol.p�ۿ�) = varCol(7)
                    
                Next
            End If
        Next
        
    End If
    
    If vsf.TextMatrix(intRow, mCol.��Ŀ���) = "����" Then
        Call SetRowData(Val(vsf.RowData(intRow)), intRow, "�Ƽ���Ŀ")
    ElseIf vsf.TextMatrix(intRow, mCol.��Ŀ���) = "���" Then
        Call SetRowData(Val(vsf.RowData(intRow)), intRow, "�Ƽ���Ŀ")
    End If
    
    ReadPrice = True
    
End Function

Private Function WritePrice(ByVal intRow As Integer) As Boolean
    Dim strTmp As String
    Dim lngRow As Long
    Dim varCol As Variant
    
    On Error GoTo errHand
    
    For lngRow = 1 To vsfPrice.Rows - 1
        If Val(vsfPrice.TextMatrix(lngRow, mCol.�շ���Ŀid)) > 0 Then
            
            varCol = Split(String(8, ":"), ":")
                                
            varCol(0) = vsfPrice.TextMatrix(lngRow, mCol.�շ���Ŀ)
            varCol(1) = vsfPrice.TextMatrix(lngRow, mCol.���㵥λ)
            varCol(2) = vsfPrice.TextMatrix(lngRow, mCol.�շ�����)
            varCol(3) = vsfPrice.TextMatrix(lngRow, mCol.�շѵ���)
            varCol(4) = vsfPrice.TextMatrix(lngRow, mCol.�շ���Ŀid)
            varCol(5) = vsfPrice.TextMatrix(lngRow, mCol.�Ƽ�����)
            varCol(6) = vsfPrice.TextMatrix(lngRow, mCol.��쵥��)
            varCol(7) = vsfPrice.TextMatrix(lngRow, mCol.p�ۿ�)
                        
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

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    For lngLoop = 1 To vsf.Rows - 1
        If lngLoop <> vsf.Rows - 1 Then
            If Val(vsf.RowData(lngLoop)) = 0 Then
                ShowSimpleMsg "�� " & lngLoop & " ���������벻����������������Ч�������Ŀ��"
                LocationGrid vsf, lngLoop, mCol.��Ŀ����
                Exit Function
            End If
        End If
    Next
    
    ValidEdit = True
    
End Function

Private Function SaveEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��������
    '����:  True        ����ɹ�
    '       False       ����ʧ��
    '------------------------------------------------------------------------------------------------------------------
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngRow As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    strSQL(ReDimArray(strSQL)) = "ZL_�������Ŀ¼_DELETE(" & mlngKey & ")"
    
    For lngLoop = 1 To vsf.Rows - 1
        If Val(vsf.RowData(lngLoop)) > 0 Then
            
            
            strTmp = ""
            If vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ) <> "" Then
                varRow = Split(vsf.TextMatrix(lngLoop, mCol.�Ʒ���ϸ), ";")
                For lngRow = 0 To UBound(varRow)
                    
                    varCol = Split(varRow(lngRow), ":")
                    
                    If strTmp <> "" Then strTmp = strTmp & ";"
                    strTmp = strTmp & varCol(4) & ":" & varCol(2) & ":" & varCol(5) & ":" & Format(varCol(7) / 10, "0.00000")
                            
                Next
            End If
            
            strSQL(ReDimArray(strSQL)) = "ZL_�������Ŀ¼_INSERT(" & mlngKey & "," & _
                                                                Val(vsf.RowData(lngLoop)) & ",'" & _
                                                                vsf.TextMatrix(lngLoop, mCol.��鲿λ) & "'," & _
                                                                Val(vsf.TextMatrix(lngLoop, mCol.�ɼ���ʽid)) & ",'" & _
                                                                vsf.TextMatrix(lngLoop, mCol.��鲿λid) & "','" & _
                                                                vsf.TextMatrix(lngLoop, mCol.����걾) & "','" & _
                                                                strTmp & "')"
        End If
    Next
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    gcnOracle.CommitTrans
    blnTran = False
    
    SaveEdit = True
    
    Exit Function
    
errHand:
    
    If ErrCenter = 1 Then
        Resume
    End If
    
    If blnTran Then gcnOracle.RollbackTrans
    
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


'���������弰��ؼ����¼�����******************************************************************************************
Private Sub cmd_Click()
    
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    gstrSQL = GetPublicSQL(SQL.�������ѡ��)
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If ShowTxtSelect(Me, txt, "����,1080,0,1;����,2400,0,0;����,900,0,0;˵��,1500,0,0", Me.Name & "\�������ѡ��", "����б���ѡ��һ��������͡�", rsData, rs, 8790, 5100) Then
        
        Call ClearData
        
        txt.Text = zlCommFun.NVL(rs("����"))
        mlngKey = zlCommFun.NVL(rs("ID"))
        
        Call ReadData(mlngKey)
        Call ReadPrice(vsf.Row)
        
        txt.Tag = ""
        mstrName = txt.Text
        
        EditChanged = False
        
    End If

    Call LocationObj(txt)

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    Call WritePrice(vsf.Row)
    
    If ValidEdit = False Then Exit Sub
    
    If SaveEdit Then
       
        Call mfrmMain.EditRefresh("�������", mlngKey)
        
        If mlngKey = 0 Then
            
            EditChanged = False
        Else
            EditChanged = False
        End If
        
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With fraTitle
        .Left = 0
        .Top = -90
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picBack
        .Left = 0
        .Top = fraTitle.Top + fraTitle.Height
        .Width = fraTitle.Width
        .Height = Me.ScaleHeight - .Top - stbThis.Height - fraButton.Height
    End With
    
    With fraButton
        .Left = picBack.Left
        .Top = picBack.Top + picBack.Height - 90
        .Width = picBack.Width
    End With
    
    cmdCancel.Left = fraButton.Width - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 45

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If EditChanged Then
        Cancel = (MsgBox("�������޸ĵ����ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub picBack_Resize()
    On Error Resume Next
    
    With vsf
        
        .Left = 60
        .Top = txt.Top + txt.Height + 45
        .Width = picBack.Width - .Left - 60
        .Height = picBack.Height - .Top - vsfPrice.Height - 60 - 45
        
    End With
    
    With vsfPrice
        .Left = vsf.Left
        .Top = vsf.Top + vsf.Height + 45
        .Width = picBack.Width - .Left - 60
    End With
    
End Sub

Private Sub txt_Change()
    txt.Tag = "Changed"
End Sub

Private Sub txt_GotFocus()
    zlControl.TxtSelAll txt
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim strText As String
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    
    If KeyAscii = vbKeyReturn Then
        If txt.Tag = "Changed" Then
            
            txt.Tag = ""
            strText = UCase(txt.Text)
            
            gstrSQL = GetPublicSQL(SQL.������͹���ѡ��)
            
            If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                strTmp = strText & "%"
            Else
                strTmp = "%" & strText & "%"
            End If
                    
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText & "%", strTmp)
            
            If ShowTxtFilter(Me, txt, "����,1080,0,1;����,2400,0,0;����,900,0,0;˵��,1500,0,0", Me.Name & "\������͹���ѡ��", "����б���ѡ��һ��������͡�", rsData, rs) Then
                
                Call ClearData
                
                txt.Text = zlCommFun.NVL(rs("����"))
                mlngKey = zlCommFun.NVL(rs("ID"))
                
                Call ReadData(mlngKey)
                Call ReadPrice(vsf.Row)
                
                txt.Tag = ""
                mstrName = txt.Text
                
                zlCommFun.PressKey vbKeyTab
                zlCommFun.PressKey vbKeyTab
            Else
                txt.Text = mstrName
            End If
            txt.Tag = ""
            Call LocationObj(txt)
            
        Else
            zlCommFun.PressKey vbKeyTab
            zlCommFun.PressKey vbKeyTab
        End If
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
    End If
    
End Sub

Private Sub txt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 And txt.Locked Then
        glngTXTProc = GetWindowLong(txt.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txt.Locked Then
        Call SetWindowLong(txt.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Cancel As Boolean)

    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)

    If txt.Tag = "Changed" Then txt.Text = mstrName
    
End Sub

Private Sub txtSum_GotFocus(Index As Integer)
    
    zlControl.TxtSelAll txtSum(Index)
        
End Sub

Private Sub txtSum_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        
        Call WritePrice(vsf.Row)
        
        If Index = 1 Then Call ChangeTotal(Val(txtSum(0).Text), Val(txtSum(1).Text), 1)
        If Index = 2 Then Call ChangeTotal(Val(txtSum(0).Text), Val(txtSum(2).Text), 2)
        
        Call ReadPrice(vsf.Row)
   
        zlCommFun.PressKey vbKeyTab
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
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

Private Sub txtSum_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txtSum(Index).Text, txtSum(Index).MaxLength)
    
    If Index = 1 Then
        If InStr(txtSum(1).Text, ".") > 0 Then
            If Len(Mid(txtSum(1).Text, InStr(txtSum(1).Text, ".") + 1)) > 2 Then
                MsgBox "ֻ����������λС��λ����", vbExclamation, gstrSysName
                Cancel = True
            End If
        End If
    End If
    
End Sub

Private Sub vsf_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    If vsf.Rows = 2 And Val(vsf.RowData(1)) = 0 Then
        Call ResetVsf(vsfPrice)
    Else
        Call ReadPrice(vsf.Row)
    End If
    
    Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
    
    EditChanged = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim db�ۿ� As Double
    Dim lngLoop As Long
    
    Select Case Col
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.��Ŀ���
        If vsf.EditText <> vsf.Cell(flexcpData, Row, Col) Then
            vsf.RowData(Row) = 0
            vsf.Cell(flexcpText, Row, mCol.��Ŀ����, Row, vsf.Cols - 1) = ""
            
            EditChanged = True
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case mCol.���۸�
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
        Call ReadPrice(Row)

    '------------------------------------------------------------------------------------------------------------------
    Case mCol.�ۿ�
        
        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.�ۿ�)), 2)
        Call ReadPrice(Row)
        
    End Select
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    '���ñ༭״̬
    
    If NewRow = OldRow Then Exit Sub
    
    Select Case vsf.TextMatrix(NewRow, mCol.��Ŀ���)
        Case "���"
            vsf.EditMode(mCol.�ɼ���ʽ) = 0
            vsf.EditMode(mCol.����걾) = 0
            vsf.EditMode(mCol.��鲿λ) = 1
            
            vsf.ComboList(mCol.�ɼ���ʽ) = ""
            vsf.ComboList(mCol.����걾) = ""
            vsf.ComboList(mCol.��鲿λ) = "..."
        Case "����"
            vsf.EditMode(mCol.�ɼ���ʽ) = 1
            vsf.EditMode(mCol.����걾) = 1
            vsf.EditMode(mCol.��鲿λ) = 0
            
            vsf.ComboList(mCol.�ɼ���ʽ) = "..."
            vsf.ComboList(mCol.����걾) = " "
            vsf.ComboList(mCol.��鲿λ) = ""
    End Select
    
    '�Ƽ���Ŀ�б�
    
    mstrSQL = "Select * From �������Ŀ¼ where ���=[1] And ������Ŀid=1"
    
    Call ReadPrice(NewRow)

End Sub

Private Sub vsf_BeforeComboList(ByVal OldCol As Long, ByVal NewCol As Long, ComboList As String, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    
    '�����½������б�
    
    Select Case NewCol
    Case mCol.����걾
        
        ComboList = ReadSample(Val(vsf.RowData(vsf.Row)))
        
    End Select
    
    If ComboList = "" Then ComboList = " |"
End Sub

Private Sub vsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = (Val(vsf.RowData(Row)) <= 0)
End Sub

Private Sub vsf_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    On Error GoTo errHand
    
    Call WritePrice(OldRow)
    
    Exit Sub
    
errHand:
End Sub

Private Sub vsf_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsData As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Select Case Col
        Case mCol.��Ŀ����
            
            gstrSQL = GetPublicSQL(SQL.�����Ŀѡ��)
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, 1, 2)
            If ShowGrdSelect(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 5100) Then
                
                If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                    ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                    Exit Sub
                End If
                
                vsf.EditText = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.��Ŀ���) = zlCommFun.NVL(rs("���").Value)
                vsf.TextMatrix(Row, mCol.��Ŀ����) = zlCommFun.NVL(rs("����").Value)
                vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                                
                If vsf.TextMatrix(Row, mCol.��Ŀ���) = "����" Then
                    Call SetDefault(Val(vsf.RowData(Row)), "ִ�п���", "����걾", "�ɼ���ʽ", "�ɼ�ִ��", "�Ƽ���Ŀ")
                Else
                    Call SetDefault(Val(vsf.RowData(Row)), "ִ�п���", "�Ƽ���Ŀ")
                End If
                
                Call CreatePriceList(Row)
                Call WritePrice(Row)
                
                Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
                
                EditChanged = True
                
            End If
        
        Case mCol.��鲿λ
                        
            gstrSQL = "select B.�걾��λ AS ����,B.ID,0 AS ѡ�� from ������Ŀ��� A,������ĿĿ¼ B WHERE (B.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or B.����ʱ�� is NULL) AND A.������ĿID=B.ID AND A.�������ID=[1]"
            
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(Row)))
            
            If ShowGrdSelect(Me, vsf, "����,3300,0,0", Me.Name & "\��鲿λѡ��", "����б���ѡ��һ����鲿λ��", rsData, rs, 8790, 5100) Then
                
                vsf.TextMatrix(Row, vsf.Col) = ""
                vsf.TextMatrix(Row, mCol.��鲿λid) = ""
                
                rs.Filter = ""
                rs.Filter = "ѡ��=1"
                If rs.RecordCount > 0 Then
                    rs.MoveFirst
                    Do While Not rs.EOF
                        vsf.TextMatrix(Row, vsf.Col) = vsf.TextMatrix(Row, vsf.Col) & zlCommFun.NVL(rs("����").Value) & ","
                        vsf.TextMatrix(Row, mCol.��鲿λid) = vsf.TextMatrix(Row, mCol.��鲿λid) & zlCommFun.NVL(rs("ID").Value) & ","
                        rs.MoveNext
                    Loop
                    
                    If vsf.TextMatrix(Row, mCol.��鲿λ) <> "" Then vsf.TextMatrix(Row, mCol.��鲿λ) = Mid(vsf.TextMatrix(Row, mCol.��鲿λ), 1, Len(vsf.TextMatrix(Row, mCol.��鲿λ)) - 1)
                    If vsf.TextMatrix(Row, mCol.��鲿λid) <> "" Then vsf.TextMatrix(Row, mCol.��鲿λid) = Mid(vsf.TextMatrix(Row, mCol.��鲿λid), 1, Len(vsf.TextMatrix(Row, mCol.��鲿λid)) - 1)
                    
                End If
                                
                Call CreatePriceList(Row)
                
                EditChanged = True
                
            End If
            
        Case mCol.�ɼ���ʽ
        
            
            gstrSQL = "SELECT 1 As ĩ��,A.ID,A.����,A.���� " & _
                "FROM ������ĿĿ¼ A,�����÷����� B " & _
                "WHERE (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) AND A.���='E' AND A.��������='6' AND A.ID=B.�÷�id AND B.��Ŀid=[1]"
            
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf.RowData(Row)))
            If rs.BOF Then
            
                gstrSQL = "SELECT 1 As ĩ��,A.ID,A.����,A.���� " & _
                    "FROM ������ĿĿ¼ A WHERE (A.����ʱ�� = To_Date('30000101', 'YYYYMMDD') Or A.����ʱ�� is NULL) AND A.���='E' AND A.��������='6' "
            End If
               
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
            If ShowGrdSelect(Me, vsf, "����,1200,0,1;����,3300,0,0", Me.Name & "\�ɼ���ʽѡ��", "����б���ѡ��һ���ɼ���ʽ��", rsData, rs, 8790, 5100) Then
                
                vsf.Cell(flexcpData, Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, vsf.Col) = zlCommFun.NVL(rs("����").Value)
                vsf.TextMatrix(Row, mCol.�ɼ���ʽid) = zlCommFun.NVL(rs("ID").Value)
                
                Call CreatePriceList(Row)
                
                EditChanged = True
            End If
    End Select

End Sub

Private Sub vsf_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim strTmp As String
    Dim strText As String
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsf.EditText, "'") > 0 Then
                KeyCode = 0
                vsf.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.��Ŀ����
                    
                    strText = UCase(vsf.EditText)
                    gstrSQL = GetPublicSQL(SQL.�����Ŀ����ѡ��, strText)
                    
                    If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If

                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "C", "D", strText & "%", strTmp, 1, 2)
                    
                    If ShowGrdFilter(Me, vsf, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;���,900,0,0", Me.Name & "\�����Ŀ����ѡ��", "����б���ѡ��һ�������Ŀ��", rsData, rs, 8790, 5100) Then

                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If

                        vsf.EditText = zlCommFun.NVL(rs("����").Value)
                        vsf.TextMatrix(Row, mCol.��Ŀ���) = zlCommFun.NVL(rs("���").Value)
                        vsf.TextMatrix(Row, mCol.��Ŀ����) = zlCommFun.NVL(rs("����").Value)
                        vsf.Cell(flexcpData, Row, Col) = vsf.TextMatrix(Row, Col)
                        vsf.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                        
                        If vsf.TextMatrix(Row, mCol.��Ŀ���) = "����" Then
                            Call SetDefault(Val(vsf.RowData(Row)), "ִ�п���", "����걾", "�ɼ���ʽ", "�ɼ�ִ��", "�Ƽ���Ŀ")
                        Else
                            Call SetDefault(Val(vsf.RowData(Row)), "ִ�п���", "�Ƽ���Ŀ")
                        End If
                        
                        Call CreatePriceList(Row)
                        
                        Call ChangeItem(Val(vsf.TextMatrix(Row, mCol.�����۸�)), Val(vsf.TextMatrix(Row, mCol.���۸�)), 1)
                
                        EditChanged = True
                    Else
                        KeyCode = 0
                        Cancel = True
                        
                        vsf.Cell(flexcpData, Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        vsf.EditText = vsf.Cell(flexcpData, Row, Col)
                        vsf.TextMatrix(Row, Col) = vsf.Cell(flexcpData, Row, Col)
                        
                    End If
            End Select
        End If
    Else
        EditChanged = True
    End If
End Sub


Private Sub vsfPrice_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)

    Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.�շѵ���)), Val(vsfPrice.TextMatrix(Row, mCol.��쵥��)), 1)
    
    EditChanged = True
    
End Sub

Private Sub vsfPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Select Case Col
    Case mCol.�Ƽ���Ŀ
    
        If Left(vsfPrice.TextMatrix(Row, mCol.�Ƽ���Ŀ), 4) = "�ɼ���ʽ" Then
            vsfPrice.TextMatrix(Row, mCol.�Ƽ�����) = "2"
        Else
            vsfPrice.TextMatrix(Row, mCol.�Ƽ�����) = "1"
        End If
        
    Case mCol.�շ�����
        vsfPrice.TextMatrix(Row, mCol.�շѽ��) = Val(vsfPrice.TextMatrix(Row, mCol.�շѵ���)) * Val(vsfPrice.TextMatrix(Row, mCol.�շ�����))
        vsfPrice.TextMatrix(Row, mCol.�����) = Val(vsfPrice.TextMatrix(Row, mCol.��쵥��)) * Val(vsfPrice.TextMatrix(Row, mCol.�շ�����))
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.�շѵ���)), Val(vsfPrice.TextMatrix(Row, mCol.��쵥��)), 1)
                
    Case mCol.��쵥��
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.�շѵ���)), Val(vsfPrice.TextMatrix(Row, mCol.��쵥��)), 1)
    
    Case mCol.p�ۿ�
        
        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.�շѵ���)), Val(vsfPrice.TextMatrix(Row, mCol.p�ۿ�)), 2)
        
    End Select
    
    EditChanged = True
    
End Sub

Private Sub vsfPrice_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str�Ƽ���Ŀ As String
    Dim str�Ƽ����� As String
    
    If vsfPrice.Rows = 2 Then
        
        str�Ƽ���Ŀ = vsfPrice.TextMatrix(1, mCol.�Ƽ���Ŀ)
        str�Ƽ����� = vsfPrice.TextMatrix(1, mCol.�Ƽ�����)
        
        vsfPrice.Body.Cell(flexcpText, 1, mCol.�Ƽ���Ŀ + 1, 1, vsfPrice.Cols - 1) = ""
        vsfPrice.RowData(1) = 0

        vsfPrice.TextMatrix(1, mCol.�Ƽ���Ŀ) = str�Ƽ���Ŀ
        vsfPrice.TextMatrix(1, mCol.�Ƽ�����) = str�Ƽ�����
        Call vsfPrice_AfterDeleteRow(1, Col)
        
        Cancel = True
    End If
End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    If Col = mCol.�շ���Ŀ Then
            
        gstrSQL = GetPublicSQL(SQL.�շ���Ŀѡ��)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If ShowGrdSelect(Me, vsfPrice, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;����,900,0,0;���,900,0,0", Me.Name & "\�շ���Ŀѡ��", "����б���ѡ��һ���շ���Ŀ��", rsData, rs, 8790, 5100) Then

            If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                Exit Sub
            End If

            vsfPrice.EditText = zlCommFun.NVL(rs("����").Value)
            vsfPrice.TextMatrix(Row, mCol.��Ŀ���) = zlCommFun.NVL(rs("���").Value)
            vsfPrice.TextMatrix(Row, mCol.��Ŀ����) = zlCommFun.NVL(rs("����").Value)
            vsfPrice.TextMatrix(Row, mCol.���㵥λ) = zlCommFun.NVL(rs("��λ").Value)
            vsfPrice.TextMatrix(Row, mCol.�շѵ���) = zlCommFun.NVL(rs("����").Value, 0)
            vsfPrice.TextMatrix(Row, mCol.��쵥��) = zlCommFun.NVL(rs("����").Value, 0)
            vsfPrice.TextMatrix(Row, mCol.�շ���Ŀid) = zlCommFun.NVL(rs("ID").Value, 0)
            
            If Val(vsfPrice.TextMatrix(Row, mCol.�շ�����)) < 1 Then vsfPrice.TextMatrix(Row, mCol.�շ�����) = 1
            
            vsfPrice.TextMatrix(Row, mCol.�շѽ��) = zlCommFun.NVL(rs("����").Value, 0) * Val(vsfPrice.TextMatrix(Row, mCol.�շ�����))
            vsfPrice.TextMatrix(Row, mCol.�����) = zlCommFun.NVL(rs("����").Value, 0) * Val(vsfPrice.TextMatrix(Row, mCol.�շ�����))
            
            vsfPrice.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
            
            vsf.TextMatrix(vsf.Row, mCol.�����۸�) = SumPrice(mCol.�շѽ��)
            vsf.TextMatrix(vsf.Row, mCol.���۸�) = SumPrice(mCol.�����)
            
            Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.�շѵ���)), Val(vsfPrice.TextMatrix(Row, mCol.��쵥��)), 1)
                
            EditChanged = True

        End If
        
    End If
            
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim strText As String
    Dim strTmp As String
    
    If KeyCode = vbKeyReturn Then
        If ComboList = "..." Then
            
            If InStr(vsfPrice.EditText, "'") > 0 Then
                KeyCode = 0
                vsfPrice.EditText = ""
                Cancel = True
                Exit Sub
            End If
    
            Select Case Col
                Case mCol.�շ���Ŀ
                    
                    strText = UCase(vsfPrice.EditText)
                    
                    gstrSQL = GetPublicSQL(SQL.�շ���Ŀ����, strText)
                    
                    If ParamInfo.��Ŀ����ƥ�䷽ʽ = 1 Then
                        strTmp = strText & "%"
                    Else
                        strTmp = "%" & strText & "%"
                    End If
                    
                    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strText & "%", strTmp)
                    
                    If ShowGrdFilter(Me, vsfPrice, "����,1200,0,1;����,2700,0,0;��λ,900,0,0;����,900,0,0;���,900,0,0", Me.Name & "\�շ���Ŀ����", "����б���ѡ��һ���շ���Ŀ��", rsData, rs, 8790, 5100) Then
                        
                        
                        If CheckHave(zlCommFun.NVL(rs("ID").Value)) Then
                            ShowSimpleMsg "ѡ�����Ŀ��" & zlCommFun.NVL(rs("����").Value) & "���ѱ�ѡ��"
                            Exit Sub
                        End If
            
                        vsfPrice.EditText = zlCommFun.NVL(rs("����").Value)
                        vsfPrice.TextMatrix(Row, mCol.��Ŀ���) = zlCommFun.NVL(rs("���").Value)
                        vsfPrice.TextMatrix(Row, mCol.��Ŀ����) = zlCommFun.NVL(rs("����").Value)
                        vsfPrice.TextMatrix(Row, mCol.���㵥λ) = zlCommFun.NVL(rs("��λ").Value)
                        vsfPrice.TextMatrix(Row, mCol.�շѵ���) = zlCommFun.NVL(rs("����").Value, 0)
                        vsfPrice.TextMatrix(Row, mCol.��쵥��) = zlCommFun.NVL(rs("����").Value, 0)
                        vsfPrice.TextMatrix(Row, mCol.�շ���Ŀid) = zlCommFun.NVL(rs("ID").Value, 0)
                        
                        If Val(vsfPrice.TextMatrix(Row, mCol.�շ�����)) < 1 Then vsfPrice.TextMatrix(Row, mCol.�շ�����) = 1
                        
                        vsfPrice.TextMatrix(Row, mCol.�շѽ��) = zlCommFun.NVL(rs("����").Value, 0) * Val(vsfPrice.TextMatrix(Row, mCol.�շ�����))
                        vsfPrice.TextMatrix(Row, mCol.�����) = zlCommFun.NVL(rs("����").Value, 0) * Val(vsfPrice.TextMatrix(Row, mCol.�շ�����))
                        
                        vsfPrice.RowData(Row) = zlCommFun.NVL(rs("ID").Value)
                        
                        Call ChangePrice(Val(vsfPrice.TextMatrix(Row, mCol.�շѵ���)), Val(vsfPrice.TextMatrix(Row, mCol.��쵥��)), 1)
                        
                        EditChanged = True
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
        EditChanged = True
    End If
End Sub




