VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMultiBalance 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�շѽ���"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMultiBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtOwe 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   210
      Width           =   2010
   End
   Begin VB.TextBox txtTmp 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   1560
      MaxLength       =   10
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   6330
      TabIndex        =   12
      Top             =   4530
      Width           =   1400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   4800
      TabIndex        =   11
      Top             =   4530
      Width           =   1400
   End
   Begin VB.Frame Frame1 
      Height          =   150
      Left            =   -90
      TabIndex        =   13
      Top             =   4215
      Width           =   7845
   End
   Begin VB.TextBox txt�Ҳ� 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   4800
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   3720
      Width           =   1995
   End
   Begin VB.TextBox txt�ɿ� 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   1215
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   3720
      Width           =   1995
   End
   Begin VB.TextBox txtPay 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   135
      Width           =   1995
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPay 
      Height          =   2385
      Left            =   195
      TabIndex        =   6
      Top             =   1080
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   4207
      _Version        =   393216
      Rows            =   7
      Cols            =   4
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "^ ���㷽ʽ |^  ������  |^    �������    |^          ��ע    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Label lblOwe 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      Height          =   240
      Left            =   3690
      TabIndex        =   2
      Top             =   255
      Width           =   960
   End
   Begin VB.Label lbl�Ҳ� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֽ��Ҳ�"
      Height          =   240
      Left            =   3690
      TabIndex        =   9
      Top             =   3810
      Width           =   960
   End
   Begin VB.Label lbl�ɿ� 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֽ�ɿ�"
      Height          =   240
      Left            =   195
      TabIndex        =   7
      Top             =   3810
      Width           =   960
   End
   Begin VB.Label lblBalance 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������ϸ"
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   720
      Width           =   960
   End
   Begin VB.Label lblPay 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����ϼ�"
      Height          =   240
      Left            =   195
      TabIndex        =   0
      Top             =   255
      Width           =   960
   End
End
Attribute VB_Name = "frmMultiBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mintDefault As Integer 'ȱʡ���㷽ʽ��(Ϊ0��ʾû��)
Private mintInsure As Integer '��:��Ϊҽ������ʱ�Ŵ���
Private mlng����ID As Long '��:��Ϊҽ������ʱ�Ŵ���
Private mcurPay As Currency '��:�ſ����ռ���Ԥ�����Ӧ�ɺϼ�(δ����ֱҺ�С����)����:ʵ��֧���ϼ�(����֮��)
Private mstrBalance As String '��/��:���㷽ʽ|������|�������|ժҪ
Private mcurError As Currency '��:�����
Private mrs���㷽ʽ As ADODB.Recordset
Private mlngPayRow As Long, mstrӦ�����㷽ʽ As String
Private mcur�ɿ� As Currency    '��¼���ν���Ľɿ���Ҳ����
Private mcur�Ҳ� As Currency
Private mcur�ֽ� As Currency  '35135
Private mcurOneCard As Currency 'һ��ͨ���,��һ��ͨ����ɿ���ʱ�Ŵ���
Private mblnHotKey As Boolean

Private Enum COLS
    C0��ʽ = 0
    C1��� = 1
    C2���� = 2
    C3��ע = 3
End Enum


Public Function ShowMe(frmParent As Object, _
    ByVal intInsure As Integer, ByVal lng����ID As Long, curPay As Currency, _
    strBalance As String, curError As Currency, rs���㷽ʽ As ADODB.Recordset, _
    cur�ɿ� As Currency, cur�Ҳ� As Currency, CurOneCard As Currency, _
    cur�ֽ� As Currency) As Boolean
    
    mintInsure = intInsure
    mlng����ID = lng����ID
    mcurPay = curPay
    mstrBalance = strBalance
    mcurError = curError
    mcurOneCard = CurOneCard
    mcur�ֽ� = 0
    Set mrs���㷽ʽ = rs���㷽ʽ
    
    Me.Show 1, frmParent
    
    If mblnOK Then
        curPay = mcurPay
        strBalance = mstrBalance
        curError = mcurError
        cur�ɿ� = mcur�ɿ�
        cur�Ҳ� = mcur�Ҳ�
        cur�ֽ� = mcur�ֽ�
    End If
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, j As Long
    Dim str֧Ʊ���㷽ʽ As String
    Dim lngCashRow As Long
    
    If Val(txtOwe.Text) <> 0 Then
        If Val(txtOwe.Text) > 0 Then
            MsgBox "����֧������,�밴����ʾ�Ĳ��", vbExclamation, gstrSysName
            mshPay.SetFocus: Exit Sub
        Else
            MsgBox "����֧��������,�밴����ʾ�Ĳ���˿", vbExclamation, gstrSysName
            mshPay.SetFocus: Exit Sub
        End If
    End If
    '���˺�:28947
    If mintInsure <> 0 Then
        If gclsInsure.CheckInsureValid(mintInsure) = False Then
            Exit Sub
        End If
    End If
    

            
    With mshPay
        mcur�ֽ� = 0
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COLS.C1���)) = 0 And i <> mlngPayRow Then
                If (.TextMatrix(i, COLS.C2����) <> "" Or .TextMatrix(i, COLS.C3��ע) <> "") Then
                    If MsgBox("ע��:��" & i & "��û��������,��������" & IIf(.TextMatrix(i, COLS.C2����) <> "", "�����", "��ע") & _
                        vbCrLf & "!����Ϣ���ᱣ��!ȷ��Ҫ������?", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        .SetFocus: .Row = i: .Col = COLS.C1���: Exit Sub
                    End If
                End If
            ElseIf Val(.RowData(i)) = 7 Then
                j = j + 1
            End If
            If .RowData(i) = 1 Then
                mcur�ֽ� = mcur�ֽ� + Val(mshPay.TextMatrix(i, 1))
                lngCashRow = i
            End If
        Next
        
        If j > 1 Then
            MsgBox "��֧��һ��ʹ�ö���һ��֧ͨ����", vbInformation
            Exit Sub
        End If
        '���˺�:35204,�ɿ������
        Select Case gTy_Module_Para.byt�ɿ����
        Case 1  '1-��������ɿ��Ž��������ۼ�
        Case 2  '2-�շ�ʱ����Ҫ����ɿ���
            If Val(mcur�ֽ�) > 0 And Val(txt�ɿ�.Text) = 0 Then
                MsgBox "ע��:" & vbCrLf & _
                "    �ò���δ����ɿ���,���ܽ����շ�!", vbInformation + vbDefaultButton1, gstrSysName
                If txt�ɿ�.Enabled And txt�ɿ�.Visible Then txt�ɿ�.SetFocus
                Exit Sub
            End If
        Case Else   ',0-�������нɿ�������ۼƿ���
        End Select
        '37642
        If Val(mcur�ֽ�) < 0 Then
            If MsgBox("ע��:" & vbCrLf & "   �ò��˵��ֽ�Ϊ������,���Ƿ����Ҫ�˲����ֽ�(" & Format(mcur�ֽ�, "####0.00;-####0.00;0.00;0.00") & ") ?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If lngCashRow > 0 Or lngCashRow < .Rows Then
                    .Row = lngCashRow: .Col = COLS.C1���
                End If
                .SetFocus
                Exit Sub
            End If
        End If
        mcurPay = 0: mstrBalance = ""
        '33722
        str֧Ʊ���㷽ʽ = ""
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, COLS.C1���)) <> 0 Then
                '֧Ʊ�Ľ��㷽ʽ,��Ҫ��������,��Ҫ���ڷ�̯ʱ,Ӧ������̯֧Ʊ,����������
                mcurPay = mcurPay + Val(.TextMatrix(i, COLS.C1���))
                If (.TextMatrix(i, COLS.C0��ʽ) Like "*֧Ʊ*" Or i = mlngPayRow) And mlngPayRow > 0 Then
                    str֧Ʊ���㷽ʽ = str֧Ʊ���㷽ʽ & "||" & .TextMatrix(i, COLS.C0��ʽ) & "|" & .TextMatrix(i, COLS.C1���) & _
                        "|" & IIf(.TextMatrix(i, COLS.C2����) = "", " ", .TextMatrix(i, COLS.C2����)) & _
                        "|" & IIf(.TextMatrix(i, COLS.C3��ע) = "", " ", .TextMatrix(i, COLS.C3��ע))
                Else
                    mstrBalance = mstrBalance & "||" & .TextMatrix(i, COLS.C0��ʽ) & "|" & .TextMatrix(i, COLS.C1���) & _
                        "|" & IIf(.TextMatrix(i, COLS.C2����) = "", " ", .TextMatrix(i, COLS.C2����)) & _
                        "|" & IIf(.TextMatrix(i, COLS.C3��ע) = "", " ", .TextMatrix(i, COLS.C3��ע))
                End If
                '�ո���������ַָ���
            End If
        Next
    
        '֧Ʊ����̯
        mstrBalance = mstrBalance & str֧Ʊ���㷽ʽ
        mstrBalance = Mid(mstrBalance, 3)
        mcurPay = Format(mcurPay, "0.00")
        mcur�ɿ� = Val(txt�ɿ�.Text)
        mcur�Ҳ� = Val(txt�Ҳ�.Text)
    End With
    
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mcurOneCard > 0 Then
        If txt�ɿ�.Enabled Then txt�ɿ�.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If txtTmp.Visible Then
            txtTmp.Visible = False
            mshPay.SetFocus
        Else
            Call cmdCancel_Click
        End If
    Case vbKeyF2
        If cmdOK.Enabled And cmdOK.Visible Then
            Call cmdOK.SetFocus
            Call cmdOK_Click
        End If
    Case vbKeyF12
        If Shift = vbCtrlMask Then
            'ǿ����LED����,(�ϼ�)
            If gblnLED Then
                mblnHotKey = True: txt�ɿ�.SetFocus
                If ActiveControl Is txt�ɿ� Then txt�ɿ�_GotFocus
            End If
        End If
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 _
        And Not ActiveControl Is mshPay _
        And Not ActiveControl Is txtTmp _
        And Not ActiveControl Is txt�ɿ� Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("'|", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strSQL As String, i As Long
    Dim arrPay As Variant, j As Long
    
    mblnOK = False
    mintDefault = 0
    mcurError = 0
    
    txtPay.Text = Format(mcurPay, "0.00")
    arrPay = Array()
    If mstrBalance <> "" Then
        arrPay = Split(mstrBalance, "||")
    End If
    
    On Error GoTo errH
    mrs���㷽ʽ.Filter = "����=1 or ����=2 or ����=7"
    
    With mshPay
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .Rows = mrs���㷽ʽ.RecordCount + 1
        For i = 1 To mrs���㷽ʽ.RecordCount
            .RowData(i) = NVL(mrs���㷽ʽ!����, 1)
            .TextMatrix(i, COLS.C0��ʽ) = mrs���㷽ʽ!����
            
            'ȱʡ���㷽ʽ(û�������ֽ�)
            If mrs���㷽ʽ!���� = gstr���㷽ʽ Then mintDefault = i
            If NVL(mrs���㷽ʽ!ȱʡ, 0) = 1 And mintDefault = 0 Then mintDefault = i
            If NVL(mrs���㷽ʽ!����, 1) = 1 And mintDefault = 0 Then mintDefault = i
            'ȱʡֵ(��һ�ε�)
            For j = 0 To UBound(arrPay)
                If Split(arrPay(j), "|")(0) = mrs���㷽ʽ!���� Then
                    .TextMatrix(i, COLS.C1���) = Format(Split(arrPay(j), "|")(1), "0.00")
                    .TextMatrix(i, COLS.C2����) = Trim(Split(arrPay(j), "|")(2))  'ȥ����Ϊ�Ŀո����
                    .TextMatrix(i, COLS.C3��ע) = Trim(Split(arrPay(j), "|")(3))
                    Exit For
                End If
            Next
            If Val(NVL(mrs���㷽ʽ!Ӧ����)) = 1 Then
                mlngPayRow = i: mstrӦ�����㷽ʽ = mrs���㷽ʽ!����
                .RowHeight(i) = 0
            End If
            mrs���㷽ʽ.MoveNext
        Next
        If mintDefault > 0 Then .CellFontBold = True
        '����Ӧ�����ȱʡ����֧Ʊ�µ��к�����:33722
        j = -1
        For i = 1 To .Rows - 1
            If .RowData(i) = 2 And InStr(1, .TextMatrix(i, COLS.C0��ʽ), "֧Ʊ") > 0 And i <> mlngPayRow Then
                '��ȡ����֧Ʊ��
                j = i
            End If
        Next
        
        If j <> -1 And mlngPayRow > 0 Then
             If mlngPayRow <> j And j < .Rows - 1 Then
                '��Ҫ��Ӧ�����з���֧Ʊ��
                .RowPosition(mlngPayRow) = j + 1
                mlngPayRow = j + 1
             End If
        End If
        If mlngPayRow > 0 Then
            .Row = mlngPayRow
            For i = 0 To .COLS - 1
                .Col = i: .CellFontBold = True
            Next
        End If
        .Row = 1: .Col = COLS.C1���
    End With
    Call ShowMoney(mstrBalance = "", False)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mshPay_DblClick()
    With mshPay
        If Not txtTmp.Visible And mshPay.Row > 0 And mshPay.Col > COLS.C0��ʽ Then
            If mshPay.Row <> mlngPayRow Then
                Call SetTxtTmp
                txtTmp.Text = mshPay.TextMatrix(mshPay.Row, mshPay.Col)
                txtTmp.SelStart = 0: txtTmp.SelLength = Len(txtTmp.Text)
            End If
        End If
    End With
End Sub
Private Sub mshPay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call LocateMshpay
    ElseIf KeyCode = vbKeyDelete Then
        If mshPay.Row > 0 And mshPay.Col > COLS.C0��ʽ Then
            mshPay.TextMatrix(mshPay.Row, mshPay.Col) = ""
            If mshPay.Col = COLS.C1��� Then Call ShowMoney(False, mshPay.Row <> mintDefault)
        End If
    End If
End Sub

Private Sub mshPay_KeyPress(KeyAscii As Integer)
    If Not txtTmp.Visible And mshPay.Row > 0 And mshPay.Col > COLS.C0��ʽ And KeyAscii <> 13 Then
        If mshPay.Col = COLS.C1��� Then
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        If mshPay.Col <> COLS.C3��ע Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If mshPay.Col = COLS.C2���� Then
            If zlCommFun.IsCharChinese(Chr(KeyAscii)) Then Call Beep: Exit Sub
        End If
        If mshPay.Row <> mlngPayRow Then
            Call SetTxtTmp
            txtTmp.Text = Chr(KeyAscii)
            txtTmp.SelStart = 1
        End If
    End If
End Sub

Private Sub LocateMshpay()
    Dim j As Long
    Dim a As Long
    Dim lngRow As Long
    
    With mshPay
        lngRow = .Row
        If mlngPayRow > 0 And mlngPayRow = .Rows - 1 Then
            If lngRow = .Rows - 2 Then lngRow = .Rows - 1
        End If
        
        '��ĩ�е����һ�л���,��ĩ�н��Ϊ�㻻��
        If lngRow < .Rows - 1 And (.Col = .COLS - 1 Or .TextMatrix(.Row, COLS.C1���) = "" And .Col <> COLS.C0��ʽ) Then
            a = .Row
            For j = .Row + 1 To .Rows - 1
                If .RowHeight(j) > 0 Then
                    .Row = j: Exit For
                End If
            Next
            .Col = COLS.C1���
            If .Row - (.Height \ .RowHeight(0) - 2) > 1 Then
                .TopRow = .Row - (.Height \ .RowHeight(1) - 2)
            End If
            If a = j Then
                
            End If
            Call mshPay_EnterCell
        ElseIf lngRow = .Rows - 1 And (.Col = .COLS - 1 Or .TextMatrix(.Row, COLS.C1���) = "" And .Col <> COLS.C0��ʽ) Then
        'ĩ�е����һ��,ĩ�н��Ϊ��,Tab
            Call zlCommFun.PressKey(vbKeyTab)
        Else
             If .RowData(.Row) = 1 And .Col = COLS.C1��� Then '�ֽ�������������
                .Col = .Col + 2
            Else
                .Col = .Col + 1
            End If
            Call mshPay_EnterCell
        End If
    End With
End Sub

Private Sub SetTxtTmp()
    With txtTmp
        .MaxLength = Val("" & Choose(mshPay.Col, 10, 30, 25))   'ժҪ�50λ,�����������ֻ��25������
        .Left = mshPay.Left + mshPay.CellLeft + 15
        .Top = mshPay.Top + mshPay.CellTop + (mshPay.CellHeight - txtTmp.Height) / 2 - 15
        .Width = mshPay.CellWidth - 60
        .ForeColor = mshPay.CellForeColor
        .BackColor = mshPay.CellBackColor
        .Alignment = IIf(mshPay.Col = COLS.C1���, 1, 0)
        .ZOrder: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub mshPay_LeaveCell()
    txtTmp.Visible = False
End Sub

Private Sub mshPay_Scroll()
    txtTmp.Visible = False
End Sub


Private Sub txtTmp_KeyPress(KeyAscii As Integer)
    Dim blnCent As Boolean, i As Long
    
    If KeyAscii <> 13 Then
        If mshPay.Col = COLS.C1��� Then
            If InStr(txtTmp.Text, ".") > 0 And Chr(KeyAscii) = "." Then KeyAscii = 0:  Exit Sub
            If InStr("0123456789.-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        ElseIf mshPay.Col = COLS.C2���� Then  '������������ַ�����
            If InStr("'|,", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Else    '��ע
            If InStr("'|", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Beep: Exit Sub
        End If
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        If mshPay.Col = COLS.C1��� Then
            If Not IsNumeric(txtTmp.Text) And txtTmp.Text <> "" Then
                MsgBox "��������ȷ����ֵ��", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txtTmp): Exit Sub
            End If
            If Val(txtTmp.Text) > 100 Then
                If Val(txtTmp.Text) > mcurPay * 2 Then
                    If MsgBox("��������ֳ����˸���ϼƵ�����(" & Format(mcurPay * 2, "0.00") & ")����ȷ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                        Call zlControl.TxtSelAll(txtTmp)
                        Exit Sub
                    End If
                End If
            End If
            
            txtTmp.Text = Format(Val(txtTmp.Text), "0.00")
            If Val(txtTmp.Text) <> 0 Then
                If Val(mshPay.RowData(mshPay.Row)) = 1 Then
                    '��������ֽ���������,����зֱҴ���
                    blnCent = True
                    If gBytMoney = 0 Then blnCent = False
                    If blnCent And mintInsure <> 0 And mlng����ID <> 0 Then
                        If gclsInsure.GetCapability(support����Ԥ��, mlng����ID, mintInsure) Then
                            If Not gclsInsure.GetCapability(support�ֱҴ���, mlng����ID, mintInsure) Then
                                blnCent = False
                            End If
                        End If
                    End If
                    If blnCent Then
                        txtTmp.Text = Format(CentMoney(Val(txtTmp.Text)), "0.00")
                    End If
                ElseIf Val(mshPay.RowData(mshPay.Row)) = 7 And mcurOneCard > 0 Then 'һ��ͨ
                    If Val(txtTmp.Text) > mcurOneCard Then
                        txtTmp.Text = Format(mcurOneCard, "0.00")
                    End If
                End If
            End If
        
            If Val(txtTmp.Text) = 0 Then
                mshPay.TextMatrix(mshPay.Row, mshPay.Col) = ""
            Else
                mshPay.TextMatrix(mshPay.Row, mshPay.Col) = Format(Val(txtTmp.Text), "0.00")
            End If
            
            '��������
            Call ShowMoney(False, mshPay.Row <> mintDefault)
        Else
        '���ַ�����
            If mshPay.Col = COLS.C2���� Then
                If InStr(txtTmp.Text, ",") > 0 Then Call Beep: Exit Sub
                If zlCommFun.IsCharChinese(txtTmp.Text) Then Call Beep: Exit Sub
            End If
            If InStr(txtTmp.Text, "'") > 0 Or InStr(txtTmp.Text, "|") > 0 Then Call Beep: Exit Sub
            
            mshPay.TextMatrix(mshPay.Row, mshPay.Col) = txtTmp.Text
        End If
        mshPay.SetFocus
        txtTmp.Visible = False
        '�����뻻��
        Call LocateMshpay
    End If
End Sub

Private Sub txtTmp_GotFocus()
    If mshPay.Col = COLS.C3��ע Then
        txtTmp.IMEMode = 1
        zlCommFun.OpenIme True
    Else
        txtTmp.IMEMode = 3
    End If
End Sub

Private Sub txtTmp_LostFocus()
    txtTmp.Visible = False
    If mshPay.Col = COLS.C3��ע Then
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub mshPay_EnterCell()
    If mshPay.Col = COLS.C3��ע Then
       zlCommFun.OpenIme True
       Exit Sub
    End If
    zlCommFun.OpenIme False
End Sub

Private Sub txtTmp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtTmp.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtTmp.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtTmp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Call SetWindowLong(txtTmp.hWnd, GWL_WNDPROC, glngTXTProc)
End Sub

Private Sub txtTmp_Validate(Cancel As Boolean)
    txtTmp.Visible = False
End Sub

Private Sub txtPay_GotFocus()
    Call zlControl.TxtSelAll(txtPay)
End Sub

Private Sub txt�ɿ�_Change()
    Dim cur�ֽ� As Currency, i As Long
    
    For i = 1 To mshPay.Rows - 1
        If mshPay.RowData(i) = 1 Then
            cur�ֽ� = Val(mshPay.TextMatrix(i, 1))
            Exit For
        End If
    Next
    If Val(txt�ɿ�.Text) = 0 Then txt�Ҳ�.Text = "0.00": Exit Sub
    txt�Ҳ�.Text = Format(Val(txt�ɿ�.Text) - cur�ֽ�, "0.00")
End Sub

Private Sub txt�ɿ�_GotFocus()
    Dim cur�ֽ� As Currency
    Dim i As Long
    '35204
    Call zlControl.TxtSelAll(txt�ɿ�)
    cur�ֽ� = 0
    With mshPay
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                cur�ֽ� = cur�ֽ� + Val(mshPay.TextMatrix(i, 1))
            End If
        Next
    End With
    'LED��ʾ:Ӧ�ɽ��
     If gblnLED And cur�ֽ� <> 0 Then
        '�Զ����ۻ��ֹ�����ʱ���ȼ�����
        If (Not gbln�ֹ����� And ActiveControl Is txt�ɿ�) Or (gbln�ֹ����� And mblnHotKey) Then
            mblnHotKey = False
            zl9LedVoice.Speak "#21 " & cur�ֽ�
        End If
    End If
    
End Sub
Private Function get�ֽ�() As Currency
    Dim i As Long, cur�ֽ� As Double
    cur�ֽ� = 0
    With mshPay
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                cur�ֽ� = cur�ֽ� + Val(mshPay.TextMatrix(i, 1))
            End If
        Next
    End With
    get�ֽ� = cur�ֽ�
End Function

Private Sub txt�ɿ�_KeyPress(KeyAscii As Integer)
    Dim cur�ֽ� As Currency
    If KeyAscii = 13 Then
        If Val(txt�ɿ�.Text) = 0 Then txt�ɿ�.Text = "0.00"
        If txt�ɿ�.Text <> "0.00" Then
            If Val(txt�Ҳ�.Text) >= 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                MsgBox "�ɿ����,�벹��Ӧ�ɽ�", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txt�ɿ�): txt�ɿ�.SetFocus
                Exit Sub
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab) '�����ۼӽɿ�
        End If
        
        'LED��ʾ
        cur�ֽ� = get�ֽ�
        If gblnLED Then
            mblnHotKey = False
            Call zl9LedVoice.DisplayBank( _
                "�ϼ�:" & mcurPay & "Ԫ,Ӧ��:" & cur�ֽ� & "Ԫ", _
                "����:" & txt�ɿ�.Text & "Ԫ" & IIf(Val(txt�Ҳ�.Text) = 0, "", ",����:" & txt�Ҳ�.Text & "Ԫ"))
            zl9LedVoice.Speak "#22 " & txt�ɿ�.Text
            zl9LedVoice.Speak "#23 " & txt�Ҳ�.Text
            zl9LedVoice.Speak "#3"
        End If
    Else
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        ElseIf KeyAscii = Asc(".") And InStr(txt�ɿ�.Text, ".") > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt�ɿ�_LostFocus()
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt�ɿ�_Validate(Cancel As Boolean)
    txt�ɿ�.Text = Format(Val(txt�ɿ�.Text), "0.00")
End Sub

Private Sub txt�Ҳ�_GotFocus()
    Call zlControl.TxtSelAll(txt�Ҳ�)
End Sub

Private Sub ShowMoney(blnFirst As Boolean, blnAutoCalc As Boolean, Optional bln֧Ʊ As Boolean = False)
'���ܣ����ú���ʾ����ĸ��ֽ��
'������blnFirst=��һ�ε���ʱ�Զ�����ȱʡ���㷽ʽ�����
'      blnAutoCalc=�Ƿ���ݲ���Զ���ƽȱʡ���㷽ʽ
    Dim curPay As Currency, curOwn As Currency
    Dim blnCent As Boolean, i As Long, blnSet As Boolean
    
    txt�ɿ�.Text = "0.00"
    '�ж��Ƿ�Ӧ�ý��зֱҴ���
    blnCent = True
    If gBytMoney = 0 Then blnCent = False
    If blnCent And mintInsure <> 0 And mlng����ID <> 0 Then
        If gclsInsure.GetCapability(support����Ԥ��, mlng����ID, mintInsure) Then
            If Not gclsInsure.GetCapability(support�ֱҴ���, mlng����ID, mintInsure) Then
                blnCent = False
            End If
        End If
    End If
    
    '��һ�ε���ʱ�Զ�����ȱʡ���㷽ʽ�����
    '-----------------------------------------------------------------------------------------------------
    If blnFirst Then
        If mcurOneCard > 0 Then
            For i = 1 To mshPay.Rows - 1
                If mshPay.RowData(i) = 7 Then
                    mshPay.TextMatrix(i, COLS.C1���) = Format(mcurOneCard, "0.00")
                    blnSet = True
                End If
            Next
        End If
        
        If mintDefault > 0 Then
            If mshPay.RowData(mintDefault) = 1 And blnCent Then '�ֽ�ʱҪ���зֱҴ���
                mshPay.TextMatrix(mintDefault, COLS.C1���) = Format(CentMoney(mcurPay - IIf(blnSet, mcurOneCard, 0)), "0.00")
            Else
                mshPay.TextMatrix(mintDefault, COLS.C1���) = Format(mcurPay - IIf(blnSet, mcurOneCard, 0), "0.00")
            End If
        End If
    End If
    
    Call Calc��֧Ʊ
    '��ʾ�ɿ���
    '-----------------------------------------------------------------------------------------------------
    curPay = 0
    For i = 1 To mshPay.Rows - 1
        curPay = curPay + Val(mshPay.TextMatrix(i, 1))
    Next
    curOwn = mcurPay - curPay
    txtOwe.Text = Format(mcurPay - curPay, "0.00") '�����ǲ��,��һ�����ֽ�,���Բ�����ֱ�
    
    '���ݲ���Զ���ƽ������
    '-----------------------------------------------------------------------------------------------------
    If blnAutoCalc And Val(txtOwe.Text) <> 0 Then
        'ʣ�ಿ�ݳ������õ�ȱʡ���㷽ʽ��
        If mlngPayRow >= 0 And bln֧Ʊ Then
             mshPay.TextMatrix(mlngPayRow, 1) = Format(Val(mshPay.TextMatrix(mlngPayRow, 1)) + curOwn, "0.00")
             If Val(mshPay.TextMatrix(mlngPayRow, 1)) <> 0 Then
                mshPay.RowHeight(mlngPayRow) = mshPay.RowHeight(0)
             Else
                mshPay.RowHeight(mlngPayRow) = 0
             End If
        Else
            If mintDefault > 0 Then
                If mshPay.RowData(mintDefault) = 1 And blnCent Then '�ֽ�ʱҪ���зֱҴ���
                    mshPay.TextMatrix(mintDefault, 1) = _
                        Format(CentMoney(Val(mshPay.TextMatrix(mintDefault, 1)) + curOwn), "0.00")
                Else
                    mshPay.TextMatrix(mintDefault, 1) = _
                        Format(Val(mshPay.TextMatrix(mintDefault, 1)) + curOwn, "0.00")
                End If
                If Val(mshPay.TextMatrix(mintDefault, 1)) = 0 Then
                    mshPay.TextMatrix(mintDefault, 1) = ""
                End If
                txtOwe.Text = "0.00"
            End If
        End If
    End If
    
    '���������(������-���ʽ��)
    '-----------------------------------------------------------------------------------------------------
    curPay = 0
    For i = 1 To mshPay.Rows - 1
        curPay = curPay + Val(mshPay.TextMatrix(i, 1))
    Next
    mcurError = Format(curPay - mcurPay, gstrDec)
    
    '�п��ܽɿ��������Ǵ���ֱҵ�����,�Ͳ���ʾ��(��������������ʱ��������0.29�����,0.79��0.5,0.29��0)
    If Val(txtOwe.Text) <> 0 And (Abs(Val(Val(txtOwe.Text))) < 0.1 Or gBytMoney = 5 And Abs(Val(Val(txtOwe.Text))) < 0.3) And mintDefault > 0 Then
        If mshPay.RowData(mintDefault) = 1 And blnCent Then
            If CentMoney(Val(mshPay.TextMatrix(mintDefault, 1)) + Val(txtOwe.Text)) = Val(mshPay.TextMatrix(mintDefault, 1)) Then
                txtOwe.Text = "0.00"
            End If
        End If
    End If
    
    '���ܽɿ������С�������������,�����������С��1��,�Ͳ�����
    If Val(txtOwe.Text) <> 0 And mcurError + curOwn = 0 And curOwn < 0.005 And curOwn >= -0.005 Then
        txtOwe.Text = "0.00"
    End If
End Sub
Private Sub Calc��֧Ʊ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Զ�������֧Ʊ
    '����:���˺�
    '����:2010-11-08 14:37:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dbl֧Ʊ�� As Double, dbl�ֽ� As Double
    '33722
    With mshPay
        '�����,����֧Ʊ����,������Ҫ������֧Ʊ
     '  If Not (.Col = COLS.C1��� And InStr(1, .TextMatrix(.Row, COLS.C0��ʽ), "֧Ʊ") > 0 _
            And .RowData(mshPay.Row) = 2 And .Row <> mlngPayRow) Then Exit Sub
        If mlngPayRow <= 0 Then Exit Sub
        dbl֧Ʊ�� = 0: dbl�ֽ� = 0
        For i = 1 To .Rows - 1
             If InStr(1, .TextMatrix(i, COLS.C0��ʽ), "֧Ʊ") > 0 _
                And .RowData(i) = 2 And i <> mlngPayRow Then
                dbl֧Ʊ�� = dbl֧Ʊ�� + Val(.TextMatrix(i, COLS.C1���))
             ElseIf i <> mintDefault And i <> mlngPayRow Then
                    dbl�ֽ� = dbl�ֽ� + Val(.TextMatrix(i, COLS.C1���))
            End If
        Next
        If RoundEx(mcurPay - dbl�ֽ� - dbl֧Ʊ��, 2) >= 0 Or dbl֧Ʊ�� = 0 Then
            .TextMatrix(mlngPayRow, COLS.C1���) = "": .RowHeight(mlngPayRow) = 0
        Else
            .TextMatrix(mlngPayRow, COLS.C1���) = Format(RoundEx(mcurPay - dbl�ֽ� - dbl֧Ʊ��, 2), "0.00"): .RowHeight(mlngPayRow) = .RowHeight(0)
        End If
    End With
End Sub
