VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmStuffQualityCard 
   Caption         =   "������������༭"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10890
   Icon            =   "frmStuffQualityCard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10890
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   4275
      Left            =   90
      TabIndex        =   2
      Top             =   960
      Width           =   10695
      _cx             =   18865
      _cy             =   7541
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   15132390
      BackColorAlternate=   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   19
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmStuffQualityCard.frx":000C
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
   Begin VB.TextBox txtCheck 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   10
      Top             =   5430
      Width           =   1000
   End
   Begin VB.TextBox txtVerify 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6360
      TabIndex        =   9
      Top             =   5430
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8400
      TabIndex        =   8
      Top             =   6240
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   7
      Top             =   6240
      Width           =   1100
   End
   Begin VB.TextBox txt��ע 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   6
      Top             =   5820
      Width           =   9975
   End
   Begin VB.ComboBox cboStock 
      Enabled         =   0   'False
      Height          =   300
      Left            =   585
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1920
   End
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      Height          =   315
      IMEMode         =   2  'OFF
      Left            =   9360
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   165
      Width           =   1425
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "�Ǽ���"
      Height          =   180
      Left            =   120
      TabIndex        =   17
      Top             =   5490
      Width           =   540
   End
   Begin VB.Label lblVerify 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   5640
      TabIndex        =   16
      Top             =   5490
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lbl��ע 
      AutoSize        =   -1  'True
      Caption         =   "��ע"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   360
   End
   Begin VB.Label txtCheckDate 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   300
      Left            =   2955
      TabIndex        =   14
      Top             =   5400
      Width           =   1800
   End
   Begin VB.Label lblCheckDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ǽ�����"
      Height          =   180
      Left            =   2100
      TabIndex        =   13
      Top             =   5475
      Width           =   720
   End
   Begin VB.Label txtVerifyDate 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   300
      Left            =   8730
      TabIndex        =   12
      Top             =   5400
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblVerifyDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Left            =   7875
      TabIndex        =   11
      Top             =   5475
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblStore 
      AutoSize        =   -1  'True
      Caption         =   "�ⷿ"
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   660
      Width           =   360
   End
   Begin VB.Label lblNo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO."
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   8880
      TabIndex        =   4
      Top             =   202
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "������������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11175
   End
End
Attribute VB_Name = "frmStuffQualityCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint�༭״̬ As Integer '1-���� 2-�޸� 3-���� 4-�鿴
Private mlng����id As Long      '����id,�޸ĺͲ鿴״̬����ֵ
Private mstrMatch As String     'ƥ�䷽ʽ
Private mintUnit As Long        '0-ɢװ��λ��1-��װ��λ
Private mlng�ⷿid As Long      '�ⷿid
Private mblnChange As Boolean   '�Ƿ���й��༭
Private mint����� As Integer  '��ʾ���ĳ���ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mFMT As g_FmtString     'С��λ���ĸ�ʽ��
Private ArrNum As Variant       '��¼��������
Private mblnUsableNum As Boolean '�Ƿ�������¿��ÿ��
Private mintcboIndex As Integer
Private mblnValidateEdit As Boolean

Public Sub ShowMe(ByVal int�༭״̬ As Integer, ByVal fraPar As Form, ByVal lng�ⷿID As Long, ByVal lng����id As Long, ByVal intUnit As Integer)
    mint�༭״̬ = int�༭״̬
    mlng�ⷿid = lng�ⷿID
    mlng����id = lng����id
    mintUnit = intUnit
    
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
        .FM_ɢװ���ۼ� = GetFmtString(2, g_�ۼ�)
    End With
    
    Me.Show vbModal, fraPar
End Sub

'�������������
Private Sub CheckDepend()
    Dim rsDepend As New ADODB.Recordset
    Dim strStock As String
    
    On Error GoTo ErrHandle
    
    '��ȡ�ɲ����Ŀⷿ���ʱ���
    strStock = "VKW"
    
    '��鵱ǰ��Ա���������Ƿ�Ϊ�����Ŀ⡱�����Ƽ��ҡ��������ϲ��š�
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = [3] Or a.վ�� is Null) And c.�������� = b.���� " _
            & "  AND Instr([2],b.����,1) > 0 " _
            & "  AND a.id = c.����id " _
            & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
            & IIf(InStr(1, gstrPrivs, ";���пⷿ;") > 0, "", " and a.id in (Select ����id from ������Ա where ��Աid =[1])")
    Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.Id, strStock, gstrNodeNo)
    
    If rsDepend.EOF Then
        MsgBox "û���������Ŀ����ʵĲ��Ż򲻾߱���ص�Ȩ��,��鿴���Ź������ϵͳ����Ա��Ȩ��", vbInformation, gstrSysName
        If rsDepend.State = 1 Then rsDepend.Close
        Exit Sub
    End If
    
    'װ��ⷿ����
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!����
            .ItemData(.NewIndex) = rsDepend!Id
            If rsDepend!Id = UserInfo.����ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        .Text = frmStuffQualityList.cboStock.Text
        If .ListIndex = -1 Then .ListIndex = 0
        mintcboIndex = .ListIndex
        rsDepend.Close
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cboStock_Click()
    mint����� = Get������(cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
        
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To VSFDetail.Rows - 1
                If Val(VSFDetail.TextMatrix(i, VSFDetail.ColIndex("����id"))) <> 0 Then
                    Exit For
                End If
            Next
            If i <> VSFDetail.Rows Then
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���Ӧ���ĵĵ�λ����Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '�������ĵ�λ�ı�
                    mintcboIndex = .ListIndex
                    VSFDetail.Rows = 1
                    VSFDetail.Rows = 2
                    VSFDetail.Row = 1
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        
        mint����� = Get������(cboStock.ItemData(cboStock.ListIndex))
        
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        If ValiData = False Then Exit Sub
        
        If SaveCard = True Then
            MsgBox "����ɹ���", vbInformation, gstrSysName
            VSFDetail.Rows = 1
            VSFDetail.Rows = 2
            VSFDetail.SetFocus
            VSFDetail.Row = 1
            VSFDetail.Col = 0
            mblnChange = False
            If mint�༭״̬ = 2 Then Unload Me
            Exit Sub
        End If
    End If
    If mint�༭״̬ = 3 Then
        If mlng����id = 0 Then
            Exit Sub
        End If
        
        If SaveCheck(mlng����id) = True Then
            MsgBox "����ɹ���", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
End Sub

Private Function SaveCheck(ByVal lng����id As Long) As Boolean
    '������
    Dim strVerifyDate As String
    Dim lngRow As Long
    Dim arrSQL As Variant
    Dim dblTemp As Double
    Dim dbl��װϵ�� As Double
    Dim blnTran As Boolean
    
    On Error GoTo ErrHandle
    
    With VSFDetail
        arrSQL = Array()
        strVerifyDate = txtVerifyDate.Caption
        
        gstrSQL = "Zl_������������_Verify("
        '����id
        gstrSQL = gstrSQL & lng����id & ","
        '������
        gstrSQL = gstrSQL & "'" & txtVerify.Text & "',"
        '��������
        gstrSQL = gstrSQL & "to_date('" & strVerifyDate & "','yyyy-mm-dd HH24:MI:SS'))"
        
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = gstrSQL
        
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("����id"))) <> 0 Then
                dbl��װϵ�� = Val(.TextMatrix(lngRow, .ColIndex("����ϵ��")))
                
                gstrSQL = "Zl_������������_Verify("
                '���
                gstrSQL = gstrSQL & lngRow & ","
                'NO
                gstrSQL = gstrSQL & "'" & TxtNo.Text & "',"
                '�ⷿid
                gstrSQL = gstrSQL & Val(cboStock.ItemData(cboStock.ListIndex)) & ","
                '����id
                gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("����id"))) & ","
                '����
                gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("����"))) & ","
                '��������
                dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("��������"))) * dbl��װϵ��, g_С��λ��.obj_���С��.����С��)
                gstrSQL = gstrSQL & dblTemp & ","
                '�ɱ���
                dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) / dbl��װϵ��, g_С��λ��.obj_���С��.�ɱ���С��)
                gstrSQL = gstrSQL & dblTemp & ","
                '�ɱ����
                dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) * Val(.TextMatrix(lngRow, .ColIndex("��������"))), g_С��λ��.obj_���С��.���С��)
                gstrSQL = gstrSQL & dblTemp & ","
                '�ۼ۽��
                dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) * Val(.TextMatrix(lngRow, .ColIndex("��������"))), g_С��λ��.obj_���С��.���С��)
                gstrSQL = gstrSQL & dblTemp & ","
                '����
                dblTemp = Round((Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) - Val(.TextMatrix(lngRow, .ColIndex("�ɱ���")))) * Val(.TextMatrix(lngRow, .ColIndex("��������"))), g_С��λ��.obj_���С��.���С��)
                gstrSQL = gstrSQL & dblTemp & ","
                '������ID
                gstrSQL = gstrSQL & "19,"
                '������
                gstrSQL = gstrSQL & "'" & txtVerify.Text & "',"
                '��������
                gstrSQL = gstrSQL & "to_date('" & strVerifyDate & "','yyyy-mm-dd HH24:MI:SS'),"
                'ҩ��(ҩ��)ҵ��
                gstrSQL = gstrSQL & "1)"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            End If
        Next
    End With
    
    blnTran = True
    gcnOracle.BeginTrans
    For lngRow = 0 To UBound(arrSQL)
        Call zldatabase.ExecuteProcedure(CStr(arrSQL(lngRow)), "SaveCheck")
    Next
    gcnOracle.CommitTrans
    SaveCheck = True
    
    Exit Function
ErrHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveCard() As Boolean
    '�������޸ı���
    Dim rsTemp As ADODB.Recordset
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strNo As String
    Dim lng����id As Long
    Dim strCheckDate As String
    Dim dbl��װϵ�� As Double
    Dim blnTran As Boolean
    Dim arrSQL As Variant
    Dim dblTemp As Double
    
    On Error GoTo ErrHandle
    
    arrSQL = Array()
    
    With VSFDetail
        
        If TxtNo.Text = "" Then
            strNo = zldatabase.GetNextNo(74, cboStock.ItemData(cboStock.ListIndex)) '����ҩƷ��������ȡNO
        Else
            strNo = TxtNo.Text
        End If
        
        If .Rows > 1 And .TextMatrix(1, .ColIndex("����id")) <> "" Then
            If mint�༭״̬ = 2 Then '�޸�
                gstrSQL = "Zl_������������_Delete(" & mlng����id & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
                
                gstrSQL = "zl_������������_Delete('" & strNo & "')"
            
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = gstrSQL
            Else
                gstrSQL = "Select ������������_ID.Nextval as id From Dual"
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "������������")
                mlng����id = rsTemp!Id
            End If
            
            strCheckDate = Format(txtCheckDate.Caption, "yyyy-mm-dd hh:mm:ss")
            
            '�������
            gstrSQL = "Zl_������������_Insert ("
            'Id
            gstrSQL = gstrSQL & mlng����id & ","
            'No
            gstrSQL = gstrSQL & "'" & strNo & "',"
            '�Ǽ���
            gstrSQL = gstrSQL & "'" & txtCheck.Text & "',"
            '�Ǽ�����
            gstrSQL = gstrSQL & "to_date('" & strCheckDate & "','yyyy-mm-dd HH24:MI:SS'),"
            '������
            gstrSQL = gstrSQL & "Null,"
            '��������
            gstrSQL = gstrSQL & "Null,"
            '��ע
            gstrSQL = gstrSQL & IIf(txt��ע.Text = "", "NULL", "'" & txt��ע.Text & "'") & ")"
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = gstrSQL
            
            '�α����
            For lngRow = 1 To .Rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("����id"))) <> 0 Then
                    dbl��װϵ�� = Val(.TextMatrix(lngRow, .ColIndex("����ϵ��")))
                    
                    gstrSQL = "Zl_����������¼_Insert ("
                    '����id
                    gstrSQL = gstrSQL & mlng����id & ","
                    '�ⷿid
                    gstrSQL = gstrSQL & Val(cboStock.ItemData(cboStock.ListIndex)) & ","
                    '����id
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("����id"))) & ","
                    '����
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("����"))) & ","
                    '����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("����")) & "',"
                    '����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("����")) & "',"
                    '�ɱ���
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) / dbl��װϵ��, g_С��λ��.obj_���С��.�ɱ���С��)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '���ۼ�
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) / dbl��װϵ��, g_С��λ��.obj_���С��.���ۼ�С��)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '��������
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("��������"))) * dbl��װϵ��, g_С��λ��.obj_���С��.����С��)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '����ԭ��
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("����ԭ��")) & "',"
                    '����취
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("����취")) & "',"
                    '��ҩ��λid
                    gstrSQL = gstrSQL & IIf(Trim(.TextMatrix(lngRow, .ColIndex("��Ӧ��"))) = "", "Null", Val(.TextMatrix(lngRow, .ColIndex("��ҩ��λid")))) & ")"
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                    
                    gstrSQL = "Zl_������������_Insert("
                    '������ID
                    gstrSQL = gstrSQL & "19,"
                    'NO
                    gstrSQL = gstrSQL & "'" & strNo & "',"
                    '���
                    gstrSQL = gstrSQL & lngRow & ","
                    '�ⷿid
                    gstrSQL = gstrSQL & Val(cboStock.ItemData(cboStock.ListIndex)) & ","
                    '����id
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("����id"))) & ","
                    '����
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("����"))) & ","
                    '��������
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("��������"))) * dbl��װϵ��, g_С��λ��.obj_���С��.����С��)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '�ɱ���
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) / dbl��װϵ��, g_С��λ��.obj_���С��.�ɱ���С��)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '�ɱ����
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) * Val(.TextMatrix(lngRow, .ColIndex("��������"))), g_С��λ��.obj_���С��.���С��)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '���ۼ�
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) / dbl��װϵ��, g_С��λ��.obj_���С��.���ۼ�С��)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '�ۼ۽��
                    dblTemp = Round(Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) * Val(.TextMatrix(lngRow, .ColIndex("��������"))), g_С��λ��.obj_���С��.���С��)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '����
                    dblTemp = Round((Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) - Val(.TextMatrix(lngRow, .ColIndex("�ɱ���")))) * Val(.TextMatrix(lngRow, .ColIndex("��������"))), g_С��λ��.obj_���С��.���С��)
                    gstrSQL = gstrSQL & dblTemp & ","
                    '������
                    gstrSQL = gstrSQL & "'" & txtCheck.Text & "',"
                    '��������
                    gstrSQL = gstrSQL & "to_date('" & strCheckDate & "','yyyy-mm-dd HH24:MI:SS'),"
                    '����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("����")) & "',"
                    '����
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("����")) & "',"
                    'Ч�ڣ����Ч��
                    gstrSQL = gstrSQL & "Null,Null,"
                    'ժҪ
                    gstrSQL = gstrSQL & IIf(txt��ע.Text = "", "NULL", "'" & txt��ע.Text & "'") & ","
                    '����ۣ������λ,��ֵ˰��
                    gstrSQL = gstrSQL & "Null,Null,Null,"
                    'ҩ��(ҩ��)ҵ��
                    gstrSQL = gstrSQL & "1,"
                    '����id
                    gstrSQL = gstrSQL & mlng����id & ")"
                    
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
            Next
            
            blnTran = True
            gcnOracle.BeginTrans
            For lngRow = 0 To UBound(arrSQL)
                Call zldatabase.ExecuteProcedure(CStr(arrSQL(lngRow)), "SaveCard")
            Next
            gcnOracle.CommitTrans
            SaveCard = True
        Else
            SaveCard = False
            Exit Function
        End If
    End With

    Exit Function
ErrHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ValiData() As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    
    '����ʱ���ݼ��
    With VSFDetail
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("����id")) <> "" And Val(.TextMatrix(lngRow, .ColIndex("��������"))) = 0 Then
                MsgBox "��" & lngRow & "�����ݻ�����������Ϊ0��գ�", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("��������")
                .SetFocus
                Exit Function
            End If
        Next
    End With
    ValiData = True
End Function

Private Sub Form_Load()
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Call CheckDepend
    mstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "") 'ƥ�䷽ʽ
    mblnUsableNum = (Val(zldatabase.GetPara("������¿��ÿ��", 100)) = 1) '�����¿��ÿ��
    
    If InStr(1, ";" & gstrPrivs & ";", ";�鿴�ɱ���;") = 0 Then 'Ȩ�ޡ��鿴�ɱ��ۡ�
        VSFDetail.ColHidden(VSFDetail.ColIndex("�ɱ���")) = True
        VSFDetail.ColHidden(VSFDetail.ColIndex("�ɱ����")) = True
    End If
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        cboStock.Enabled = True
        txtCheck.Text = UserInfo.�û���
        txt��ע.Enabled = True
        Me.Caption = "�������������޸�"
        If mint�༭״̬ = 1 Then
            Me.Caption = "����������������"
            txtCheckDate.Caption = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
    End If
    
    If mint�༭״̬ = 3 Then
        txtVerify.Text = UserInfo.�û���
        txtVerifyDate.Caption = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        Me.Caption = "��������������"
    End If
    
    If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
        Call initCard
        If mint�༭״̬ = 3 Then
            CmdSave.Caption = "����(&O)"
        End If
        If mint�༭״̬ = 4 Then
            CmdSave.Visible = False
            Me.Caption = "���������������"
        End If
    End If
    
    If Val(zldatabase.GetPara("ʹ�ø��Ի����")) = 1 Then RestoreWinState Me, App.ProductName, "������������"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With LblTitle
        .Left = 0
        .Top = 120
        .Width = Me.ScaleWidth
    End With
    
    With TxtNo
        .Move Me.ScaleWidth - .Width - 150
    End With
    lblNo.Move TxtNo.Left - lblNo.Width - 100
    
    With lblStore
        .Move 150, 720
    End With
    
    With cboStock
        .Move lblStore.Left + lblStore.Width + 50, lblStore.Top - 60
    End With
    
    With VSFDetail
        .Move lblStore.Left, lblStore.Top + lblStore.Height + 100, Me.ScaleWidth - lblStore.Left - 150, Me.ScaleHeight - .Top - txtVerify.Height - 1200
    End With
    
    lblCheck.Move VSFDetail.Left, VSFDetail.Top + VSFDetail.Height + 200
    txtCheck.Move lblCheck.Left + lblCheck.Width + 100, lblCheck.Top - 60
    lblCheckDate.Move txtCheck.Left + txtCheck.Width + 100, lblCheck.Top
    txtCheckDate.Move lblCheckDate.Left + lblCheckDate.Width + 100, txtCheck.Top
    
    If mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
        lblVerify.Visible = True
        txtVerify.Visible = True
        lblVerifyDate.Visible = True
        txtVerifyDate.Visible = True
        
        lblVerify.Move txtCheckDate.Left + txtCheckDate.Width + 500, lblCheck.Top
        txtVerify.Move lblVerify.Left + lblVerify.Width + 100, txtCheck.Top
        
        lblVerifyDate.Move txtVerify.Left + txtVerify.Width + 100, lblVerify.Top
        txtVerifyDate.Move lblVerifyDate.Left + lblVerifyDate.Width + 100, txtVerify.Top
    End If
    
    lbl��ע.Move lblCheck.Left, lblCheck.Top + lblCheck.Height + 200
    txt��ע.Move txtCheck.Left, lbl��ע.Top - 20, VSFDetail.Width - lbl��ע.Left - 530
    
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 200, lbl��ע.Top + lbl��ע.Height + 280
    CmdSave.Move cmdCancel.Left - CmdSave.Width - 200, lbl��ע.Top + lbl��ע.Height + 280
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, "������������"
        Exit Sub
    End If
    
    If MsgBox("���ݿ����Ѹı䣬��δ���棬�Ƿ��˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        mblnChange = False
        SaveWinState Me, App.ProductName, "������������"
    End If
    
End Sub

Private Sub txt��ע_GotFocus()
    zlControl.TxtSelAll txt��ע
End Sub

Private Function CheckRedo(ByVal rsTemp As ADODB.Recordset) As Boolean
    '���ܣ�����ظ���¼
    Dim i As Integer
    Dim str���� As String
    Dim str����ID As String
    
    With VSFDetail
        CheckRedo = False
        str����ID = rsTemp!����ID
        str���� = IIf(IsNull(rsTemp!����), "", rsTemp!����)
        
        For i = 1 To .Rows - 1
            If str����ID = .TextMatrix(i, .ColIndex("����ID")) And str���� = .TextMatrix(i, .ColIndex("����")) And .TextMatrix(i, .ColIndex("����ID")) <> "" Then
                If str����ID <> .TextMatrix(.Row, .ColIndex("����ID")) Then
                    MsgBox "[" & rsTemp!���� & "]" & rsTemp!���� & "���ò����б����Ѵ��ڣ�", vbInformation, gstrSysName
                    If .TextMatrix(.Row, .ColIndex("����id")) = "" Then .TextMatrix(.Row, .ColIndex("������Ϣ")) = ""
                    CheckRedo = True
                    Exit For
                End If
            End If
        Next
        
    End With
End Function

Private Sub vsfDetail_AfterSort(ByVal Col As Long, Order As Integer)
    Dim lngRow As Long
    
    With VSFDetail
        If .Rows > 1 Then
            For lngRow = 1 To .Rows - 1
                If Val(.TextMatrix(lngRow, .ColIndex("����id"))) = 0 Then
                    .RemoveItem lngRow
                    .Rows = .Rows + 1
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfDetail_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As Recordset
    Dim dblTop As Double
    Dim dblLeft As Double
    Dim vRect As RECT
    
    On Error GoTo ErrHandle
    
    With VSFDetail
        Select Case Col
            Case .ColIndex("������Ϣ")
                Set RecReturn = Frm����ѡ����.ShowMe(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), , , , , , , , , , , , , gstrPrivs)
                If RecReturn.RecordCount = 0 Then Exit Sub
                If CheckRedo(RecReturn) = True Then Exit Sub
                
                Call SetColValue(.Row, RecReturn!����ID, RecReturn!����, RecReturn!����, _
                                IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                IIf(IsNull(RecReturn!����), "", RecReturn!����), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                IIf(IsNull(RecReturn!��ҩ��λID), "", RecReturn!��ҩ��λID), IIf(IsNull(RecReturn!ɢװ��λ), "", RecReturn!ɢװ��λ), _
                                IIf(IsNull(RecReturn!��װ��λ), "", RecReturn!��װ��λ), IIf(IsNull(RecReturn!����ϵ��), "", RecReturn!����ϵ��), _
                                IIf(IsNull(RecReturn!�ۼ�), "", RecReturn!�ۼ�), IIf(IsNull(RecReturn!ʱ��), "", RecReturn!ʱ��), _
                                IIf(IsNull(RecReturn!���÷���), "", RecReturn!���÷���), IIf(IsNull(RecReturn!�ⷿ����), "", RecReturn!�ⷿ����), _
                                IIf(IsNull(RecReturn!��������), 0, RecReturn!��������))
                .Col = .ColIndex("��Ӧ��")
            Case .ColIndex("��Ӧ��")
                vRect = zlControl.GetControlRect(.hwnd) '��ȡλ��
                dblTop = vRect.Top + .CellTop + .CellHeight - 950
                dblLeft = vRect.Left + .CellLeft
                gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ��Ӧ�� " & _
                          "Where (վ�� = [1] Or վ�� is Null) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
                          "  And (substr(����,5,1)=1 Or Nvl(ĩ��,0)=0) " & _
                          "Start with �ϼ�ID is null connect by prior ID =�ϼ�ID order by level,ID"
                Set RecReturn = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "��ҩ��λ", False, "", "", False, False, _
                                    True, dblLeft, dblTop, 1000, False, False, True, gstrNodeNo)
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
                    .TextMatrix(Row, .ColIndex("��Ӧ��")) = RecReturn!����
                    .TextMatrix(Row, .ColIndex("��ҩ��λid")) = RecReturn!Id
                End If
                .Col = .ColIndex("��������")
            Case .ColIndex("����ԭ��")
                vRect = zlControl.GetControlRect(.hwnd)
                dblTop = vRect.Top + .CellTop + .CellHeight - 950
                dblLeft = vRect.Left + .CellLeft
                gstrSQL = "Select ���� id,'' �ϼ�ID,0 ĩ��,����,����,���� From ������ԭ�� " & _
                          " Where ���� is not null " & _
                          " order by ID"
                Set RecReturn = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "����ԭ��", False, "", "", False, False, _
                                    True, dblLeft, dblTop, 1000, False, False, True, gstrNodeNo)
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
                    .TextMatrix(Row, .ColIndex("����ԭ��")) = RecReturn!����
                End If
                .Col = .ColIndex("����취")
            Case .ColIndex("����취")
                vRect = zlControl.GetControlRect(.hwnd)
                dblTop = vRect.Top + .CellTop + .CellHeight - 950
                dblLeft = vRect.Left + .CellLeft
                gstrSQL = "Select ���� id,'' �ϼ�ID,0 ĩ��,����,����,���� From �������취 " & _
                          " Where ���� is not null " & _
                          " order by ID"
                Set RecReturn = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "����취", False, "", "", False, False, _
                                    True, dblLeft, dblTop, 1000, False, False, True, gstrNodeNo)
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
                    .TextMatrix(Row, .ColIndex("����취")) = RecReturn!����
                End If
        End Select
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetail_ChangeEdit()
    mblnChange = True
End Sub

Private Sub vsfDetail_DblClick()
    With VSFDetail
        If .Col = .ColIndex("������Ϣ") Or .Col = .ColIndex("��Ӧ��") Or .Col = .ColIndex("��������") Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.TextMatrix(.Row, .Col)) * 2
        End If
    End With
End Sub

Private Sub VSFDetail_EnterCell()
    With VSFDetail
        .Editable = flexEDNone
    End With
End Sub

Private Sub vsfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim dblLeft As Single
    Dim dblTop As Single
    Dim vRect As RECT
    Dim RecReturn As Recordset
    Dim strKey As String
    
    With VSFDetail
        If KeyCode = vbKeyDelete Then
            If MsgBox("��ɾ�����У��Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Shift = 0
            Else
                .RemoveItem VSFDetail.Row
            End If
        ElseIf KeyCode = vbKeyReturn Then
            If Trim(.TextMatrix(.Row, .ColIndex("������Ϣ"))) = "" Then KeyCode = 0: Exit Sub
            
            If .Col = .ColIndex("������Ϣ") Then
                .Col = .ColIndex("��Ӧ��")
            ElseIf .Col = .ColIndex("��Ӧ��") Then
                .Col = .ColIndex("��������")
            ElseIf .Col = .ColIndex("��������") Then
                .TextMatrix(.Row, .Col) = Format(.TextMatrix(.Row, .Col), mFMT.FM_����)
                .TextMatrix(.Row, .ColIndex("�ɱ����")) = Format(Val(.TextMatrix(.Row, .ColIndex("�ɱ���"))) * Val(.TextMatrix(.Row, .Col)), mFMT.FM_���)
                .TextMatrix(.Row, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(.Row, .ColIndex("���ۼ�"))) * Val(.TextMatrix(.Row, .Col)), mFMT.FM_���)
                .Col = .ColIndex("����ԭ��")
            ElseIf .Col = .ColIndex("����ԭ��") Then
                .Col = .ColIndex("����취")
            ElseIf .Col = .ColIndex("����취") Then
                If .Row = .Rows - 1 Then
                    .Rows = .Rows + 1
                    .Row = .Rows - 1
                    .Col = .ColIndex("������Ϣ")
                Else
                    .Row = .Row + 1
                    .Col = .ColIndex("������Ϣ")
                End If
            Else
                If InStr(1, ";" & gstrPrivs & ";", ";�鿴�ɱ���;") = 0 Then
                    If .Col = .ColIndex("��λ") Then
                        .Col = .ColIndex("���ۼ�")
                    ElseIf .Col = .ColIndex("���ۼ�") Then
                        .Col = .ColIndex("�ۼ۽��")
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    .Col = .Col + 1
                End If
            End If
        Else
            If mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
                .Editable = flexEDNone
            Else
                If .Col = .ColIndex("������Ϣ") Or .Col = .ColIndex("��Ӧ��") Or .Col = .ColIndex("��������") Then
                    .Editable = flexEDKbdMouse
                    .ColComboList(.Col) = ""
                Else
                    .Editable = flexEDNone
                End If
            End If
        End If
    End With
End Sub

Private Sub SetColValue(ByVal intRow As Integer, ByVal lng����ID As Long, ByVal str���� As String, ByVal str���� As String, _
                        ByVal str��� As String, ByVal str���� As String, ByVal str���� As String, ByVal str���� As String, _
                        ByVal str��ҩ��λid As String, ByVal strɢװ��λ As String, ByVal str��װ��λ As String, _
                        ByVal str����ϵ�� As String, ByVal str�ۼ� As String, ByVal int�Ƿ��� As Integer, _
                        ByVal int���÷��� As String, ByVal int�ⷿ���� As String, ByVal dbl�������� As Double)
    '���ֵ
    Dim rsTemp As Recordset
    Dim str��װϵ�� As String
    Dim bln���� As Boolean
    
    On Error GoTo ErrHandle
    
    With VSFDetail
        .TextMatrix(intRow, .ColIndex("����id")) = lng����ID
        .EditText = "[" & str���� & "]" & str����
        .TextMatrix(intRow, .ColIndex("������Ϣ")) = "[" & str���� & "]" & str����
        .TextMatrix(intRow, .ColIndex("���")) = str���
        .TextMatrix(intRow, .ColIndex("����")) = str����
        .TextMatrix(intRow, .ColIndex("��λ")) = IIf(mintUnit = 0, strɢװ��λ, str��װ��λ)
        .TextMatrix(intRow, .ColIndex("��ҩ��λid")) = str��ҩ��λid
        .TextMatrix(intRow, .ColIndex("����ϵ��")) = str����ϵ��
        .TextMatrix(intRow, .ColIndex("����")) = str����
        .TextMatrix(intRow, .ColIndex("����")) = str����
        .TextMatrix(intRow, .ColIndex("�Ƿ���")) = int�Ƿ���
        .TextMatrix(intRow, .ColIndex("��������")) = ""
        .TextMatrix(intRow, .ColIndex("��Ӧ��")) = ""
        .TextMatrix(intRow, .ColIndex("�ɱ����")) = ""
        .TextMatrix(intRow, .ColIndex("�ۼ۽��")) = ""
        .TextMatrix(intRow, .ColIndex("����ԭ��")) = ""
        .TextMatrix(intRow, .ColIndex("����취")) = ""
        
        str��װϵ�� = IIf(mintUnit = 0, 1, str����ϵ��)
        
        '��������
        .TextMatrix(intRow, .ColIndex("��������")) = Format(Val(dbl��������) / Val(str��װϵ��), mFMT.FM_����)
        
        '��Ӧ��
        gstrSQL = "select ���� from ��Ӧ�� where substr(����,5,1)=1 and id=[1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "��Ӧ�̲�ѯ", Val(str��ҩ��λid))
        If Not rsTemp.EOF Then
            .TextMatrix(intRow, .ColIndex("��Ӧ��")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
        End If
        
        '�ɱ���
        gstrSQL = "Select Decode(Nvl(a.ƽ���ɱ���, 0), 0, b.�ɱ���, a.ƽ���ɱ���) �ɱ���" & vbNewLine & _
                "From ҩƷ��� A, �������� B" & vbNewLine & _
                "Where a.ҩƷid = b.����id And a.�ⷿid = [1] And a.ҩƷid = [2] And Nvl(a.����, 0) = [3]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ɱ��۲�ѯ", Val(cboStock.ItemData(cboStock.ListIndex)), Val(lng����ID), Val(str����))
        If Not rsTemp.EOF Then
            .TextMatrix(intRow, .ColIndex("�ɱ���")) = Format(Val(rsTemp!�ɱ���) * Val(str��װϵ��), mFMT.FM_�ɱ���)
        End If
        
        '�ۼ�
        .TextMatrix(intRow, .ColIndex("���ۼ�")) = Format(Val(str�ۼ�) * Val(str��װϵ��), mFMT.FM_���ۼ�)
        If int�Ƿ��� = 1 Then
            If int���÷��� = 0 Then
                If int�ⷿ���� = 1 Then
                    gstrSQL = "Select Distinct 0 " & _
                            "From ��������˵�� " & _
                            "Where ((�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')) And ����id = [1]"
                    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "���Ų�ѯ", Val(cboStock.ItemData(cboStock.ListIndex)))
                    If rsTemp.RecordCount = 0 Then
                        bln���� = True
                    End If
                End If
            Else
                bln���� = True
            End If
        
            gstrSQL = "" & _
                "   Select nvl(���ۼ�,0) as �����ۼ�,nvl(ʵ�ʽ��,0)/ʵ������ as ƽ�����ۼ�" & _
                "   From ҩƷ��� " & _
                "   Where �ⷿid=[1]" & _
                "       and ҩƷid=[2]" & _
                "       and ����=1 and ʵ������>0 and " & _
                "       nvl(����,0)=[3]"
            
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "�ۼ۲�ѯ", Val(cboStock.ItemData(cboStock.ListIndex)), Val(lng����ID), Val(str����))
            If Not rsTemp.EOF Then
                If bln���� = True Then
                    .TextMatrix(intRow, .ColIndex("���ۼ�")) = Format(Val(rsTemp!�����ۼ�) * Val(str��װϵ��), mFMT.FM_���ۼ�)
                Else
                    .TextMatrix(intRow, .ColIndex("���ۼ�")) = Format(Val(rsTemp!ƽ�����ۼ�) * Val(str��װϵ��), mFMT.FM_���ۼ�)
                End If
            End If
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfDetail_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With VSFDetail
        If .Col = .ColIndex("������Ϣ") Or .Col = .ColIndex("��Ӧ��") Then
            .ColComboList(.Col) = "|..."
        ElseIf .Col = .ColIndex("����ԭ��") Or .Col = .ColIndex("����취") Then
            .ColComboList(.Col) = "..."
        End If
    End With
End Sub

Private Sub VSFDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim RecReturn As ADODB.Recordset
    Dim strKey As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intOldRow As Integer
    Dim i As Integer
    Dim intRow As Integer
    Dim intPosition As Integer

    On Error GoTo ErrHandle
    
    With VSFDetail
        intOldRow = .Row
        strKey = UCase(Trim(.EditText))
        
        Select Case Col
            Case .ColIndex("������Ϣ")
                If KeyAscii <> vbKeyReturn Then Exit Sub
                If strKey = "" Then Exit Sub
                
                dblLeft = Me.Left + .Left + .CellLeft + 130
                dblTop = Me.Top + .Top + .CellTop + .CellHeight + 500
                If dblTop + 4300 > Screen.Height Then
                    dblTop = dblTop - .CellHeight - 3680
                End If
                
                If Mid(strKey, 1, 1) = "[" Then
                    If InStr(2, strKey, "]") <> 0 Then
                        strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
                    Else
                        strKey = Mid(strKey, 2)
                    End If
                End If
                .TextMatrix(.Row, .Col) = strKey
                
                Set RecReturn = FrmMulitSel.ShowSelect(Me, 2, cboStock.ItemData(cboStock.ListIndex), , cboStock.ItemData(cboStock.ListIndex), strKey, dblLeft, dblTop, .CellWidth, .CellHeight, , , , , , , , , , , , gstrPrivs)
                
                If RecReturn.RecordCount = 0 Then .TextMatrix(.Row, .ColIndex("������Ϣ")) = "": Exit Sub
                If CheckRedo(RecReturn) = True Then Exit Sub
                
                Call SetColValue(.Row, RecReturn!����ID, RecReturn!����, RecReturn!����, _
                            IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                            IIf(IsNull(RecReturn!����), "", RecReturn!����), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                            IIf(IsNull(RecReturn!��ҩ��λID), "", RecReturn!��ҩ��λID), IIf(IsNull(RecReturn!ɢװ��λ), "", RecReturn!ɢװ��λ), _
                            IIf(IsNull(RecReturn!��װ��λ), "", RecReturn!��װ��λ), IIf(IsNull(RecReturn!����ϵ��), "", RecReturn!����ϵ��), _
                            IIf(IsNull(RecReturn!�ۼ�), "", RecReturn!�ۼ�), IIf(IsNull(RecReturn!ʱ��), "", RecReturn!ʱ��), _
                            IIf(IsNull(RecReturn!���÷���), "", RecReturn!���÷���), IIf(IsNull(RecReturn!�ⷿ����), "", RecReturn!�ⷿ����), _
                            IIf(IsNull(RecReturn!��������), 0, RecReturn!��������))
                .Col = .ColIndex("��Ӧ��")
            Case .ColIndex("��Ӧ��")
                If KeyAscii <> vbKeyReturn Or mblnValidateEdit = False Then Exit Sub
                If strKey = "" Then .Col = .ColIndex("��������"): Exit Sub
                .TextMatrix(.Row, .ColIndex("��Ӧ��")) = strKey
                
                vRect = zlControl.GetControlRect(.hwnd) '��ȡλ��
                dblTop = vRect.Top + .CellTop + .CellHeight - 950
                dblLeft = vRect.Left + .CellLeft
                gstrSQL = "Select id,����,����,���� From ��Ӧ�� " & _
                          "Where (վ�� = [2] Or վ�� is Null) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
                          "  And ĩ��=1 And (substr(����,5,1)=1 Or Nvl(ĩ��,0)=0) " & _
                          "  And (���� like [1] Or ���� like [1] or ���� like [1] )"

                Set RecReturn = zldatabase.ShowSQLSelect(Me, gstrSQL, 0, "��ҩ��λ", False, "", "", False, False, _
                                True, dblLeft, dblTop, 1000, False, False, True, IIf(gstrMatchMethod = "0", "%", "") & strKey & "%", gstrNodeNo)
                If RecReturn Is Nothing Then
                    MsgBox "δ�ҵ���Ӧ�̡�" & Trim(.EditText) & "�������������룡", vbInformation, gstrSysName
                    .EditText = ""
                    .TextMatrix(.Row, .ColIndex("��Ӧ��")) = ""
                    .TextMatrix(.Row, .ColIndex("��ҩ��λid")) = ""
                    Exit Sub
                Else
                    .EditText = RecReturn!����
                    .TextMatrix(.Row, .ColIndex("��Ӧ��")) = RecReturn!����
                    .TextMatrix(.Row, .ColIndex("��ҩ��λid")) = RecReturn!Id
                    .Col = .ColIndex("��������")
                End If
            Case .ColIndex("��������")
                If Not (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then
                    If InStr(1, strKey, ".") > 0 Then
                        If Chr(KeyAscii) = "." Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                        
                        If .EditSelLength = Len(strKey) Then Exit Sub
                        If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= Len(Mid(mFMT.FM_����, InStr(1, mFMT.FM_����, ".") + 1)) And strKey Like "*.*" Then
                            KeyAscii = 0
                            Exit Sub
                        Else
                            Exit Sub
                        End If
                    End If
                    
                    If InStr(1, strKey, "-") > 0 Then
                        If Chr(KeyAscii) = "-" Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                    
                    If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or Chr(KeyAscii) = ".") Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        If Val(strKey + Chr(KeyAscii)) > 99999999 Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                    
                End If
        End Select
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim rsTemp As ADODB.Recordset
    Dim str��װϵ�� As String
    Dim dbl�ɱ��� As Double
    Dim dbl���ۼ� As Double
    Dim dbl�ɱ���� As Double
    Dim dbl�ۼ۽�� As Double
    Dim dbl�������� As Double
    Dim dbl�������� As Double
    
    On Error GoTo ErrHandle
    
    ArrNum = Array()
    
    gstrSQL = "Select e.No, c.����id, b.����, b.����, b.���, d.���� As ��Ӧ��, a.����, a.����, a.����, a.����ԭ��, a.����취, " & IIf(mintUnit = 0, " b.���㵥λ", " c.��װ��λ") & " as ��λ, " & vbNewLine & _
            " a.��������, a.�ɱ���, a.���ۼ�,a.��ҩ��λid, c.����ϵ��, e.�Ǽ���, e.�Ǽ�����, e.������, e.��������, e.��ע, b.�Ƿ���, f.�������� " & vbNewLine & _
            " From ����������¼ A, �շ���ĿĿ¼ B, �������� C, ��Ӧ�� D, ������������ E, ҩƷ��� F, ҩƷ�շ���¼ H " & vbNewLine & _
            " Where a.����id = e.Id And a.����id = b.Id And b.Id = c.����id And a.�ⷿid=f.�ⷿid And a.����id=f.ҩƷid And nvl(a.����,0) = nvl(f.����,0) " & vbNewLine & _
            " And e.No=h.No And h.����=21 And a.�ⷿid=h.�ⷿid And a.����id=h.ҩƷid And nvl(a.����,0) = nvl(h.����,0) " & vbNewLine & _
            " And a.��ҩ��λid = d.Id(+) And a.����id = [1] " & vbNewLine & _
            " Order by h.��� "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����id)
    
    With VSFDetail
        Do While Not rsTemp.EOF
            TxtNo.Text = rsTemp!NO
            .TextMatrix(.Rows - 1, .ColIndex("������Ϣ")) = "[" & rsTemp!���� & "]" & rsTemp!����
            .TextMatrix(.Rows - 1, .ColIndex("���")) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            .TextMatrix(.Rows - 1, .ColIndex("��Ӧ��")) = IIf(IsNull(rsTemp!��Ӧ��), "", rsTemp!��Ӧ��)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.Rows - 1, .ColIndex("��λ")) = IIf(IsNull(rsTemp!��λ), "", rsTemp!��λ)
            .TextMatrix(.Rows - 1, .ColIndex("����ԭ��")) = IIf(IsNull(rsTemp!����ԭ��), "", rsTemp!����ԭ��)
            .TextMatrix(.Rows - 1, .ColIndex("����취")) = IIf(IsNull(rsTemp!����취), "", rsTemp!����취)
            .TextMatrix(.Rows - 1, .ColIndex("����id")) = IIf(IsNull(rsTemp!����ID), "", rsTemp!����ID)
            .TextMatrix(.Rows - 1, .ColIndex("��ҩ��λid")) = IIf(IsNull(rsTemp!��ҩ��λID), "", rsTemp!��ҩ��λID)
            .TextMatrix(.Rows - 1, .ColIndex("����ϵ��")) = IIf(IsNull(rsTemp!����ϵ��), "", rsTemp!����ϵ��)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.Rows - 1, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.Rows - 1, .ColIndex("�Ƿ���")) = IIf(IsNull(rsTemp!�Ƿ���), "", rsTemp!�Ƿ���)
            
            str��װϵ�� = IIf(mintUnit = 0, 1, rsTemp!����ϵ��)
            
            If IsNull(rsTemp!��������) = False Then dbl�������� = Val(rsTemp!��������) / Val(str��װϵ��)
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(dbl��������, mFMT.FM_����)
            
            If IsNull(rsTemp!��������) = False Then dbl�������� = Val(rsTemp!��������) / Val(str��װϵ��)
            .TextMatrix(.Rows - 1, .ColIndex("��������")) = Format(dbl��������, mFMT.FM_����)
            
            If IsNull(rsTemp!�ɱ���) = False Then dbl�ɱ��� = Val(rsTemp!�ɱ���) * Val(str��װϵ��)
            .TextMatrix(.Rows - 1, .ColIndex("�ɱ���")) = Format(dbl�ɱ���, mFMT.FM_�ɱ���)
            
            If IsNull(rsTemp!���ۼ�) = False Then dbl���ۼ� = Val(rsTemp!���ۼ�) * Val(str��װϵ��)
            .TextMatrix(.Rows - 1, .ColIndex("���ۼ�")) = Format(dbl���ۼ�, mFMT.FM_���ۼ�)
            
            dbl�ɱ���� = Val(rsTemp!�ɱ���) * Val(rsTemp!��������)
            .TextMatrix(.Rows - 1, .ColIndex("�ɱ����")) = Format(dbl�ɱ����, mFMT.FM_���)
            
            dbl�ۼ۽�� = Val(rsTemp!���ۼ�) * Val(rsTemp!��������)
            .TextMatrix(.Rows - 1, .ColIndex("�ۼ۽��")) = Format(dbl�ۼ۽��, mFMT.FM_���)
            
            txtCheck.Text = IIf(IsNull(rsTemp!�Ǽ���), "", rsTemp!�Ǽ���)
            If IsNull(rsTemp!�Ǽ�����) = False Then
                txtCheckDate.Caption = Format(rsTemp!�Ǽ�����, "yyyy-mm-dd hh:mm:ss")
            End If
            
            If mint�༭״̬ = 4 Then
                txtVerify.Text = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                If IsNull(rsTemp!��������) = False Then
                    txtVerifyDate.Caption = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
                End If
            End If
            
            txt��ע.Text = IIf(IsNull(rsTemp!��ע), "", rsTemp!��ע)
            
            If mint�༭״̬ = 2 And mblnUsableNum = True Then
                ReDim Preserve ArrNum(UBound(ArrNum) + 1)
                ArrNum(UBound(ArrNum)) = Val(.TextMatrix(.Rows - 1, .ColIndex("����id"))) & "," & Val(.TextMatrix(.Rows - 1, .ColIndex("����"))) & "," & dbl��������
            End If
            
            .Rows = .Rows + 1
            rsTemp.MoveNext
        Loop
        
        If .Rows > 1 Then
            .Row = 1
            .Col = 0
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfDetail_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With VSFDetail
            If mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
                .Editable = flexEDNone
            ElseIf .Col = .ColIndex("������Ϣ") Or .Col = .ColIndex("��Ӧ��") Then
                .Editable = flexEDKbdMouse
                .ColComboList(.Col) = "|..."
            ElseIf .Col = .ColIndex("����ԭ��") Or .Col = .ColIndex("����취") Then
                .Editable = flexEDKbdMouse
                .ColComboList(.Col) = "..."
            ElseIf .Col = .ColIndex("��������") Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
                .ColComboList(.Col) = ""
            End If
        End With
    End If
End Sub

Private Sub vsfDetail_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    With VSFDetail
        Select Case .ColKey(Col)
            Case "������Ϣ"
                VSFDetail_KeyPressEdit Row, Col, vbKeyReturn
            Case "��Ӧ��"
                If Val(.TextMatrix(Row, .ColIndex("����id"))) = 0 Then Exit Sub
                mblnValidateEdit = True
                VSFDetail_KeyPressEdit Row, Col, vbKeyReturn
                mblnValidateEdit = False
            Case "��������"
                If Val(.TextMatrix(Row, .ColIndex("����id"))) = 0 Then Exit Sub
                If Val(Trim(.EditText)) = 0 Then
                    .TextMatrix(.Row, .Col) = .EditText
                    MsgBox "������������Ϊ0���,��ֻ�������֣�", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    Exit Sub
                ElseIf Val(Trim(.EditText)) < 0 Then
                    If Val(.TextMatrix(.Row, .ColIndex("�Ƿ���"))) <> 0 Or Val(.TextMatrix(.Row, .ColIndex("����"))) <> 0 Then
                        .TextMatrix(.Row, .Col) = .EditText
                        MsgBox "�ò��ϲ��ܸ������⣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
                
                If Not CompareUsableQuantity(Row, Val(Trim(.EditText))) Then
                    Cancel = True
                    .TextMatrix(.Row, .Col) = .EditText
                    Exit Sub
                Else
                    .EditText = Format(Val(Trim(.EditText)), mFMT.FM_����)
                    .TextMatrix(.Row, .Col) = Format(Val(Trim(.EditText)), mFMT.FM_����)
                    .TextMatrix(.Row, .ColIndex("�ɱ����")) = Format(Val(.TextMatrix(.Row, .ColIndex("�ɱ���"))) * Val(Trim(.EditText)), mFMT.FM_���)
                    .TextMatrix(.Row, .ColIndex("�ۼ۽��")) = Format(Val(.TextMatrix(.Row, .ColIndex("���ۼ�"))) * Val(Trim(.EditText)), mFMT.FM_���)
                    .Col = .ColIndex("����ԭ��")
                End If
        End Select
    End With
End Sub

Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl��д���� As Double) As Boolean
    '������������бȽ�
    Dim dblUsableQuantity As Double '��������
    Dim dblOldNum As Double '�޸�ǰ��������
    Dim lng����ID As Long
    Dim lng���� As Long
    Dim intLop As Integer
    
    'mint�����: 0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    
    CompareUsableQuantity = False
    
    With VSFDetail
        If .TextMatrix(intRow, .ColIndex("����id")) = "" Or Val(dbl��д����) = 0 Then Exit Function
        
        If mint�༭״̬ = 2 And mblnUsableNum = True Then
            For intLop = 0 To UBound(ArrNum)
                lng����ID = Val(Split(ArrNum(intLop), ",")(0))
                lng���� = Val(Split(ArrNum(intLop), ",")(1))
                If lng����ID = Val(.TextMatrix(intRow, .ColIndex("����id"))) And lng���� = Val(.TextMatrix(intRow, .ColIndex("����"))) Then
                    dblOldNum = Val(Split(ArrNum(intLop), ",")(2))
                    Exit For
                End If
            Next
        End If
        
        dblUsableQuantity = Val(.TextMatrix(intRow, .ColIndex("��������")))
        .TextMatrix(intRow, .ColIndex("��������")) = dbl��д����
        
        If Val(.TextMatrix(intRow, .ColIndex("����"))) > 0 Or Val(.TextMatrix(intRow, .ColIndex("�Ƿ���"))) = 1 Then '���Ƴ��ⷿ�ǿⷿ�������Ƿ�����������ĵ��ж�
            If mint�༭״̬ = 1 Then
                If dbl��д���� > dblUsableQuantity Then
                    MsgBox "�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Then
                If dbl��д���� > dblUsableQuantity + dblOldNum Then
                    MsgBox "�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity + dblOldNum & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            .EditText = Format(dbl��д����, mFMT.FM_����)
            .TextMatrix(intRow, .ColIndex("��������")) = Format(dbl��д����, mFMT.FM_����)
            CompareUsableQuantity = True
            Exit Function
        End If
        
        ' ����������
        
        If mint����� = 0 Then
            '0-�����
        ElseIf mint����� = 1 Then
            '1-��飬��������
            If mint�༭״̬ = 1 Then
                If dbl��д���� > dblUsableQuantity Then
                    If MsgBox("�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Then
                If dbl��д���� > dblUsableQuantity + dblOldNum Then
                    If MsgBox("�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity + dblOldNum & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
        ElseIf mint����� = 2 Then
            '2-��飬�����ֹ
            If mint�༭״̬ = 1 Then
                If dbl��д���� > dblUsableQuantity Then
                    MsgBox "�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Then
                If dbl��д���� > dblUsableQuantity + dblOldNum Then
                    MsgBox "�������������" & dbl��д���� & "�������˸����ĵĿ��ÿ��������" & dblUsableQuantity + dblOldNum & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
        .TextMatrix(intRow, .ColIndex("��������")) = Format(dbl��д����, mFMT.FM_����)
    End With
    
    CompareUsableQuantity = True
    
End Function
