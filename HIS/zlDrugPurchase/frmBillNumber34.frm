VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBillNumber34 
   Caption         =   "���͵��ŵ���"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15855
   Icon            =   "frmBillNumber34.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   15855
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   14640
      TabIndex        =   3
      Top             =   7695
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "������ⵥ��(&O)"
      Height          =   350
      Left            =   12960
      TabIndex        =   2
      Top             =   7695
      Width           =   1575
   End
   Begin VB.PictureBox pic��Ʊ��Ϣ 
      BackColor       =   &H00FF80FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4320
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   14
      Top             =   7740
      Width           =   260
   End
   Begin VB.PictureBox pic��Ӧ����ɫ 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   6480
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   7740
      Width           =   260
   End
   Begin VB.Frame fra����ⷿ 
      Caption         =   "����ⷿ"
      Height          =   675
      Left            =   9120
      TabIndex        =   10
      Top             =   840
      Width           =   3345
      Begin VB.ComboBox cbo�ⷿ 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame fra��ȡ�������� 
      Caption         =   "��ȡ��������"
      Height          =   675
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   8505
      Begin VB.CommandButton cmd��ȡ��¼��Ʊ��Ϣ 
         Caption         =   "��ȡ��¼��Ʊ��Ϣ(&P)"
         Height          =   350
         Left            =   6360
         TabIndex        =   16
         Top             =   215
         Width           =   2055
      End
      Begin VB.CommandButton cmd��ȡ���� 
         Caption         =   "��ȡ���͵�����(&T)"
         Height          =   350
         Left            =   4440
         TabIndex        =   7
         Top             =   215
         Width           =   1695
      End
      Begin VB.TextBox txt���͵��� 
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   3075
      End
      Begin VB.Label lbl���͵��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "���͵���"
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   300
         Width           =   720
      End
   End
   Begin VB.Frame Frmline1 
      Height          =   120
      Left            =   120
      TabIndex        =   1
      Top             =   645
      Width           =   15735
   End
   Begin VB.Frame Frmline2 
      Height          =   135
      Left            =   120
      TabIndex        =   0
      Top             =   7440
      Width           =   15735
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   5685
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   15540
      _cx             =   27411
      _cy             =   10028
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   37
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   315
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBillNumber34.frx":6852
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
      ExplorerBar     =   5
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
   Begin VB.Label lbl���������� 
      AutoSize        =   -1  'True
      Caption         =   "��ʾ������0�ŵ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   17
      Top             =   7770
      Width           =   1785
   End
   Begin VB.Label lbl��Ʊ��Ϣ 
      AutoSize        =   -1  'True
      Caption         =   "��Ʊ��Ϣ����ȷ"
      Height          =   180
      Left            =   4680
      TabIndex        =   15
      Top             =   7785
      Width           =   1260
   End
   Begin VB.Label lbl��Ӧ����ɫ 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ����Ϣ����ȷ"
      Height          =   180
      Left            =   6840
      TabIndex        =   13
      Top             =   7785
      Width           =   1440
   End
   Begin VB.Image Image 
      Height          =   480
      Left            =   255
      Picture         =   "frmBillNumber34.frx":6D9C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblInfor 
      Caption         =   $"frmBillNumber34.frx":70A6
      Height          =   405
      Left            =   840
      TabIndex        =   9
      Top             =   240
      Width           =   10695
   End
End
Attribute VB_Name = "frmBillNumber34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mCanColColor As Long = &H8080FF           '��Ӧ����Ϣ�������ǳ��ɫ��ʶ
Private Const mNoColColor As Long = &H80000012           '��Ӧ����Ϣ��ȷ��Ϊ��ɫ
Private Const mFPColColor As Long = &HFF80FF              '��Ʊ��Ϣ�������ǳ��ɫ��ʶ
Private mint���ݿ����� As Integer              '0:SQLserver���ݿ⣻1��Oracle���ݿ�
Private mblnIsConn As Boolean
Private marrSql As Variant

Private Sub cbo�ⷿ_Click()
    If Val(cbo�ⷿ.ListIndex) <> Val(cbo�ⷿ.Tag) And vsfList.Rows > 1 Then
        If MsgBox("����ı�ⷿ����Ҫ������ȡ�������ݣ��Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, GSTR_MESSAGE) = vbYes Then
            vsfList.Rows = 1
            lbl����������.Caption = "��ʾ������0�ŵ���"
        Else
            cbo�ⷿ.ListIndex = Val(cbo�ⷿ.Tag)
        End If
    End If
    cbo�ⷿ.Tag = Val(cbo�ⷿ.ListIndex)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim i As Integer
    Dim blnTrans As Boolean
    
    If vsfList.Rows < 2 Then Exit Sub
    If cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex) = -1 Then
        MsgBox "��ѡ��ⷿ��", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    On Error GoTo ErrHand

    If MsgBox("�Ƿ�ȷ�����룿", vbQuestion + vbYesNo + vbDefaultButton2, GSTR_MESSAGE) = vbNo Then
        Exit Sub
    End If
    
    If cmdOk.Tag = "������ⵥ" Then
        Call Save������ⵥ
    ElseIf cmdOk.Tag = "��¼��Ʊ��Ϣ" Then
        Call Saveδ��˷�Ʊ��Ϣ
        Call Save����˷�Ʊ��Ϣ
    End If
    
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(marrSql)
        Call gobjComLib.zlDatabase.ExecuteProcedure(CStr(marrSql(i)), "�����⹺��ⵥ")
    Next
    gcnOracle.CommitTrans: blnTrans = False
    
    marrSql = Array()
    MsgBox "����ɹ���", vbInformation, GSTR_MESSAGE
    vsfList.Rows = 1

    lbl����������.Caption = "��ʾ������0�ŵ���"
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = vbDefault
    MsgBox "����ʧ�ܣ����飡", vbInformation, GSTR_MESSAGE
End Sub

Private Sub cmd��ȡ��¼��Ʊ��Ϣ_Click()
    Dim strSQL As String
    Dim rs��Ʊ��Ϣ As New ADODB.Recordset
    Dim rs��¼��Ʊ��Ϣ As New ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If mblnIsConn = False Then
        MsgBox "���������м����ݿ����ӣ�", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

    If Trim(txt���͵���.Text) = "" Then
        MsgBox "����¼�����͵��ţ�", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    strSQL = "Select Distinct g.���͵��� , a.No, a.ҩƷid, a.���, d.���� As ҩƷ����,d.����, d.���, d.���㵥λ as ��λ,a.���� as ������,  a.���� as ����, a.Ч�� as ��Ч����, a.��д���� As ����, a.�ɱ��� as ����, a.�ɱ���� As ���," & vbNewLine & _
                    "                Nvl(a.����, 0) ����,d.�Ƿ���, Decode(a.����, Null, 0, a.����) As ����, a.���ۼ� As ���ۼ�, a.���۽��, a.���, a.��׼�ĺ�, c.�������, c.��Ʊ��," & vbNewLine & _
                    "                c.��Ʊ����, c.��Ʊ����, c.��Ʊ���, a.��ҩ��λid, f.���� As ��Ӧ�̱���, f.���� As ��Ӧ��, a.ժҪ , a.��Ʒ�ϸ�֤ , a.������, a.��������,  a.�����," & vbNewLine & _
                    "                a.�������, a.�ⷿid, a.���, a.���ս���, a.��������, a.��ҩ�� As �˲���, a.��ҩ���� As �˲�����, Nvl(a.�÷�, 0) As ����, a.Ƶ�� As �ӳ���," & vbNewLine & _
                    "                a.�Է�����id, a.�ƻ�id " & vbNewLine & _
                    "From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ D, Ӧ����¼ C, ��Ӧ�� F, ���͵��Ŷ��� G " & vbNewLine & _
                    "Where a.ҩƷid = b.ҩƷid And b.ҩƷid = d.Id And a.��ҩ��λid = f.Id And a.�ⷿid=[2] And Substr(f.����, 1, 1) = 1 And a.Id = c.�շ�id(+) And" & vbNewLine & _
                    "      c.ϵͳ��ʶ(+) = 1 And c.��¼����(+) = 0 And a.��¼״̬ = 1 And a.���� = 1 And a.No = g.No And g.���� = 1" & vbNewLine & _
                    "      And Not Exists (Select 1 From Ӧ����¼ Where ID = c.Id And ������� Is Not Null) And g.���͵��� =[1] order by a.no , a.ҩƷid"

    Set rs��¼��Ʊ��Ϣ = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯҩƷ��Ϣ", txt���͵���.Text, Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)))
    

    strSQL = "SELECT FPHM ��Ʊ��,PSQYBM ��Ӧ�̱���,KPRQ ��Ʊ����,GLDJLX ����," & _
                    "GLMXBH ���͵���,YPBM ҩƷ����,DJ ��Ʊ����,SL ��Ʊ����,JE ��Ʊ��� " & _
                    "from mid_invoice where  GLDJLX='1' and GLMXBH='" & txt���͵���.Text & "'"
    
    rs��Ʊ��Ϣ.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
    
'    If mint���ݿ����� = 0 Then
'        'SQLserver���ݿ�
'        rs��Ʊ��Ϣ.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
'    Else
'        'Oracle���ݿ�
'        Set rs��Ʊ��Ϣ = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ��Ʊ��Ϣ")
'    End If
    
    If rs��¼��Ʊ��Ϣ.RecordCount > 0 And rs��Ʊ��Ϣ.RecordCount > 0 Then
        cmdOk.Caption = "���뷢Ʊ��Ϣ(&O)"
        cmdOk.Tag = "��¼��Ʊ��Ϣ"
        Call DataVsf(rs��¼��Ʊ��Ϣ, rs��Ʊ��Ϣ)
    Else
        MsgBox "û�в�ѯ����Ʊ���ݣ����飡", vbInformation, GSTR_MESSAGE
        vsfList.Rows = 1
        lbl����������.Caption = "��ʾ������0�ŵ���"
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    If mblnIsConn = False Then
        MsgBox "���������м����ݿ����ӣ�", vbInformation, GSTR_MESSAGE
    Else
        MsgBox "��ȡ�ⲿ���ݴ���", vbInformation, GSTR_MESSAGE
    End If
    vsfList.Rows = 1
End Sub

Private Sub cmd��ȡ����_Click()
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim rs��Ʊ��Ϣ As New ADODB.Recordset
    Dim rs������͵���Ϣ As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If mblnIsConn = False Then
        MsgBox "���������м����ݿ����ӣ�", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

    If Trim(txt���͵���.Text) = "" Then
        MsgBox "����¼�����͵��ţ�", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

    strSQL = "Select 1" & vbNewLine & _
                    "From ���͵��Ŷ��� A, ҩƷ�շ���¼ B" & vbNewLine & _
                    "Where a.No = b.No And a.���� = 1 And a.���͵��� = [1] And Rownum < 2"
                    
    Set rs������͵���Ϣ = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ���͵��Ƿ��Ѿ�������", Trim(txt���͵���.Text))
        
    If rs������͵���Ϣ.RecordCount > 0 Then
        If MsgBox("���͵���[" & Trim(txt���͵���.Text) & "]�Ѿ�����������¼���Ƿ������ȡ���ݣ�", vbQuestion + vbYesNo + vbDefaultButton2, GSTR_MESSAGE) = vbNo Then
            Exit Sub
        End If
    End If
    
    strSQL = "select KFBM �ⷿid,PSQYBM ��Ӧ�̱���,PSDH ���͵���,YPBM ҩƷ����,SCPH ����,SCRQ ��������," _
                  & " YXRQ ��Ч����,DJ ����,SL ����,JLDW ��λ,JE ���,SCQY ������ " _
                  & " From MID_DELIVERY_ORDER  where PSDH='" & txt���͵���.Text & "' order by ��Ӧ�̱���,ҩƷ����"
    
    rsTmp.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
    
'    If mint���ݿ����� = 0 Then
'        'SQLserver���ݿ�
'        rsTmp.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
'    Else
'        'Oracle���ݿ�
'        Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯҩƷ��Ϣ")
'    End If
    
    strSQL = "SELECT FPHM ��Ʊ��,PSQYBM ��Ӧ�̱���,KPRQ ��Ʊ����,GLDJLX ����," & _
                    "GLMXBH ���͵���,YPBM ҩƷ����,DJ ��Ʊ����,SL ��Ʊ����,JE ��Ʊ��� " & _
                    "from mid_invoice where  GLDJLX='1' and GLMXBH='" & txt���͵���.Text & "'"
    
    rs��Ʊ��Ϣ.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
    
'    If mint���ݿ����� = 0 Then
'        'SQLserver���ݿ�
'        rs��Ʊ��Ϣ.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly
'    Else
'        'Oracle���ݿ�
'        Set rs��Ʊ��Ϣ = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ��Ʊ��Ϣ")
'    End If
    
    If rsTmp.RecordCount > 0 Then
        cmdOk.Caption = "������ⵥ��(&O)"
        cmdOk.Tag = "������ⵥ"
        Call DataVsf(rsTmp, rs��Ʊ��Ϣ)
    Else
        MsgBox "û�в�ѯ�����ݣ����飡", vbInformation, GSTR_MESSAGE
        vsfList.Rows = 1
        lbl����������.Caption = "��ʾ������0�ŵ���"
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    If mblnIsConn = False Then
        MsgBox "���������м����ݿ����ӣ�", vbInformation, GSTR_MESSAGE
    Else
        MsgBox "��ȡ�ⲿ���ݴ���", vbInformation, GSTR_MESSAGE
    End If
    vsfList.Rows = 1
End Sub

Private Sub Form_Load()
    vsfList.AllowSelection = False '���ܶ�ѡ
    vsfList.Rows = 1
    Call GetUserNameInfo
    Call SetMedicalWH
    Call ConnectDatabase
    marrSql = Array()
    lbl����������.Caption = "��ʾ������0�ŵ���"
End Sub

Public Function GetUserNameInfo() As Boolean
'��ȡ�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = gobjComLib.zlDatabase.GetUserInfo
    
    With rsTmp
        If Not .EOF Then
            glngUserID = IIf(IsNull(!Id), 0, !Id)
            glngDeptID = IIf(IsNull(!����id), 0, !����id)
            gstrUserNameNew = IIf(IsNull(!����), "", !����) '��ǰ�û�����
            GetUserNameInfo = True
        Else
            glngUserID = 0
            glngDeptID = 0
            gstrUserNameNew = "" '��ǰ�û�����
        End If
    End With
    rsTmp.Close

    strSQL = "Select ������, ����ֵ, ȱʡֵ From Zlparameters Where ϵͳ = [1] And Nvl(˽��, 0) = 0 And ģ�� Is Null and ������=[2] "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "ȡϵͳ����", 100, 149)
    With rsTmp
        If Not .EOF Then
            gbytЧ�� = IIf(IsNull(rsTmp!����ֵ), rsTmp!ȱʡֵ, rsTmp!����ֵ)
        Else
            gbytЧ�� = 0
        End If
    End With
    
End Function

Private Sub SetMedicalWH()
'����ҩ��combobox��Ϣ��ͬHIS�����û�Ҫ��HIS�Ĳ���Ȩ��һ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i, j As Integer
    Dim strStock As String
    
    If InStr(1, gstrPrivs, "����ҩ���⹺���") = 0 Then
        strStock = "HIJ"
    Else
        strStock = "HIJKLMN"
    End If
            
    'ҩ����Ϣ
    strSQL = "SELECT DISTINCT a.id, a.���� " _
            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = '-' Or a.վ�� is Null) And c.�������� = b.���� " _
            & "  AND Instr([2],b.����,1) > 0 " _
            & "  AND a.id = c.����id " _
            & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
            & IIf(InStr(1, gstrPrivs, "���пⷿ") > 0, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])") _
            & " order by a.id"
            
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngUserID, strStock)
    
    cbo�ⷿ.Clear
    For i = 0 To rsTmp.RecordCount - 1
        cbo�ⷿ.AddItem rsTmp!����
        cbo�ⷿ.ItemData(i) = rsTmp!Id
        rsTmp.MoveNext
    Next
    cbo�ⷿ.Tag = IIf(gintListIndex = -1, 0, gintListIndex)
    cbo�ⷿ.ListIndex = IIf(gintListIndex = -1, 0, gintListIndex)
    rsTmp.Close
End Sub


Private Sub DataVsf(ByVal rsVal As ADODB.Recordset, ByVal rs��Ʊ��Ϣ As ADODB.Recordset)
'���������
    Dim i As Integer
    Dim str��Ӧ����Ϣ As String
    Dim str��Ӧ�� As String
    Dim strSQL As String
    Dim lng��Ӧ��id As Long
    Dim lngҩƷID As Long
    Dim strҩƷ���� As String
    Dim strҩƷ��� As String
    Dim rsҩƷ��Ϣ As New ADODB.Recordset
    Dim dbl���ۼ� As Double
    Dim int��� As Integer
    Dim int������������ As Integer
    Dim str��Ӧ��id As String
    Dim strNO�� As String
    
    On Error GoTo ErrHand
    
    If cmdOk.Tag = "������ⵥ" Then
        strSQL = "Select Distinct a.Id, a.����, a.���, a.����, a.�Ƿ���, c.�ּ�" & vbNewLine & _
                        "From �շ���ĿĿ¼ A, ҩƷ��� B, �շѼ�Ŀ C" & vbNewLine & _
                        "Where a.Id = b.ҩƷid And a.Id = c.�շ�ϸĿid And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                        "      Sysdate Between c.ִ������ And Nvl(c.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))  And Exists" & vbNewLine & _
                        "     (Select 1 From �շ�ִ�п��� D Where b.ҩƷid = d.�շ�ϸĿid And d.ִ�п���id = [1])"
   
        Set rsҩƷ��Ϣ = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯҩƷ��Ϣ", Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex)))
    
        If rsҩƷ��Ϣ.RecordCount = 0 Then
            MsgBox "û�в�ѯ��ҩƷ��Ϣ�����飡", vbInformation, GSTR_MESSAGE
            vsfList.Rows = 1
            lbl����������.Caption = "��ʾ������0�ŵ���"
            Exit Sub
        End If
    End If
    
    vsfList.Rows = 1
    With vsfList
        For i = 1 To rsVal.RecordCount
            If cmdOk.Tag = "������ⵥ" Then
                If IIf(IsNull(rsVal!������), "", rsVal!������) <> "" Then
                    Call Add������(rsVal!������)
                End If
                
                str��Ӧ����Ϣ = Check��Ӧ��(rsVal!��Ӧ�̱���)
                 str��Ӧ�� = Split(str��Ӧ����Ϣ, "|")(0)
                 lng��Ӧ��id = Val(Split(str��Ӧ����Ϣ, "|")(1))
                
                rsҩƷ��Ϣ.Filter = "����='" & rsVal!ҩƷ���� & "'"
                If rsҩƷ��Ϣ.RecordCount = 0 Then
                    lngҩƷID = 0
                Else
                    lngҩƷID = rsҩƷ��Ϣ!Id
                    strҩƷ���� = "[" & rsҩƷ��Ϣ!���� & "]" & rsҩƷ��Ϣ!����
                    strҩƷ��� = rsҩƷ��Ϣ!���
                    dbl���ۼ� = rsҩƷ��Ϣ!�ּ�
                    int��� = IIf(IsNull(rsҩƷ��Ϣ!�Ƿ���), 0, rsҩƷ��Ϣ!�Ƿ���)
                End If
            ElseIf cmdOk.Tag = "��¼��Ʊ��Ϣ" Then
                str��Ӧ�� = rsVal!��Ӧ��
                lng��Ӧ��id = Val(rsVal!��ҩ��λID)
                lngҩƷID = Val(rsVal!ҩƷID)
                strҩƷ���� = "[" & rsVal!ҩƷ���� & "]" & rsVal!����
                strҩƷ��� = rsVal!���
            End If
            
            rs��Ʊ��Ϣ.Filter = "��Ӧ�̱���='" & rsVal!��Ӧ�̱��� & "' and ҩƷ����='" & rsVal!ҩƷ���� & "'"
    
            If lngҩƷID <> 0 Then
                .Rows = .Rows + 1
                .TextMatrix(i, .ColIndex("�к�")) = i
                .TextMatrix(i, .ColIndex("ҩƷID")) = lngҩƷID
                .TextMatrix(i, .ColIndex("ҩƷ����")) = strҩƷ����
                .TextMatrix(i, .ColIndex("���")) = strҩƷ���
                .TextMatrix(i, .ColIndex("������")) = NVL(rsVal!������)
                .TextMatrix(i, .ColIndex("����")) = NVL(rsVal!����)
                .TextMatrix(i, .ColIndex("��������")) = Format(NVL(rsVal!��������), "yyyy-mm-dd")
                .TextMatrix(i, .ColIndex("��Ч����")) = Format(NVL(rsVal!��Ч����), "yyyy-mm-dd")
                .TextMatrix(i, .ColIndex("��λ")) = NVL(rsVal!��λ)
                .TextMatrix(i, .ColIndex("����")) = Format(NVL(rsVal!����, 0), "0.0000")
                .TextMatrix(i, .ColIndex("����")) = Format(NVL(rsVal!����, 0), "0.00")
                .TextMatrix(i, .ColIndex("���")) = Format(NVL(rsVal!���, 0), "0.00")
                .TextMatrix(i, .ColIndex("��Ӧ��")) = str��Ӧ��
                .TextMatrix(i, .ColIndex("��Ӧ��id")) = lng��Ӧ��id
                
                .Cell(flexcpForeColor, i, .ColIndex("�к�"), i, .ColIndex("��Ʊ���")) = IIf(lng��Ӧ��id = 0, mCanColColor, mNoColColor)
                
                If cmdOk.Tag = "������ⵥ" Then
                    .TextMatrix(i, .ColIndex("���ۼ�")) = Format(IIf(int��� = 0, dbl���ۼ�, NVL(rsVal!����, 0)), "0.0000")
                    .TextMatrix(i, .ColIndex("���۽��")) = Format(.TextMatrix(i, .ColIndex("���ۼ�")) * .TextMatrix(i, .ColIndex("����")), "0.00")
                    .TextMatrix(i, .ColIndex("�ӳ���")) = Format(IIf(.TextMatrix(i, .ColIndex("����")) = 0, 0, .TextMatrix(i, .ColIndex("���ۼ�")) / .TextMatrix(i, .ColIndex("����")) - 1), "0.0000")
                    .TextMatrix(i, .ColIndex("���")) = Format(.TextMatrix(i, .ColIndex("���۽��")) - .TextMatrix(i, .ColIndex("���")), "0.00")
                    .ColHidden(.ColIndex("NO")) = True
                    .ColHidden(.ColIndex("�����")) = True
                    .ColHidden(.ColIndex("�������")) = True
                ElseIf cmdOk.Tag = "��¼��Ʊ��Ϣ" Then
                    .TextMatrix(i, .ColIndex("���ۼ�")) = Format(NVL(rsVal!���ۼ�, 0), "0.0000")
                    .TextMatrix(i, .ColIndex("���۽��")) = Format(NVL(rsVal!���۽��, 0), "0.00")
                    .TextMatrix(i, .ColIndex("�ӳ���")) = Format(NVL(rsVal!�ӳ���, 0), "0.0000")
                    .TextMatrix(i, .ColIndex("���")) = Format(NVL(rsVal!���, 0), "0.00")
                    .TextMatrix(i, .ColIndex("NO")) = rsVal!NO
                    .TextMatrix(i, .ColIndex("���")) = NVL(rsVal!���)
                    .TextMatrix(i, .ColIndex("��Ʒ�ϸ�֤")) = NVL(rsVal!��Ʒ�ϸ�֤)
                    .TextMatrix(i, .ColIndex("�˲���")) = NVL(rsVal!�˲���)
                    .TextMatrix(i, .ColIndex("�˲�����")) = NVL(rsVal!�˲�����)
                    .TextMatrix(i, .ColIndex("����")) = NVL(rsVal!����, 0)
                    .TextMatrix(i, .ColIndex("��׼�ĺ�")) = NVL(rsVal!��׼�ĺ�)
                    .TextMatrix(i, .ColIndex("�������")) = NVL(rsVal!�������)
                    .TextMatrix(i, .ColIndex("����")) = NVL(rsVal!����, 0)
                    .TextMatrix(i, .ColIndex("��Ʊ����")) = NVL(rsVal!��Ʊ����)
                    .TextMatrix(i, .ColIndex("�ƻ�id")) = NVL(rsVal!�ƻ�id, 0)
                    .TextMatrix(i, .ColIndex("���ս���")) = NVL(rsVal!���ս���)
                    .TextMatrix(i, .ColIndex("�Է�����ID")) = NVL(rsVal!�Է�����ID, 0)
                    .TextMatrix(i, .ColIndex("�����")) = NVL(rsVal!�����)
                    .TextMatrix(i, .ColIndex("�������")) = NVL(rsVal!�������)
                    .TextMatrix(i, .ColIndex("���")) = NVL(rsVal!���)
                    .ColHidden(.ColIndex("NO")) = False
                    .ColHidden(.ColIndex("�����")) = False
                    .ColHidden(.ColIndex("�������")) = False
                End If
                '��Ʊ��Ϣ
                If rs��Ʊ��Ϣ.RecordCount = 1 Then
                    .TextMatrix(i, .ColIndex("��Ʊ��")) = NVL(rs��Ʊ��Ϣ!��Ʊ��)
                    .TextMatrix(i, .ColIndex("��Ʊ����")) = Format(NVL(rs��Ʊ��Ϣ!��Ʊ����), "yyyy-mm-dd")
                    
                    If Val(NVL(rs��Ʊ��Ϣ!��Ʊ����, 0)) = Val(.TextMatrix(i, .ColIndex("����"))) Then
                        .TextMatrix(i, .ColIndex("��Ʊ���")) = Format(NVL(rs��Ʊ��Ϣ!��Ʊ���, 0), "0.00")
                    Else
                        If Val(NVL(rs��Ʊ��Ϣ!��Ʊ����, 0)) = 0 Then
                            .TextMatrix(i, .ColIndex("��Ʊ���")) = Format(0, "0.00")
                        Else
                            .TextMatrix(i, .ColIndex("��Ʊ���")) = Format(NVL(rs��Ʊ��Ϣ!��Ʊ���, 0) / rs��Ʊ��Ϣ!��Ʊ���� * .TextMatrix(i, .ColIndex("����")), "0.00")
                        End If
                    End If
                    
                ElseIf rs��Ʊ��Ϣ.RecordCount > 1 Then
                    .Cell(flexcpForeColor, i, .ColIndex("�к�"), i, .ColIndex("��Ʊ���")) = mFPColColor
                End If
                
                If cmdOk.Tag = "��¼��Ʊ��Ϣ" Then
                    If NVL(rsVal!��Ʊ��) <> "" Then .TextMatrix(i, .ColIndex("��Ʊ��")) = NVL(rsVal!��Ʊ��)
                    If NVL(rsVal!��Ʊ����) <> "" Then .TextMatrix(i, .ColIndex("��Ʊ����")) = Format(NVL(rsVal!��Ʊ����), "yyyy-mm-dd")
                    If NVL(rsVal!��Ʊ���) <> "" Then .TextMatrix(i, .ColIndex("��Ʊ���")) = Format(NVL(rsVal!��Ʊ���, 0), "0.00")
                End If

                If cmdOk.Tag = "������ⵥ" Then
                    If .Cell(flexcpForeColor, i, .ColIndex("�к�"), i, .ColIndex("��Ʊ���")) = mNoColColor Then
                        If InStr(";" & str��Ӧ��id & ";", ";" & .TextMatrix(i, .ColIndex("��Ӧ��id")) & ";") = 0 Then
                            str��Ӧ��id = IIf(str��Ӧ��id = "", "", str��Ӧ��id & ";") & .TextMatrix(i, .ColIndex("��Ӧ��id"))
                            int������������ = int������������ + 1
                        End If
                    End If
                Else
                    If .Cell(flexcpForeColor, i, .ColIndex("�к�"), i, .ColIndex("��Ʊ���")) = mNoColColor Then
                        If InStr(";" & strNO�� & ";", ";" & .TextMatrix(i, .ColIndex("NO")) & ";") = 0 Then
                            strNO�� = IIf(strNO�� = "", "", strNO�� & ";") & .TextMatrix(i, .ColIndex("NO"))
                            int������������ = int������������ + 1
                        End If
                    End If
                End If
                
            End If
            
            rsVal.MoveNext
        Next
        
    End With

    If int������������ > 0 Then
        lbl����������.Caption = "��ʾ������" & int������������ & "�ŵ���"
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "��ȡ�ⲿ���ݴ���", vbInformation, GSTR_MESSAGE
    vsfList.Rows = 1
End Sub

Private Sub Add������(ByVal str������ As String)
    Dim int���� As Integer
    Dim strCodes As String
    Dim rs������ As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHand

    strSQL = "select ����,���� from ҩƷ������ where ����=[1]"
    Set rs������ = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ��������Ϣ", str������)
    If rs������.RecordCount = 0 Then
                
        strSQL = "SELECT Nvl(MAX(LENGTH(����)),2) As Length FROM ҩƷ������"
        Set rs������ = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "ҩƷ�����̱��볤��")
        int���� = rs������!length
        
        strSQL = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM ҩƷ������"
        Set rs������ = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "ҩƷ�����̱���")
        strCodes = rs������!Code
        
        int���� = Len(strCodes)
        strCodes = strCodes + 1
        If int���� >= Len(strCodes) Then
            strCodes = String(int���� - Len(strCodes), "0") & strCodes
        End If
    
        strSQL = "ZL_ҩƷ������_INSERT('" & strCodes & "','" & str������ & "',zlSpellCode('" & str������ & "',10))"
        
        Call gobjComLib.zlDatabase.ExecuteProcedure(strSQL, "")
    End If
    
    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "��ȡ�ⲿ���ݴ���", vbInformation, GSTR_MESSAGE
End Sub

Private Function Check��Ӧ��(ByVal str��Ӧ�̱��� As String) As String
    Dim rs��Ӧ�� As New ADODB.Recordset
    Dim str��Ӧ�� As String
    Dim strSQL As String
    Dim lng��Ӧ��id As Long
    
    On Error GoTo ErrHand

    strSQL = "Select a.Id, a.���� ,a.����" & vbNewLine & _
                    "From ��Ӧ�� A" & vbNewLine & _
                    "Where a.ĩ�� = 1 And substr(����,1,1)=1 And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) and a.����=[1]"
                    
    Set rs��Ӧ�� = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ѯ��Ӧ����Ϣ", str��Ӧ�̱���)

    If rs��Ӧ��.RecordCount > 0 Then
        Check��Ӧ�� = rs��Ӧ��!���� & "|" & rs��Ӧ��!Id
    Else
        Check��Ӧ�� = "|"
    End If
    
    Exit Function
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "��ȡ�ⲿ���ݴ���", vbInformation, GSTR_MESSAGE
End Function

Private Sub ConnectDatabase()
'�����������ݿ�
    Dim str������ As String, str���ݿ� As String, str�û��� As String, str���� As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    strSQL = "Select ����, ���� From ҩƷ��������ӿ� where ����='��ݸҩ�¹���ϵͳ' and �Ƿ�����=1"
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    '���ͣ�0--SQLserver���ݿ⣻1--Oracle���ݿ�
    'SQLserver���ݿ�   ���ã���������|���ݿ���|�û���|����
    'Oracle���ݿ�        ���ã���������|���ݿ���|�û���|����

    If rsTmp.RecordCount > 0 Then
        str������ = Split(rsTmp!����, "|")(0)
        str���ݿ� = Split(rsTmp!����, "|")(1)
        str�û��� = Split(rsTmp!����, "|")(2)
        str���� = Split(rsTmp!����, "|")(3)
    Else
        MsgBox "���ӷ�����ʧ�ܣ��������м����ݿ�����ӣ�", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

'        str������ = "����\WINCCPLUSMIG2008"
'        str���ݿ� = "master"
'        str�û��� = "sa"
'        str���� = "his"
    
    mint���ݿ����� = NVL(rsTmp!����, 0)
    If mint���ݿ����� = 0 Then
        'SQLserver���ݿ�
        mblnIsConn = MSSQLServerOpen(str������, str���ݿ�, str�û���, str����)
    Else
        'Oracle���ݿ�
        mblnIsConn = OraDataOpenTest(str���ݿ�, str�û���, str����)
    End If

    Exit Sub
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "��ȡ�ⲿ���ݴ���", vbInformation, GSTR_MESSAGE
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.Height < 8000 Then Me.Height = 8000
    If Me.Width < 14000 Then Me.Width = 14000
    
    Frmline1.Left = 0
    Frmline1.Width = Me.ScaleWidth
    Frmline2.Left = 0
    Frmline2.Width = Me.ScaleWidth
    Frmline2.Top = Me.ScaleHeight * 25 / 28
    
    vsfList.Left = Me.ScaleHeight / 80
    vsfList.Width = Me.ScaleWidth - Me.ScaleHeight / 40
    vsfList.Height = Frmline2.Top - vsfList.Top - Me.ScaleHeight / 40
    
    cmdCancel.Left = vsfList.Width - cmdCancel.Width + Me.ScaleHeight / 80
    cmdCancel.Top = Frmline2.Top + Me.ScaleHeight / 25
    
    cmdOk.Top = cmdCancel.Top
    cmdOk.Left = cmdCancel.Left - cmdOk.Width - 100

    lbl����������.Top = cmdCancel.Top + 80
    lbl����������.Left = vsfList.Left
    
    lbl��Ӧ����ɫ.Top = cmdCancel.Top + 100
    lbl��Ӧ����ɫ.Left = Me.ScaleWidth / 2
    
    pic��Ӧ����ɫ.Top = cmdCancel.Top + 60
    pic��Ӧ����ɫ.Left = lbl��Ӧ����ɫ.Left - pic��Ӧ����ɫ.Width - 50
    
    lbl��Ʊ��Ϣ.Top = cmdCancel.Top + 100
    lbl��Ʊ��Ϣ.Left = pic��Ӧ����ɫ.Left - lbl��Ʊ��Ϣ.Width - 500
    
    pic��Ʊ��Ϣ.Top = cmdCancel.Top + 60
    pic��Ʊ��Ϣ.Left = lbl��Ʊ��Ϣ.Left - pic��Ʊ��Ϣ.Width - 50
End Sub

Private Sub txt���͵���_KeyPress(KeyAscii As Integer)
    If InStr(" ~%^&|`'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���͵���_GotFocus()
    Me.txt���͵���.SelStart = 0: Me.txt���͵���.SelLength = 100
End Sub

Private Sub Save������ⵥ()
    Dim i As Integer
    Dim strSQL As String
    Dim strNO As String
    Dim str��Ӧ��id As String
    Dim strNO�� As String
    Dim strDate As String
    Dim int��� As Integer
    Dim lng�ⷿID As Long
    
    strDate = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    lng�ⷿID = Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    
    With vsfList
        For i = 1 To .Rows - 1
            If .Cell(flexcpForeColor, i, .ColIndex("�к�"), i, .ColIndex("��Ʊ���")) = mNoColColor Then
                If InStr(";" & str��Ӧ��id & ";", ";" & .TextMatrix(i, .ColIndex("��Ӧ��id")) & ";") = 0 Then
                    str��Ӧ��id = IIf(str��Ӧ��id = "", "", str��Ӧ��id & ";") & .TextMatrix(i, .ColIndex("��Ӧ��id"))
                    strNO = gobjComLib.zlDatabase.GetNextNo(21, lng�ⷿID)
                    int��� = 0
                    
                    strSQL = "Zl_���͵��Ŷ���_INSERT("
                    '�ⷿid_In
                    strSQL = strSQL & lng�ⷿID
                    '����_In
                    strSQL = strSQL & ",1"
                    'No_In
                    strSQL = strSQL & ",'" & strNO & "'"
                    '���͵���_In
                    strSQL = strSQL & ",'" & txt���͵���.Text & "'"
                    strSQL = strSQL & ")"
                    
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = strSQL
                    
                End If
                
                int��� = int��� + 1
                
                strSQL = "zl_ҩƷ�⹺_INSERT("
                'NO
                strSQL = strSQL & "'" & strNO & "'"
                '���
                strSQL = strSQL & "," & int���
                '�ⷿID
                strSQL = strSQL & "," & lng�ⷿID
                '�Է�����ID
                strSQL = strSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("�Է�����ID"))) = 0, "NULL", Val(.TextMatrix(i, .ColIndex("�Է�����ID"))))
                '��ҩ��λID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("��Ӧ��id")))
                'ҩƷID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("ҩƷid")))
                '����
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("������")) & "'"
                '����
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("����")) & "'"
                'Ч��
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("��Ч����")) & "','yyyy-mm-dd HH24:MI:SS')"
                'ʵ������
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("����")))
                '�ɱ���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("����")))
                '�ɱ����
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("���")))
                '����
                strSQL = strSQL & "," & 100
                '���ۼ�
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("���ۼ�")))
                '���۽��
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("���۽��")))
                '���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("���")))
                'ժҪ
                strSQL = strSQL & ",'�����͵���[" & Trim(txt���͵���.Text) & "]����'"
                '������
                strSQL = strSQL & ",'" & gstrUserNameNew & "'"
                '��Ʊ��
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("��Ʊ��")) & "'"
                '��Ʊ����
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("��Ʊ����")) & "','yyyy-mm-dd HH24:MI:SS')"
                '��Ʊ���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("��Ʊ���")))
                '��������
                strSQL = strSQL & "," & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS')"
                '���
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("���")) & "'"
                '��Ʒ�ϸ�֤
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("��Ʒ�ϸ�֤")) & "'"
                '�˲���
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("�˲���")) & "'"
                '�˲�����
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("�˲�����")) & "','yyyy-mm-dd HH24:MI:SS')"
                '����
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("����")))
                '�Ƿ��˻�
                strSQL = strSQL & "," & 1
                '��������
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("��������")) & "','yyyy-mm-dd HH24:MI:SS')"
                '��׼�ĺ�
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("��׼�ĺ�")) & "'"
                '�������
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("�������")) & "'"
                '����
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("����")))
                '�ӳ���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("�ӳ���")))

                strSQL = strSQL & ")"
                    
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = strSQL
            End If
        Next
    
    End With
    
End Sub

Private Sub Saveδ��˷�Ʊ��Ϣ()
    Dim i As Integer
    Dim strSQL As String
    Dim strNO As String
    Dim str��Ӧ��id As String
    Dim strNO�� As String
    Dim strDate As String
    Dim int��� As Integer
    Dim lng�ⷿID As Long
    
    strDate = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    lng�ⷿID = Val(cbo�ⷿ.ItemData(cbo�ⷿ.ListIndex))
    
    With vsfList
        For i = 1 To .Rows - 1
            If .Cell(flexcpForeColor, i, .ColIndex("�к�"), i, .ColIndex("��Ʊ���")) = mNoColColor And Trim(.TextMatrix(i, .ColIndex("�����"))) = "" Then

                If InStr(";" & strNO�� & ";", ";" & .TextMatrix(i, .ColIndex("NO")) & ";") = 0 Then
                    strNO�� = IIf(strNO�� = "", "", strNO�� & ";") & .TextMatrix(i, .ColIndex("NO"))
                    int��� = 0
                    strNO = .TextMatrix(i, .ColIndex("NO"))
                    strSQL = "zl_ҩƷ�⹺_Delete('" & strNO & "')"
                    ReDim Preserve marrSql(UBound(marrSql) + 1)
                    marrSql(UBound(marrSql)) = strSQL
                End If
                
                int��� = int��� + 1
                
                strSQL = "zl_ҩƷ�⹺_INSERT("
                'NO
                strSQL = strSQL & "'" & strNO & "'"
                '���
                strSQL = strSQL & "," & int���
                '�ⷿID
                strSQL = strSQL & "," & lng�ⷿID
                '�Է�����ID
                strSQL = strSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("�Է�����ID"))) = 0, "NULL", Val(.TextMatrix(i, .ColIndex("�Է�����ID"))))
                '��ҩ��λID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("��Ӧ��id")))
                'ҩƷID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("ҩƷid")))
                '����
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("������")) & "'"
                '����
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("����")) & "'"
                'Ч��
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("��Ч����")) & "','yyyy-mm-dd HH24:MI:SS')"
                'ʵ������
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("����")))
                '�ɱ���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("����")))
                '�ɱ����
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("���")))
                '����
                strSQL = strSQL & "," & 100
                '���ۼ�
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("���ۼ�")))
                '���۽��
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("���۽��")))
                '���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("���")))
                'ժҪ
                strSQL = strSQL & ",'�����͵���[" & Trim(txt���͵���.Text) & "]���뷢Ʊ��Ϣ'"
                '������
                strSQL = strSQL & ",'" & gstrUserNameNew & "'"
                '��Ʊ��
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("��Ʊ��")) & "'"
                '��Ʊ����
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("��Ʊ����")) & "','yyyy-mm-dd HH24:MI:SS')"
                '��Ʊ���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("��Ʊ���")))
                '��������
                strSQL = strSQL & "," & "to_date('" & strDate & "','yyyy-mm-dd HH24:MI:SS')"
                '���
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("���")) & "'"
                '��Ʒ�ϸ�֤
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("��Ʒ�ϸ�֤")) & "'"
                '�˲���
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("�˲���")) & "'"
                '�˲�����
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("�˲�����")) & "','yyyy-mm-dd HH24:MI:SS')"
                '����
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("����")))
                '�Ƿ��˻�
                strSQL = strSQL & "," & 1
                '��������
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("��������")) & "','yyyy-mm-dd HH24:MI:SS')"
                '��׼�ĺ�
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("��׼�ĺ�")) & "'"
                '�������
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("�������")) & "'"
                '����
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("����")))
                '�ӳ���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("�ӳ���")))
                '��Ʊ����
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("��Ʊ����")) & "'"
                '�ƻ�id
                strSQL = strSQL & "," & IIf(Val(.TextMatrix(i, .ColIndex("�ƻ�id"))) = 0, "NULL", Val(.TextMatrix(i, .ColIndex("�ƻ�id"))))
                '�������
                strSQL = strSQL & "," & 0
                '���ս���
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("���ս���")) & "'"
                strSQL = strSQL & ")"
                    
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = strSQL
            End If
        Next
    
    End With
End Sub

Private Sub Save����˷�Ʊ��Ϣ()
    Dim i As Integer
    Dim strSQL As String
    
    With vsfList
        For i = 1 To .Rows - 1
            If .Cell(flexcpForeColor, i, .ColIndex("�к�"), i, .ColIndex("��Ʊ���")) = mNoColColor And Trim(.TextMatrix(i, .ColIndex("�����"))) <> "" Then

                strSQL = "zl_ҩƷ�⹺��Ʊ��Ϣ_UPDATE("
                'NO
                strSQL = strSQL & "'" & .TextMatrix(i, .ColIndex("NO")) & "'"
                '���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("���")))
                '��Ʊ��
                strSQL = strSQL & ",'" & .TextMatrix(i, .ColIndex("��Ʊ��")) & "'"
                '��Ʊ����
                strSQL = strSQL & "," & "to_date('" & .TextMatrix(i, .ColIndex("��Ʊ����")) & "','yyyy-mm-dd HH24:MI:SS')"
                '��Ʊ���
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("��Ʊ���")))
                '��ҩ��λID
                strSQL = strSQL & "," & Val(.TextMatrix(i, .ColIndex("��Ӧ��id")))
                '������־
                strSQL = strSQL & "," & 1
                strSQL = strSQL & ")"
                    
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = strSQL
                
            End If
        Next
    
    End With
End Sub


