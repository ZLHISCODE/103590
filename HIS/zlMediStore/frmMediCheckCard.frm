VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMediCheckCard 
   Caption         =   "ҩƷ���ձ༭"
   ClientHeight    =   7530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10830
   Icon            =   "frmMediCheckCard.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   10830
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      Height          =   315
      IMEMode         =   2  'OFF
      Left            =   9360
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   1425
   End
   Begin VB.TextBox txt��ע 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   14
      Top             =   6660
      Width           =   9975
   End
   Begin VB.TextBox txtProvider 
      Enabled         =   0   'False
      Height          =   300
      Left            =   7500
      TabIndex        =   11
      Top             =   660
      Width           =   2895
   End
   Begin VB.CommandButton cmdProvider 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   300
      Left            =   10455
      TabIndex        =   10
      Top             =   660
      Width           =   300
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9720
      TabIndex        =   9
      Top             =   7080
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8400
      TabIndex        =   8
      Top             =   7080
      Width           =   1100
   End
   Begin VB.TextBox txtVerify 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6360
      TabIndex        =   7
      Top             =   6270
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCheck 
      Enabled         =   0   'False
      Height          =   300
      Left            =   840
      TabIndex        =   5
      Top             =   6270
      Width           =   1215
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfBill 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   10695
      _cx             =   18865
      _cy             =   8916
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediCheckCard.frx":030A
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
   Begin VB.ComboBox cboStock 
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   660
      Width           =   1920
   End
   Begin VB.Label LblVerifyDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Left            =   7920
      TabIndex        =   20
      Top             =   6300
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TxtVerifyDate 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   300
      Left            =   8730
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Label LblCheckDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Left            =   2400
      TabIndex        =   18
      Top             =   6300
      Width           =   720
   End
   Begin VB.Label TxtCheckDate 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   300
      Left            =   3180
      TabIndex        =   17
      Top             =   6240
      Width           =   1875
   End
   Begin VB.Label LblNo 
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
      TabIndex        =   16
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl˵�� 
      AutoSize        =   -1  'True
      Caption         =   "��ע"
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   6720
      Width           =   360
   End
   Begin VB.Label LblProvider 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ҩ��λ"
      Height          =   180
      Left            =   6750
      TabIndex        =   12
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblVerify 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   5640
      TabIndex        =   6
      Top             =   6330
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lblCheck 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   6330
      Width           =   540
   End
   Begin VB.Label lblStore 
      AutoSize        =   -1  'True
      Caption         =   "���տⷿ"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   720
   End
   Begin VB.Label LblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ҩƷ���յ�"
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
Attribute VB_Name = "frmMediCheckCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint�༭״̬ As Integer '1-���� 2-�޸� 3-���� 4-�鿴
Private mlng����id As Long  '����id,�޸ĺͲ鿴״̬���и�ֵ
Private mstrMatch As String         'ƥ�䷽ʽ
Private mlng�ⷿID As Long
Private mstr���ս��� As String '��¼Ĭ�����ս���

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��

Public Sub showMe(ByVal int�༭״̬ As Integer, ByVal fraPar As Form, ByVal lng�ⷿID As Long, ByVal lng����id As Long)
    mint�༭״̬ = int�༭״̬
    mlng�ⷿID = lng�ⷿID
    mlng����id = lng����id
    
    Me.Show vbModal, fraPar
End Sub

'�������������
Private Function CheckDepend() As Boolean
    Dim rsDepend As New ADODB.Recordset
    Dim strStock As String, strCaption As String
    
    CheckDepend = False
    On Error GoTo errHandle
    
    '��ȡ�ɲ����Ŀⷿ
    strStock = "HIJKLMN"
    
    '�����ҩƷ���ã����鵱ǰ�����Ƿ������ò��ţ���������ⷿ��ҩ
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = [3] Or a.վ�� is Null) And c.�������� = b.���� " _
            & "  AND Instr([2],b.����,1) > 0 " _
            & "  AND a.id = c.����id " _
            & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
            & " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])"
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�������", UserInfo.�û�ID, strStock, gstrNodeNo)
        
    If rsDepend.EOF Then
        MsgBox "����Ա��ҩƷ������ա�Ȩ�ޣ��������Ա��ϵ��", vbInformation, gstrSysName
        If rsDepend.State = 1 Then rsDepend.Close
        Exit Function
    End If
    
    'װ��ⷿ����
    With cboStock
        .Clear
        Do While Not rsDepend.EOF
            .AddItem rsDepend!����
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = mlng�ⷿID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0
        rsDepend.Close
    End With
    
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboStock_Click()
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        Call SetSelectorRS(1, "ҩƷ������չ���", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0)
    End If
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str�������� As String
    '��ȡ�ɲ����Ŀⷿ
    str�������� = "H,I,J,K,L,M,N"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfBill): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfBill, True)
            Exit Sub
        End If
    End If
    
    If Select����ѡ����(Me, cboStock, Trim(cboStock.Text), str��������, True) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdProvider_Click()
    Dim rsProvider As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo errHandle
    vRect = GetControlRect(txtProvider.hWnd) '��ȡλ��
    dblLeft = vRect.Left
    dblTop = vRect.Top - 700
    
    gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ��Ӧ�� " & _
              "Where (վ�� = [1] Or վ�� is Null) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
              "  And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
              "Start with �ϼ�ID is null connect by prior ID =�ϼ�ID order by level,ID"
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "��ҩ��λ", True, "", "", False, False, _
                        True, dblLeft, dblTop, 1000, blnCancel, False, True, gstrNodeNo)
    If rsProvider Is Nothing Then
        Exit Sub
    Else
        txtProvider.Text = rsProvider!����
        txtProvider.Tag = rsProvider!id
        vsfBill.SetFocus
        vsfBill.Row = 1
        vsfBill.Col = vsfBill.ColIndex("ҩƷ����")
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub CmdSave_Click()
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        If ValidData = False Then Exit Sub
        
        If SaveCard = True Then
            MsgBox "����ɹ���", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
    If mint�༭״̬ = 3 Then
        If mlng����id = 0 Then
            Exit Sub
        End If
        
        If SaveCheck(mlng����id) = True Then
            MsgBox "���˳ɹ���", vbInformation, gstrSysName
            Unload Me
            Exit Sub
        End If
    End If
End Sub

Private Function SaveCheck(ByVal lng����id As Long) As Boolean
    '��˵���
    Dim strVerifyDate As String
    
    On Error GoTo errHandle
    
    strVerifyDate = TxtVerifyDate.Caption
    
    gstrSQL = "Zl_ҩƷ���ռ�¼_Verify("
    '����id_In   In ҩƷ���ռ�¼.Id%Type,
    gstrSQL = gstrSQL & lng����id & ","
    '������_In   In ҩƷ���ռ�¼.������%Type,
    gstrSQL = gstrSQL & "'" & txtVerify.Text & "',"
    '��������_In In ҩƷ���ռ�¼.��������%Type
    gstrSQL = gstrSQL & "to_date('" & strVerifyDate & "','yyyy-mm-dd HH24:MI:SS'))"
    
    Call zlDataBase.ExecuteProcedure(gstrSQL, "SaveCard")
    SaveCheck = True
            
    Exit Function
errHandle:
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
    Dim int�ϸ� As Integer
    Dim blnִ�й��� As Boolean
    Dim arrSql As Variant
    Dim str�������� As String
    Dim strЧ�� As String
    Dim str��ҩ���� As String
        
    arrSql = Array()
        
    On Error GoTo errHandle

    If txtNo.Text = "" Then
        strNo = zlDataBase.GetNextNo(148, Me.cboStock.ItemData(Me.cboStock.ListIndex))
    Else
        strNo = txtNo.Text
    End If
    
    With vsfBill
        If .rows > 1 Then
            If mint�༭״̬ = 2 Then '�޸�
                gstrSQL = "Zl_ҩƷ���ռ�¼_Delete(" & mlng����id & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            Else
                gstrSQL = "Select ҩƷ���ռ�¼_Id.Nextval as id From Dual"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "���ձ���")
                mlng����id = rsTemp!id
            End If
            
            strCheckDate = Format(txtCheckDate.Caption, "yyyy-mm-dd hh:mm:ss")
            '����Ƿ񶼺ϸ��� 1��ʾ���ϸ�0��ʾ�ϸ�
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("ҩƷid")) <> "" Then
                    If .TextMatrix(lngRow, .ColIndex("���ս��")) = "���ϸ�" Then
                        int�ϸ� = 1
                        Exit For
                    End If
                End If
            Next
            '�������
            gstrSQL = "Zl_ҩƷ���ռ�¼_Insert ("
            'Id_In         In ҩƷ���ռ�¼.Id%Type,
            gstrSQL = gstrSQL & mlng����id & ","
            'No_In         In ҩƷ���ռ�¼.No%Type,
            gstrSQL = gstrSQL & "'" & strNo & "',"
            '�ⷿid_In     In ҩƷ���ռ�¼.�ⷿid%Type,
            gstrSQL = gstrSQL & cboStock.ItemData(Me.cboStock.ListIndex) & ","
            '��ҩ��λid_In In ҩƷ���ռ�¼.��ҩ��λid%Type,
            gstrSQL = gstrSQL & Val(txtProvider.Tag) & ","
            '������_In     In ҩƷ���ռ�¼.������%Type,
            gstrSQL = gstrSQL & "'" & txtCheck.Text & "',"
            '��������_In   In ҩƷ���ռ�¼.��������%Type,
            gstrSQL = gstrSQL & "to_date('" & strCheckDate & "','yyyy-mm-dd HH24:MI:SS'),"
            '�Ƿ�ϸ�_In   In ҩƷ���ռ�¼.�Ƿ�ϸ�%Type,
            gstrSQL = gstrSQL & int�ϸ� & ","
            '��ע_in     in ҩƷ���ռ�¼.��ע%type
            gstrSQL = gstrSQL & IIf(txt��ע.Text = "", "NULL", "'" & txt��ע.Text & "'") & ")"
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = gstrSQL
            
            '�α����
            For lngRow = 1 To .rows - 1
                If .TextMatrix(lngRow, .ColIndex("ҩƷid")) <> "" Then
                    str�������� = .TextMatrix(lngRow, .ColIndex("��������"))
                    strЧ�� = .TextMatrix(lngRow, .ColIndex("Ч��"))
                    str��ҩ���� = .TextMatrix(lngRow, .ColIndex("��ҩ����"))
                    
                    gstrSQL = "Zl_ҩƷ������ϸ_Insert ("
                    '����id_In   In ҩƷ������ϸ.����id%Type,
                    gstrSQL = gstrSQL & mlng����id & ","
                    'ҩƷid_In   In ҩƷ������ϸ.ҩƷid%Type,
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("ҩƷid"))) & ","
                    '�ɱ���_In   In ҩƷ������ϸ.�ɱ���%Type,
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("�ɱ���"))) & ","
                    '���ۼ�_In   In ҩƷ������ϸ.���ۼ�%Type,
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) & ","
                    '��ҩ����_In In ҩƷ������ϸ.��ҩ����%Type,
                    gstrSQL = gstrSQL & Val(.TextMatrix(lngRow, .ColIndex("��ҩ����"))) & ","
                    '����_In     In ҩƷ������ϸ.����%Type,
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("ҩƷ����")) & "',"
                    '��������_In In ҩƷ������ϸ.��������%Type,
                    gstrSQL = gstrSQL & IIf(str�������� = "", "NULL", "to_date('" & str�������� & "','yyyy-mm-dd')") & ","
                    'Ч��_In     In ҩƷ������ϸ.Ч��%Type,
                    gstrSQL = gstrSQL & IIf(strЧ�� = "", "NULL", "to_date('" & strЧ�� & "','yyyy-mm-dd')") & ","
                    '����_In     In ҩƷ������ϸ.����%Type,
                    gstrSQL = gstrSQL & IIf(.TextMatrix(lngRow, .ColIndex("����")) = "", "NULL", "'" & .TextMatrix(lngRow, .ColIndex("����")) & "'") & ","
                    '��׼�ĺ�_In In ҩƷ������ϸ.��׼�ĺ�%Type,
                    gstrSQL = gstrSQL & IIf(.TextMatrix(lngRow, .ColIndex("��׼�ĺ�")) = "", "NULL", "'" & .TextMatrix(lngRow, .ColIndex("��׼�ĺ�")) & "'") & ","
                    '��ҩ����_In In ҩƷ������ϸ.��ҩ����%Type,
                    gstrSQL = gstrSQL & IIf(str��ҩ���� = "", "NULL", "to_date('" & str��ҩ���� & "','yyyy-mm-dd')") & ","
                    '�Ƿ�ϸ�_In In ҩƷ������ϸ.�Ƿ�ϸ�%Type
                    gstrSQL = gstrSQL & IIf(.TextMatrix(lngRow, .ColIndex("���ս��")) = "���ϸ�", 1, 0) & ","
                    '���ս���_In In ҩƷ������ϸ.���ս���%Type
                    gstrSQL = gstrSQL & "'" & .TextMatrix(lngRow, .ColIndex("���ս���")) & "')"
                    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = gstrSQL
                End If
            Next
            
            blnִ�й��� = True
            gcnOracle.BeginTrans
            For lngRow = 0 To UBound(arrSql)
                Call zlDataBase.ExecuteProcedure(CStr(arrSql(lngRow)), "SaveCard")
            Next
            gcnOracle.CommitTrans
            SaveCard = True
        Else
            SaveCard = False
            Exit Function
        End If
    End With

    Exit Function
errHandle:
    If blnִ�й��� = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ValidData() As Boolean
    Dim lngRow As Long
    Dim lngCol As Long
    
    '����ʱ���ݼ��
    If Val(txtProvider.Tag) = 0 Then
        MsgBox "��ѡ��һ����Ӧ�̣�", vbInformation, gstrSysName
        txtProvider.SetFocus
        Exit Function
    End If
    
    With vsfBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, .ColIndex("ҩƷid")) <> "" And Val(.TextMatrix(lngRow, .ColIndex("���ۼ�"))) = 0 Then
                MsgBox "��" & lngRow & "���������ۼ۲���Ϊ�㣡", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("���ۼ�")
                .SetFocus
                Exit Function
            End If
            
            If .TextMatrix(lngRow, .ColIndex("ҩƷid")) <> "" And Val(.TextMatrix(lngRow, .ColIndex("��ҩ����"))) = 0 Then
                MsgBox "��" & lngRow & "�����ݽ�ҩ��������Ϊ�㣡", vbInformation, gstrSysName
                .Row = lngRow
                .Col = .ColIndex("��ҩ����")
                .SetFocus
                Exit Function
            End If
        Next
    End With
    ValidData = True
End Function

Private Sub Form_Load()
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    If CheckDepend = False Then Exit Sub
    mstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")  'ƥ�䷽ʽ
    
    Call GetDrugDigit(mlng�ⷿID, "ҩƷ���չ���", 4, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initGrid
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        cboStock.Enabled = True
        txtProvider.Enabled = True
        cmdProvider.Enabled = True
        txtCheck.Text = UserInfo.�û�����
        txt��ע.Enabled = True
                
        If mint�༭״̬ = 1 Then
            txtNo.Text = zlDataBase.GetNextNo(148, Me.cboStock.ItemData(Me.cboStock.ListIndex))
            txtCheckDate.Caption = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        End If
    End If
    If mint�༭״̬ = 3 Then
        txtVerify.Text = UserInfo.�û�����
        TxtVerifyDate.Caption = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    End If
    
    If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
        Call initCard
        If mint�༭״̬ = 3 Then
            CmdSave.Caption = "����(&O)"
        End If
        If mint�༭״̬ = 4 Then
            CmdSave.Visible = False
        End If
    End If
    RestoreWinState Me, App.ProductName, "ҩƷ�������"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With LblTitle
        .Left = 0
        .Top = 120
        .Width = Me.Width
    End With
    
    With txtNo
        .Move Me.Width - .Width - 300
    End With
    LblNo.Move txtNo.Left - LblNo.Width - 100
    
    With lblStore
        .Move 100, 720
    End With
    
    With cboStock
        .Move lblStore.Left + lblStore.Width + 50, lblStore.Top - 60
    End With
    
    With cmdProvider
        .Move Me.Width - cmdProvider.Width - 300, lblStore.Top - 60
    End With
    
    With txtProvider
        .Move cmdProvider.Left - .Width - 10, lblStore.Top - 60
    End With
    
    With LblProvider
        .Move txtProvider.Left - .Width - 100, lblStore.Top
    End With
    
    With vsfBill
        .Move lblStore.Left, lblStore.Top + lblStore.Height + 100, Me.Width - lblStore.Left - 300, Me.Height - .Top - txtVerify.Height - 1500
    End With
    
    lblCheck.Move vsfBill.Left, vsfBill.Top + vsfBill.Height + 200
    txtCheck.Move lblCheck.Left + lblCheck.Width + 100, lblCheck.Top - 60
    lblCheckDate.Move txtCheck.Left + txtCheck.Width + 100, vsfBill.Top + vsfBill.Height + 200
    txtCheckDate.Move lblCheckDate.Left + lblCheckDate.Width + 100, lblCheck.Top - 60
    
    If mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
        lblVerify.Visible = True
        txtVerify.Visible = True
        LblVerifyDate.Visible = True
        TxtVerifyDate.Visible = True
        
        lblVerify.Move txtCheckDate.Left + txtCheckDate.Width + 500, vsfBill.Top + vsfBill.Height + 200
        txtVerify.Move lblVerify.Left + lblVerify.Width + 100, lblCheck.Top - 60
        
        LblVerifyDate.Move txtVerify.Left + txtVerify.Width + 100, lblVerify.Top
        TxtVerifyDate.Move LblVerifyDate.Left + LblVerifyDate.Width + 200, lblVerify.Top - 60
    End If
    
    lbl˵��.Move lblCheck.Left, lblCheck.Top + lblCheck.Height + 100
    txt��ע.Move txtCheck.Left, lbl˵��.Top - 20, vsfBill.Width - lbl˵��.Left - 530
    
    CmdCancel.Move Me.Width - CmdCancel.Width - 300, lbl˵��.Top + lbl˵��.Height + 180
    CmdSave.Move CmdCancel.Left - CmdSave.Width - 200, lbl˵��.Top + lbl˵��.Height + 180
End Sub

Private Sub initGrid()
    '��ʼ�����
    With vsfBill
        .ColComboList(.ColIndex("ҩƷ����")) = "|..."
        .ColComboList(.ColIndex("����")) = "|..."
        .ColComboList(.ColIndex("���ս���")) = "|..."
        .ColDataType(.ColIndex("��ҩ����")) = flexDTDate
        .ColDataType(.ColIndex("��������")) = flexDTDate
        .ColDataType(.ColIndex("Ч��")) = flexDTDate
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, "ҩƷ�������"
    Call ReleaseSelectorRS  'ж�����ݼ�
End Sub

Private Sub txtProvider_GotFocus()
    zlControl.TxtSelAll txtProvider
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsProvider As ADODB.Recordset
    Dim strProviderText As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    vRect = GetControlRect(txtProvider.hWnd) '��ȡλ��
    dblLeft = vRect.Left
    dblTop = vRect.Top - 700
    
    With txtProvider
        If Trim(.Text) = "" Then Exit Sub
        strProviderText = UCase(.Text)
        gstrSQL = "Select id,����,����,���� From ��Ӧ�� " & _
                  "Where (վ�� = [2] Or վ�� is Null) And (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null) " & _
                  "  And ĩ��=1 And (substr(����,1,1)=1 Or Nvl(ĩ��,0)=0) " & _
                  "  And (���� like [1] Or ���� like [1] or ���� like [1] )"
             
        Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "��ҩ��λ", False, "", "", False, False, _
                        True, dblLeft, dblTop, 1000, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & strProviderText & "%", gstrNodeNo)
        
        If blnCancel Then txtProvider.SetFocus: Exit Sub
        
        If rsProvider Is Nothing Then
            MsgBox "δƥ�䵽������Ĺ�ҩ��λ", vbOKOnly + vbInformation, gstrSysName
            txtProvider.SelStart = 0
            txtProvider.SelLength = Len(txtProvider)
            Exit Sub
        Else
            txtProvider.Text = rsProvider!����
            txtProvider.Tag = rsProvider!id
            vsfBill.SetFocus
            vsfBill.Row = 1
            vsfBill.Col = vsfBill.ColIndex("ҩƷ����")
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt��ע_GotFocus()
    zlControl.TxtSelAll txt��ע
End Sub


Private Sub vsfBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim RecReturn As ADODB.Recordset
    Dim i As Integer
    Dim intRow As Integer
    Dim intOldRow As Integer
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo errHandle
    With vsfBill
        intOldRow = vsfBill.Row
        vRect = GetControlRect(vsfBill.hWnd) '��ȡλ��
        dblLeft = vRect.Left + vsfBill.CellLeft
        dblTop = vRect.Top - vsfBill.Height + vsfBill.CellTop + vsfBill.CellHeight
        
        Select Case .ColKey(Col)
            Case "����"
                gstrSQL = "Select ���� as id,����,���� From ҩƷ������ Order By ���� "
                    
                Set RecReturn = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "����", False, "", "", False, False, _
                    True, dblLeft, dblTop, .Height, blnCancel, False, True)
                
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
'                    .TextMatrix(Row, .ColIndex("����id")) = RecReturn!id
                    .TextMatrix(Row, .ColIndex("����")) = RecReturn!����
                End If
            Case "ҩƷ����"
                If grsMaster.State = adStateClosed Then
                    Call SetSelectorRS(1, "ҩƷ������չ���", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0)
                End If
                Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0, True, True, True)
                If RecReturn.RecordCount > 0 Then
                    Set RecReturn = CheckRedo(RecReturn) '����ظ���¼�����ظ��ļ�¼���˵�Ȼ�󷵻ع��˺�����ݼ�
                End If
                
                If RecReturn.RecordCount > 0 Then
                    RecReturn.MoveFirst
                    For i = 1 To RecReturn.RecordCount
                        With vsfBill
                            intRow = .Row
                            SetColValue .Row, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                                RecReturn!ҩƷid, _
                                IIf(IsNull(RecReturn!���), "", RecReturn!���), RecReturn!����, _
                                RecReturn!ҩ�ⵥλ

                            .Col = .ColIndex("��ҩ����")
                                                    
                            If (.TextMatrix(intRow, .ColIndex("ҩƷid")) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, .ColIndex("ҩƷid")) <> "" Then
                                .rows = .rows + 1
                            End If
        
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        End With
                    Next
                    vsfBill.Row = intOldRow
                    RecReturn.Close
                End If
            Case "���ս���"
                gstrSQL = "Select ���� as id, ����, ���� as ���� From ������ս��� Order By ���� "
                    
                Set RecReturn = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "���ս���", False, "", "", False, False, _
                    True, dblLeft, dblTop, .Height, blnCancel, False, True)
                
                If RecReturn Is Nothing Then
                    Exit Sub
                Else
                    .TextMatrix(Row, .ColIndex("���ս���")) = RecReturn!����
                End If
        End Select
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckRedo(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ����ظ��ļ�¼���˵��������ع��˺�����ݼ���
    Dim i As Integer
    Dim strTemp As String
    Dim strҩƷid As String
    Dim str�ظ�ҩ�� As String
    Dim strDub As String
    Dim strsql As String
    
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If InStr(1, strTemp, rsTemp!ҩƷid) = 0 Then
            strTemp = strTemp & rsTemp!ҩƷid & "|"
        End If
        rsTemp.MoveNext
    Loop
    
    With vsfBill
        For i = 1 To .rows - 1
            If InStr(1, strTemp, .TextMatrix(i, .ColIndex("ҩƷid"))) > 0 And .TextMatrix(i, .ColIndex("ҩƷid")) <> "" Then
                strҩƷid = strҩƷid & .TextMatrix(i, .ColIndex("ҩƷid")) & "," & .TextMatrix(i, .ColIndex("ҩƷ����")) & "|"
            End If
        Next
        
        If strҩƷid <> "" Then   'Ϊ��������ƴ��sql
            strDub = ""
            For i = 0 To UBound(Split(strҩƷid, "|")) - 1
                strDub = strDub & "ҩƷid<>" & Split(Split(strҩƷid, "|")(i), ",")(0) & " and "
                If UBound(Split(str�ظ�ҩ��, ",")) <= 2 Then
                    str�ظ�ҩ�� = str�ظ�ҩ�� & Split(Split(strҩƷid, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        
        If str�ظ�ҩ�� <> "" Then
            MsgBox str�ظ�ҩ�� & "�б����Ѿ������ˣ�" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strsql = strDub
        End If
        rsTemp.Filter = strsql
        Set CheckRedo = rsTemp
    End With
End Function

Private Function SetColValue(ByVal intRow As Integer, _
    ByVal strҩƷ���� As String, _
    ByVal strͨ���� As String, _
    ByVal str��Ʒ�� As String, _
    ByVal lngҩƷID As Long, _
    ByVal str��� As String, _
    ByVal str���� As String, _
    ByVal str��λ As String) As Boolean
    Dim strҩ�� As String
    Dim rsTemp As Recordset
    '��ѡ�������ҩƷ��ӵ�vsf���
    '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
    On Error GoTo ErrHand
    With vsfBill
        .TextMatrix(intRow, .ColIndex("���ս��")) = "�ϸ�"
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = strͨ����
        Else
            strҩ�� = IIf(str��Ʒ�� <> "", str��Ʒ��, strͨ����)
        End If
                
        .TextMatrix(intRow, .ColIndex("ҩƷ����")) = strҩƷ���� & strҩ��
        .TextMatrix(intRow, .ColIndex("ҩƷid")) = lngҩƷID
        .TextMatrix(intRow, .ColIndex("���")) = str���
        .TextMatrix(intRow, .ColIndex("����")) = str����
        .TextMatrix(intRow, .ColIndex("��λ")) = str��λ
        
        If mstr���ս��� = "" Then
            gstrSQL = "Select ����  From ������ս��� where ȱʡ��־=1"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "SetColValue")
            
            If Not rsTemp.EOF Then
                .TextMatrix(intRow, .ColIndex("���ս���")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                mstr���ս��� = rsTemp!����
            End If
        Else
            .TextMatrix(intRow, .ColIndex("���ս���")) = mstr���ս���
        End If
    End With
    
    SetColValue = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfBill_DblClick()
    With vsfBill
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            If .Col = .ColIndex("���ս��") And .TextMatrix(.Row, .ColIndex("ҩƷid")) <> "" Then
                If .TextMatrix(.Row, .Col) = "�ϸ�" Then
                    .TextMatrix(.Row, .Col) = "���ϸ�"
                Else
                    .TextMatrix(.Row, .Col) = "�ϸ�"
                End If
            End If
                
            If Not (.Col = .ColIndex("���ս��") Or .Col = .ColIndex("���") Or .Col = .ColIndex("����") Or .Col = .ColIndex("��λ")) Then
                .EditCell
                .EditSelStart = 0
                .EditSelLength = Len(.TextMatrix(.Row, .Col)) * 2
            End If
        End If
    End With
End Sub

Private Sub vsfBill_EnterCell()
    With vsfBill
        If .Col = .ColIndex("���ս��") Or .Col = .ColIndex("���") Or .Col = .ColIndex("����") Or .Col = .ColIndex("��λ") Then
            .Editable = flexEDNone
        Else
            If mint�༭״̬ = 3 Or mint�༭״̬ = 4 Then
                .Editable = flexEDNone
            Else
                .Editable = flexEDKbdMouse
            End If
            If .Col = .ColIndex("ҩƷ����") Or .Col = .ColIndex("����") Or .Col = .ColIndex("���ս���") Then
                .ColComboList(.Col) = ""
            End If
        End If
    End With
End Sub

Private Sub vsfBill_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If MsgBox("��ɾ�����У��Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Shift = 0
        Else
            vsfBill.RemoveItem vsfBill.Row
        End If
    ElseIf KeyCode = vbKeyReturn Then
        With vsfBill
            If .Col <> .Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row = .rows - 1 Then
                    If .TextMatrix(.Row, .ColIndex("ҩƷid")) = "" Then
                        KeyCode = 0
                    Else
                        .rows = .rows + 1
                        .Row = .rows - 1
                    End If
                Else
                    .Row = .Row + 1
                    .Col = 1
                End If
            End If
        End With
    End If
End Sub



Private Sub vsfBill_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With vsfBill
        If .Col = .ColIndex("ҩƷ����") Or .Col = .ColIndex("����") Or .Col = .ColIndex("���ս���") Then
            .ColComboList(.Col) = "|..."
        End If
    End With
End Sub

Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim RecReturn As ADODB.Recordset
    Dim strkey As String
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intOldRow As Integer
    Dim i As Integer
    Dim intRow As Integer
    Dim intPosition As Integer
    
    On Error GoTo errHandle
    With vsfBill
        intOldRow = .Row
        
        strkey = UCase(.EditText)
                
        Select Case .ColKey(Col)
            Case "ҩƷ����"
                If Trim(strkey) = "" Then Exit Sub
                If KeyAscii <> vbKeyReturn Then Exit Sub
                
                vRect = GetControlRect(vsfBill.hWnd) '��ȡλ��
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top - vsfBill.Height + vsfBill.CellTop + vsfBill.CellHeight
                
                dblTop = dblTop + vsfBill.Height
                If strkey <> "" Then
                    If grsMaster.State = adStateClosed Then '��ȡ���ݼ�
                        Call SetSelectorRS(1, "ҩƷ������չ���", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0)
                    End If
                    Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, dblLeft, dblTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), , 0, True, True, True)
                                                        
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckRedo(RecReturn) '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
                    End If
                                                        
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        
                        For i = 1 To RecReturn.RecordCount
                            intRow = .Row
                            If SetColValue(.Row, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                                RecReturn!ҩƷid, _
                                IIf(IsNull(RecReturn!���), "", RecReturn!���), RecReturn!����, RecReturn!ҩ�ⵥλ) = False Then
                                 KeyAscii = 0
                                 Exit Sub
                             End If
                            .EditText = .TextMatrix(.Row, .Col)
                            If (.TextMatrix(intRow, .ColIndex("ҩƷid")) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, .ColIndex("ҩƷid")) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        KeyAscii = 0
                    End If
                End If
            Case "����"
                If Trim(strkey) = "" Then Exit Sub
                If KeyAscii <> vbKeyReturn Then Exit Sub
                vRect = GetControlRect(vsfBill.hWnd) '��ȡλ��
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top - vsfBill.Height + vsfBill.CellTop + vsfBill.CellHeight
                
                gstrSQL = "Select ���� as id,����,����" & _
                            " From ҩƷ������" & _
                            " where ���� Like [1] " & _
                            "       Or ���� Like [2] " & _
                            "       Or ���� Like [2] Order By ���� "
                Set RecReturn = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "������Ŀ", False, "", "", False, False, _
                    True, dblLeft, dblTop, .Height, blnCancel, False, True, strkey & "%", mstrMatch & strkey & "%")
                If RecReturn Is Nothing Then
                    .Text = ""
                    Exit Sub
                Else
                    .Text = RecReturn!����
                    .EditText = RecReturn!����
                End If
            Case "���ս���"
                If Trim(strkey) = "" Then Exit Sub
                If KeyAscii <> vbKeyReturn Then Exit Sub
                vRect = GetControlRect(vsfBill.hWnd) '��ȡλ��
                dblLeft = vRect.Left + vsfBill.CellLeft
                dblTop = vRect.Top - vsfBill.Height + vsfBill.CellTop + vsfBill.CellHeight
                
                gstrSQL = "Select ���� as id, ����, ����" & _
                            " From ������ս���" & _
                            " where ���� Like [1] " & _
                            "       Or ���� Like [2] Order By ���� "
                Set RecReturn = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "���ս���", False, "", "", False, False, _
                    True, dblLeft, dblTop, .Height, blnCancel, False, True, strkey & "%", mstrMatch & strkey & "%")
                If RecReturn Is Nothing Then
                    .Text = ""
                    Exit Sub
                Else
                    .Text = RecReturn!����
                End If
            Case "��ҩ����", "��������", "Ч��"
                If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or (KeyAscii = vbKeyDelete Or KeyAscii = vbKeyReturn Or Chr(KeyAscii) = "-")) Then
                    If KeyAscii <> vbKeyBack Then
                        KeyAscii = 0
                    End If
                End If
            Case "�ɱ���"
                If Not (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then
                    If InStr(1, strkey, ".") > 0 Then
                        If Chr(KeyAscii) = "." Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                        intPosition = InStr(1, strkey, ".") + 1
                        If Len(Mid(strkey, intPosition)) >= mintCostDigit Then
                            If strkey = .TextMatrix(.Row, .Col) Then
                                strkey = Chr(KeyAscii)
                            Else
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or Chr(KeyAscii) = ".") Then
                        If KeyAscii <> vbKeyBack Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        If Val(strkey + Chr(KeyAscii)) > 99999999 Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                End If
                If KeyAscii = vbKeyReturn Then
                    .EditText = zlStr.FormatEx(strkey, mintCostDigit, True, True)
                    .TextMatrix(.Row, .Col) = .EditText
                End If
            Case "���ۼ�"
                If Not (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then
                    If InStr(1, strkey, ".") > 0 Then
                        If Chr(KeyAscii) = "." Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                        intPosition = InStr(1, strkey, ".") + 1
                        If Len(Mid(strkey, intPosition)) >= mintCostDigit Then
                            If strkey = .TextMatrix(.Row, .Col) Then
                                strkey = Chr(KeyAscii)
                            Else
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or Chr(KeyAscii) = ".") Then
                        If KeyAscii <> vbKeyBack Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        If Val(strkey + Chr(KeyAscii)) > 99999999 Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                End If
                If KeyAscii = vbKeyReturn Then
                    .EditText = zlStr.FormatEx(strkey, mintPriceDigit, True, True)
                    .TextMatrix(.Row, .Col) = .EditText
                End If
            Case "��ҩ����"
                If Not (KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then
                    If InStr(1, strkey, ".") > 0 Then
                        If Chr(KeyAscii) = "." Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                        intPosition = InStr(1, strkey, ".") + 1
                        If Len(Mid(strkey, intPosition)) >= mintCostDigit Then
                            If strkey = .TextMatrix(.Row, .Col) Then
                                strkey = Chr(KeyAscii)
                            Else
                                KeyAscii = 0
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If Not ((Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9) Or Chr(KeyAscii) = ".") Then
                        If KeyAscii <> vbKeyBack Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    Else
                        If Val(strkey + Chr(KeyAscii)) > 99999999 Then
                            KeyAscii = 0
                            Exit Sub
                        End If
                    End If
                End If
                If KeyAscii = vbKeyReturn Then
                    .EditText = zlStr.FormatEx(strkey, mintNumberDigit, True, True)
                    .TextMatrix(.Row, .Col) = .EditText
                End If
        End Select
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initCard()
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select f.Id, f.No, f.�ⷿid, f.��ҩ��λid, g.���� As ��Ӧ������, f.������, f.��������, f.������, f.��������, f.��ע, b.����, b.Id As ҩƷid, b.����, b.���," & vbNewLine & _
                "       c.ҩ�ⵥλ, c.ҩ���װ, a.��ҩ����, e.���� As ����, a.�ɱ���, a.���ۼ�, a.��ҩ����, a.����, a.��������, a.Ч��, a.����, a.��׼�ĺ�, a.���ս���," & vbNewLine & _
                "       Nvl(a.�Ƿ�ϸ�, 0) As �Ƿ�ϸ�" & vbNewLine & _
                "From ҩƷ���ռ�¼ F, ҩƷ������ϸ A, �շ���ĿĿ¼ B, ҩƷ��� C, ҩƷ���� D, ҩƷ���� E, ��Ӧ�� G" & vbNewLine & _
                "Where f.Id = a.����id And a.ҩƷid = b.Id And b.Id = c.ҩƷid And c.ҩ��id = d.ҩ��id And d.ҩƷ���� = e.����(+) And f.��ҩ��λid = g.Id(+) And" & vbNewLine & _
                "      a.����id = [1]"

    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "������ϸ��ѯ", mlng����id)
            
    With vsfBill
        Do While Not rsTemp.EOF
            txtNo.Text = rsTemp!NO
            .TextMatrix(.rows - 1, .ColIndex("���ս��")) = IIf(rsTemp!�Ƿ�ϸ� = 0, "�ϸ�", "���ϸ�")
            .TextMatrix(.rows - 1, .ColIndex("ҩƷid")) = rsTemp!ҩƷid
            .TextMatrix(.rows - 1, .ColIndex("ҩƷ����")) = "[" & rsTemp!���� & "]" & rsTemp!����
            .TextMatrix(.rows - 1, .ColIndex("���")) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            .TextMatrix(.rows - 1, .ColIndex("��λ")) = IIf(IsNull(rsTemp!ҩ�ⵥλ), "", rsTemp!ҩ�ⵥλ)
            .TextMatrix(.rows - 1, .ColIndex("��ҩ����")) = IIf(IsNull(rsTemp!��ҩ����), "", Format(rsTemp!��ҩ����, "yyyy-mm-dd hh:mm:ss"))
            .TextMatrix(.rows - 1, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.rows - 1, .ColIndex("�ɱ���")) = IIf(IsNull(rsTemp!�ɱ���), "", zlStr.FormatEx(rsTemp!�ɱ���, mintCostDigit, True, True))
            .TextMatrix(.rows - 1, .ColIndex("���ۼ�")) = IIf(IsNull(rsTemp!���ۼ�), "", zlStr.FormatEx(rsTemp!���ۼ�, mintPriceDigit, True, True))
            .TextMatrix(.rows - 1, .ColIndex("��ҩ����")) = IIf(IsNull(rsTemp!��ҩ����), "", zlStr.FormatEx(rsTemp!��ҩ����, mintNumberDigit, True, True))
            .TextMatrix(.rows - 1, .ColIndex("ҩƷ����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.rows - 1, .ColIndex("��������")) = IIf(IsNull(rsTemp!��������), "", Format(rsTemp!��������, "yyyy-mm-dd"))
            .TextMatrix(.rows - 1, .ColIndex("Ч��")) = IIf(IsNull(rsTemp!Ч��), "", Format(rsTemp!Ч��, "yyyy-mm-dd"))
            .TextMatrix(.rows - 1, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.rows - 1, .ColIndex("��׼�ĺ�")) = IIf(IsNull(rsTemp!��׼�ĺ�), "", rsTemp!��׼�ĺ�)
            .TextMatrix(.rows - 1, .ColIndex("���ս���")) = IIf(IsNull(rsTemp!���ս���), "", rsTemp!���ս���)
            
            txtProvider.Text = rsTemp!��Ӧ������
            txtProvider.Tag = rsTemp!��ҩ��λID
            
            If mint�༭״̬ = 4 Then
                txtCheck.Text = rsTemp!������
            End If
            txtCheckDate.Caption = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
            txt��ע.Text = IIf(IsNull(rsTemp!��ע), "", rsTemp!��ע)
            
            If mint�༭״̬ = 4 Then
                txtVerify.Text = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                If IsNull(rsTemp!��������) = False Then
                    TxtVerifyDate.Caption = Format(rsTemp!��������, "yyyy-mm-dd hh:mm:ss")
                End If
            End If
            
            .rows = .rows + 1
            rsTemp.MoveNext
        Loop
        
        If .rows > 1 Then
            .Row = 1
            .Col = .ColIndex("ҩƷ����")
        End If
    End With
End Sub

Private Sub vsfBill_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsfBill
        If .Col = .ColIndex("ҩƷ����") Or .Col = .ColIndex("����") Or .Col = .ColIndex("���ս���") Then
            .ColComboList(.Col) = "|..."
        Else
            .ColComboList(.ColIndex("ҩƷ����")) = ""
            .ColComboList(.ColIndex("����")) = ""
            .ColComboList(.ColIndex("���ս���")) = ""
        End If
    End With
End Sub

Private Sub vsfBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strkey As String
    
    On Error GoTo errHandle
    With vsfBill
        
        strkey = UCase(Trim(.Text))
'        If strkey = "" Then
            strkey = UCase(.EditText)
'        End If
        
        If Trim(strkey) = "" Then Exit Sub
        
        Select Case .ColKey(Col)
            Case "��ҩ����"
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "�Բ��𣬽�ҩ���ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(.Row, .ColIndex("��ҩ����")) = strkey
                Else
                    If Not IsDate(strkey) Then
                        MsgBox "�Բ��𣬽�ҩ���ڱ���Ϊ��������(2015-10-10) ��20151010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Case "��������"
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "�Բ����������ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(.Row, .ColIndex("��������")) = strkey
                Else
                    If Not IsDate(strkey) Then
                        MsgBox "�Բ����������ڱ���Ϊ��������(2015-10-10) ��20151010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
            Case "Ч��"
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "�Բ���Ч�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(.Row, .ColIndex("Ч��")) = strkey
                Else
                    If Not IsDate(strkey) Then
                        MsgBox "�Բ���Ч�ڱ���Ϊ��������(2015-10-10) ��20151010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                End If
        End Select
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


