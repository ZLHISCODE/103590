VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmEarnSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ����"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmEarnSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsfBill 
      Height          =   930
      Left            =   1200
      TabIndex        =   16
      Top             =   2760
      Width           =   2830
      _cx             =   4992
      _cy             =   1640
      Appearance      =   2
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4530
      TabIndex        =   21
      Top             =   3360
      Width           =   1100
   End
   Begin VB.ComboBox cmb���� 
      Height          =   300
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1920
      Width           =   2025
   End
   Begin VB.ComboBox cmb�վ� 
      Height          =   300
      Left            =   1680
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2320
      Width           =   2025
   End
   Begin VB.CheckBox chk���� 
      Alignment       =   1  'Right Justify
      Caption         =   "����(&G)"
      Height          =   255
      Left            =   450
      TabIndex        =   10
      Top             =   1605
      Width           =   945
   End
   Begin VB.CommandButton cmd�ϼ� 
      Caption         =   "��"
      Height          =   240
      Left            =   2970
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Width           =   255
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   4290
      TabIndex        =   20
      Top             =   -150
      Width           =   30
   End
   Begin VB.CheckBox chkĩ�� 
      Caption         =   "ĩ��(&M)"
      Height          =   225
      Left            =   480
      TabIndex        =   19
      Top             =   4440
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   180
      Index           =   1
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   5
      Tag             =   "����"
      Text            =   "111111"
      Top             =   555
      Width           =   1125
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   7
      Tag             =   "����"
      Top             =   870
      Width           =   2025
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4530
      TabIndex        =   18
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4530
      TabIndex        =   17
      Top             =   150
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   9
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   2025
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1200
      MaxLength       =   10
      TabIndex        =   9
      Tag             =   "����"
      Top             =   1260
      Width           =   1305
   End
   Begin VB.TextBox txtTemp 
      Enabled         =   0   'False
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1200
      MaxLength       =   8
      TabIndex        =   4
      TabStop         =   0   'False
      Tag             =   "����"
      Text            =   "11"
      Top             =   510
      Width           =   1305
   End
   Begin VB.Label lblEdit 
      Caption         =   "��ͬ����  �վݷ�Ŀ(&F)"
      Height          =   900
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   1065
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ȱʡ������Ŀ(&B)"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   1965
      Width           =   1350
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�վݷ�Ŀ(&T)"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   2400
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&U)"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   3
      Top             =   570
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   2
      Left            =   480
      TabIndex        =   6
      Top             =   930
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&S)"
      Height          =   180
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   1290
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�(&D)"
      Height          =   180
      Index           =   9
      Left            =   480
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
End
Attribute VB_Name = "frmEarnSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr�ϼ���ĿID As String     '��ǰ�༭���ϼ���ĿID
Dim mstrID As String         '��ǰ�༭����ĿID

Dim mstr�ϼ����� As String    'ԭʼ���ϼ������ֵ
Dim mstr���� As String        'ԭʼ�ı��������ֵ
Dim mint���� As Integer       '�޸�ǰ�����¼����ڵı�����ĳ���
Dim mintSuccess As Integer
Dim mblnChange As Boolean  '���޸�
Dim mblnҩ�� As Boolean
Dim mstr�վݷ�Ŀ As String

Private Sub cmb����_KeyPress(KeyAscii As Integer)
    '-----------------------------------------------------------------------------------
    '���̶�λ
    '-----------------------------------------------------------------------------------
    Dim intI As Integer
    intI = zlControl.CboMatchIndex(cmb����.hwnd, KeyAscii)
    '���ݹ�������CboSetIndex��λ��ָ������
    Call zlControl.CboSetIndex(cmb����.hwnd, intI)
End Sub

Private Sub cmb�վ�_KeyPress(KeyAscii As Integer)
    '-----------------------------------------------------------------------------------
    '���̶�λ
    '-----------------------------------------------------------------------------------
    Dim intI As Integer
    intI = zlControl.CboMatchIndex(cmb�վ�.hwnd, KeyAscii)
    '���ݹ�������CboSetIndex��λ��ָ������
    Call zlControl.CboSetIndex(cmb�վ�.hwnd, intI)
End Sub

Private Sub cmb�վ�_Validate(Cancel As Boolean)
    With vsfBill
        If .TextMatrix(1, 1) = "" Then .TextMatrix(1, 1) = cmb�վ�.Text
        If .TextMatrix(2, 1) = "" Then .TextMatrix(2, 1) = cmb�վ�.Text
        If .TextMatrix(3, 1) = "" Then .TextMatrix(3, 1) = cmb�վ�.Text
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    If chkĩ��.Value = 1 Then
        ShowHelp App.ProductName, Me.hwnd, "frm������Ŀ����2", Int((glngSys) / 100)
    Else
        ShowHelp App.ProductName, Me.hwnd, "frm������Ŀ����1", Int((glngSys) / 100)
    End If
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save��Ŀ() = False Then Exit Sub
    mintSuccess = mintSuccess + 1
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstrID = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(1).Text = GetMaxLocalCode(mstr�ϼ���ĿID, "������Ŀ")
    cmdOK.Enabled = False
    frmEarnManage.FillList frmEarnManage.tvwMain_S.SelectedItem.Key
    txtEdit(1).SetFocus
    txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ���ĿID, "������Ŀ")
    txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8, txtTemp.MaxLength) - Len(txtTemp.Text)
    mblnChange = False
'    Unload Me
End Sub

Private Function IsValid() As Boolean
'����:��������������Ŀ�������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = 1 To 3
        strTemp = Trim(txtEdit(i).Text)
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox "���������ݲ��ܳ���" & Int(txtEdit(i).MaxLength / 2) & "������" & "��" & txtEdit(i).MaxLength & "����ĸ��", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
            txtEdit(i).SetFocus
            txtEdit(i).SelStart = 0
            txtEdit(i).SelLength = 100
            Exit Function
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)
    
    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(1).Text) = 0 Then
            MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            Exit Function
        End If
    Else
        If Len(txtEdit(1).Text) < txtEdit(1).MaxLength Then
            MsgBox "����ĳ��Ȳ�����", vbExclamation, gstrSysName
            txtEdit(1).SetFocus
            Exit Function
        End If
    End If
    If Not IsNumeric(txtEdit(1).Text) Or InStr(txtEdit(1).Text, ",") > 0 Or InStr(txtEdit(1).Text, ".") > 0 Then
        MsgBox "����Ӧ��������ɡ�", vbExclamation, gstrSysName
        txtEdit(1).SetFocus
        Exit Function
    End If
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(2).Text = ""
        txtEdit(2).SetFocus
        Exit Function
    End If
    If chkĩ��.Value And cmb�վ�.ListIndex < 1 Then
        MsgBox "�վݷ�Ŀ����Ϊ�ա�", vbExclamation, gstrSysName
        cmb�վ�.SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function Save��Ŀ() As Boolean
'����:����༭�����ݵ�������Ŀ����
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim lng������ĿID As Long
    Dim str�վݷ�Ŀ���� As String
    
    On Error GoTo ErrHandle
    
    With vsfBill
        str�վݷ�Ŀ���� = .TextMatrix(1, 1) & "|" & .TextMatrix(2, 1) & "|" & .TextMatrix(3, 1)
    End With
    
    If mstrID = "" Then       '����һ����¼
        lng������ĿID = zlDatabase.GetNextId("������Ŀ")
        gstrSQL = "zl_������Ŀ_insert(" & lng������ĿID & "," & IIF(mstr�ϼ���ĿID = "", "null", mstr�ϼ���ĿID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & txtEdit(3).Text & "'," & chk����.Value & ",'" & cmb�վ�.Text & _
            "','" & cmb����.Text & "'," & chkĩ��.Value & ",'" & str�վݷ�Ŀ���� & "')"
    Else    '�޸�
        gstrSQL = "zl_������Ŀ_update(" & mstrID & "," & IIF(mstr�ϼ���ĿID = "", "null", mstr�ϼ���ĿID) & _
            ",'" & txtTemp.Text & txtEdit(1).Text & "','" & txtEdit(2).Text & _
            "','" & txtEdit(3).Text & "'," & chk����.Value & ",'" & cmb�վ�.Text & _
            "','" & cmb����.Text & "'," & Len(mstr����) + 1 & "," & chkĩ��.Value & ",'" & str�վݷ�Ŀ���� & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Save��Ŀ = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function �༭��Ŀ(ByVal str�ϼ���Ŀ As String, ByVal str�ϼ���ĿID As String, ByVal str�ϼ����� As String, _
    Optional strID As String = "", Optional ByVal blnĩ����Ŀ As Boolean) As Boolean
'����:��������õ�������Ŀ�����ڽ���ͨѶ�ĳ���
'����:str�ϼ���Ŀ     �ϼ�������Ŀ������
'     str�ϼ���ĿID   �ϼ�������Ŀ��ID
'     str�ϼ�����     �ϼ�������Ŀ�ı���
'     strID           ��������Ŀ�ĵ�ID
'     blnĩ����Ŀ     ��������Ŀ�Ƿ�ĩ��
'����ֵ:�༭�ɹ�����True,����ΪFalse
    
    Dim rs������Ŀ As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim rs�վݷ�Ŀ��Ӧ As ADODB.Recordset
    
    Dim i As Integer
    
    mintSuccess = 0
    mstrID = strID
    
    Load frmEarnSet
    
    mblnҩ�� = (glngSys \ 100 = 8)
    
    On Error GoTo ErrHandle
    chkĩ��.Value = 0
    If strID <> "" Then
        rs������Ŀ.CursorLocation = adUseClient
        gstrSQL = "select A.ID,A.����,A.���� from ������Ŀ A,������Ŀ B " & _
                " where A.ID(+)=B.�ϼ�ID and B.ID=[1]"
        Set rs������Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
                
        mstr�ϼ���ĿID = IIF(IsNull(rs������Ŀ("ID")), "", rs������Ŀ("ID"))
        mstr�ϼ����� = IIF(IsNull(rs������Ŀ("����")), "", rs������Ŀ("����"))
        
        txtTemp.Text = mstr�ϼ�����
        txtEdit(9).Text = IIF(IsNull(rs������Ŀ("����")), "��", rs������Ŀ("����"))
        'ȡ���ϼ����룬�������볤�ȵ�ֵ
        txtTemp.MaxLength = GetLocalCodeLength(mstr�ϼ���ĿID, "������Ŀ")
        'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
        
        gstrSQL = "select ID,�ϼ�ID,����,����,����,ĩ��,����,�վݷ�Ŀ,������Ŀ from ������Ŀ  " & _
            "where ID =[1]"
        Set rs������Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
        
        txtEdit(1).Text = Mid(rs������Ŀ("����"), Len(txtTemp.Text) + 1)
        mstr���� = rs������Ŀ("����")
        '��������ӽڵ����ڵ������
        mint���� = GetDownCodeLength(mstrID, "������Ŀ")
        ' 8 - (mint���� - Len(mstr����))�����ʽ����˼��ҪΪ���ĺ��ӵı����������
        txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8 - (mint���� - Len(mstr����)), txtTemp.MaxLength) - Len(mstr�ϼ�����)
        txtEdit(2).Text = rs������Ŀ("����")
        txtEdit(3).Text = IIF(IsNull(rs������Ŀ("����")), "", rs������Ŀ("����"))
        chk����.Value = IIF(rs������Ŀ("����") = 1, 1, 0)
        chkĩ��.Value = IIF(rs������Ŀ("ĩ��") = 1, 1, 0)
        chkĩ��.Enabled = False
    Else
        mstr�ϼ���ĿID = str�ϼ���ĿID
        mstr�ϼ����� = str�ϼ�����
        
        txtTemp.Text = str�ϼ�����
        txtEdit(9).Text = str�ϼ���Ŀ
        'ȡ���ϼ����룬�������볤�ȵ�ֵ
        txtTemp.MaxLength = GetLocalCodeLength(str�ϼ���ĿID, "������Ŀ")
        'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
        '�жϱ����Ƿ�����
        If Len(mstr�ϼ�����) = 8 Then
            MsgBox "�����������Ӽ��ˣ����볤���Ѿ��þ���", vbExclamation, gstrSysName
            mblnChange = False
            Unload frmEarnSet
            Exit Function
        End If
        txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8, txtTemp.MaxLength) - Len(mstr�ϼ�����)
        txtEdit(1).Text = GetMaxLocalCode(str�ϼ���ĿID, "������Ŀ")
        mstr���� = mstr�ϼ����� & txtEdit(1).Text
        If blnĩ����Ŀ Then chkĩ��.Value = 1
        
    End If
    If chkĩ��.Value = 1 Then
        rsTemp.CursorLocation = adUseClient
        
        gstrSQL = "select ���� from �վݷ�Ŀ order By ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        If rsTemp.RecordCount = 0 Then
            MsgBox "������ɡ��վݷ�Ŀ�������á�", vbExclamation, gstrSysName
            �༭��Ŀ = False
            mblnChange = False
            Unload frmEarnSet
            Exit Function
        End If
        cmb�վ�.Clear
        cmb�վ�.AddItem ""
        Do Until rsTemp.EOF
            cmb�վ�.AddItem rsTemp("����")
            mstr�վݷ�Ŀ = IIF(mstr�վݷ�Ŀ = "", rsTemp("����"), mstr�վݷ�Ŀ & "|" & rsTemp("����"))
            rsTemp.MoveNext
        Loop
        cmb�վ�.ListIndex = 0
        rsTemp.Close
        
        If mblnҩ�� = True Then
            lblEdit(6).Visible = False
            cmb����.Visible = False
            txtEdit(9).Top = cmb����.Top
            lblEdit(9).Top = lblEdit(6).Top
            cmd�ϼ�.Top = txtEdit(9).Top + 30
            frmEarnSet.Height = 3000
        Else
            'ҩ��ϵͳ����������Ŀ
            '���˺�:2007/05/17:���ڲ����еĲ�����Ŀ�������¼���ϵ,���ͳһ�����˲�����Ŀ,������Ŀ�Ĳ�����Ŀֻ��ͳ��ĩ��Ϊ1�ļ�¼.
            gstrSQL = "select ���� from ������Ŀ where ĩ��=1"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

            cmb����.Clear
            Do Until rsTemp.EOF
                cmb����.AddItem rsTemp("����")
                rsTemp.MoveNext
            Loop
            If rsTemp.RecordCount = 0 Then
                mblnChange = False
                MsgBox "������ɡ�������Ŀ�������á�", vbExclamation, gstrSysName
                mblnChange = False
                Unload frmEarnSet
                �༭��Ŀ = False
                Exit Function
            End If
            cmb����.ListIndex = 0
            rsTemp.Close
        End If
        
        If mstrID <> "" Then
            On Error Resume Next
            cmb�վ�.Text = rs������Ŀ("�վݷ�Ŀ")
            If Err <> 0 Then
                cmb�վ�.AddItem rs������Ŀ("�վݷ�Ŀ")
                cmb�վ�.Text = rs������Ŀ("�վݷ�Ŀ")
                Err.Clear
            End If
            cmb����.Text = rs������Ŀ("������Ŀ")
            If Err <> 0 Then
                cmb����.AddItem rs������Ŀ("������Ŀ")
                cmb����.Text = rs������Ŀ("������Ŀ")
                Err.Clear
            End If
        End If
        
        '�վݷ�Ŀ��Ӧ
        With vsfBill
            .Rows = 4
            .Cols = 2
            .Editable = flexEDNone
            
            .FixedAlignment(-1) = flexAlignCenterCenter
            .ColAlignment(0) = flexAlignLeftCenter
            .ColAlignment(1) = flexAlignLeftCenter
            .ColWidth(0) = 1200
            .ColWidth(1) = 1600
            
            .TextMatrix(0, 0) = "����"
            .TextMatrix(0, 1) = "�վݷ�Ŀ"
            
            .TextMatrix(1, 0) = "0-����"
            .TextMatrix(2, 0) = "1-סԺ"
            .TextMatrix(3, 0) = "2-���ＰסԺ"
        End With
        
        gstrSQL = "Select ������Ŀid, ����, �վݷ�Ŀ From �վݷ�Ŀ��Ӧ Where ������ĿID = [1]"
        Set rs�վݷ�Ŀ��Ӧ = zlDatabase.OpenSQLRecord(gstrSQL, "�վݷ�Ŀ��Ӧ", Val(strID))
        
        With rs�վݷ�Ŀ��Ӧ
            Do While Not .EOF
                If !���� = 0 Then
                    vsfBill.TextMatrix(1, 1) = !�վݷ�Ŀ
                ElseIf !���� = 1 Then
                    vsfBill.TextMatrix(2, 1) = !�վݷ�Ŀ
                Else
                    vsfBill.TextMatrix(3, 1) = !�վݷ�Ŀ
                End If
                .MoveNext
            Loop
        End With
        
    Else
        lblEdit(5).Visible = False
        lblEdit(6).Visible = False
        lblEdit(7).Visible = False
        chk����.Visible = False
        cmb����.Visible = False
        cmb�վ�.Visible = False
        vsfBill.Visible = False
'        txtEdit(9).Top = chk����.Top
'        lblEdit(9).Top = txtEdit(9).Top + 75
'        cmd�ϼ�.Top = txtEdit(9).Top + 30
        frmEarnSet.Height = 2300
        cmdHelp.Top = txtEdit(3).Top
    End If
    
    frmEarnSet.Caption = IIF(chkĩ��.Value = 1, "������Ŀ����", "�����������")
    
    mblnChange = False
    frmEarnSet.Show vbModal
    �༭��Ŀ = mintSuccess > 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmd�ϼ�_Click()
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim str���� As String
    Dim int����  As Integer
    
    strSQL = "select ID,�ϼ�ID,����,���� from ������Ŀ  " & _
        "where ĩ�� <> 1 start with �ϼ�ID is null connect by prior ID =�ϼ�ID"
    strID = mstr�ϼ���ĿID
    str���� = txtEdit(9).Text
    str���� = txtTemp.Text
    blnRe = frmTreeSel.ShowTree(strSQL, strID, str����, str����, mstrID, "������Ŀ", "����������Ŀ", , mstr����)
    '�ɹ�����
    If blnRe Then
        '�ж��Ƿ����
        If Len(str����) >= Len(mstr����) Then
            If Mid(str����, 1, Len(str����)) = mstr���� Then
                MsgBox "����ϼ������ʣ���Ϊѡ��������������¼���", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        '�µı����Ŀ��
        int���� = GetLocalCodeLength(strID, "������Ŀ")
        'ֻ���޸Ĳ��б�Ҫ���
        If mstrID <> "" Then
            '�䴿�¼�����+�µı�������<=8
            If mint���� - Len(mstr����) + IIF(int���� = 0, Len(str����) + 1, int����) > 8 Then
                MsgBox "����ϼ������ʣ���Ϊ���ı���̫���ˡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        mstr�ϼ���ĿID = strID
        txtEdit(9).Text = str����
        txtTemp.MaxLength = int����
        txtTemp.Text = str����
        If mstrID <> "" Then
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8 - (mint���� - Len(mstr����)), txtTemp.MaxLength) - Len(str����)
            txtEdit(1).Text = GetMaxLocalCode(mstr�ϼ���ĿID, "������Ŀ")
        Else
            txtEdit(1).MaxLength = IIF(txtTemp.MaxLength = 0, 8, txtTemp.MaxLength) - Len(str����)
            txtEdit(1).Text = GetMaxLocalCode(mstr�ϼ���ĿID, "������Ŀ")
        End If
        mblnChange = True
        'txtEdit(1).Text = Mid(txtEdit(1).Text, Len(txtTemp.Text) + 1)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Form_Activate()
    txtEdit(2).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtEdit_Change(Index As Integer)
    cmdOK.Enabled = True
    If Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(txtEdit(2).Text)
    End If
    mblnChange = True
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 2 Then
        zlCommFun.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Then
        'Ҫ���������ƣ����Բ����й��ַ�
        If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Then
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(1).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(1).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

Private Sub vsfBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfBill
        If Row = 0 Then Exit Sub
        If Col <> 1 Then Exit Sub
        
        .ColComboList(1) = mstr�վݷ�Ŀ
    End With
End Sub

Private Sub vsfBill_EnterCell()
    With vsfBill
        .Editable = flexEDKbd
        If .Row = 0 Then Exit Sub
        If .Col = 0 Then Exit Sub
        
        .Editable = flexEDKbdMouse
    End With
End Sub


Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim i As Long
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    For i = 0 To vsfBill.ComboCount - 1
        If zlStr.GetCodeByVB(vsfBill.ComboItem(i)) Like UCase(Chr(KeyAscii)) & "*" Then
            vsfBill.ComboIndex = i: Exit For
        End If
    Next
End Sub
