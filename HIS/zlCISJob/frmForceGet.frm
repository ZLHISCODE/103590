VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.0#0"; "zlIDKind.ocx"
Begin VB.Form frmForceGet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ǿ������"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   Icon            =   "frmForceGet.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Visible         =   0   'False
   Begin VB.Frame fraKind 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   280
      Left            =   960
      TabIndex        =   12
      Top             =   160
      Width           =   2190
      Begin zlIDKind.PatiIdentify PatiIdentify 
         Height          =   270
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   476
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         IDKindStr       =   $"frmForceGet.frx":058A
         BeginProperty IDKindFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowSortName    =   -1  'True
         DefaultCardType =   "���￨"
         IDKindWidth     =   555
         BeginProperty CardNoShowFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "���Ϊ����"
      Height          =   255
      Left            =   2760
      TabIndex        =   11
      Top             =   2303
      Width           =   1215
   End
   Begin VB.ComboBox cbo������� 
      Height          =   300
      Left            =   885
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   2
      Left            =   -120
      TabIndex        =   8
      Top             =   585
      Width           =   7635
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Index           =   1
      Left            =   -120
      TabIndex        =   7
      Top             =   2130
      Width           =   7635
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5205
      TabIndex        =   3
      Top             =   2250
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6300
      TabIndex        =   4
      Top             =   2250
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsRegist 
      Height          =   1080
      Left            =   135
      TabIndex        =   0
      Top             =   945
      Width           =   7200
      _cx             =   12700
      _cy             =   1905
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmForceGet.frx":0651
      ScrollTrack     =   -1  'True
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
   Begin VB.Image imgSentence 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3120
      Picture         =   "frmForceGet.frx":0730
      ToolTipText     =   "ѡ�񱾿�������Ĳ���"
      Top             =   120
      Width           =   360
   End
   Begin VB.Image imgStaKB 
      Height          =   330
      Left            =   3430
      Picture         =   "frmForceGet.frx":0E1A
      ToolTipText     =   "���������Ļ����"
      Top             =   120
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lbl������� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������"
      Height          =   180
      Left            =   90
      TabIndex        =   10
      Top             =   2340
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   0
      Left            =   3240
      TabIndex        =   9
      Top             =   195
      Width           =   90
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Һż�¼��"
      Height          =   180
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   705
      Width           =   900
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "���ﲡ��"
      Height          =   180
      Index           =   2
      Left            =   210
      TabIndex        =   5
      Top             =   195
      Width           =   720
   End
End
Attribute VB_Name = "frmForceGet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String 'IN
Private mstr�Һŵ� As String 'Out
Private mlng����ID As Long
Private mbytSize As Byte
Private Enum COL_REGIST
    COL_NO = 0
    col_���� = 1
    COL_��Ŀ = 2
    COL_ҽ�� = 3
    COL_���� = 4
    COL_ʱ�� = 5
    COL_״̬ = 6
    COL_���� = 7
End Enum
Private mblnStaKB As Boolean '�Ƿ��Զ�������Ļ����
Private mlng�����ID As Long
Private mobjSquare As Object     '���˺� ����:2011-12-25 16:37:31
Private mblnCard As Boolean
Private mobjKeyBoard As Object '��Ļ���̶���̬����
Private mblnˢ���س� As Boolean
Private msinTime As Single
Private mlng�������ID As Long

Public Function ShowMe(frmParent As Object, ByVal strPrivs As String, lng�������ID As Long, _
    ByVal objSquare As Object) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ǿ���������
    '���:objSquare-�����㲿������
    '����:
    '����:
    '����:���˺�
    '����:2011-12-25 16:14:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjSquare = objSquare
    mstrPrivs = strPrivs
    mlng�������ID = lng�������ID
    Me.Show 1, frmParent
    ShowMe = mstr�Һŵ�
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As ADODB.Recordset
    Dim arrSQL As Variant
    Dim strTime As String
    Dim i As Integer
    Dim blnTrans As Boolean
    
    If mlng����ID = 0 Then
        MsgBox "��������Ҫ����Ĳ��ˡ�", vbInformation, gstrSysName
        PatiIdentify.SetFocus: Exit Sub
    End If
    If vsRegist.TextMatrix(1, 0) = "" Then
        MsgBox "�ò���û�п���������ĹҺż�¼��", vbInformation, gstrSysName
        PatiIdentify.SetFocus: Exit Sub
    End If
    If cbo�������.ListIndex = -1 Then
        MsgBox "��ȷ���Բ��˽�������Ŀ��ҡ�", vbInformation, gstrSysName
        cbo�������.SetFocus: Exit Sub
    End If
    On Error GoTo errH
    With vsRegist
        If BillExpend(.TextMatrix(.Row, COL_NO)) Then
            MsgBox "�ò��˹Һ��ѳ�����Ч�����������ٽ���ת�", vbInformation, gstrSysName
            Exit Sub
        End If
        arrSQL = Array()
        If Val(.RowData(.Row)) = 2 Then
            '���˺� ����:2011-12-25 16:37:31
            '��ԤԼ����ǿ������
            If Val(zlDatabase.GetPara("�Һ�ģʽ", glngSys, 9000, 1)) <> 1 And Not mobjSquare Is Nothing Then
                If Not mobjSquare.zlRegisterIncept(Me, p����ҽ��վ, Trim(.TextMatrix(.Row, COL_NO)), cbo�������.Text, IIf(mblnCard, mlng�����ID, 0), "") Then Exit Sub
            Else
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_����ԤԼ�Һ�_����('" & Trim(.TextMatrix(.Row, COL_NO)) & "','" & cbo�������.Text & "')"
            End If
        End If
        '��¼����䶯��¼
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_����䶯��¼_Insert('" & .TextMatrix(.Row, COL_NO) & "',3,'ǿ������','" & UserInfo.���� & "','" & UserInfo.��� & "',NULL," & cbo�������.ItemData(cbo�������.ListIndex) & ",NULL," & UserInfo.ID & ",'" & UserInfo.���� & "')"
         
        'ִ��
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_���˽���(" & mlng����ID & ",'" & .TextMatrix(.Row, COL_NO) & "'," & cbo�������.ItemData(cbo�������.ListIndex) & ",'" & UserInfo.���� & "',Null," & IIf(chk����.Visible And chk����.Value = 1, "1", "0") & ")"
        '��������
        gcnOracle.BeginTrans: blnTrans = True
        For i = LBound(arrSQL) To UBound(arrSQL)
            zlDatabase.ExecuteProcedure CStr(arrSQL(i)), Me.Caption
        Next
        gcnOracle.CommitTrans: blnTrans = False
        
        mstr�Һŵ� = .TextMatrix(.Row, COL_NO)
    End With
    '����ţ�42196
    'Call zlDatabase.SetPara("�����������", cbo�������.ItemData(cbo�������.ListIndex), glngSys, p����ҽ��վ, InStr(mstrPrivs, "��������") > 0)
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If mblnˢ���س� And KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        mblnˢ���س� = False
    End If
End Sub

Private Sub Form_Resize()
    Frame1(1).Width = Me.Width + 100
    Frame1(2).Width = Frame1(1).Width
End Sub

Private Sub imgSentence_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim objCardData As Object
    Dim n As Long
    
    If gint����Һ����� = 0 And gint��ͨ�Һ����� = 0 Then
        n = 1
    Else
        If gint��ͨ�Һ����� - gint����Һ����� > 0 Then
            n = gint��ͨ�Һ�����
        Else
            n = gint����Һ�����
        End If
    End If
    
    vRect = zlControl.GetControlRect(fraKind.hwnd)
    
    blnCancel = True
    On Error GoTo errH
    strSQL = "Select A.����ID as ID,A.�����,A.���� as ����,A.�Ա�,A.����,a.����ʱ�� as �Һ�ʱ�� From ���˹Һż�¼ A,������Ϣ B" & _
    " Where A.����ID=B.����ID" & IIf(mlng�������ID = 0, "", " And A.ִ�в���ID+0=[1]") & _
    " And A.��¼���� <> 2 And A.��¼״̬ = 1 " & _
    " And A.����ʱ�� Between Sysdate-" & n & " And trunc(Sysdate)+1-1/24/60/60 order by A.����ʱ�� desc"

    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ѡ�����7��Ĳ���", False, "", "", False, False, True, _
                vRect.Left, vRect.Top + 50, PatiIdentify.Height, blnCancel, False, True, mlng�������ID)
    If blnCancel = True Then
        MsgBox "δ���ҵ������ҽ��ڵĲ���!", vbInformation, gstrSysName
        Exit Sub
    End If
    If (Not rsTmp Is Nothing) And blnCancel = False Then
        Set objCardData = New zlIDKind.PatiInfor
        mlng����ID = Val(rsTmp!ID & "")
        PatiIdentify.Text = rsTmp!���� & ""
        objCardData.����ID = mlng����ID
        objCardData.���� = rsTmp!���� & ""
        objCardData.����� = rsTmp!����� & ""
        objCardData.�Ա� = rsTmp!�Ա� & ""
        objCardData.���� = rsTmp!���� & ""
        Call SetPati(objCardData)
    Else
        MsgBox "δ���ҵ������ҽ��ڵĲ���!", vbInformation, gstrSysName
        blnCancel = True
        Exit Sub
    End If
    Exit Sub
errH:
    MsgBox "δ���ҵ������ҽ��ڵĲ���!", vbInformation, gstrSysName
    blnCancel = True
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgStaKB_Click()
    On Error Resume Next
    If mobjKeyBoard Is Nothing Then Set mobjKeyBoard = CreateObject("zlScreenKeyboard.clsKeyBoard")
    Call mobjKeyBoard.StartUp
    Call mobjKeyBoard.SetPos
    err.Clear: On Error GoTo 0
End Sub
 
Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim lng��������ID As Long
    Dim strSQL As String
    
    mblnˢ���س� = False
    
    '������Ļ����
    mblnStaKB = Val(zlDatabase.GetPara("������Ļ����", glngSys, p����ҽ��վ)) <> 0
    '��������
    mbytSize = zlDatabase.GetPara("����", glngSys, p����ҽ��վ, "0")
    Call initCardSquareData
    '����ȱʡ���ҷ�ʽ
    On Error Resume Next
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0 Then
        PatiIdentify.objIDKind.IDKind = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "IDKind", 0))
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo errH
    mlng����ID = 0
    mstr�Һŵ� = ""
    '����ţ�42196
    'lng��������ID = Val(zlDatabase.GetPara("�����������", glngSys, p����ҽ��վ, , Array(lbl�������, cbo�������), InStr(mstrPrivs, "��������") > 0))
    'If lng��������ID = 0 Then
        '���ﷶΧ��1=�ұ��˺ŵĲ���,2=�����Ҳ���,3=�����Ҳ���   (�������������������ҽ���ù��û���������ܸÿ��Ҳ��ǵ�ǰҽ���Ŀ���)
        '����ţ�42196��ԭ��10.28���ڷ���кŵĵ������Ѹ�Ϊ������ұ�������ԣ�ǿ������ʱ��ȱʡ���ҿɲ��жϽ��ﷶΧ��ֱ��ȡ���ز����Ľ�����ҡ�
        'If Val(zlDatabase.GetPara("���ﷶΧ", glngSys, p����ҽ��վ, "2")) = 3 Then
            lng��������ID = Val(zlDatabase.GetPara("�������", glngSys, p����ҽ��վ))
        'End If
    'End If
    
    'ȷ��ȱʡ���������
    On Error GoTo errH
    strSQL = "Select Distinct A.ID,A.����,B.ȱʡ" & _
        " From ���ű� A,������Ա B,��������˵�� C" & _
        " Where A.ID=B.����ID And A.ID=C.����ID And C.��������||''='�ٴ�' And C.������� IN(1,3)" & _
        " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) And B.��ԱID=[1]" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    Do While Not rsTmp.EOF
        cbo�������.AddItem rsTmp!����
        cbo�������.ItemData(cbo�������.NewIndex) = rsTmp!ID
                
        If rsTmp!ID = lng��������ID Then
            cbo�������.ListIndex = cbo�������.NewIndex
        
        ElseIf Nvl(rsTmp!ȱʡ, 0) = 1 And cbo�������.ListIndex = -1 Then
            cbo�������.ListIndex = cbo�������.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If cbo�������.ListIndex = -1 And cbo�������.ListCount > 0 Then cbo�������.ListIndex = 0
    
    If mblnStaKB Then
        On Error Resume Next
        Set mobjKeyBoard = Nothing
        Set mobjKeyBoard = CreateObject("zlScreenKeyboard.clsKeyBoard")
        err.Clear: On Error GoTo 0
        If Not mobjKeyBoard Is Nothing Then
            imgStaKB.Visible = True
            Call mobjKeyBoard.StartUp
        Else
            MsgBox "��Ļ���̲���δ����ȷ��װ������ʹ�ã�", vbInformation, gstrSysName
        End If
    End If
    Call SetFontSize(mbytSize)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Call PatiIdentify.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjKeyBoard = Nothing
    Set mobjSquare = Nothing
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) <> 0 Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "IDKind", PatiIdentify.objIDKind.IDKind)
    End If
End Sub

Private Sub PatiIdentify_FindPatiArfter(ByVal objCard As zlIDKind.Card, ByVal blnCard As Boolean, ShowName As String, objHisPati As zlIDKind.PatiInfor, objCardData As zlIDKind.PatiInfor, strErrMsg As String, blnCancel As Boolean)
    
    If objHisPati Is Nothing Then
        MsgBox "û���ҵ���صĲ�����Ϣ", vbInformation, gstrSysName
        Exit Sub
    End If
    If objHisPati.����ID = 0 Then
        MsgBox "û���ҵ���صĲ�����Ϣ", vbInformation, gstrSysName
        Exit Sub
    End If
    Call SetPati(objHisPati)
End Sub

Private Sub SetPati(ByVal objHisPati As zlIDKind.PatiInfor)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNO As String
    Dim str����IDs As String, i As Long
    Dim strMsg As String, blnDo As Boolean
    
    lblInfo(1).Caption = "�����:" & objHisPati.����� & "  �Ա�:" & objHisPati.�Ա� & "  ����:" & objHisPati.����
    'Ϊ����ȷ��ʾ����ͨ��SQL��ͨ������������
    strSQL = "Select A.NO,A.��¼����,D.ID as ����ID,D.���� as ����," & _
        " C.ID as ��ĿID,C.���� as ��Ŀ,A.ִ����,A.����,A.����ʱ��,A.ִ��״̬,Decode(A.����,1,'��','��') as ����" & _
        " From ���˹Һż�¼ A,������ü�¼ B,�շ���ĿĿ¼ C,���ű� D" & _
        " Where A.NO=B.NO And B.��¼����=4 And B.��¼״̬ in (1,0) And B.�շ����='1' And a.��¼���� in (1,2) And a.��¼״̬ =1" & _
        "           And B.�۸񸸺� is Null And B.�������� is Null And B.�շ�ϸĿID=C.ID And A.ִ�в���ID=D.ID" & _
        "           And A.����ʱ��<=trunc(Sysdate)+1-1/24/60/60  And A.����ID=[1]" & _
        IIf(Val(zlDatabase.GetPara(210, glngSys)) = 1, "", " And A.����ʱ�� Between Sysdate - Decode(A.����,1," & IIf(gint����Һ����� = 0, 1, gint����Һ�����) & "," & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����) & ") And trunc(Sysdate)+1-1/24/60/60") & _
        " Order by ����ʱ�� Desc "
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, objHisPati.����ID)
    With vsRegist
        If Not rsTmp.EOF Then
            .Rows = .FixedRows
            str����IDs = GetUser����IDs
            mlng����ID = objHisPati.����ID
            Do While Not rsTmp.EOF
                blnDo = True
                If Nvl(rsTmp!ִ����) = UserInfo.���� Then
                    strMsg = strMsg & vbCrLf & "�Һż�¼" & rsTmp!NO & "��ҽ���Ǳ��˵�" & Decode(Nvl(rsTmp!ִ��״̬, 0), 0, "����", 1, "����", 2, "���ھ����") & "�ţ�����ʹ�����﹦�ܡ�"
                    blnDo = False 'ҽ���������������ĺţ�����ͨ�����﹦�ܡ�
                End If
                If blnDo Then
                    '������������ڵ�ǰҽ���������ҹҵĺţ�����״̬������
                    'ȫԺ����������κο��ҹҵĺţ����Ʋ��������ھ���
                    If InStr("," & str����IDs & ",", "," & rsTmp!����ID & ",") = 0 Then
                        If InStr(mstrPrivs, "ȫԺ��������") = 0 Then
                            strMsg = strMsg & vbCrLf & "�Һż�¼" & rsTmp!NO & "���Һſ���Ϊ""" & rsTmp!���� & """�����Ǳ��ƹҺţ�û��Ȩ�޽������"
                            blnDo = False
                        ElseIf Nvl(rsTmp!ִ��״̬, 0) = 2 And InStr(GetInsidePrivs(p����ҽ��վ), ";����ǿ���������ھ���Ĳ���;") = 0 Then
                            strMsg = strMsg & vbCrLf & "�Һż�¼" & rsTmp!NO & "��""" & rsTmp!���� & """��ҽ��""" & rsTmp!ִ���� & """���ھ�����ܽ������"
                            blnDo = False
                        End If
                    End If
                End If
                If blnDo Then
                    .AddItem "": i = .Rows - 1
                    .TextMatrix(i, COL_NO) = rsTmp!NO
                    .RowData(i) = Val(Nvl(rsTmp!��¼����))
                    .TextMatrix(i, col_����) = rsTmp!����
                    .Cell(flexcpData, i, col_����) = Val(rsTmp!����ID)
                    .TextMatrix(i, COL_��Ŀ) = rsTmp!��Ŀ
                    .Cell(flexcpData, i, COL_��Ŀ) = Val(rsTmp!��ĿID)
                    .TextMatrix(i, COL_ҽ��) = Nvl(rsTmp!ִ����)
                    .TextMatrix(i, COL_����) = Nvl(rsTmp!����)
                    .TextMatrix(i, COL_ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                    .TextMatrix(i, COL_״̬) = Decode(Nvl(rsTmp!ִ��״̬, 0), 0, "����", 1, "����", 2, "���ھ���")
                    .TextMatrix(i, COL_����) = rsTmp!����
                End If
                rsTmp.MoveNext
            Loop
            If .Rows = .FixedRows Then
                .Rows = .FixedRows + 1
                strMsg = "����""" & objHisPati.���� & """û�п�������ĹҺż�¼��" & vbCrLf & strMsg
            Else
                strMsg = ""
            End If
        Else
            .Rows = .FixedRows
            .Rows = .FixedRows + 1
            strMsg = "����""" & objHisPati.���� & """�ڹҺ���Ч������û�йҺż�¼��"
        End If
        .Row = .FixedRows
    End With
    
    If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
    
    If blnDo Then
        mblnˢ���س� = True
        msinTime = Timer
        Do
            If (Timer - msinTime) > 0.25 Then Exit Do
            If mblnˢ���س� Then
                DoEvents
            Else
                Exit Do
            End If
        Loop
        mblnˢ���س� = False
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub PatiIdentify_FindPatiBefore(ByVal objCard As zlIDKind.Card, blnCard As Boolean, strShowText As String, objCardData As zlIDKind.PatiInfor, blnFindPatied As Boolean, blnCancel As Boolean)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strNO As String
    Dim str����IDs As String, i As Long
    Dim strMsg As String, blnDo As Boolean
    Dim lng�����ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String
    Dim vRect As RECT
    Dim blnLikeCode As Boolean '���ݼ������
    
    On Error GoTo errH
    If strShowText = "" Then blnCancel = True: Exit Sub
    If zlCommFun.IsCharAlpha(strShowText) Then
        blnLikeCode = True
        strWhere = " Instr(',' || Zlpinyincode(����), ',' || [1]) > 0"
        strWhere = strWhere & "And ��¼���� <> 2 And ��¼״̬ = 1 And ����ʱ�� Between Sysdate - Decode(����,1," & IIf(gint����Һ����� = 0, 1, gint����Һ�����) & "," & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����) & ") And trunc(Sysdate)+1-1/24/60/60"
        strShowText = UCase(strShowText)
    ElseIf Left(strShowText, 1) = "." Then '�Һŵ�
        strNO = GetFullNO(Mid(UCase(strShowText), 2), 12)
        strSQL = "Select ����ID,�����,����,�Ա�,���� From ���˹Һż�¼ Where NO=[2] And ��¼����=1 "
    Else
        Select Case objCard.����
            Case "����"
                strSQL = "Select A.����ID,A.�����,A.����,A.�Ա�,A.���� From ���˹Һż�¼ A,������Ϣ B Where A.����ID=B.����ID And A.��¼���� <> 2 And A.��¼״̬ = 1 And b.����=[1]" & _
                            IIf(Val(zlDatabase.GetPara(210, glngSys)) = 1, "", " And a.����ʱ�� Between Sysdate - Decode(A.����,1," & IIf(gint����Һ����� = 0, 1, gint����Һ�����) & "," & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����) & ") And trunc(Sysdate)+1-1/24/60/60")
                strSQL = strSQL & " Order by A.����ʱ�� desc"
            Case "�Һŵ���"
                strNO = GetFullNO(UCase(strShowText), 12)
                strSQL = "Select ����ID,�����,����,�Ա�,���� From ���˹Һż�¼ Where NO=[2] And ��¼����=1  "
        End Select
    End If
    If strWhere <> "" Then
        strSQL = "Select /*+ RULE */ ����ID,�����,����,�Ա�,���� From ���˹Һż�¼ Where " & strWhere
    End If
    If strSQL = "" Then Exit Sub
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strShowText, strNO)

    If rsTmp.EOF = False Then
        If rsTmp.RecordCount > 1 Then
            vRect = zlControl.GetControlRect(fraKind.hwnd)
            If blnLikeCode Then
                strSQL = "Select /*+ RULE */ ����ID as ID,�����,���� as ����,�Ա�,����,����ʱ�� as �Һ�ʱ�� From ���˹Һż�¼ Where " & strWhere & " order by ����ʱ�� desc"
            Else
                strSQL = "Select A.����ID as ID,A.�����,A.���� as ����,A.�Ա�,A.����,a.����ʱ�� as �Һ�ʱ�� From ���˹Һż�¼ A,������Ϣ B Where A.����ID=B.����ID And B.����=[1] And A.��¼���� <> 2 And A.��¼״̬ = 1" & IIf(Val(zlDatabase.GetPara(210, glngSys)) = 1, "", " And a.����ʱ�� Between Sysdate - Decode(A.����,1," & IIf(gint����Һ����� = 0, 1, gint����Һ�����) & "," & IIf(gint��ͨ�Һ����� = 0, 1, gint��ͨ�Һ�����) & ") And trunc(Sysdate)+1-1/24/60/60") & " order by a.����ʱ�� desc"
            End If
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "ѡ����", False, "", "", False, False, True, _
                        vRect.Left, vRect.Top + 50, PatiIdentify.Height, blnCancel, False, True, strShowText, strNO)
            
            If Not rsTmp Is Nothing And blnCancel = False Then
                mlng����ID = Val(rsTmp!ID & "")
                strShowText = rsTmp!���� & ""
                Set objCardData = New zlIDKind.PatiInfor
                objCardData.����ID = mlng����ID
                objCardData.���� = rsTmp!���� & ""
                objCardData.����� = rsTmp!����� & ""
                objCardData.�Ա� = rsTmp!�Ա� & ""
                objCardData.���� = rsTmp!���� & ""
                Call SetPati(objCardData)
                blnFindPatied = True
            Else
                blnCancel = True
                Exit Sub
            End If
        Else
            mlng����ID = Val(rsTmp!����ID & "")
            strShowText = rsTmp!���� & ""
            Set objCardData = New zlIDKind.PatiInfor
            objCardData.����ID = mlng����ID
            objCardData.���� = rsTmp!���� & ""
            objCardData.����� = rsTmp!����� & ""
            objCardData.�Ա� = rsTmp!�Ա� & ""
            objCardData.���� = rsTmp!���� & ""
            Call SetPati(objCardData)
            blnFindPatied = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub PatiIdentify_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    mlng�����ID = objCard.�ӿ����
End Sub


Private Sub vsRegist_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    chk����.Visible = (vsRegist.TextMatrix(NewRow, COL_����) = "��")
End Sub

Private Sub vsRegist_GotFocus()
    vsRegist.BackColorSel = &HFFEBD7
    vsRegist.ForeColorSel = &H0&
End Sub

Private Sub vsRegist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsRegist.TextMatrix(1, 0) <> "" And cbo�������.ListCount = 1 Then
            Call cmdOK_Click
        End If
    End If
End Sub

Private Sub vsRegist_LostFocus()
    vsRegist.BackColorSel = &HC0C0C0
    vsRegist.ForeColorSel = &H0&
End Sub

Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ���㿨����������Ϣ
    '���:blnClosed:�رն���
    '����:���˺�
    '����:2010-01-05 14:51:23
    '����:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjSquare Is Nothing Then Exit Sub
    Call PatiIdentify.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjSquare, , "")
    PatiIdentify.objIDKind.AllowAutoICCard = True
    PatiIdentify.objIDKind.AllowAutoIDCard = True
End Sub

Private Sub SetFontSize(ByVal bytSize As Byte)
'���ܣ����н��������ͳһ����
'������bytSize  0-9�����壬1-12������
    
    Me.Width = IIf(bytSize = 0, 7200, 9500)
    Me.Height = IIf(bytSize = 0, 3100, 4100)
    Call zlControl.SetPubFontSize(Me, bytSize)
    vsRegist.Height = 5 * vsRegist.RowHeight(0)
    Call SetCtlPos
    vsRegist.Width = Me.Width - 2 * vsRegist.Left
End Sub

Private Sub SetCtlPos()
'���ܣ����ý���ؼ�λ��
    Dim lngDis1 As Long, lngDis2 As Long
    lngDis1 = 30: lngDis2 = 120
    
    Call zlControl.SetPubCtrlPos(False, 0, lblInfo(2), lngDis1, fraKind, 0, imgSentence, lngDis1, imgStaKB)
    imgSentence.Top = imgSentence.Top - 10
    Call zlControl.SetPubCtrlPos(True, -1, lblInfo(2), 120 + Frame1(1).Height + 90, lblInfo(1), 30, vsRegist, 90 + Frame1(2).Height + 180, lbl�������)
    Frame1(1).Top = lblInfo(2).Top + lblInfo(2).Height + 120
    Frame1(2).Top = vsRegist.Top + vsRegist.Height + 90
    Call zlControl.SetPubCtrlPos(False, 0, lbl�������, lngDis1, cbo�������, lngDis2, chk����, lngDis2, cmdOK, lngDis2, cmdCancel)
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - lngDis2
    cmdOK.Left = cmdCancel.Left - 60 - cmdOK.Width
End Sub
