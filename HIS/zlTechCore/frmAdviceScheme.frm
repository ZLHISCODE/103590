VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdviceScheme 
   AutoRedraw      =   -1  'True
   Caption         =   "����Ϊ����ҽ��"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   Icon            =   "frmAdviceScheme.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8955
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   8940
      Begin VB.OptionButton opt���� 
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   1
         Left            =   6570
         TabIndex        =   19
         Top             =   1785
         Width           =   930
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "˽��(&P)"
         Height          =   180
         Index           =   0
         Left            =   5520
         TabIndex        =   18
         Top             =   1785
         Value           =   -1  'True
         Width           =   930
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   255
         Left            =   3030
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "�� * ��ѡ�����з���"
         Top             =   690
         Width           =   285
      End
      Begin VB.Frame fraLine 
         Height          =   60
         Left            =   -60
         TabIndex        =   28
         Top             =   510
         Width           =   9510
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         Height          =   255
         Left            =   7125
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "�� * ��ѡ��"
         Top             =   690
         Width           =   285
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   4350
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   4
         Top             =   660
         Width           =   3090
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Index           =   0
         Left            =   1095
         MaxLength       =   60
         TabIndex        =   7
         Top             =   1005
         Width           =   2250
      End
      Begin VB.TextBox txtƴ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   4350
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1005
         Width           =   960
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   5970
         MaxLength       =   12
         TabIndex        =   10
         Top             =   1005
         Width           =   960
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   5970
         MaxLength       =   12
         TabIndex        =   15
         Top             =   1350
         Width           =   960
      End
      Begin VB.TextBox txtƴ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   4350
         MaxLength       =   12
         TabIndex        =   14
         Top             =   1350
         Width           =   960
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Index           =   1
         Left            =   1095
         MaxLength       =   40
         TabIndex        =   12
         Top             =   1350
         Width           =   2250
      End
      Begin VB.TextBox txt˵�� 
         Height          =   300
         Left            =   1095
         MaxLength       =   60
         TabIndex        =   17
         Top             =   1695
         Width           =   4215
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   20
         TabIndex        =   1
         Top             =   660
         Width           =   2250
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3690
         TabIndex        =   3
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   405
         TabIndex        =   0
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   405
         TabIndex        =   6
         Top             =   1065
         Width           =   630
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&S)           (ƴ��)            (���)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3690
         TabIndex        =   8
         Top             =   1065
         Width           =   3780
      End
      Begin VB.Label lblnote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdviceScheme.frx":058A
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   1095
         TabIndex        =   29
         Top             =   75
         Width           =   6555
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&M)           (ƴ��)            (���)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3690
         TabIndex        =   13
         Top             =   1410
         Width           =   3780
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   405
         TabIndex        =   11
         Top             =   1410
         Width           =   630
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   435
         Picture         =   "frmAdviceScheme.frx":061C
         Top             =   0
         Width           =   480
      End
      Begin VB.Label lbl˵�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "˵��(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   405
         TabIndex        =   16
         Top             =   1740
         Width           =   630
      End
   End
   Begin VB.Frame fraCommand 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      TabIndex        =   26
      Top             =   7005
      Width           =   9390
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   5850
         TabIndex        =   21
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6960
         TabIndex        =   22
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         Height          =   350
         Left            =   405
         TabIndex        =   25
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   1575
         TabIndex        =   23
         ToolTipText     =   "Ctrl+A"
         Top             =   135
         Width           =   1100
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "ȫ��(&R)"
         Height          =   350
         Left            =   2685
         TabIndex        =   24
         ToolTipText     =   "Ctrl+R"
         Top             =   135
         Width           =   1100
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4920
      Left            =   0
      TabIndex        =   20
      Top             =   2085
      Width           =   8955
      _cx             =   15796
      _cy             =   8678
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceScheme.frx":0EE6
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      FrozenCols      =   1
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmAdviceScheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mint��Դ As Integer 'IN:1-����,2-סԺ
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mstr�Һŵ� As String
Private mintӤ�� As Integer
Private mblnOK As Boolean
Private Enum COL���׷���
    colѡ�� = 0
    col��Ч = 1
    col���� = 2
    col���� = 3
    col������λ = 4
    col���� = 5
    col������λ = 6
    colƵ�� = 7
    col�÷� = 8
    col���� = 9
    colִ��ʱ�� = 10
    colִ�п��� = 11
    colִ������ = 12
    col��� = 13
    col������ = 14
    col������� = 15
    col������ĿID = 16
    col�շ�ϸĿID = 17
    col�걾��λ = 18
    colƵ�ʴ��� = 19
    colƵ�ʼ�� = 20
    col�����λ = 21
End Enum

Public Function ShowMe(ByVal strPrivs As String, ByVal int��Դ As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal str�Һŵ� As String, ByVal intӤ�� As Integer, frmParent As Object) As Boolean
    
    mstrPrivs = strPrivs
    mint��Դ = int��Դ
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mstr�Һŵ� = str�Һŵ�
    mintӤ�� = intӤ��
    mblnOK = False
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdAll_Click()
    Call Form_KeyDown(vbKeyA, vbCtrlMask)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Call Form_KeyDown(vbKeyR, vbCtrlMask)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim arrSQL() As Variant
    Dim colSerial As New Collection, lng����ID As Long
    Dim i As Long, j As Long
    
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt����.Text) > txt����.MaxLength Then
        MsgBox "����ĳ��������" & txt����.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt����.SetFocus: Exit Sub
    End If
    
    If Trim(Me.txt����(0).Text) = "" Then
        MsgBox "���������ƣ�", vbInformation, gstrSysName
        Me.txt����(0).SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt����(0).Text) > txt����(0).MaxLength Then
        MsgBox "���Ƴ�����" & txt����(0).MaxLength & "���ַ���" & txt����(0).MaxLength \ 2 & "�����֣���", vbInformation, gstrSysName
        Me.txt����(0).SetFocus: Exit Sub
    End If
    
    If Val(txt����.Tag) = 0 Then
        MsgBox "��Ϊ�ó��׷���ȷ��һ�����ࡣ", vbInformation, gstrSysName
        txt����.SetFocus: Exit Sub
    End If
    
    If zlCommFun.ActualLen(txt����(1).Text) > txt����(1).MaxLength Then
        MsgBox "����������" & txt����(1).MaxLength & "���ַ���" & txt����(1).MaxLength \ 2 & "�����֣���", vbInformation, gstrSysName
        Me.txt����(1).SetFocus: Exit Sub
    End If
    If zlCommFun.ActualLen(txt˵��.Text) > txt˵��.MaxLength Then
        MsgBox "˵��������" & txt˵��.MaxLength & "���ַ���" & txt˵��.MaxLength \ 2 & "�����֣���", vbInformation, gstrSysName
        Me.txt˵��.SetFocus: Exit Sub
    End If
    
    If Val(vsAdvice.TextMatrix(vsAdvice.FixedRows, col������ĿID)) = 0 Then
        MsgBox "û�п��Ա���Ϊ���׷�����ҽ����", vbInformation, gstrSysName
        vsAdvice.SetFocus: Exit Sub
    End If
    
    '���ݱ���
    If Val(txt����.Tag) = 0 Then
        lng����ID = zlDatabase.GetNextId("������ĿĿ¼")
        If zlClinicCodeRepeat(Trim(Me.txt����.Text)) Then Exit Sub
    Else
        lng����ID = Val(txt����.Tag)
        If zlClinicCodeRepeat(Trim(Me.txt����.Text), lng����ID) Then Exit Sub
    End If
    
    arrSQL = Array()
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "ZL_���׷�����Ŀ_Update(" & _
        lng����ID & "," & Val(Me.txt����.Tag) & ",'" & Trim(Me.txt����.Text) & "'," & _
        "'" & Trim(Me.txt����(0).Text) & "','" & Trim(Me.txtƴ��(0).Text) & "','" & Trim(Me.txt���(0).Text) & "'," & _
        "'" & Trim(Me.txt����(1).Text) & "','" & Trim(Me.txtƴ��(1).Text) & "','" & Trim(Me.txt���(1).Text) & "'," & _
        "'" & Trim(Me.txt˵��.Text) & "'," & IIF(opt����(0).Value, UserInfo.ID, "Null") & ")"
    With vsAdvice
        '��¼ԭ����ID�����������
        j = 1
        colSerial.Add 0, "_0"
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col������ĿID)) <> 0 And Val(.TextMatrix(i, colѡ��)) <> 0 Then
                colSerial.Add j, "_" & Val(.TextMatrix(i, col���))
                j = j + 1
            End If
        Next
        
        j = 1
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col������ĿID)) <> 0 And Val(.TextMatrix(i, colѡ��)) <> 0 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_���׷�������_Insert(" & _
                    lng����ID & "," & j & "," & ZVal(colSerial("_" & Val(.TextMatrix(i, col������)))) & "," & _
                    IIF(.TextMatrix(i, col��Ч) = "����", 0, 1) & "," & Val(.TextMatrix(i, col������ĿID)) & "," & _
                    ZVal(Val(.TextMatrix(i, col����))) & "," & ZVal(Val(.TextMatrix(i, col����))) & "," & _
                    ZVal(Val(.TextMatrix(i, col�շ�ϸĿID))) & ",'" & .TextMatrix(i, col�걾��λ) & "'," & _
                    "'" & .TextMatrix(i, colƵ��) & "'," & ZVal(.TextMatrix(i, colƵ�ʴ���)) & "," & _
                    ZVal(.TextMatrix(i, colƵ�ʼ��)) & ",'" & .TextMatrix(i, col�����λ) & "'," & _
                    "'" & .TextMatrix(i, col����) & "'," & Val(.Cell(flexcpData, i, colִ������)) & "," & _
                    ZVal(Val(.Cell(flexcpData, i, colִ�п���))) & ",'" & .TextMatrix(i, colִ��ʱ��) & "')"
                j = j + 1
            End If
        Next
    End With

    If UBound(arrSQL) = 0 Then
        MsgBox "û��ѡ��Ҫ����Ϊ���׷�����ҽ����", vbInformation, gstrSysName
        vsAdvice.SetFocus: Exit Sub
    End If
    
    '�ύSQL���
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    On Error GoTo 0
        
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim objTmp As Object
    
    strSQL = _
        " Select ID,�ϼ�ID,0 as ĩ��,����,����,NULL as ˵��" & _
        " From ���Ʒ���Ŀ¼ Where ����=6" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Union ALL " & _
        " Select ID,����ID as �ϼ�ID,1 as ĩ��,����,����,�걾��λ as ˵��" & _
        " From ������ĿĿ¼ Where ���='9'"
        If InStr(mstrPrivs, "������׷���") > 0 Then
            strSQL = strSQL & " And (��ԱID is Null Or ��ԱID=[1])"
        Else
            strSQL = strSQL & " And ��ԱID=[1]"
        End If
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "���׷���", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, UserInfo.ID)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "��ǰ��û���������׷�������ѡ��", vbInformation, gstrSysName
        End If
        txt����.SetFocus
    Else
        txt����.Tag = rsTmp!ID
        txt����.Text = rsTmp!����
        txt����(0).Text = rsTmp!����
        
        On Error GoTo errH
        
        '���༰˵��
        strSQL = "Select A.�걾��λ,A.����ID,'['||B.����||']'||B.���� as ����" & _
            " From ������ĿĿ¼ A,���Ʒ���Ŀ¼ B Where A.����ID=B.ID(+) And A.ID=[1]"
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(txt����.Tag))
        txt����.Tag = Nvl(rsTmp!����ID)
        txt����.Text = Nvl(rsTmp!����)
        txt˵��.Text = Nvl(rsTmp!�걾��λ)
        
        '����������
        strSQL = "Select ����,����,����,���� From ������Ŀ���� Where ������ĿID=[1]"
        'Set rsTmp = New ADODB.Recordset
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(txt����.Tag))
        With rsTmp
            Do While Not .EOF
                If !���� = 1 And !���� = 1 Then Me.txtƴ��(0).Text = !����
                If !���� = 1 And !���� = 2 Then Me.txt���(0).Text = !����
                If !���� = 9 Then Me.txt����(1).Text = !����
                If !���� = 9 And !���� = 1 Then Me.txtƴ��(1).Text = !����
                If !���� = 9 And !���� = 2 Then Me.txt���(1).Text = !����
                .MoveNext
            Loop
        End With
        
        '�ؼ���ɫ��ʶ
        For Each objTmp In Me.Controls
            If TypeName(objTmp) = "TextBox" Then
                objTmp.ForeColor = &HC00000
            End If
        Next
        
        vsAdvice.SetFocus
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd����_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    
    strSQL = "Select ID,�ϼ�ID,����,����,����" & _
        " From ���Ʒ���Ŀ¼ Where ����=6" & _
        " Start With �ϼ�ID is Null Connect by Prior ID=�ϼ�ID"
    vPoint = GetCoordPos(fraEdit.Hwnd, txt����.Left, txt����.Top)
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 1, "���׷���", , txt����.Text, , , , True, vPoint.x, vPoint.y, txt����.Height, blnCancel)
    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "û�н����������Ʒ��࣬���ȵ�������Ŀ�����н�����", vbInformation, gstrSysName
        End If
    Else
        txt����.Tag = rsTmp!ID '��¼����ID
        txt����.Text = "[" & rsTmp!���� & "]" & rsTmp!����
        
        If gint���Ʊ��� = 1 And Val(txt����.Tag) = 0 Then
            Call GetMaxCode
        End If
    End If

    txt����.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    Else
        With vsAdvice
            If KeyCode = vbKeyA And Shift = vbCtrlMask Then
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, col������ĿID)) <> 0 Then
                        .TextMatrix(i, colѡ��) = -1
                    End If
                Next
            ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
                For i = .FixedRows To .Rows - 1
                    .TextMatrix(i, colѡ��) = 0
                Next
            End If
        End With
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not Me.ActiveControl Is vsAdvice Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    If InStr(mstrPrivs, "������׷���") = 0 Then
        opt����(0).Enabled = False
        opt����(1).Enabled = False
        opt����(0).Value = True
    End If
    
    Call GetMaxCode
    Call LoadAdvice
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        
    fraEdit.Left = 0
    fraEdit.Top = 0
    fraEdit.Width = Me.ScaleWidth
    fraLine.Left = -15
    fraLine.Width = Me.ScaleWidth + 30
    
    vsAdvice.Left = 0
    vsAdvice.Top = fraEdit.Top + fraEdit.Height
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - fraEdit.Height - fraCommand.Height
    
    fraCommand.Left = 0
    fraCommand.Top = vsAdvice.Top + vsAdvice.Height
    fraCommand.Width = Me.ScaleWidth
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3)
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Function LoadAdvice() As Boolean
'���ܣ���ȡ��ǰ����ָ����ҽ��
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    
    On Error GoTo errH
    
    '����ֻ�����ﲡ���ܹ�����,������δת��
    'סԺ����ѡ��ʱ��������סԺ����δת��
    strSQL = "Select Distinct A.ID,A.���,A.���ID,A.ҽ����Ч,A.������ĿID,A.ҽ������," & _
        " A.��������,A.ִ��Ƶ��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ,A.ҽ������,A.ִ������," & _
        " Nvl(C.����,Decode(Nvl(A.ִ������,0),0,'<����>',5,'<Ժ��ִ��>')) as ִ�п���,A.ִ��ʱ�䷽��," & _
        " A.ִ�п���ID,A.�걾��λ,B.���,B.����,B.���㵥λ,A.�ܸ����� as ����,D.���㵥λ as ������λ,D.id as �շ�ϸĿID" & _
        " From ����ҽ����¼ A,������ĿĿ¼ B,���ű� C,�շ���ĿĿ¼ D" & _
        " Where A.������ĿID=B.ID And A.ִ�п���ID=C.ID(+) And A.�շ�ϸĿID=D.ID(+)" & _
        " And A.ҽ��״̬ Not IN(2,4) And A.��ʼִ��ʱ�� is Not Null And A.������Դ<>3 And Nvl(A.Ӥ��,0)=[2]" & _
        IIF(mlng��ҳID = 0, " And A.����ID+0=[1] And A.�Һŵ�=[3]", " And A.����ID=[1] And A.��ҳID=[4]") & _
        " Order by A.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mintӤ��, mstr�Һŵ�, mlng��ҳID)
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '����������
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, colѡ��) = -1
                .TextMatrix(i, col��Ч) = IIF(Nvl(rsTmp!ҽ����Ч, 0) = 0, "����", "����")
                .TextMatrix(i, col����) = rsTmp!ҽ������
                .TextMatrix(i, col�걾��λ) = Nvl(rsTmp!�걾��λ)  '����걾
                .TextMatrix(i, col����) = FormatEx(Nvl(rsTmp!��������), 4)
                If Not IsNull(rsTmp!��������) Then
                    .TextMatrix(i, col������λ) = Nvl(rsTmp!���㵥λ)
                End If
                If .TextMatrix(i, col��Ч) = "����" Then
                    If Not IsNull(rsTmp!����) Then
                        .TextMatrix(i, col����) = FormatEx(Nvl(rsTmp!����), 4)
                        If Not IsNull(rsTmp!������λ) Then
                            .TextMatrix(i, col������λ) = Nvl(rsTmp!������λ)
                        Else
                            .TextMatrix(i, col������λ) = Nvl(rsTmp!���㵥λ)
                        End If
                    End If
                End If
                .TextMatrix(i, colƵ��) = Nvl(rsTmp!ִ��Ƶ��)
                .TextMatrix(i, colƵ�ʴ���) = Nvl(rsTmp!Ƶ�ʴ���)
                .TextMatrix(i, colƵ�ʼ��) = Nvl(rsTmp!Ƶ�ʼ��)
                .TextMatrix(i, col�����λ) = Nvl(rsTmp!�����λ)
                .TextMatrix(i, col����) = Nvl(rsTmp!ҽ������)
                .TextMatrix(i, colִ��ʱ��) = Nvl(rsTmp!ִ��ʱ�䷽��)
                .TextMatrix(i, colִ�п���) = Nvl(rsTmp!ִ�п���)
                .Cell(flexcpData, i, colִ�п���) = CLng(Nvl(rsTmp!ִ�п���ID, 0))
                .Cell(flexcpData, i, colִ������) = Val(Nvl(rsTmp!ִ������, 0))
                .TextMatrix(i, col���) = rsTmp!ID
                .TextMatrix(i, col������) = Nvl(rsTmp!���ID)
                .TextMatrix(i, col������ĿID) = rsTmp!������ĿID
                .TextMatrix(i, col�������) = rsTmp!���
                .TextMatrix(i, col�շ�ϸĿID) = zlCommFun.Nvl(rsTmp!�շ�ϸĿID)
                
                '���������ؼ��÷���ʾ
                If InStr(",C,D,F,G,E,", rsTmp!���) > 0 And Not IsNull(rsTmp!���ID) Then
                    .RowHidden(i) = True
                ElseIf rsTmp!��� = "7" Then
                    .RowHidden(i) = True
                ElseIf rsTmp!��� = "E" And IsNull(rsTmp!���ID) _
                    And Val(.TextMatrix(i - 1, col������)) = rsTmp!ID _
                    And InStr(",5,6,", .TextMatrix(i - 1, col�������)) > 0 Then
                    '��ҩ;��
                    .RowHidden(i) = True
                    '��ʾ��ҩ;��
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col������)) = rsTmp!ID Then
                            .TextMatrix(j, col�÷�) = rsTmp!����
                            
                            '��ʾ��ҩִ������
                            If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                                .TextMatrix(j, colִ������) = "�Ա�ҩ"
                            ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                                .TextMatrix(j, colִ������) = "��Ժ��ҩ"
                            End If
                        Else
                            Exit For
                        End If
                    Next
                ElseIf rsTmp!��� = "E" And IsNull(rsTmp!���ID) _
                    And Val(.TextMatrix(i - 1, col������)) = rsTmp!ID _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col�������)) > 0 Then
                    '��ҩ�÷������ɼ�����
                    .TextMatrix(i, col�÷�) = rsTmp!����
                    
                    '��ҩ������ִ�п���
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col������)) = rsTmp!ID Then
                            If InStr(",7,C,", .TextMatrix(j, col�������)) > 0 Then
                                .TextMatrix(i, colִ�п���) = .TextMatrix(j, colִ�п���)
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    '��ҩ����
                    If .TextMatrix(i - 1, col�������) <> "C" Then
                        .TextMatrix(i, col������λ) = "��"
                        
                        '��ʾ��ҩ�䷽ִ������:��ҩƷΪ׼�ж�
                        j = .FindRow(CStr(rsTmp!ID), , col������)
                        If Val(.Cell(flexcpData, j, colִ������)) = 5 And Val(.Cell(flexcpData, i, colִ������)) <> 5 Then
                            .TextMatrix(i, colִ������) = "�Ա�ҩ"
                        ElseIf Val(.Cell(flexcpData, j, colִ������)) <> 5 And Val(.Cell(flexcpData, i, colִ������)) = 5 Then
                            .TextMatrix(i, colִ������) = "��Ժ��ҩ"
                        End If
                    End If
                End If
                rsTmp.MoveNext
            Next
        End If
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col����
        .Redraw = flexRDDirect
    End With
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmd����_Click
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlControl.TxtSelAll(txt����)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmd����_Click
    End If
End Sub

Private Sub txt����_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt����(Index))
End Sub

Private Sub txt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Me.txtƴ��(Index).Text = zlCommFun.zlGetSymbol(Me.txt����(Index).Text, 0)
    Me.txt���(Index).Text = zlCommFun.zlGetSymbol(Me.txt����(Index).Text, 1)
End Sub

Private Sub txtƴ��_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtƴ��(Index))
End Sub

Private Sub txtƴ��_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_GotFocus()
    Call zlControl.TxtSelAll(txt˵��)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt���(Index))
End Sub

Private Sub txt���_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = colѡ�� Then Call RowSelectSame(Row)
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col���� Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colѡ�� Then
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_DblClick()
    Call vsAdvice_KeyPress(32)
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsAdvice
        If KeyAscii = 32 Then
            If .Col <> colѡ�� Then
                KeyAscii = 0
                If Val(.TextMatrix(.Row, col������ĿID)) <> 0 Then
                    .TextMatrix(.Row, colѡ��) = IIF(Val(.TextMatrix(.Row, colѡ��)) = 0, -1, 0)
                    Call RowSelectSame(.Row)
                End If
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> colѡ�� Then
        Cancel = True
    ElseIf Val(vsAdvice.TextMatrix(vsAdvice.Row, col������ĿID)) = 0 Then
        Cancel = True
    End If
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '����һ����ҩ������еı��߼�����
        lngLeft = col��Ч: lngRight = col��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = colƵ��: lngRight = col�÷�
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowInһ����ҩ(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, col�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col�������)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col������)) = Val(.TextMatrix(lngRow, col������)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col������)) = Val(.TextMatrix(lngRow, col������)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col������)) = Val(.TextMatrix(lngRow, col������)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col������)) = Val(.TextMatrix(lngRow, col������)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub RowSelectSame(ByVal lngRow As Long)
'���ܣ�����ָ����(����Ϊ������)��ѡ��״̬,�����ҽ��һ��ѡ��
    Dim i As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, col������)) <> 0 Then
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col������)) = Val(.TextMatrix(lngRow, col������)) _
                    Or Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col������)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col������)) = Val(.TextMatrix(lngRow, col������)) _
                    Or Val(.TextMatrix(i, col���)) = Val(.TextMatrix(lngRow, col������)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col������)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col������)) = Val(.TextMatrix(lngRow, col���)) Then
                    .TextMatrix(i, colѡ��) = .TextMatrix(lngRow, colѡ��)
                Else
                    Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub GetMaxCode()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strTmp As String
    
    On Error GoTo errH
    
    If gint���Ʊ��� = 1 And Val(txt����.Tag) <> 0 Then
        '����+����+˳����
        strTmp = Mid(txt����.Text, 2, InStr(1, txt����.Text, "]") - 2)
        strSQL = "Select Nvl(Max(����),'0000000') as ���� From ������ĿĿ¼ Where ���='9' And ���� Like [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "9" & strTmp & "%")
        On Error Resume Next
        txt����.Text = "9" & strTmp & Right(String(10, "0") & Val(rsTmp!����) + 1, Len(rsTmp!����) - 1 - Len(strTmp))
    Else
        '˳����
        strSQL = "Select Nvl(Max(����),'0000000') as ���� From ������ĿĿ¼ Where ���='9'"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        txt����.Text = Right(String(10, "0") & Val(rsTmp!����) + 1, Len(rsTmp!����))
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
