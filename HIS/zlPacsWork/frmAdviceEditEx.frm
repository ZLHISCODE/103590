VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdviceEditEx 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4275
   ControlBox      =   0   'False
   Icon            =   "frmAdviceEditEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.OptionButton optPosition 
      Caption         =   "���벿λ(&I)"
      Height          =   180
      Index           =   1
      Left            =   1455
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1995
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.OptionButton optPosition 
      Caption         =   "ѡ��λ(&S)"
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1995
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "��"
      Height          =   240
      Left            =   2475
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "ѡ����Ŀ(*)"
      Top             =   1950
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Left            =   525
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3555
      Picture         =   "frmAdviceEditEx.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "ȡ��(Esc)"
      Top             =   1920
      Width           =   450
   End
   Begin VB.ComboBox cboData 
      Height          =   300
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExt 
      Align           =   1  'Align Top
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4275
      _cx             =   7541
      _cy             =   3254
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
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdviceEditEx.frx":0596
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      Begin VB.CommandButton cmd 
         Caption         =   "��"
         Height          =   240
         Left            =   3435
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "ѡ����Ŀ(*)"
         Top             =   1035
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.ComboBox cbo�걾 
      Height          =   300
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdOK 
      Height          =   315
      Left            =   3015
      Picture         =   "frmAdviceEditEx.frx":06A2
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "ȷ��(F2)"
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   105
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmAdviceEditEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'��ڲ�����
Public mstrPrivs As String
Public mlngHwnd As Long '���ڶ�λ�Ŀؼ����
Public mint��Ч As Integer
Public mstr�Ա� As String
Public mint������� As Integer '1-����,2-סԺ
'0-������,1-��������,2-��ҩ�䷽,3-����걾,4-�������
Public mintType As Integer
'��/��:���="��λID1,��λID2,..."
'      ����="����ID1,����ID2,...;����ID",���п���û�и�������������
'      ��ҩ="��ҩID1,����1,��ע1;��ҩID2,����2,��ע2;...|�巨ID"
'      ����걾="��ĿID1,��ĿID2,...;����걾"
'      �������="��ĿID1,��ĿID2,...;����걾"
Public mstrExtData As String '����ʱΪ��;ҽ����������ʱΪ"��ĿID;"
'��������ĿID,��ҩ�䷽ʱΪ�䷽ID��ζ��ҩID,�������ʱ��ʾ���Ƶ���ID
Public mlng��ĿID As Long

'����ְ����Ҫ��
Public mbln��ʿվ As Boolean '�Ƿ�ʿվ����
Public mblnҽ�� As Boolean '�Ƿ�ҽ���򹫷Ѳ���

'���ڲ�����
Public mblnOK As Boolean '��

'�������
Private mlng��ҩ�� As Long
Private mint���� As Integer
Private mstrLike As String
Private mblnReturn As Boolean '�Ƿ��˻س�ȷ��
Private mblnNotAddNew As Boolean '�Ƿ���������
'-----------------------------------------------------------------------------------------------------
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long

Private Sub cboData_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboData.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = zlControl.CboMatchIndex(cboData.Hwnd, KeyAscii)
        If lngIdx = -1 And cboData.ListCount > 0 Then lngIdx = 0
        cboData.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo�걾_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo�걾.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = zlControl.CboMatchIndex(cbo�걾.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo�걾.ListCount > 0 Then lngIdx = 0
        cbo�걾.ListIndex = lngIdx
    End If
End Sub

Private Sub cmd_Click()
'���ܣ�����Ŀѡ����
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim strSQL As String, int�Ա� As Integer, strSQLItem As String
    Dim strStock As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, strҩƷ As String
    Dim strSamples As String
    
    If mstr�Ա� Like "*��*" Then
        int�Ա� = 1
    ElseIf mstr�Ա� Like "*Ů*" Then
        int�Ա� = 2
    End If
    
    On Error GoTo errH
    
    If mintType = 0 And optPosition(1).Value Then
        '�����鲿λ
        strSQL = _
            "Select A.ID, A.����, A.�걾��λ As ��鲿λ" & vbNewLine & _
            "From ������ĿĿ¼ A, ������Ŀ��� B" & vbNewLine & _
            "Where A.ID = B.������Ŀid And B.�������id = [1] And A.������� In ([2], 3) And" & vbNewLine & _
            "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & vbNewLine & _
            "Order By B.���"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��λ", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, mlng��ĿID, mint�������)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ����õļ�鲿λ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        
        '����ظ�����
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "�ü�鲿λ�Ѿ���������¼�롣", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Set��λ����(vsExt.Row, rsTmp)
    ElseIf mintType = 1 Then
        '���븽������:���ﲻ�ǵ���Ӧ��,��˲�����
        strSQLItem = _
            " From ������ĿĿ¼ A Where A.���='F' And A.ID<>" & mlng��ĿID & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And A.������� IN([1],3) And Nvl(A.ִ��Ƶ��,0) IN(0,[2]) And Nvl(A.�����Ա�,0) IN(0,[3])"
        
        strSQL = "Select Distinct 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ��ģ" & _
            " From ���Ʒ���Ŀ¼ Where ����=5" & _
            " Start With ID In (Select ����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID"
        strSQL = strSQL & " Union ALL" & _
            " Select Distinct 1 as ĩ��,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ" & _
            strSQLItem & " Order By ����"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "����", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mint�������, IIF(mint��Ч = 0, 2, 1), int�Ա�)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ����õ�������Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        
        '����ظ�����
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "�ø��������Ѿ���������¼�롣", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Set��������(vsExt.Row, rsTmp)
    ElseIf mintType = 2 And CellCanEdit(vsExt.Row, vsExt.Col) Then
        If vsExt.Col Mod 4 = 0 Then
            '��ҩ���,��ҩ��δָ��ʱ,����������¼
            If mlng��ҩ�� <> 0 Then
                strStock = _
                    "Select ҩƷID,Sum(Nvl(��������,0)) as ��� From ҩƷ���" & _
                    " Where (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate))" & _
                    " And ���� = 1 And �ⷿID=" & mlng��ҩ�� & _
                    " Group by ҩƷID" & _
                    " Having Sum(Nvl(��������,0))<>0"
            Else
                strStock = "Select NULL as ҩƷID,NULL as ��� From Dual"
            End If
            
            '����ҩƷȨ��
            strҩƷ = ""
            If InStr(mstrPrivs, "�´�����ҩ��") = 0 Then
                strҩƷ = strҩƷ & " And E.�������<>'����ҩ'"
            End If
            If InStr(mstrPrivs, "�´ﶾ��ҩ��") = 0 Then
                strҩƷ = strҩƷ & " And E.�������<>'����ҩ'"
            End If
            If InStr(mstrPrivs, "�´����ҩ��") = 0 Then
                strҩƷ = strҩƷ & " And E.��ֵ���� Not IN('����','����')"
            End If
            
            'ѡ��ζ�в�ҩ:���ﲻ�ǵ���Ӧ��,��˲�����
            strSQL = "Select 0 as ĩ��,-1 as ID,-NULL as �ϼ�ID,NULL as ����," & _
                " CHR(13)||'������ҩ' as ����,NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ����ְ��ID From Dual"
            strSQL = strSQL & " Union ALL" & _
                " Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ���,NULL as ����,NULL as ���,NULL as ����ְ��ID" & _
                " From ���Ʒ���Ŀ¼ Where ����=3" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            strSQL = strSQL & " Union ALL" & _
                " Select Distinct 1 as ĩ��,A.ID,A.����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,D.���,D.����," & _
                " Decode(X.���,NULL,NULL,X.���/C.סԺ��װ||C.סԺ��λ) AS ���,E.����ְ�� as ����ְ��ID" & _
                " From ������ĿĿ¼ A,ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,(" & strStock & ") X" & _
                " Where A.���='7' And A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) And A.������� IN([1],3)" & _
                    " And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([1],3)" & _
                    " And Nvl(A.ִ��Ƶ��,0) IN(0,[2])" & strҩƷ & _
                    " And Nvl(A.�����Ա�,0) IN(0,[3]) And C.ҩƷID=X.ҩƷID(+)"
            strSQL = strSQL & " Union ALL" & _
                " Select Distinct 1 as ĩ��,A.ID,-1 as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,D.���,D.����," & _
                " Decode(X.���,NULL,NULL,X.���/C.סԺ��װ||C.סԺ��λ) AS ���,E.����ְ�� as ����ְ��ID" & _
                " From ������ĿĿ¼ A,ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,���Ƹ�����Ŀ T,(" & strStock & ") X" & _
                " Where A.���='7' And A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) And A.������� IN([1],3)" & _
                    " And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([1],3)" & _
                    " And Nvl(A.ִ��Ƶ��,0) IN(0,[2])" & strҩƷ & _
                    " And Nvl(A.�����Ա�,0) IN(0,[3]) And C.ҩƷID=X.ҩƷID(+)" & _
                    " And T.������ĿID=A.ID And T.��ԱID=[4]"
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "��ҩ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
                mint�������, IIF(mint��Ч = 0, 2, 1), int�Ա�, UserInfo.ID)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ����õ���ҩ��Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            
            '����ظ�����
            If ItemExist(rsTmp!ID, vsExt.Row, vsExt.Col) Then
                MsgBox "��ζ��ҩ���䷽���Ѿ�¼�롣", vbInformation, gstrSysName
                Exit Sub
            End If
            
            '����ְ����
            If Not mbln��ʿվ Then
                strSQL = CheckOneDuty(rsTmp!����, Nvl(rsTmp!����ְ��ID), UserInfo.����, mblnҽ��)
                If strSQL <> "" Then
                    MsgBox strSQL, vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            
            '��ȡ����ֵ
            vsExt.TextMatrix(vsExt.Row, vsExt.Col) = rsTmp!����
            vsExt.TextMatrix(vsExt.Row, vsExt.Col + 2) = rsTmp!��λ
            vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = vsExt.TextMatrix(vsExt.Row, vsExt.Col)
            vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col + 2) = CLng(rsTmp!ID) '��¼��ҩID
            
            Call EnterNextCell(vsExt.Row, vsExt.Col)
        ElseIf vsExt.Col Mod 4 = 3 Then
            'ѡ���ע
            strSQL = "Select Rownum as ID,����,����,���� From ��ҩ�����ע Order by ����"
            vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "��ע", , , , , , True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ����õļ����ע�����ȵ�����������������á�", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            
            '��ȡ����ֵ
            vsExt.TextMatrix(vsExt.Row, vsExt.Col) = rsTmp!����
            vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = vsExt.TextMatrix(vsExt.Row, vsExt.Col)
            
            Call EnterNextCell(vsExt.Row, vsExt.Col)
        End If
    ElseIf mintType = 4 Then
        '������Ŀ
        With Me.cbo�걾
            For i = 0 To .ListCount - 1
                strSamples = strSamples & ",'" & .List(i) & "'"
            Next
        End With
        If Len(strSamples) > 0 Then
            strSamples = Mid(strSamples, 2)
        Else
            strSamples = "''"
        End If
        If mlng��ĿID > 0 Then 'ָ�������Ƶ���
            strSQLItem = "From ������ĿĿ¼ A,���Ƶ���Ӧ�� B,������Ŀ�ο� C,���鱨����Ŀ D " & _
                "Where A.ID=B.������ĿID And A.id=D.������Ŀid(+) And D.������ĿID=C.��Ŀid(+)" & _
                " And B.Ӧ�ó���=[2] And B.�����ļ�ID=[1]" & _
                " And A.���='C' And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) In (0,[3])" & _
                " And A.������� IN([2],3) And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And (C.�걾���� In (" & strSamples & ") Or C.�걾���� Is Null)"
        Else
            strSQLItem = "From ������ĿĿ¼ A,���Ƶ���Ӧ�� B,������Ŀ�ο� C,���鱨����Ŀ D " & _
                "Where A.ID=B.������ĿID(+) And A.id=D.������Ŀid(+) And D.������ĿID=C.��Ŀid(+)" & _
                " And A.���='C' And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) In (0,[3])" & _
                " And A.������� IN([2],3) And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And (C.�걾���� In (" & strSamples & ") Or C.�걾���� Is Null)"
'                " And (B.������ĿID is Null Or B.Ӧ�ó���=" & mint������� & ")"
        End If
        
        strSQL = "Select Distinct 0 as ĩ��,ID,�ϼ�ID,����,����,' ' As ��������,0 As �����ļ�ID,' ' As �걾��λ" & _
            " From ���Ʒ���Ŀ¼ Where ����=5" & _
            " Start With ID In (Select A.����ID " & strSQLItem & ") Connect by Prior �ϼ�ID=ID"
        strSQL = strSQL & " Union ALL" & _
            " Select Distinct 1 as ĩ��,A.ID,����ID as �ϼ�ID,A.����,A.����,A.�������� as ��������,B.�����ļ�ID,A.�걾��λ " & strSQLItem & " Order By ����"
        
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "������Ŀ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mlng��ĿID, mint�������, int�Ա�)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ����õļ�����Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        If rsTmp("��������") = "΢����" And vsExt.Rows > 2 Then
            If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '��������ֻ�ܿ�һ��΢������Ŀ
                MsgBox "΢������Ŀֻ�ܵ������룡", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        If mlng��ĿID = 0 Then mlng��ĿID = Nvl(rsTmp("�����ļ�ID"), 0)
        
        '����ظ�����
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "�ü�����Ŀ�Ѿ�¼�룡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '�����������Ƿ���ͬ
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 And i <> vsExt.Row Then
                If Not (vsExt.TextMatrix(i, 1) = Nvl(rsTmp!��������) _
                    Or vsExt.TextMatrix(i, 1) = "" Or Nvl(rsTmp!��������) = "") Then
                    MsgBox "��������ͬ�������͵���Ŀ����������Ŀ�ļ�������Ϊ""" & vsExt.TextMatrix(i, 1) & """��", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '���³�ʼ�걾
        If Not InitCombox(rsTmp("ID"), Nvl(rsTmp("�걾��λ"))) Then Exit Sub
        
        Call Set������Ŀ(vsExt.Row, rsTmp)
        If rsTmp("��������") = "΢����" Then
            mblnNotAddNew = True
            vsExt.Rows = 2
        Else
            mblnNotAddNew = False
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdData_Click()
'���ܣ�����Ŀѡ����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str�Ա� As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    If mstr�Ա� Like "*��*" Then
        str�Ա� = "0,1"
    ElseIf mstr�Ա� Like "*Ů*" Then
        str�Ա� = "0,2"
    Else
        str�Ա� = "0"
    End If
    
    If mintType = 1 Then
        '����������Ŀ:���ﲻ�ǵ���Ӧ��,��˲�����
        strSQLItem = " From ������ĿĿ¼ A Where A.���='G'" & _
                " And A.������� IN([2],3) And A.ID<>[1]" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)"

        strSQL = "Select Distinct 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��λ,NULL as ��������" & _
            " From ���Ʒ���Ŀ¼ Where ����=5" & _
            " Start With ID In (Select ����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID"
        strSQL = strSQL & " Union ALL" & _
            " Select Distinct 1 as ĩ��,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��������" & _
            strSQLItem & " Order By ����"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "������Ŀ", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mlng��ĿID, mint�������)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
            End If
            txtData.SetFocus: Exit Sub
        End If
        txtData.Tag = rsTmp!ID
        txtData.Text = "[" & rsTmp!���� & "]" & rsTmp!����
        cmdData.Tag = txtData.Text
        
        txtData.SetFocus
    ElseIf mintType = 4 Then
        '����걾
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim str��ҩIDs As String, blnSkip As Boolean
    Dim strMsg As String, strTmp As String
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    
    If mintType = 0 Then '��鲿λ���
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 Then
                If optPosition(0).Value Then
                    If Val(vsExt.TextMatrix(i, vsExt.Cols - 1)) <> 0 Then
                        strTmp = strTmp & "," & vsExt.RowData(i)
                    End If
                Else
                    strTmp = strTmp & "," & vsExt.RowData(i)
                End If
            End If
        Next
        strTmp = Mid(strTmp, 2)
        If strTmp = "" Then
            MsgBox "������Ҫһ����鲿λ��", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
    ElseIf mintType = 1 Or mintType = 4 Then '����������������Ŀ��������Ŀ���걾
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 Then
                strTmp = strTmp & "," & vsExt.RowData(i)
            End If
        Next
        strTmp = Mid(strTmp, 2)
        If strTmp = "" And mintType = 4 Then
            MsgBox "����Ҫѡ��һ��������Ŀ��", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
        strTmp = strTmp & ";" & IIF(mintType = 4, Me.cbo�걾.Text, IIF(Val(txtData.Tag) = 0, "", Val(txtData.Tag)))
    ElseIf mintType = 2 Then '��ҩ�䷽��巨
        blnSkip = False
        For i = vsExt.FixedRows To vsExt.Rows - 1
            For j = 0 To vsExt.Cols - 1 Step 4
                If CLng(vsExt.Cell(flexcpData, i, j + 2)) <> 0 Then
                    If Val(vsExt.TextMatrix(i, j + 1)) = 0 Then
                        If Not blnSkip Then
                            If MsgBox("""" & vsExt.TextMatrix(i, j) & """û�����뵥ζ������Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                vsExt.Row = i: vsExt.Col = j + 1
                                Call vsExt.ShowCell(i, j + 1)
                                vsExt.SetFocus: Exit Sub
                            End If
                            blnSkip = True
                        End If
                    End If
                    If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                        strTmp = strTmp & ";" & vsExt.Cell(flexcpData, i, j + 2) & "," & vsExt.TextMatrix(i, j + 1) & "," & vsExt.TextMatrix(i, j + 3)
                        str��ҩIDs = str��ҩIDs & "," & CLng(vsExt.Cell(flexcpData, i, j + 2))
                    End If
                End If
            Next
        Next
        strTmp = Mid(strTmp, 2)
        str��ҩIDs = Mid(str��ҩIDs, 2)
        
        If strTmp = "" Then
            MsgBox "�����䷽����������һζ��ҩ��", vbInformation, gstrSysName
            vsExt.Row = vsExt.FixedRows: vsExt.Col = 0
            vsExt.SetFocus: Exit Sub
        End If
        If cboData.ListIndex = -1 Then
            MsgBox "��ȷ����ҩ�䷽�ļ巨��", vbInformation, gstrSysName
            cboData.SetFocus: Exit Sub
        End If
        
        '����ְ����
        If Not mbln��ʿվ Then
            strSQL = "Select ҩ��ID,����ְ�� From ҩƷ���� Where ҩ��ID IN(" & str��ҩIDs & ")"
            On Error GoTo errH
            Set rsTmp = New ADODB.Recordset
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption) 'IN
            For i = vsExt.FixedRows To vsExt.Rows - 1
                For j = 0 To vsExt.Cols - 1 Step 4
                    If CLng(vsExt.Cell(flexcpData, i, j + 2)) <> 0 Then
                        If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                            rsTmp.Filter = "ҩ��ID=" & CLng(vsExt.Cell(flexcpData, i, j + 2))
                            If Not rsTmp.EOF Then
                                strMsg = CheckOneDuty(vsExt.TextMatrix(i, j), Nvl(rsTmp!����ְ��), UserInfo.����, mblnҽ��)
                                If strMsg <> "" Then
                                    vsExt.Row = i: vsExt.Col = j
                                    Call vsExt.ShowCell(i, j)
                                    MsgBox strMsg, vbInformation, gstrSysName
                                    vsExt.SetFocus: Exit Sub
                                End If
                            End If
                        End If
                    End If
                Next
            Next
        End If
        
        'ҩƷ���ɼ��
        If Not Check��ҩ����(str��ҩIDs) Then Exit Sub
        
        strTmp = strTmp & "|" & cboData.ItemData(cboData.ListIndex)
    ElseIf mintType = 3 Then '����걾
        For i = 1 To vsExt.Rows - 1
            If Val(vsExt.TextMatrix(i, 1)) <> 0 Then
                strTmp = vsExt.TextMatrix(i, 0)
                Exit For
            End If
        Next
        If strTmp = "" Then
            MsgBox "��ѡ��ü�����Ŀ�ļ���걾��", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
    End If
    
    mstrExtData = strTmp
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mintType = 0 Then
        optPosition(0).TabStop = False: optPosition(1).TabStop = False '��Ȼ��Ч
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    ElseIf mintType = 0 And Shift = vbCtrlMask And KeyCode = vbKeyA Then
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 Then vsExt.TextMatrix(i, 2) = 1
        Next
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '����������ָ�����������
    If InStr(",;|'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim blnMulti As Boolean, vRect As RECT
    
    Call zlControl.CboSetHeight(cboData, Me.Height * 2)
    Call zlControl.CboSetWidth(cboData.Hwnd, cboData.Width * 1.2)
    
    '����ƥ��
    mstrLike = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
    mint���� = Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "��������", 0)) '����ƥ�䷽ʽ��0-ƴ��,1-���
    If mint������� = 0 Then mint������� = 2 'ȱʡΪסԺ
    mblnOK = False
    mblnNotAddNew = False
            
    '��ʼ�������ʽ
    If mintType = 0 Then
        optPosition(0).Visible = True: optPosition(1).Visible = True
        optPosition(Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "��λȷ����ʽ", 0))).Value = True
        If Not Init������ Then Unload Me: Exit Sub
    ElseIf mintType = 1 Then
        lblData.Visible = True
        txtData.Visible = True
        cmdData.Visible = True
        lblData.Caption = "����"
        If Not Init������Ŀ Then Unload Me: Exit Sub
    ElseIf mintType = 2 Then
        mlng��ҩ�� = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, IIF(mint������� = 1, "����", "סԺ") & "ȱʡ��ҩ��", 0))
        
        lblData.Visible = True
        cboData.Visible = True
        lblData.Caption = "�巨"
        If Not Init��ҩ�䷽ Then Unload Me: Exit Sub
    ElseIf mintType = 3 Then
        If Not Init����걾 Then Unload Me: Exit Sub
    ElseIf mintType = 4 Then
        lblData.Visible = True
        lblData.Caption = "�걾"
        With cbo�걾
            .Left = txtData.Left: .Top = txtData.Top: .Width = txtData.Width
            .Visible = True
        End With
        If Not Init������� Then Unload Me: Exit Sub
        If Not InitCombox(DefaultValue:=Me.txtData) Then Unload Me: Exit Sub
    
        blnMulti = GetSysParVal(84) = "1" '�Ƿ�����һ��ҽ��������������Ŀ
        
        If Len(Trim(mstrExtData)) > 0 Then
            If Len(Trim(Split(mstrExtData, ";")(0))) > 0 And Not blnMulti Then
                vsExt.Enabled = False
                '���ֻ��һ���걾����ʾ������
                If cbo�걾.ListCount < 2 Then cmdOK_Click: Exit Sub
            End If
        End If
    End If
    
    '���嶨λ
    GetWindowRect mlngHwnd, vRect
    Me.Left = (vRect.Left - 1) * Screen.TwipsPerPixelX
    Me.Top = (vRect.Top - 1) * Screen.TwipsPerPixelY - Me.Height
    
    Call Form_Resize
End Sub

Private Function Init��ҩ�䷽() As Boolean
'���ܣ���ʼ����ҩ�䷽����ʽ������
'������mstrExtData=����ÿζ��ҩ��Ϣ���巨��Ϣ�Ĵ�,Ϊ��ʱ��ʾ��������ҩ�䷽
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim lngRow As Long, lngCol As Long
    Dim str��ҩIDs As String, lng�巨ID As Long
    Dim arr��ҩ As Variant

    vsExt.Clear
    vsExt.Cols = 12: vsExt.Rows = 7
    vsExt.FixedCols = 0: vsExt.FixedRows = 1
    vsExt.ColWidth(0) = 795: vsExt.ColAlignment(0) = 1 '��ζ��ҩ
    vsExt.ColWidth(1) = 450: vsExt.ColAlignment(1) = 7 '��ζ����
    vsExt.ColWidth(2) = 285: vsExt.ColAlignment(2) = 1 '��λ
    vsExt.ColWidth(3) = 750: vsExt.ColAlignment(3) = 1 '��ע
    For i = 4 To vsExt.Cols - 1
        vsExt.ColWidth(i) = vsExt.ColWidth(i - 4)
        vsExt.ColAlignment(i) = vsExt.ColAlignment(i - 4)
    Next
    vsExt.MergeCells = flexMergeFixedOnly
    vsExt.MergeRow(0) = True
    vsExt.Cell(flexcpAlignment, 0, 0, 0, vsExt.Cols - 1) = 4
    vsExt.Cell(flexcpText, 0, 0, 0, vsExt.Cols - 1) = "�����������в�ҩ,��ζ����,��ע����*��ѡȡ��ҩ���ע��"
    
    Me.Width = (Me.Width - Me.ScaleWidth) + 2280 * 3 + 250
    vsExt.GridColor = vsExt.BackColor
    vsExt.Editable = flexEDKbdMouse
    
    On Error GoTo errH
    
    If mstrExtData <> "" Then '�޸�
        lng�巨ID = Val(Split(mstrExtData, "|")(1))
        arr��ҩ = Split(Split(mstrExtData, "|")(0), ";")
        
        For i = 0 To UBound(arr��ҩ)
            str��ҩIDs = str��ҩIDs & "," & CStr(Split(arr��ҩ(i), ",")(0))
        Next
        str��ҩIDs = Mid(str��ҩIDs, 2)
        
        strSQL = "Select A.ID,A.����,A.���㵥λ From ������ĿĿ¼ A Where ID IN(" & str��ҩIDs & ")"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'IN
        
        If vsExt.Rows < -Int(rsTmp.RecordCount / -3) + 1 Then
            vsExt.Rows = -Int(rsTmp.RecordCount / -3) + 1
        End If
        lngRow = vsExt.FixedRows: lngCol = 0
        
        '�������ڵ����ݺʹ�����ʾ
        For i = 0 To UBound(arr��ҩ)
            rsTmp.Filter = "ID=" & CStr(Split(arr��ҩ(i), ",")(0))
            If Not rsTmp.EOF Then
                vsExt.TextMatrix(lngRow, lngCol) = rsTmp!����
                vsExt.TextMatrix(lngRow, lngCol + 1) = CStr(Split(arr��ҩ(i), ",")(1))
                vsExt.TextMatrix(lngRow, lngCol + 2) = Nvl(rsTmp!���㵥λ)
                vsExt.TextMatrix(lngRow, lngCol + 3) = CStr(Split(arr��ҩ(i), ",")(2))
                
                '���ڻָ���ʾ�ļ�¼
                vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
                vsExt.Cell(flexcpData, lngRow, lngCol + 1) = vsExt.TextMatrix(lngRow, lngCol + 1)
                vsExt.Cell(flexcpData, lngRow, lngCol + 2) = CLng(rsTmp!ID) '��¼��ҩID
                vsExt.Cell(flexcpData, lngRow, lngCol + 3) = vsExt.TextMatrix(lngRow, lngCol + 3)
                                
                '��һλ��
                If lngCol + 4 > vsExt.Cols - 1 Then
                    lngRow = lngRow + 1: lngCol = 0
                Else
                    lngCol = lngCol + 4
                End If
            End If
        Next
    Else '����
        strSQL = "Select ID,���,����,���㵥λ From ������ĿĿ¼ Where ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID)
        If rsTmp!��� = "7" Then
            '�����˵�ζ�в�ҩ
            vsExt.TextMatrix(vsExt.FixedRows, 0) = rsTmp!����
            vsExt.TextMatrix(vsExt.FixedRows, 2) = Nvl(rsTmp!���㵥λ)
            
            '���ڻָ���ʾ�ļ�¼
            vsExt.Cell(flexcpData, vsExt.FixedRows, 0) = vsExt.TextMatrix(vsExt.FixedRows, 0)
            vsExt.Cell(flexcpData, vsExt.FixedRows, 2) = CLng(rsTmp!ID) '��¼��ҩID
        Else
            '�������䷽��Ŀ
            strSQL = "Select A.ID,A.����,A.���㵥λ,B.��������,B.ҽ������" & _
                " From ������ĿĿ¼ A,������Ŀ��� B,ҩƷ��� C,�շ���ĿĿ¼ D" & _
                " Where A.ID=B.������ĿID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID And B.�������ID=[1]" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) And A.������� IN([2],3)" & _
                " And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� is NULL) And D.������� IN([2],3)" & _
                " Order by B.���"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID, mint�������)
            If rsTmp.EOF Then
                MsgBox "����ҩ�䷽��ǰ����Ч���䷽��ɣ����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
                Exit Function
            End If
            
            If vsExt.Rows < -Int(rsTmp.RecordCount / -3) + 1 Then
                vsExt.Rows = -Int(rsTmp.RecordCount / -3) + 1
            End If
            lngRow = vsExt.FixedRows: lngCol = 0
            
            '�������õ����ݵĴ�����ʾ
            For i = 1 To rsTmp.RecordCount
                vsExt.TextMatrix(lngRow, lngCol) = rsTmp!����
                vsExt.TextMatrix(lngRow, lngCol + 1) = Nvl(rsTmp!��������)
                vsExt.TextMatrix(lngRow, lngCol + 2) = Nvl(rsTmp!���㵥λ)
                vsExt.TextMatrix(lngRow, lngCol + 3) = Nvl(rsTmp!ҽ������)
                
                '���ڻָ���ʾ�ļ�¼
                vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
                vsExt.Cell(flexcpData, lngRow, lngCol + 1) = vsExt.TextMatrix(lngRow, lngCol + 1)
                vsExt.Cell(flexcpData, lngRow, lngCol + 2) = CLng(rsTmp!ID) '��¼��ҩID
                vsExt.Cell(flexcpData, lngRow, lngCol + 3) = vsExt.TextMatrix(lngRow, lngCol + 3)
                
                '��һλ��
                If lngCol + 4 > vsExt.Cols - 1 Then
                    lngRow = lngRow + 1: lngCol = 0
                Else
                    lngCol = lngCol + 4
                End If
                rsTmp.MoveNext
            Next
            
            '��ȡ�䷽��Ŀ��ȱʡ�巨
            strSQL = "Select �÷�ID From �����÷����� Where ����=1 And ��ĿID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID)
            If Not rsTmp.EOF Then lng�巨ID = rsTmp!�÷�ID
        End If
    End If
        
    '��ҩ�巨
    strSQL = "Select A.ID,A.����,A.���� From ������ĿĿ¼ A" & _
        " Where A.���='E' And A.��������='3' And A.������� IN([1],3)" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " Order by A.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint�������)
    If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
    If rsTmp.EOF Then
        MsgBox "δ�ҵ���Ч����ҩ�巨�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    
    For i = 1 To rsTmp.RecordCount
        cboData.AddItem rsTmp!���� & "-" & rsTmp!����
        cboData.ItemData(cboData.NewIndex) = rsTmp!ID
        If rsTmp!ID = lng�巨ID Then
            Call zlControl.CboSetIndex(cboData.Hwnd, cboData.NewIndex)
        End If
        rsTmp.MoveNext
    Next
    
    Call SetSplitLine
    vsExt.Row = vsExt.FixedRows: vsExt.Col = 0
    Init��ҩ�䷽ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetSplitLine()
'���ܣ�������ҩ�䷽��������зָ���
    Dim lngRow As Long, lngCol As Long
        
    vsExt.Redraw = False
    lngRow = vsExt.Row: lngCol = vsExt.Col
    
    vsExt.Select vsExt.FixedRows, 3, vsExt.Rows - 1, 3
    vsExt.CellBorder &HC0C0C0, 0, 0, 1, 0, 0, 0
    vsExt.Select vsExt.FixedRows, 7, vsExt.Rows - 1, 7
    vsExt.CellBorder &HC0C0C0, 0, 0, 1, 0, 0, 0

    vsExt.Row = lngRow: vsExt.Col = lngCol
    vsExt.Redraw = True
End Sub

Private Function Init������Ŀ() As Boolean
'���ܣ���ʼ����������ʽ������
'������mstrExtData=��������������������Ŀ����Ϣ,���п���û�и���������Ϊ��ʱ��ʾ������������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, lng����ID As Long
    Dim arr����IDs As Variant, str����IDs As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = mstrExtData
    If strSQL = "" Then strSQL = ";"
    str����IDs = CStr(Split(strSQL, ";")(0))
    lng����ID = Val(Split(strSQL, ";")(1))
    
    '��������
    If str����IDs <> "" Then
        strSQL = "Select A.ID,A.����,A.����,A.��������" & _
            " From ������ĿĿ¼ A" & _
            " Where A.���='F' And A.ID IN(" & str����IDs & ")" & _
            " Order by A.����"
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'IN
        i = rsTmp.RecordCount
    End If
        
    vsExt.Clear
    vsExt.Rows = IIF(i = 0, 2, i + 1)
    vsExt.Cols = 2
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 0) = "��������"
    vsExt.TextMatrix(0, 1) = "��ģ"
    vsExt.ColWidth(0) = 3200: vsExt.ColWidth(1) = 800
    vsExt.FixedAlignment(0) = 4: vsExt.FixedAlignment(1) = 4
    vsExt.ColAlignment(0) = 1: vsExt.ColAlignment(1) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If str����IDs <> "" And i <> 0 Then
        arr����IDs = Split(str����IDs, ",") '����ԭ������˳��
        For i = 0 To UBound(arr����IDs)
            rsTmp.Filter = "ID=" & CStr(arr����IDs(i))
            If Not rsTmp.EOF Then
                j = j + 1
                vsExt.RowData(j) = CLng(rsTmp!ID)
                vsExt.TextMatrix(j, 0) = "[" & rsTmp!���� & "]" & rsTmp!����
                vsExt.Cell(flexcpData, j, 0) = vsExt.TextMatrix(j, 0) '���ڻָ���ʾ
                vsExt.TextMatrix(j, 1) = Nvl(rsTmp!��������, 0)
            End If
        Next
    End If
    
    '������Ŀ
    If lng����ID <> 0 Then
        strSQL = "Select A.ID,A.����,A.����,�������� From ������ĿĿ¼ A Where A.���='G' And A.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
        If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
        If Not rsTmp.EOF Then
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!���� & "]" & rsTmp!����
            cmdData.Tag = txtData.Text '���ڻָ���ʾ
        End If
    End If
    
    vsExt.Row = 1: vsExt.Col = 0
    Init������Ŀ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init������() As Boolean
'���ܣ���ʼ����鲿λ����ʽ������
'������mstrExtData=������鲿λ����Ϣ,Ϊ��ʱ��ʾ�������������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strPosition As String, arrPosition As Variant
    
    On Error GoTo errH
    
    If Not Visible Then
        strPosition = mstrExtData
    Else
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 Then
                If vsExt.Cols = 3 Then
                    If Val(vsExt.TextMatrix(i, 2)) <> 0 Then
                        strPosition = strPosition & "," & vsExt.RowData(i)
                    End If
                Else
                    strPosition = strPosition & "," & vsExt.RowData(i)
                End If
            End If
        Next
        strPosition = Mid(strPosition, 2)
    End If
    
    '�������õĲ�λ˳���
    strSQL = "Select A.���,A.����,A.����,A.�걾��λ,B.������ĿID" & _
        " From ������ĿĿ¼ A,������Ŀ��� B" & _
        " Where A.ID=B.������ĿID And B.�������ID=[1]" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " And A.������� IN([2],3)" & _
        " Order by B.���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng��ĿID, mint�������)
    If rsTmp.EOF Then
        MsgBox "�ü�������Ŀ��ǰ����Ч��λ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsExt
        .FixedRows = 0: .FixedCols = 0
        .Rows = 0: .Cols = 0
        If optPosition(0).Value Then '��λѡ��ģʽ
            .Rows = IIF(rsTmp.EOF, 2, rsTmp.RecordCount + 1)
            .FixedRows = 1: .Cols = 3: .FixedCols = 0
            
            .TextMatrix(0, 0) = "�����Ŀ"
            .TextMatrix(0, 1) = "��鲿λ"
            .TextMatrix(0, 2) = "ѡ��"
            .FixedAlignment(0) = 4: .ColAlignment(0) = 1: .ColWidth(0) = 2000
            .FixedAlignment(1) = 4: .ColAlignment(1) = 1: .ColWidth(1) = 1500
            .FixedAlignment(2) = 4: .ColAlignment(2) = 4: .ColWidth(2) = 500
            .ColDataType(2) = flexDTBoolean
            .Editable = flexEDKbdMouse
            
            For i = 1 To rsTmp.RecordCount
                .RowData(i) = CLng(rsTmp!������ĿID) 'һ��Ҫ��ȷ����
                .TextMatrix(i, 0) = rsTmp!���� '"[" & rsTmp!���� & "]" & rsTmp!����
                .TextMatrix(i, 1) = Nvl(rsTmp!�걾��λ)
                If InStr("," & strPosition & ",", "," & rsTmp!������ĿID & ",") > 0 Then
                    .TextMatrix(i, 2) = 1
                End If
                rsTmp.MoveNext
            Next
                
            .Row = 1: .Col = 2
        Else '��λ����ģʽ
            arrPosition = Split(strPosition, ",")
            .Rows = 1 + (UBound(arrPosition) + 1) + 1
            .FixedRows = 1: .Cols = 2: .FixedCols = 0
            
            .TextMatrix(0, 0) = "�����Ŀ"
            .TextMatrix(0, 1) = "��鲿λ"
            .FixedAlignment(0) = 4: .ColAlignment(0) = 1: .ColWidth(0) = 2000
            .FixedAlignment(1) = 4: .ColAlignment(1) = 1: .ColWidth(1) = 2000
            .Editable = flexEDKbdMouse
                        
            For i = 0 To UBound(arrPosition)
                rsTmp.Filter = "������ĿID=" & arrPosition(i)
                If Not rsTmp.EOF Then
                    .RowData(i + 1) = CLng(rsTmp!������ĿID)
                    .TextMatrix(i + 1, 0) = rsTmp!����
                    .TextMatrix(i + 1, 1) = Nvl(rsTmp!�걾��λ)
                    .Cell(flexcpData, i + 1, 1) = .TextMatrix(i + 1, 1) '���ڻָ���ʾ
                End If
            Next
            
            rsTmp.Filter = 0
            .TextMatrix(.Rows - 1, 0) = rsTmp!����
            .Row = .Rows - 1: .Col = .Cols - 1
        End If
        .ShowCell .Row, .Col
        .LeftCol = 0 'Ҫ����ShowCell��,��Ȼѡ��ģʽ������
    End With
    
    Init������ = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init����걾() As Boolean
'���ܣ���ʼ������걾����ʽ���걾����
'������mstrExtData=����ȱʡ�ļ���걾����Ϣ,Ϊ��ʱ��ʾ�����������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    Dim strTmp As String, lngItemCount As Long
    Dim aTmp() As String, strSample As String, blnChecked As Boolean
    
    On Error GoTo errH
    
    If Len(mstrExtData) > 0 Then
        aTmp = Split(mstrExtData, ";")
        strTmp = aTmp(0)
        lngItemCount = UBound(Split(strTmp, ",")) + 1
        If UBound(aTmp) > 0 Then strSample = aTmp(1)
    End If
    If lngItemCount = 0 Then
        strSQL = "Select ���� From ���Ƽ���걾"
    Else
        strSQL = _
            " Select �걾����,Sum(1) From (" & _
            "   Select Distinct A.ID,B.���� As �걾����" & _
            "   From ������ĿĿ¼ A,���Ƽ���걾 B,������Ŀ�ο� C,���鱨����Ŀ D" & _
            "   Where A.ID=D.������ĿID(+) And D.������ĿID=C.��ĿID(+)" & _
            "       And (C.�걾���� Is Null Or C.�걾����=B.����) And A.ID In (" & strTmp & ")" & _
            " ) Group By �걾���� Having Sum(1)=" & lngItemCount
    End If
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
    If rsTmp.EOF Then
        MsgBox Switch(lngItemCount = 0, "δ���ü���걾���뵽�ֵ�����������á�", _
            lngItemCount = 1, "ѡȡ�ļ�����Ŀδ�������걾�����ȵ�������Ŀ����������", _
            lngItemCount > 1, "ѡȡ�ļ�����Ŀ�ļ���걾��һ�£����ȵ�������Ŀ����������"), vbInformation, gstrSysName
        Exit Function
    ElseIf rsTmp.RecordCount = 1 And mstrExtData = "" Then
        '��������Ŀʱ,���ֻ��һ���걾ʱ,ֱ��ѡ���˳�
        mstrExtData = rsTmp(0)
        mblnOK = True: Exit Function
    End If
    
    vsExt.Clear
    vsExt.Rows = IIF(rsTmp.EOF, 2, rsTmp.RecordCount + 1)
    vsExt.FixedRows = 1: vsExt.Cols = 2: vsExt.FixedCols = 0
    vsExt.Row = 1: vsExt.Col = 1
    
    vsExt.TextMatrix(0, 0) = "����걾"
    vsExt.TextMatrix(0, 1) = "ѡ��"
    vsExt.FixedAlignment(0) = 4: vsExt.ColAlignment(0) = 1: vsExt.ColWidth(0) = 3500
    vsExt.FixedAlignment(1) = 4: vsExt.ColAlignment(1) = 4: vsExt.ColWidth(1) = 500
    vsExt.ColDataType(1) = flexDTBoolean
    vsExt.Editable = flexEDKbdMouse
    
    For i = 1 To rsTmp.RecordCount
        vsExt.TextMatrix(i, 0) = rsTmp(0)
        If strSample = vsExt.TextMatrix(i, 0) Then
            vsExt.TextMatrix(i, 1) = 1
            vsExt.Row = i
            blnChecked = True
        End If
        rsTmp.MoveNext
    Next
    If Not blnChecked Then
        vsExt.TextMatrix(1, 1) = 1
        vsExt.Row = 1
    End If
    
    Init����걾 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init�������() As Boolean
'���ܣ���ʼ��������Ŀ
'������mstrExtData=����ȱʡ�ļ�����Ŀ����Ϣ,Ϊ��ʱ��ʾ�����������Ŀ
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim arrItems As Variant, strItems As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSQL = mstrExtData
    If strSQL = "" Then strSQL = ";"
    strItems = CStr(Split(strSQL, ";")(0))
    Me.txtData = Split(strSQL, ";")(1)
    cmdData.Tag = txtData.Text
    
    If strItems <> "" Then
        If mlng��ĿID > 0 Then 'ָ�������Ƶ���
            strSQL = "Select A.* From ������ĿĿ¼ A,���Ƶ���Ӧ�� B " & _
                " Where A.ID=B.������ĿID And B.Ӧ�ó���=" & mint������� & " And B.�����ļ�ID=" & mlng��ĿID & _
                " And A.���='C' And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) In (0" & IIF(Len(Trim(mstr�Ա�)) = 0, ") ", IIF(mstr�Ա� Like "*��*", ",1) ", ",2) ")) & _
                " And A.������� IN(" & mint������� & ",3) And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And A.ID In(" & strItems & ")" & _
                " Order by A.����"
        Else
            strSQL = "Select A.*,B.�����ļ�ID From ������ĿĿ¼ A,���Ƶ���Ӧ�� B " & _
                " Where A.ID=B.������ĿID(+)" & _
                " And A.���='C' And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) In (0" & IIF(Len(Trim(mstr�Ա�)) = 0, ") ", IIF(mstr�Ա� Like "*��*", ",1) ", ",2) ")) & _
                " And A.������� IN(" & mint������� & ",3) And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And A.ID In(" & strItems & ")" & _
                " Order by A.����"
'                " And (B.������ĿID is Null Or B.Ӧ�ó���=" & mint������� & ")"
        End If
        zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
        i = rsTmp.RecordCount
        If i > 0 And mlng��ĿID = 0 Then mlng��ĿID = Nvl(rsTmp("�����ļ�ID"), 0)
    End If
        
    vsExt.Clear
    vsExt.Rows = IIF(i = 0, 2, i + 1)
    vsExt.Cols = 2
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 0) = "������Ŀ"
    vsExt.ColWidth(0) = 4000: vsExt.ColHidden(1) = True
    vsExt.FixedAlignment(0) = 4
    vsExt.ColAlignment(0) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If i > 0 Then
        arrItems = Split(strItems, ",") '����ԭ������˳��
        For i = 0 To UBound(arrItems)
            rsTmp.Filter = "ID=" & arrItems(i)
            If Not rsTmp.EOF Then
                j = j + 1
                vsExt.RowData(j) = CLng(rsTmp!ID)
                vsExt.TextMatrix(j, 0) = "[" & rsTmp!���� & "]" & rsTmp!����
                vsExt.Cell(flexcpData, j, 0) = vsExt.TextMatrix(j, 0) '���ڻָ���ʾ
                vsExt.TextMatrix(j, 1) = Nvl(rsTmp!��������)
                If rsTmp("��������") = "΢����" Then mblnNotAddNew = True '΢����ֻ�ܿ�һ��������Ŀ
            End If
        Next
    End If
    
    vsExt.Row = 1: vsExt.Col = 0
    Init������� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitCombox(Optional ByVal strNewItemID As String = "", Optional ByVal DefaultValue As String = "") As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strTmp As String, lngItemCount As Long
    InitCombox = False
    
    On Error GoTo DBError
    strTmp = "": lngItemCount = 0
    For i = 1 To vsExt.Rows - 1
        If vsExt.RowData(i) <> 0 And (i <> vsExt.Row Or Len(strNewItemID) = 0) Then
            lngItemCount = lngItemCount + 1
            strTmp = strTmp & "," & vsExt.RowData(i)
        End If
    Next
    If Len(strNewItemID) > 0 Then
        lngItemCount = lngItemCount + 1
        strTmp = strTmp & "," & strNewItemID
    End If
    If Len(strTmp) > 0 Then strTmp = Mid(strTmp, 2)

    If lngItemCount = 0 Then
        strSQL = "Select ���� From ���Ƽ���걾"
    Else
        strSQL = "Select �걾����,Sum(1) From (" & _
            "   Select Distinct A.ID,B.���� As �걾����" & _
            "   From ������ĿĿ¼ A,���Ƽ���걾 B,������Ŀ�ο� C,���鱨����Ŀ D" & _
            "   Where A.ID=D.������ĿID(+) And D.������ĿID=C.��ĿID(+)" & _
            "       And (C.�걾���� Is Null Or C.�걾����=B.����) And A.ID In (" & strTmp & ")" & _
            " ) Group By �걾���� Having Sum(1)=" & lngItemCount
    End If
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
    If rsTmp.EOF Then
        MsgBox Switch(lngItemCount = 0, "δ���ü���걾���뵽�ֵ�����������á�", _
            lngItemCount = 1, "ѡȡ�ļ�����Ŀδ�������걾�����ȵ�������Ŀ����������", _
            lngItemCount > 1, "ѡȡ�ļ�����Ŀ�ļ���걾��������Ŀ�Ĳ�һ�£����ȵ�������Ŀ����������"), vbInformation, gstrSysName
        Exit Function
    End If
    
    With cbo�걾
        strTmp = .Text
        
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp(0)
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
        On Error Resume Next
        If Len(DefaultValue) > 0 Then
            .Text = DefaultValue
        Else
            .Text = strTmp
        End If
    End With
    InitCombox = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Height - cmdCancel.Width
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    mlngHwnd = 0
    mint��Ч = 0
    mstr�Ա� = ""
    mintType = 0
    mlng��ĿID = 0
    mint������� = 0
    mbln��ʿվ = False
    mblnҽ�� = False
    
    SaveSetting "ZLSOFT", "����ģ��\" & App.ProductName, "��λȷ����ʽ", IIF(optPosition(0).Value, 0, 1)
End Sub

Private Sub optPosition_Click(Index As Integer)
    If Visible Then
        Call Init������: vsExt.SetFocus
        optPosition(0).TabStop = False: optPosition(1).TabStop = False '��Ȼ��Ч
    End If
End Sub

Private Sub txtData_GotFocus()
    zlControl.TxtSelAll txtData
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset, vRect As RECT
    Dim strSQL As String, str�Ա� As String
    Dim strLike As String, blnCancel As Boolean
    
    If mstr�Ա� Like "*��*" Then
        str�Ա� = "0,1"
    ElseIf mstr�Ա� Like "*Ů*" Then
        str�Ա� = "0,2"
    Else
        str�Ա� = "0"
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtData.Text = "" Then
            If mintType = 1 Then '�������Բ�����������Ŀ
                Call zlCommFun.PressKey(vbKeyTab)
            End If
            Exit Sub
        ElseIf txtData.Text = cmdData.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        '�Ż�
        strLike = mstrLike
        If Len(txtData.Text) < 2 Then strLike = ""
        
        If mintType = 1 Then
            '����������Ŀ
            strSQL = _
                " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��������" & _
                " From ������ĿĿ¼ A,������Ŀ���� B" & _
                " Where A.ID=B.������ĿID And A.���='G' And A.������� IN([3],3)" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]) And B.����=[4]" & _
                " Order by A.����"
            vRect = GetControlRect(txtData.Hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������Ŀ", False, "", "", False, False, True, vRect.Left, vRect.Top, txtData.Height, blnCancel, False, True, _
                UCase(txtData.Text) & "%", strLike & UCase(txtData.Text) & "%", mint�������, mint���� + 1)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                txtData.Text = cmdData.Tag
                zlControl.TxtSelAll txtData
                Exit Sub
            End If
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!���� & "]" & rsTmp!����
            cmdData.Tag = txtData.Text
            
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf mintType = 4 Then
            '����걾
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmdData_Click
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtData_Validate(Cancel As Boolean)
'���ܣ��ָ���ʾԭ����
    If txtData.Text <> cmdData.Tag Then
        txtData.Text = cmdData.Tag
    End If
End Sub

Private Sub vsExt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'����:��ʾѡ��ť,����֤��ǰ��Ԫ��ɼ�
    
    '��֤��ǰ��Ԫ��ɼ�
    If NewRow >= vsExt.FixedRows And NewRow <= vsExt.Rows - 1 Then
        If vsExt.LeftCol >= vsExt.FixedCols And vsExt.LeftCol <= vsExt.Cols - 1 Then
            Call vsExt.ShowCell(NewRow, vsExt.LeftCol)
        End If
    End If
    
    If mintType = 0 And optPosition(1).Value Then
        If NewCol = 1 Then
            cmd.Height = vsExt.CellHeight - 30
            cmd.Left = vsExt.CellLeft + vsExt.CellWidth - cmd.Width - 15
            cmd.Top = vsExt.CellTop + 15
            
            cmd.Visible = True
        Else
            cmd.Visible = False
        End If
    ElseIf mintType = 1 Or mintType = 4 Then
        '��ʾ/��������ѡ��ť
        If NewCol = 0 Then
            cmd.Height = vsExt.CellHeight - 30
            cmd.Left = vsExt.CellLeft + vsExt.CellWidth - cmd.Width - 15
            cmd.Top = vsExt.CellTop + 15
            
            cmd.Visible = True
        Else
            cmd.Visible = False
        End If
    End If
End Sub

Private Sub vsExt_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'����:����ĳЩ�п�ķ�Χ
    If Row = -1 Then
        If mintType = 0 Then
            'ѡ���п�Ȳ���
            If optPosition(0).Value Then
                If 3500 - vsExt.ColWidth(0) <= 0 Then vsExt.ColWidth(0) = 3000
                vsExt.ColWidth(1) = 3500 - vsExt.ColWidth(0)
            Else
                If 4000 - vsExt.ColWidth(0) <= 0 Then vsExt.ColWidth(0) = 2000
                vsExt.ColWidth(1) = 4000 - vsExt.ColWidth(0)
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) 'ʹ��ť�ɼ���������ťλ��
        End If
    End If
End Sub

Private Sub vsExt_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    '��λ����겻�ɽ���
    If mintType = 2 And Button = 1 And vsExt.MouseCol Mod 4 = 2 Then Cancel = True
End Sub

Private Sub vsExt_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mintType = 2 Then
        '��λ�а������ɽ���
        If NewCol Mod 4 = 2 Then
            Cancel = True
            If OldCol > NewCol Then '�����ƶ�ʱ����
                vsExt.Col = NewCol - 1
            Else
                vsExt.Col = NewCol + 1
            End If
            vsExt.Row = NewRow
        End If
    End If
End Sub

Private Sub vsExt_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If cmd.Visible Then cmd.Visible = False
End Sub

Private Sub vsExt_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ�����ĳЩ�в��ܸı��п�
    If Row = -1 Then
        If mintType = 0 Then
            'ֻ����ı�ǰ�����п�
            If Col <> 0 Then Cancel = True
        ElseIf mintType = 3 Then
            Cancel = True
        End If
    End If
End Sub

Private Sub vsExt_GotFocus()
    Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) 'ʹ��ť�ɼ�
End Sub

Private Sub vsExt_KeyDown(KeyCode As Integer, Shift As Integer)
'���ܣ�ɾ��������
    Dim i As Long, j As Long, k As Long
    
    If KeyCode = vbKeyDelete Then
        If (mintType = 0 And optPosition(1).Value Or mintType = 1 Or mintType = 4) And vsExt.RowData(vsExt.Row) <> 0 Then
            If MsgBox("Ҫɾ����ǰ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            vsExt.RowData(vsExt.Row) = 0
            If mintType = 0 Then
                vsExt.TextMatrix(vsExt.Row, vsExt.Cols - 1) = ""
                vsExt.Cell(flexcpData, vsExt.Row, vsExt.Cols - 1) = ""
            Else
                For i = 0 To vsExt.Cols - 1
                    vsExt.TextMatrix(vsExt.Row, i) = ""
                    vsExt.Cell(flexcpData, vsExt.Row, i) = ""
                Next
            End If
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And vsExt.Row = vsExt.FixedRows) Then
                vsExt.RemoveItem vsExt.Row
            End If
            
            '���³�ʼ�걾
            If mintType = 4 Then InitCombox
        ElseIf mintType = 2 Then
            If CLng(vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)) <> 0 Then
                If MsgBox("Ҫɾ��""" & vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4) & """��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
                '�����ǰζҩ��Ϣ
                For i = 0 To 3
                    vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + i) = ""
                    vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + i) = Empty
                Next
                '�����������ǰ��
                For i = vsExt.Row To vsExt.Rows - 1
                    For j = 0 To vsExt.Cols - 1 Step 4
                        If Not (i = vsExt.Row And j <= (vsExt.Col \ 4) * 4) Then
                            For k = 0 To 3
                                If j = 0 Then
                                    vsExt.TextMatrix(i - 1, vsExt.Cols - (4 - k)) = vsExt.TextMatrix(i, j + k)
                                    vsExt.Cell(flexcpData, i - 1, vsExt.Cols - (4 - k)) = vsExt.Cell(flexcpData, i, j + k)
                                Else
                                    vsExt.TextMatrix(i, j + k - 4) = vsExt.TextMatrix(i, j + k)
                                    vsExt.Cell(flexcpData, i, j + k - 4) = vsExt.Cell(flexcpData, i, j + k)
                                End If
                                vsExt.TextMatrix(i, j + k) = ""
                                vsExt.Cell(flexcpData, i, j + k) = Empty
                            Next
                        End If
                    Next
                Next
                'ɾ������Ŀ���(���ٱ������������ʾ������7)
                If vsExt.Rows > 7 Then
                    For i = vsExt.Rows - 1 To 7 Step -1
                        If CLng(vsExt.Cell(flexcpData, i - 1, 2)) = 0 Then
                            vsExt.RemoveItem i
                        End If
                    Next
                End If
                Call vsExt.ShowCell(vsExt.Row, vsExt.Col)
            End If
        End If
    End If
End Sub

Private Sub vsExt_LostFocus()
    If Not ActiveControl Is cmd Then cmd.Visible = False
End Sub

Private Sub vsExt_KeyPress(KeyAscii As Integer)
'���ܣ��Ǳ༭״̬ʱ���Զ��ƶ���Ԫ��
    If KeyAscii = 13 Then
        KeyAscii = 0
        '��λ����һӦ���뵥Ԫ��
        If mintType = 0 Then
            If vsExt.Col <> vsExt.Cols - 1 Then
                vsExt.Col = vsExt.Cols - 1
            ElseIf vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = vsExt.Cols - 1
            Else
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            If vsExt.Row = vsExt.Rows - 1 Then
                If vsExt.RowData(vsExt.Row) = 0 Or mblnNotAddNew Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                Else
                    vsExt.AddItem ""
                End If
            End If
            If vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = 0
            End If
        ElseIf mintType = 2 Then
            If CLng(vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)) = 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            Else
                Call EnterNextCell(vsExt.Row, vsExt.Col)
            End If
        ElseIf mintType = 3 Then
            If vsExt.Col <> 1 Then
                vsExt.Col = 1
            ElseIf vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = 1
            Else
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        If mintType = 0 Or mintType = 1 Or mintType = 4 Then
            KeyAscii = 0
            If cmd.Visible Then cmd_Click
        ElseIf mintType = 2 Then
            KeyAscii = 0
            cmd_Click 'ѡ��ζ�в�ҩ���ע
        End If
    End If
End Sub

Private Sub vsExt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'���ܣ��ǻس�ȷ�����༭�Ĵ���(����Text:=EditText,��ValidateEdit�¼��л�û��)
    Dim i As Long
    If Not mblnReturn Then
        If mintType = 0 And optPosition(1).Value Then
            If Col = vsExt.Cols - 1 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            If Col = 0 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                            
                '���³�ʼ�걾
                If mintType = 4 Then InitCombox
            End If
        ElseIf mintType = 2 Then
            If Col Mod 4 = 0 Then '��ҩ
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
            ElseIf Col Mod 4 = 1 Then '��ζ����
                If Not IsNumeric(vsExt.TextMatrix(Row, Col)) _
                    Or Val(vsExt.TextMatrix(Row, Col)) <= 0 _
                    Or Val(vsExt.TextMatrix(Row, Col)) > LONG_MAX Then
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Else
                    vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
                End If
            ElseIf Col Mod 4 = 3 Then '��ע
                If zlCommFun.ActualLen(vsExt.TextMatrix(Row, Col)) > 100 Then
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Else
                    vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
                End If
            End If
        ElseIf mintType = 3 Then
            'ȡ�������걾ѡ��(��ѡ)
            If Val(vsExt.TextMatrix(Row, 1)) <> 0 Then
                For i = vsExt.FixedRows To vsExt.Rows - 1
                    If i <> Row And Val(vsExt.TextMatrix(i, 1)) <> 0 Then
                        vsExt.TextMatrix(i, 1) = 0
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub vsExt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'���ܣ���������ȷ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, int�Ա� As Integer, strҩƷ As String
    Dim strStock As String, blnCancel As Boolean, i As Long
    Dim vPoint As POINTAPI, strSamples As String, strLike As String
    
    If mstr�Ա� Like "*��*" Then
        int�Ա� = 1
    ElseIf mstr�Ա� Like "*Ů*" Then
        int�Ա� = 2
    End If
    
    If KeyAscii = 13 Then
        mblnReturn = True '����ǰ��س�ȷ�ϱ༭
        KeyAscii = 0
        
        '�Ż�
        strLike = mstrLike
        If Len(vsExt.EditText) < 2 Then strLike = ""
        
        On Error GoTo errH
        
        If mintType = 0 And optPosition(1).Value Then
            '�����鲿λ
            strSQL = _
                "Select A.ID, A.����, A.�걾��λ As ��鲿λ" & vbNewLine & _
                "From ������ĿĿ¼ A, ������Ŀ��� B" & vbNewLine & _
                "Where A.ID = B.������Ŀid And B.�������id = [1] And A.������� In ([4], 3) And" & vbNewLine & _
                "      (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) And" & vbNewLine & _
                "      (A.���� Like [2] Or Upper(A.�걾��λ) Like [3] Or " & IIF(mint���� = 0, "zlSpellCode", "zlWBCode") & "(A.�걾��λ) Like [3])" & vbNewLine & _
                "Order By B.���"
            vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��λ", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                mlng��ĿID, UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mint�������)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            '����ظ�����
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "�ü�鲿λ�Ѿ���������¼�롣", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            Call Set��λ����(Row, rsTmp)
        ElseIf mintType = 1 Then
            '���븽������:���ﲻ�ǵ���Ӧ��,��˲�����
            strSQL = _
                " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,A.�������� as ��ģ" & _
                " From ������ĿĿ¼ A,������Ŀ���� B" & _
                " Where A.ID=B.������ĿID And A.���='F' And A.ID<>[3]" & IIF(strLike = "", "", " And Rownum<=100") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                    " And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]) And B.����=[4]" & _
                    " And A.������� IN([5],3) And Nvl(A.ִ��Ƶ��,0) IN(0,[6]) And Nvl(A.�����Ա�,0) IN(0,[7])" & _
                " Order by A.����"
            vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "����", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mlng��ĿID, mint���� + 1, mint�������, IIF(mint��Ч = 0, 2, 1), int�Ա�)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            '����ظ�����
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "�ø��������Ѿ���������¼�롣", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            Call Set��������(Row, rsTmp)
        ElseIf mintType = 2 Then
            '��ȡ�س���,�����MsgboxʹEdit���㶪ʧ,�����ɱ༭,�����ἤ��AfterEdit�¼�
            If Col Mod 4 = 0 Then '��ҩ
                '��ҩ���,��ҩ��δָ��ʱ,����������¼
                If mlng��ҩ�� <> 0 Then
                    strStock = _
                        "Select ҩƷID,Sum(Nvl(��������,0)) as ��� From ҩƷ���" & _
                        " Where (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч��>Trunc(Sysdate))" & _
                        " And ����=1 And �ⷿID=[3] Group by ҩƷID" & _
                        " Having Sum(Nvl(��������,0))<>0"
                End If
                
                '����ҩƷȨ��
                strҩƷ = ""
                If InStr(mstrPrivs, "�´�����ҩ��") = 0 Then
                    strҩƷ = strҩƷ & " And E.�������<>'����ҩ'"
                End If
                If InStr(mstrPrivs, "�´ﶾ��ҩ��") = 0 Then
                    strҩƷ = strҩƷ & " And E.�������<>'����ҩ'"
                End If
                If InStr(mstrPrivs, "�´����ҩ��") = 0 Then
                    strҩƷ = strҩƷ & " And E.��ֵ���� Not IN('����','����')"
                End If
                
                '���뵥ζ��ҩ:���ﲻ�ǵ���Ӧ��,��˲�����
                strSQL = "Select A.ID,A.����,A.����,A.���㵥λ" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B" & _
                    " Where A.ID=B.������ĿID And A.���='7'" & _
                    " And (A.���� Like [1] And B.����=[5] Or B.���� Like [2] And B.����=[5] Or B.���� Like [2] And B.���� IN([5],3))" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) And A.������� IN([4],3)" & _
                    " And Nvl(A.ִ��Ƶ��,0) IN(0,[6]) And Nvl(A.�����Ա�,0) IN(0,[7])"
                If strLike = "" And strStock <> "" Then
                    '���������ü�������ʱ(����ƥ��),�����(+)����(ҩƷ���),����ҪGroup Byһ��(���)
                    strSQL = strSQL & " Group by A.ID,A.����,A.����,A.���㵥λ"
                End If
                If strStock = "" Then
                    strSQL = _
                        " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,D.���,D.����,NULL AS ���,E.����ְ�� as ����ְ��ID" & _
                        " From ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,(" & strSQL & ") A" & _
                        " Where A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID" & _
                            " And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([4],3)" & _
                            IIF(strLike = "", "", " And Rownum<=100") & strҩƷ & _
                        " Order by A.����"
                Else
                    strSQL = _
                        " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ,D.���,D.����," & _
                        " Decode(X.���,NULL,NULL,X.���/C.סԺ��װ||C.סԺ��λ) AS ���,E.����ְ�� as ����ְ��ID" & _
                        " From ҩƷ���� E,ҩƷ��� C,�շ���ĿĿ¼ D,(" & strSQL & ") A,(" & strStock & ") X" & _
                        " Where A.ID=E.ҩ��ID And A.ID=C.ҩ��ID And C.ҩƷID=D.ID And C.ҩƷID=X.ҩƷID(+)" & _
                            " And (D.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or D.����ʱ�� IS NULL) And D.������� IN([4],3)" & _
                            IIF(strLike = "", "", " And Rownum<=100") & strҩƷ & _
                        " Order by A.����"
                End If
                vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ҩ", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                    UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mlng��ҩ��, mint�������, mint���� + 1, IIF(mint��Ч = 0, 2, 1), int�Ա�)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                    End If
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                
                '����ظ�����
                If ItemExist(rsTmp!ID, Row, Col) Then
                    MsgBox "��ζ��ҩ���䷽���Ѿ�¼�롣", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                
                '����ְ����
                If Not mbln��ʿվ Then
                    strSQL = CheckOneDuty(rsTmp!����, Nvl(rsTmp!����ְ��ID), UserInfo.����, mblnҽ��)
                    If strSQL <> "" Then
                        MsgBox strSQL, vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Exit Sub
                    End If
                End If
                
                '��ȡ����ֵ
                vsExt.EditText = rsTmp!���� 'ֱ������ƥ��ʱ��Ҫ
                vsExt.TextMatrix(Row, Col) = rsTmp!����
                vsExt.TextMatrix(Row, Col + 2) = rsTmp!��λ
                vsExt.Cell(flexcpData, Row, Col + 2) = CLng(rsTmp!ID) '��¼��ҩID
            ElseIf Col Mod 4 = 1 Then '����
                If Not IsNumeric(vsExt.EditText) Or Val(vsExt.EditText) <= 0 Or Val(vsExt.EditText) > LONG_MAX Then
                    MsgBox "��ζ����������󣬲��Ǵ���������ֻ�������ֵ����", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                vsExt.TextMatrix(Row, Col) = vsExt.EditText
            ElseIf Col Mod 4 = 3 Then '��ע
                If vsExt.EditText <> "" Then
                    strSQL = "Select Rownum as ID,����,����,���� From ��ҩ�����ע" & _
                        " Where Upper(����) Like [1] Or Upper(����) Like [2] Or Upper(����) Like [2]" & _
                        " Order by ����"
                    vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��ע", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                        UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%")
                End If
                If rsTmp Is Nothing Then
                    If blnCancel Then
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Exit Sub
                    End If
                    '��ƥ�䵱��ֱ������
                    If zlCommFun.ActualLen(vsExt.EditText) > 100 Then
                        MsgBox "��ע�������ݹ��������ֻ���� 50 �����ֻ� 100 ���ַ���", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Exit Sub
                    End If
                    vsExt.TextMatrix(Row, Col) = vsExt.EditText
                Else
                    vsExt.EditText = rsTmp!���� 'ֱ������ƥ��ʱ��Ҫ
                    vsExt.TextMatrix(Row, Col) = rsTmp!����
                End If
            End If
            vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
            Call EnterNextCell(Row, Col)
        ElseIf mintType = 4 Then
            '������Ŀ
            With Me.cbo�걾
                For i = 0 To .ListCount - 1
                    strSamples = strSamples & ",'" & .List(i) & "'"
                Next
            End With
            If Len(strSamples) > 0 Then
                strSamples = Mid(strSamples, 2)
            Else
                strSamples = "''"
            End If
            strSQL = "Select A.ID,A.����,A.����,A.��������,A.�걾��λ" & _
                " From ������ĿĿ¼ A,������Ŀ���� C Where A.ID=C.������ĿID" & _
                " And (A.���� Like [1] Or C.���� Like [2] Or C.���� Like [2]) And C.����=[3]" & _
                " And A.���='C' And Nvl(A.����Ӧ��,0)=1 And Nvl(A.�����Ա�,0) In (0,[6])" & _
                " And A.������� IN([4],3) And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)"
            If strLike = "" Then
                '���������ü�������ʱ(����ƥ��),�����(+)����,����ҪGroup Byһ��(���)
                strSQL = strSQL & " Group by A.ID,A.����,A.����,A.��������,A.�걾��λ"
            End If
            If mlng��ĿID > 0 Then 'ָ�������Ƶ���
                strSQL = "Select Distinct A.ID,A.����,A.����,A.�������� as ��������,A.�걾��λ" & _
                    " From ���Ƶ���Ӧ�� B,������Ŀ�ο� D,���鱨����Ŀ E,(" & strSQL & ") A" & _
                    " Where A.ID=B.������ĿID And A.ID=E.������Ŀid(+) And E.������ĿID=D.��Ŀid(+)" & _
                    " And B.Ӧ�ó���+0=[4] And B.�����ļ�ID+0=[5]" & _
                    " And (D.�걾���� In (" & strSamples & ") Or D.�걾���� Is Null)" & _
                    " Order by A.����"
            Else
                strSQL = "Select Distinct A.ID,A.����,A.����,A.�������� as ��������,B.�����ļ�ID,A.�걾��λ" & _
                    " From ���Ƶ���Ӧ�� B,������Ŀ�ο� D,���鱨����Ŀ E,(" & strSQL & ") A" & _
                    " Where A.ID=B.������ĿID(+) And A.ID=E.������Ŀid(+) And E.������ĿID=D.��Ŀid(+)" & _
                    " And 0+B.Ӧ�ó���(+)=[4] And (D.�걾���� In (" & strSamples & ") Or D.�걾���� Is Null)" & _
                    " Order by A.����"
            End If
            vPoint = GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������Ŀ", False, "", "", False, False, True, vPoint.x, vPoint.y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mint���� + 1, mint�������, mlng��ĿID, int�Ա�)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "δ�ҵ�ƥ����Ŀ��", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            If rsTmp("��������") = "΢����" And vsExt.Rows > 2 Then
                If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '��������ֻ�ܿ�һ��΢������Ŀ
                    MsgBox "΢������Ŀֻ�ܵ������룡", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                    Exit Sub
                End If
            End If
            If mlng��ĿID = 0 Then mlng��ĿID = Nvl(rsTmp("�����ļ�ID"), 0)
            
            '����ظ�����
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "�ü�����Ŀ�Ѿ�¼�룡", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            '�����������Ƿ���ͬ
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And i <> Row Then
                    If Not (vsExt.TextMatrix(i, 1) = Nvl(rsTmp!��������) _
                        Or vsExt.TextMatrix(i, 1) = "" Or Nvl(rsTmp!��������) = "") Then
                        MsgBox "��������ͬ�������͵���Ŀ����������Ŀ�ļ�������Ϊ""" & vsExt.TextMatrix(i, 1) & """��", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                        Exit Sub
                    End If
                End If
            Next
            
            '���³�ʼ�걾
            If Not InitCombox(rsTmp("ID"), Nvl(rsTmp("�걾��λ"))) Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '����ʹ��ť�ɼ�
                Exit Sub
            End If
            
            Call Set������Ŀ(Row, rsTmp)
            If rsTmp("��������") = "΢����" Then
                mblnNotAddNew = True
                vsExt.Rows = 2
            Else
                mblnNotAddNew = False
            End If
        End If
    Else
        If mintType = 2 Then
            '��ζ����ֻ������������
            If Col Mod 4 = 1 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set��λ����(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    vsExt.EditText = rsInput!��鲿λ '��������ֱ��ƥ��ʱ�б�Ҫ
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, vsExt.Cols - 1) = rsInput!��鲿λ
    vsExt.Cell(flexcpData, lngRow, vsExt.Cols - 1) = vsExt.TextMatrix(lngRow, vsExt.Cols - 1)
    
    '��һ������
    If vsExt.RowData(vsExt.Rows - 1) <> 0 Then
        vsExt.AddItem ""
        vsExt.TextMatrix(vsExt.Rows - 1, 0) = vsExt.TextMatrix(vsExt.Rows - 2, 0)
    End If
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = vsExt.Cols - 1
End Sub

Private Sub Set��������(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    '��������
    vsExt.EditText = "[" & rsInput!���� & "]" & rsInput!���� '��������ֱ��ƥ��ʱ�б�Ҫ
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 0) = "[" & rsInput!���� & "]" & rsInput!����
    vsExt.Cell(flexcpData, lngRow, 0) = vsExt.TextMatrix(lngRow, 0)
    vsExt.TextMatrix(lngRow, 1) = Nvl(rsInput!��ģ)
    
    '��һ������
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 0
End Sub

Private Sub Set������Ŀ(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    '������Ŀ
    vsExt.EditText = "[" & rsInput!���� & "]" & rsInput!���� '��������ֱ��ƥ��ʱ�б�Ҫ
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 0) = "[" & rsInput!���� & "]" & rsInput!����
    vsExt.Cell(flexcpData, lngRow, 0) = vsExt.TextMatrix(lngRow, 0)
    vsExt.TextMatrix(lngRow, 1) = Nvl(rsInput!��������)
    
    '��һ������
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 0
End Sub

Private Sub vsExt_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsExt.EditSelStart = 0
    vsExt.EditSelLength = zlCommFun.ActualLen(vsExt.EditText)
End Sub

Private Sub vsExt_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'���ܣ�����ĳЩ�в�����༭(���¼�����BeforeEdit,��EditText��ֵ֮ǰ)
    mblnReturn = False
        
    If mintType = 0 Then
        'ֻ����ѡ��λ
        If optPosition(0).Value Then
            If Col <> 2 Or vsExt.RowData(Row) = 0 Then Cancel = True
        Else
            If cmd.Visible Then cmd.Visible = False '��ʼ�༭�������ذ�ť
            If Col <> 1 Then Cancel = True
        End If
    ElseIf mintType = 1 Or mintType = 4 Then
        'ֻ����༭��������
        If cmd.Visible Then cmd.Visible = False '��ʼ�༭�������ذ�ť
        If Col <> 0 Then Cancel = True
    ElseIf mintType = 2 Then
        '������������
        If Not CellCanEdit(Row, Col) Then Cancel = True
        
        If Col Mod 4 = 1 Then
            vsExt.EditMaxLength = 8
        Else
            vsExt.EditMaxLength = 0
        End If
    ElseIf mintType = 3 Then
        'ֻ����ѡ��걾
        If Col <> 1 Then
            Cancel = True
        ElseIf Val(vsExt.TextMatrix(Row, Col)) <> 0 Then
            Cancel = True '������ȡ��ѡ��(��ѡ)
        End If
    End If
End Sub

Private Function CellCanEdit(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'���ܣ�������ҩ�䷽ʱ,�ж�ָ���ĵ�Ԫ��ǰ�Ƿ���������
'˵�������䷽��������,���ǰһ��δ����,��ǰ����������
    '��λ����һ����ҩ���뵥Ԫ
    On Error Resume Next
    lngCol = (lngCol \ 4) * 4
    If lngCol - 4 >= vsExt.FixedCols Then
        lngCol = lngCol - 4
    Else
        If lngRow - 1 >= vsExt.FixedRows Then
            lngRow = lngRow - 1
            lngCol = vsExt.Cols - 4
        Else
            CellCanEdit = True
            Exit Function
        End If
    End If
    CellCanEdit = CLng(vsExt.Cell(flexcpData, lngRow, lngCol + 2)) <> 0
End Function

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ�������һ����ҩ�䷽�����뵥Ԫ��

    '��ǰλ��δ������ҩ
    If CLng(vsExt.Cell(flexcpData, lngRow, (lngCol \ 4) * 4 + 2)) = 0 Then Exit Sub
    
    '����δ����
    If lngCol Mod 4 = 1 And vsExt.TextMatrix(lngRow, lngCol) = "" Then Exit Sub
    
    If lngCol + 1 <= vsExt.Cols - 1 Then
        lngCol = lngCol + 1
    Else
        If lngRow + 1 > vsExt.Rows - 1 Then
            vsExt.AddItem "", vsExt.Rows
            Call SetSplitLine
        End If
        lngRow = lngRow + 1
        lngCol = vsExt.FixedCols
    End If
    
    vsExt.Row = lngRow: vsExt.Col = lngCol
End Sub

Private Function ItemExist(ByVal lng��ҩID As Long, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'���ܣ��ж���ҩ�䷽��������,ָ������ҩ�Ƿ��Ѿ�����
    Dim i As Long, j As Long
    
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            If Not (lngRow = i And (lngCol \ 4) * 4 = j) Then
                If CLng(vsExt.Cell(flexcpData, i, j + 2)) = lng��ҩID Then
                    ItemExist = True
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Private Function Check��ҩ����(ByVal str��ҩIDs As String) As Boolean
'���ܣ����һ���䷽�е���ҩ�������
'������str��ҩIDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str���� As String, str���� As String, lng���� As Long
    
    On Error GoTo errH
    
    strSQL = "Select ���� From ���ƻ�����Ŀ" & _
        " Where ��ĿID+0 IN(" & str��ҩIDs & ") Group by ���� Having Count(*)>1"
    strSQL = "Select A.����,A.����,B.����" & _
        " From ���ƻ�����Ŀ A,������ĿĿ¼ B" & _
        " Where A.��ĿID=B.ID And A.���� IN(" & strSQL & ")" & _
        " And A.��ĿID+0 IN(" & str��ҩIDs & ")" & _
        " Order by A.����,B.����"
    zlDatabase.OpenRecordset rsTmp, strSQL, Me.Caption 'In
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!���� <> lng���� Then
                If rsTmp!���� = 1 Then
                    str���� = str���� & vbCrLf & "��"
                Else
                    str���� = str���� & vbCrLf & "��"
                End If
                lng���� = rsTmp!����
            End If
            If rsTmp!���� = 1 Then
                str���� = str���� & "��" & rsTmp!����
            Else
                str���� = str���� & "��" & rsTmp!����
            End If
            rsTmp.MoveNext
        Next
        If str���� <> "" Then
            MsgBox "��ǰ�䷽�з�������ҩƷ������ã�" & Replace(str����, "��", "�� "), vbInformation, gstrSysName
            Exit Function
        ElseIf str���� <> "" Then
            If MsgBox("��ǰ�䷽�з�������ҩƷ�������ã�" & Replace(str����, "��", "�� ") & vbCrLf & vbCrLf & "Ҫ������", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    Check��ҩ���� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
