VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmChargeRequest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsMoney 
      Height          =   1290
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   6000
      _cx             =   10583
      _cy             =   2275
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      Begin VB.Line lnX0 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY0 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDetail 
      Height          =   1485
      Left            =   240
      TabIndex        =   1
      Top             =   2985
      Width           =   5625
      _cx             =   9922
      _cy             =   2619
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483644
      GridColorFixed  =   -2147483644
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   2
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
      Begin VB.Line lnX1 
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   1785
         Y1              =   135
         Y2              =   135
      End
      Begin VB.Line lnY1 
         Index           =   0
         Visible         =   0   'False
         X1              =   825
         X2              =   825
         Y1              =   0
         Y2              =   1215
      End
   End
   Begin VB.Image imgX 
      Height          =   45
      Left            =   2685
      MousePointer    =   7  'Size N S
      Top             =   1470
      Width           =   5085
   End
End
Attribute VB_Name = "frmChargeRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private mrsPrice As ADODB.Recordset 'δ�Ʒ�ҽ����������

Private mfrmParent As Form
Private pgbLoad As Object
Private AdviceID As Long
Private lngSendNO As Long
Private iPatientType As Integer
Private lngPatientID As Long
Private lngPatientDept As Long
Private lngPageId As Long
Private strCheckNo As String
Private str�ѱ� As String
Private int��¼���� As Integer
Private intִ��״̬ As Integer

Private lng��������ID As Long
Private mstrPrivs As String

Private mSysName As String
Private mstrSys As String

Private msgl����ۿ� As Single
Private mblnDataMoved As Boolean
Private mblnChargeDataMoved As Boolean
Public mblnCash As Boolean  '���Ѿ��շѣ�ֻҪ���ڼ��ʼ�¼��һ��Ϊ�շѵļ�¼������Ϊδ�շѣ�


Public Property Get CashState() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    CashState = mblnCash
End Property

Public Sub zlRefresh(ByVal objParent As Object, ByVal lngAdviceID As Long, ByVal SendNO As Long, Optional ByVal strPrivs As String = "", Optional ByVal strClass As String = "����", Optional ByVal strSys As String)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:  lngAdviceID         ��ҽ��id
    '------------------------------------------------------------------------------------------------------------------
    
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
        
    '��ʼ������
    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "ҽ��ID", adBigInt
    mrsPrice.Fields.Append "��������ID", adBigInt
    mrsPrice.Fields.Append "���", adVarChar, 1
    mrsPrice.Fields.Append "�շ�ϸĿID", adBigInt
    mrsPrice.Fields.Append "���㵥λ", adVarChar, 100, adFldIsNullable
    mrsPrice.Fields.Append "��������", adInteger
    mrsPrice.Fields.Append "ִ�п���", adInteger
    mrsPrice.Fields.Append "������ĿID", adBigInt
    mrsPrice.Fields.Append "�վݷ�Ŀ", adVarChar, 20, adFldIsNullable
    mrsPrice.Fields.Append "����", adDouble
    mrsPrice.Fields.Append "����", adDouble
    mrsPrice.Fields.Append "Ӧ��", adCurrency
    mrsPrice.Fields.Append "ʵ��", adCurrency
    mrsPrice.Fields.Append "���͵���", adVarChar, 30
    mrsPrice.Fields.Append "���ͺ�", adVarChar, 30
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
    
       
    iPatientType = 1
    lngPatientID = 0
    lngPageId = 0
    strCheckNo = ""
    lngPatientDept = 0
'    int�Ʒ�״̬ = 0
    str�ѱ� = ""
    int��¼���� = 1
    mstrPrivs = strPrivs
    intִ��״̬ = 0
'    strNo = ""
    lng��������ID = 0
            
    '�ӿڲ�������
    mSysName = strClass
    mstrSys = strSys
    mstrPrivs = strPrivs
    AdviceID = lngAdviceID
    lngSendNO = SendNO
    Set mfrmParent = objParent
        
    On Error GoTo DBError
    
    '����ת������
    mblnDataMoved = False
    strSQL = "Select b.����ʱ�� From ����ҽ����¼ b Where b.ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID)
    If rsTmp.BOF = False Then
        mblnDataMoved = False
        mblnChargeDataMoved = zlDatabase.DateMoved(Format(rsTmp("����ʱ��").Value, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
    Else
        mblnDataMoved = True
        mblnChargeDataMoved = True
    End If
            
    strSQL = _
            "Select A.��¼����," & _
                   "A.ִ��״̬," & _
                   "B.����ID," & _
                   "B.��ҳID," & _
                   "B.�Һŵ�," & _
                   "B.���˿���ID," & _
                   "Nvl(F.�ѱ�, D.�ѱ�) as �ѱ�," & _
                   "Decode(B.������Դ, 1, '����', 2, 'סԺ', 3, '����', 4, '���') as ��Դ, "
    strSQL = strSQL & _
                   "A.ִ�в���ID " & _
              "From ����ҽ������ A," & _
                   "����ҽ����¼ B," & _
                   "������Ϣ     D," & _
                   "������ҳ     F " & _
             "Where A.ҽ��ID = B.ID " & _
                   "And B.����ID = D.����ID " & _
                   "And B.����ID = F.����ID(+) " & _
                   "And B.��ҳID = F.��ҳID(+) " & _
                   "And A.ҽ��ID IN (SELECT ID FROM ����ҽ����¼ WHERE ID = [1] OR ���ID = [1]) " & _
                   "And A.���ͺ� = [2] " & _
             "Order by A.����ʱ�� Desc,B.���"
             
    '����ת������
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
    End If
                         
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, SendNO)
        
    If rsTmp.BOF = False Then
        
        iPatientType = Decode(rsTmp("��Դ"), "����", 1, "���", 1, 2)
        
        lngPatientID = rsTmp("����ID")
        lngPageId = Nvl(rsTmp("��ҳID"), 0)
        strCheckNo = Nvl(rsTmp("�Һŵ�"), "")
        lngPatientDept = Nvl(rsTmp("���˿���ID"), 0)
        str�ѱ� = Nvl(rsTmp!�ѱ�)
        int��¼���� = Nvl(rsTmp!��¼����, 1)
        intִ��״̬ = Nvl(rsTmp!ִ��״̬, 0)
        lng��������ID = Nvl(rsTmp!ִ�в���ID, 0)
    End If
    
    
    '��������ۿ�
    If mstrSys = "���" Then
        
        str�ѱ� = ""
        
        strSQL = "SELECT NVL(B.���۸�,1) AS �����ۿ� FROM �����Ŀҽ�� A,�����Ŀ�嵥 B WHERE A.�嵥id=B.ID AND A.ҽ��ID=[1] "
        
        '����ת������
        If mblnDataMoved Then
            strSQL = Replace(strSQL, "�����Ŀҽ��", "H�����Ŀҽ��")
            strSQL = Replace(strSQL, "�����Ŀ�嵥", "H�����Ŀ�嵥")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, AdviceID)
        If rsTmp.BOF = False Then
            msgl����ۿ� = rsTmp("�����ۿ�").Value
        End If
        
    End If
    
    Call LoadMoneyList(AdviceID, lngSendNO, 0, str�ѱ�, int��¼����)
    
    Exit Sub
    
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlClearData()
    
    '----------------------------------------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------------------------------------
    
    vsMoney.Rows = 2
    vsMoney.Cell(flexcpText, 1, 0, 1, vsMoney.Cols - 1) = ""
    
    vsDetail.Rows = 2
    vsDetail.Cell(flexcpText, 1, 0, 1, vsDetail.Cols - 1) = ""
    
End Sub

Public Function zlMenuClick(ByVal strFunc As String) As Boolean

    '----------------------------------------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------------------------------------
    Select Case strFunc
    Case "����������"
        zlMenuClick = MoneyMain
    Case "�޸ĸ��ӷ���"
        zlMenuClick = MoneyModi
    Case "ɾ�����ӷ���"
        zlMenuClick = MoneyDel
    Case "�շѵ���"
        zlMenuClick = MoneyNewBilling(1)
    Case "���ʵ���"
        zlMenuClick = MoneyNewBilling(2)
    Case "��Ѻ��õǼ�"
        zlMenuClick = MoneyNewBilling(2, True)
    End Select
    
End Function


Public Property Get Body(Optional ByVal lngIndex As Long) As Object
    Set Body = vsMoney
End Property

Private Function GetMaxNo(ByVal strNO As String, ByRef lngNO As Long, ByRef strDate As String) As Boolean
    
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "SELECT NVL(MAX(���),0) AS ���,NVL(MAX(�Ǽ�ʱ��),SYSDATE) AS �Ǽ�ʱ�� FROM ���˷��ü�¼ WHERE NO='" & strNO & "'"
    Call zlDatabase.OpenRecordset(rs, strSQL, Me.Caption)
    If rs.BOF = False Then
        lngNO = rs("���").Value
        strDate = Format(rs("�Ǽ�ʱ��").Value, "yyyy-MM-dd HH:mm:ss")
    End If
            
    GetMaxNo = True
    
End Function

Private Function MoneyMain() As Boolean
    '----------------------------------------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------------------------------------
    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim lng����ID As Long
    Dim lng��ҳID As Long
    Dim lng���ͺ�
    Dim lngҽ��ID As Long
    Dim int��Դ As Integer
    Dim int���� As Integer
    Dim lng��ĿID As Long
    Dim lngִ�в���ID As Long
    Dim lng���˲���ID As Long
    Dim lng���˿���ID As Long
    Dim lng���ID As Long
    Dim arrSQL As Variant
    Dim strSQL As String
    Dim strDate As String
    Dim i As Long
    Dim int������Ŀ�� As Integer
    Dim lng���մ���ID As Long
    Dim str���ձ��� As String
    Dim curͳ���� As Currency
    Dim strMsg As String
    Dim strNO As String
    
    If mrsPrice Is Nothing Then Exit Function
    If AdviceID = 0 Then Exit Function

    
    If vsMoney.TextMatrix(vsMoney.Row, 2) <> "[δ�Ʒ�]" Then
        MsgBox "��ִ����Ŀ����Ʒѻ��Ѿ��Ʒѡ�" & vbCrLf & "�����Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsPrice.Filter = "���͵���='" & vsMoney.TextMatrix(vsMoney.Row, 9) & "'"
    If mrsPrice.RecordCount = 0 Then
        MsgBox "��ִ����Ŀû�п��ԼƷѵ������á�" & vbCrLf & "�����Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
        Exit Function
    End If
    
    mrsPrice.MoveFirst
    
    If intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    
    If MsgBox("ȷʵҪ���ɸ���Ŀ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Function
    End If
        
    Screen.MousePointer = 11
    
    
    'lng���ͺ� = lngSendNO
    lng���ͺ� = mrsPrice("���ͺ�").Value
    lng����ID = lngPatientID
    lng��ҳID = lngPageId
    int��Դ = iPatientType
    
    '��ȡ���˵���Ϣ
    strSQL = "Select A.����,A.�Ա�,A.����,Nvl(B.�ѱ�,A.�ѱ�) as �ѱ�," & _
        " A.�����,A.סԺ��,Nvl(A.��ǰ����,B.��Ժ����) as ����," & _
        " Nvl(A.��ǰ����ID,B.��ǰ����ID) as ���˲���ID," & _
        " Nvl(A.��ǰ����ID,B.��Ժ����ID) as ���˿���ID," & _
        " Nvl(B.����,A.����) as ����,C.���� as ������" & _
        " From ������Ϣ A,������ҳ B,ҽ�Ƹ��ʽ C" & _
        " Where A.����ID=" & lng����ID & " And A.����ID=B.����ID(+)" & _
        " And B.��ҳID(+)=" & lng��ҳID & " And A.ҽ�Ƹ��ʽ=C.����(+)"
    Call zlDatabase.OpenRecordset(rsPati, strSQL, Me.Caption)
    
    '���ܶ��շ���ΪҩƷ����
    If int��¼���� = 1 Then
        lng���ID = ExistIOClass(8) '���ﻮ�۵�
    Else
        lng���ID = ExistIOClass(9) '����/סԺ���ʵ�
    End If
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    
    arrSQL = Array()
    With mrsPrice
        .MoveFirst
        
        Dim lngMaxNo As Long
        
        strNO = mrsPrice("���͵���").Value
        
        Call GetMaxNo(strNO, lngMaxNo, strDate)
        
        If int��¼���� = 1 Then
            If BillExistBalance(strNO) Then
                MsgBox "Ҫ���ɵ��շѵ���" & strNO & "��ͬһ�ŵ��ݣ�" & strNO & "�Ѿ��շѣ����������ɣ�", vbInformation, gstrSysName
                Exit Function
            End If
            strDate = "TO_DATE('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
        Else
            strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        End If
        
                       
        For i = lngMaxNo + 1 To lngMaxNo + .RecordCount
            '��ȡ��Ӧ��ҽ����Ϣ
            If lngҽ��ID <> !ҽ��ID Then
                strSQL = "Select ҽ����Ч,���˿���ID,��������ID,Ӥ��,ִ��Ƶ��,�Ƽ�����" & _
                    " From ����ҽ����¼ Where ID=" & !ҽ��ID
                Call zlDatabase.OpenRecordset(rsAdvice, strSQL, Me.Caption)
                
                '����ǰ�����Ʒ�ҽ�����Ϊ�ѼƷ�
'                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'                arrSQL(UBound(arrSQL)) = "ZL_����ҽ������_�Ʒ�(" & !ҽ��ID & "," & lng���ͺ� & ")"
            End If
            lngҽ��ID = !ҽ��ID
            
            '���˲�������
            lng���˲���ID = Nvl(rsPati!���˲���ID, 0)
            lng���˿���ID = Nvl(rsPati!���˿���ID, 0)
            If lng���˿���ID = 0 Then
                lng���˲���ID = Nvl(rsAdvice!���˿���ID, 0)
                lng���˿���ID = Nvl(rsAdvice!���˿���ID, 0)
            End If
            If lng���˿���ID = 0 Then
                lng���˲���ID = UserInfo.����ID
                lng���˿���ID = UserInfo.����ID
            End If
            
            'ÿ���շ���Ŀ�Ĵ���
            If lng��ĿID <> !�շ�ϸĿID Then
                int���� = i '��ȡ�۸񸸺�
                lngִ�в���ID = Get�շ�ִ�п���ID(lng����ID, lng��ҳID, !���, !�շ�ϸĿID, !ִ�п���, Nvl(rsAdvice!���˿���ID, 0), Nvl(rsAdvice!��������ID, 0), int��Դ)
                            
                '��ȡ������Ŀ��Ϣ
                If int��Դ = 2 And Not IsNull(rsPati!����) Then
                    strMsg = gclsInsure.GetItemInsure(lng����ID, !�շ�ϸĿID, !ʵ��, False, rsPati!����)
                    If strMsg <> "" Then
                        int������Ŀ�� = Val(Split(strMsg, ";")(0))
                        lng���մ���ID = Val(Split(strMsg, ";")(1))
                        curͳ���� = Format(Val(Split(strMsg, ";")(2)), gstrDec)
                        str���ձ��� = CStr(Split(strMsg, ";")(3))
                    End If
                End If
            End If
            lng��ĿID = !�շ�ϸĿID
            
            
            Select Case mstrSys
            Case "���"
                str�ѱ� = ""
            End Select
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If int��Դ = 1 Then
                If int��¼���� = 1 Then
                    '�������ﻮ�۵���
                    arrSQL(UBound(arrSQL)) = _
                        "zl_���ﻮ�ۼ�¼_Insert('" & strNO & "'," & i & "," & lng����ID & ",NULL," & _
                        ZVal(Nvl(rsPati!�����, 0)) & ",'" & Nvl(rsPati!������) & "','" & Nvl(rsPati!����) & "'," & _
                        "'" & Nvl(rsPati!�Ա�) & "','" & Nvl(rsPati!����) & "','" & str�ѱ� & "',NULL," & _
                        lng���˲���ID & "," & lng���˿���ID & "," & UserInfo.����ID & ",'" & UserInfo.���� & "'," & _
                        "NULL," & lng��ĿID & ",'" & !��� & "','" & !���㵥λ & "',NULL,1," & !���� & "," & _
                        !�������� & "," & ZVal(lngִ�в���ID) & "," & IIF(int���� = i, "NULL", int����) & "," & _
                        !������ĿID & ",'" & Nvl(!�վݷ�Ŀ) & "'," & !���� & "," & !Ӧ�� & "," & !ʵ�� & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.���� & "'," & ZVal(lng���ID) & ",NULL," & _
                        !ҽ��ID & ",'" & Nvl(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & Nvl(rsAdvice!ҽ����Ч, 0) & "," & _
                        Nvl(rsAdvice!�Ƽ�����, 0) & ",1)"
                Else
                    '����������ʵ���
                    arrSQL(UBound(arrSQL)) = _
                        "zl_������ʼ�¼_Insert('" & strNO & "'," & i & "," & lng����ID & "," & _
                        ZVal(Nvl(rsPati!�����, 0)) & ",'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�Ա�) & "'," & _
                        "'" & Nvl(rsPati!����) & "','" & str�ѱ� & "',NULL," & ZVal(Nvl(rsAdvice!Ӥ��, 0)) & "," & _
                        lng���˲���ID & "," & lng���˿���ID & "," & UserInfo.����ID & "," & _
                        "'" & UserInfo.���� & "',NULL," & lng��ĿID & ",'" & !��� & "'," & _
                        "'" & !���㵥λ & "',1," & !���� & "," & !�������� & "," & ZVal(lngִ�в���ID) & "," & _
                        IIF(int���� = i, "NULL", int����) & "," & !������ĿID & ",'" & Nvl(!�վݷ�Ŀ) & "'," & !���� & "," & _
                        !Ӧ�� & "," & !ʵ�� & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.��� & "'," & _
                        "'" & UserInfo.���� & "'," & ZVal(lng���ID) & ",NULL,NULL," & !ҽ��ID & "," & _
                        "'" & Nvl(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & Nvl(rsAdvice!ҽ����Ч, 0) & "," & _
                        Nvl(rsAdvice!�Ƽ�����, 0) & ")"
                End If
            Else
                '����סԺ���ʵ���
                arrSQL(UBound(arrSQL)) = _
                    "zl_סԺ���ʼ�¼_Insert('" & strNO & "'," & i & "," & lng����ID & "," & ZVal(lng��ҳID) & "," & _
                    ZVal(Nvl(rsPati!סԺ��, 0)) & ",'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�Ա�) & "'," & _
                    "'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!����) & "','" & str�ѱ� & "'," & _
                    lng���˲���ID & "," & lng���˿���ID & ",NULL," & ZVal(Nvl(rsAdvice!Ӥ��, 0)) & "," & _
                    UserInfo.����ID & ",'" & UserInfo.���� & "',NULL," & lng��ĿID & ",'" & !��� & "'," & _
                    "'" & !���㵥λ & "'," & int������Ŀ�� & "," & ZVal(lng���մ���ID) & ",'" & str���ձ��� & "'," & _
                    "1," & !���� & "," & !�������� & "," & ZVal(lngִ�в���ID) & "," & _
                    IIF(int���� = i, "NULL", int����) & "," & !������ĿID & ",'" & Nvl(!�վݷ�Ŀ) & "'," & !���� & "," & _
                    !Ӧ�� & "," & !ʵ�� & "," & curͳ���� & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.��� & "','" & UserInfo.���� & "',NULL," & ZVal(lng���ID) & ",NULL,NULL,NULL," & _
                    !ҽ��ID & ",'" & Nvl(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & Nvl(rsAdvice!ҽ����Ч, 0) & "," & _
                    Nvl(rsAdvice!�Ƽ�����, 0) & ",NULL)"
            End If
            
            .MoveNext
        Next
    End With
    
    '����ҽ���Ʒѱ�־
'    strSQL = _
'            "SELECT ID FROM ����ҽ����¼ WHERE ID=" & AdviceID
'    strSQL = strSQL & " UNION ALL " & _
'            "SELECT ID FROM ����ҽ����¼ WHERE ���ID=" & AdviceID
    
    strSQL = "SELECT ҽ��ID FROM ����ҽ������ WHERE NVL(�Ʒ�״̬,0)=0 AND NO='" & strNO & "'"
    
    Call zlDatabase.OpenRecordset(rsAdvice, strSQL, Me.Caption)
    If rsAdvice.BOF = False Then
        Do While Not rsAdvice.EOF
            
            '����ǰ�����Ʒ�ҽ�����Ϊ�ѼƷ�
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_����ҽ������_�Ʒ�(" & rsAdvice("ҽ��ID").Value & "," & lng���ͺ� & ")"
            
            rsAdvice.MoveNext
        Loop
    End If
    
    On Error GoTo errH
    
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    
    Dim strSQL1() As String
    ReDim strSQL1(0 To 1)
    strSQL1(0) = ""
    strSQL1(1) = ""
    
    If int��¼���� <> 1 Then
        If �����Զ�����(strSQL1, strNO) = False Then GoTo errH
    End If
    
    For i = 1 To UBound(strSQL1)
        If Trim(strSQL1(i)) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL1(i), Me.Caption)
    Next
    
    '���ύǰ����ҽ������
    If int��Դ = 2 And Not IsNull(rsPati!����) Then
        If gclsInsure.GetCapability(support�����ϴ�, , rsPati!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , rsPati!����) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, , rsPati!����) Then
                gcnOracle.RollbackTrans
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    gcnOracle.CommitTrans
    
    '���ύ�����ҽ������
    If int��Դ = 2 And Not IsNull(rsPati!����) Then
        If gclsInsure.GetCapability(support�����ϴ�, , rsPati!����) And gclsInsure.GetCapability(support������ɺ��ϴ�, , rsPati!����) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, , rsPati!����) Then
                If strMsg <> "" Then
                    MsgBox strMsg, vbInformation, gstrSysName
                Else
                    MsgBox "����""" & strNO & """��������ҽ������ʧ��,�õ����ѱ��棡", vbInformation, gstrSysName
                End If
            End If
        End If
    End If
    On Error GoTo 0
    Screen.MousePointer = 0
    
    MsgBox "ִ����Ŀ�����������ɳɹ���", vbInformation, gstrSysName
    'ˢ��
    
    MoneyMain = True
    
    Me.Tag = "Loading": Call Form_Activate
    
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function �����Զ�����(ByRef strSQL() As String, ByVal strNO As String) As Boolean
    
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim bln�Զ����� As Boolean
    
    On Error GoTo ErrHand
    
    bln�Զ����� = GetSysParVal(63) = "1"
    If bln�Զ����� = False Then
        �����Զ����� = True
        Exit Function
    End If
    
    strTmp = "SELECT DISTINCT A.ִ�в���id FROM ���˷��ü�¼ A,�������� B WHERE A.�շ�ϸĿid=B.����id AND NVL(B.��������,0)=1 AND A.�շ����='4' and A.NO='" & strNO & "'"
    Call zlDatabase.OpenRecordset(rs, strTmp, Me.Caption)
    If rs.BOF = False Then
        Do While Not rs.EOF
            If zlCommFun.Nvl(rs("ִ�в���id").Value, 0) > 0 Then
                strSQL(1) = "zl_�����շ���¼_��������(" & rs("ִ�в���id").Value & ",25,'" & strNO & "','" & UserInfo.���� & "','" & UserInfo.���� & "','" & UserInfo.���� & "',1,Sysdate)"
            End If
            rs.MoveNext
        Loop
    End If
    
    �����Զ����� = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function MoneyModi() As Boolean
    '----------------------------------------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------------------------------------
    
    Dim lng����ID As Long, lng��ҳID As Long
    Dim lng���˿���ID As Long
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim int������Դ As Integer, int��¼���� As Integer
    Dim strNO As String, bln��� As Boolean
    
    
    If AdviceID = 0 Then Exit Function
    If vsMoney.TextMatrix(vsMoney.Row, 0) = "������" Then
        MsgBox "ִ����Ŀ�������ò����޸ġ������Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
        Exit Function
    End If
    If intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsMoney
        strNO = .TextMatrix(.Row, 2)
        If strNO = "" Or strNO = "[δ�Ʒ�]" Then Exit Function
        int��¼���� = .Cell(flexcpData, .Row, 1)
        
        If InStr(.TextMatrix(.Row, 1), "��") > 0 Then
            MsgBox "�õ����Ѿ��շѣ��������޸ġ�", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If zlDatabase.NOMoved("���˷��ü�¼", strNO) Then
        MsgBox "�õ����Ѿ�ת�����������޸ġ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    If mSysName = "����" Then
        strSQL = "SELECT MIN(ID) AS ID FROM ����ҽ����¼ WHERE ���id=" & AdviceID
        zlDatabase.OpenRecordset rs, strSQL, Me.Caption
        If rs.BOF Then Exit Function
        lngҽ��ID = zlCommFun.Nvl(rs("ID").Value)
    Else
        lngҽ��ID = AdviceID
    End If
    
    
    lng���ͺ� = lngSendNO
    lng����ID = lngPatientID
    lng��ҳID = lngPageId
    int������Դ = iPatientType
    lng���˿���ID = lngPatientDept
    
    If int��¼���� = 2 Then
       bln��� = BillisZeroLog(strNO)
    End If
    
    
    frmTechnicExpense.mstrPrivs = mstrPrivs
    frmTechnicExpense.mbytInState = 0
    frmTechnicExpense.mbln���õǼ� = bln���
    frmTechnicExpense.mstrInNO = strNO
    frmTechnicExpense.mlngҽ��ID = lngҽ��ID
    frmTechnicExpense.mlng���ͺ� = lng���ͺ�
    frmTechnicExpense.mlng����ID = lng����ID
    frmTechnicExpense.mlng��ҳID = lng��ҳID
    frmTechnicExpense.mint������Դ = int������Դ
    frmTechnicExpense.mint��¼���� = int��¼����
    frmTechnicExpense.mlng��������ID = lng��������ID
    frmTechnicExpense.mlng���˿���id = lng���˿���ID
    On Error Resume Next
    frmTechnicExpense.Show 1, Me
    On Error GoTo 0
    If gblnOK Then
        'ˢ��
        MoneyModi = True
        Me.Tag = "Loading": Call Form_Activate
        
    End If
End Function

Private Function MoneyDel() As Boolean
    '----------------------------------------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------------------------------------
    
    Dim int������Դ As Integer, int��¼���� As Integer
    Dim strNO As String
    
    If AdviceID = 0 Then Exit Function
    If vsMoney.TextMatrix(vsMoney.Row, 0) = "������" Then
        MsgBox "ִ����Ŀ�������ò���ɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    If intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    
    With vsMoney
        strNO = .TextMatrix(.Row, 2)
        If strNO = "" Or strNO = "[δ�Ʒ�]" Then Exit Function
        int��¼���� = .Cell(flexcpData, .Row, 1)
    
        If InStr(.TextMatrix(.Row, 1), "��") > 0 Then
            MsgBox "�õ����Ѿ��շѣ�������ɾ����", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    If zlDatabase.NOMoved("���˷��ü�¼", strNO) Then
        MsgBox "�õ����Ѿ�ת����������ɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    int������Դ = iPatientType
    
    frmTechnicExpense.mstrPrivs = mstrPrivs
    frmTechnicExpense.mbytInState = 3
    frmTechnicExpense.mstrInNO = strNO
    frmTechnicExpense.mint������Դ = int������Դ
    frmTechnicExpense.mint��¼���� = int��¼����
    On Error Resume Next
    frmTechnicExpense.Show 1, Me
    On Error GoTo 0
    If gblnOK Then
        'ˢ��
        MoneyDel = True
        Me.Tag = "Loading": Call Form_Activate
    End If
End Function

Private Function MoneyNewBilling(ByVal iRecordType As Integer, Optional OnlyRecord As Boolean = False) As Boolean
    '----------------------------------------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------------------------------------
    
    Dim lng����ID As Long, lng��ҳID As Long
    Dim lng���˿���ID As Long
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim int������Դ As Integer
    
    If AdviceID = 0 Then Exit Function
    
    If intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Function
    End If
    
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    If mSysName = "����" Then
        strSQL = "SELECT MIN(ID) AS ID FROM ����ҽ����¼ WHERE ���id=" & AdviceID
        zlDatabase.OpenRecordset rs, strSQL, Me.Caption
        If rs.BOF Then Exit Function
        lngҽ��ID = zlCommFun.Nvl(rs("ID").Value)
    Else
        lngҽ��ID = AdviceID
    End If
        
    lng���ͺ� = lngSendNO
    lng����ID = lngPatientID
    lng��ҳID = lngPageId
    int������Դ = iPatientType
    lng���˿���ID = lngPatientDept
    
    frmTechnicExpense.mstrPrivs = mstrPrivs
    frmTechnicExpense.mbytInState = 0
    frmTechnicExpense.mlngҽ��ID = lngҽ��ID
    frmTechnicExpense.mlng���ͺ� = lng���ͺ�
    frmTechnicExpense.mlng����ID = lng����ID
    frmTechnicExpense.mlng��ҳID = lng��ҳID
    frmTechnicExpense.mint������Դ = int������Դ
    frmTechnicExpense.mint��¼���� = iRecordType
    frmTechnicExpense.mbln���õǼ� = OnlyRecord
    frmTechnicExpense.mlng��������ID = lng��������ID
    frmTechnicExpense.mlng���˿���id = lng���˿���ID
    On Error Resume Next
    frmTechnicExpense.Show 1, Me
    On Error GoTo 0
    If gblnOK Then
        'ˢ��
        MoneyNewBilling = True
        Me.Tag = "Loading": Call Form_Activate
    End If
End Function

Private Sub Form_Activate()
    On Error Resume Next

    If Me.Tag = "Loading" Then
        mfrmParent.Refresh
        Me.Tag = ""
        Call LoadMoneyList(AdviceID, lngSendNO, 0, str�ѱ�, int��¼����)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHand
    
    Call mfrmParent.ActiveFormKeyDown(KeyCode, Shift)

ErrHand:

End Sub

Private Sub Form_Load()
    
    On Error GoTo ShowError
    
'    Set mrsPrice = Nothing
    
    Call InitMoneyTable
    Call InitDetailTable
    
    Exit Sub
    
ShowError:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
'    If imgX.Top > Me.ScaleHeight - 1000 Then imgX.Top = Me.ScaleHeight - 1000
        
    With vsMoney
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = imgX.Top - .Top
    End With
    
    With imgX
        .Left = 0
        .Top = vsMoney.Top + vsMoney.Height
        .Width = Me.ScaleWidth
    End With
    
    With vsDetail
        .Left = 0
        .Top = imgX.Top + imgX.Height
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
            
    Call AppendRows(vsMoney, lnX0, lnY0)
    Call AppendRows(vsDetail, lnX1, lnY1)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Set mrsPrice = Nothing
End Sub


Private Sub imgX_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 1 Then Exit Sub
    
    imgX.Top = imgX.Top + y
    
    If imgX.Top < 1500 Then imgX.Top = 1500
    If Me.Height - imgX.Top - imgX.Height < 1000 Then imgX.Top = Me.Height - imgX.Height - 1000

    Form_Resize
End Sub

Private Sub vsDetail_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsDetail, lnX1, lnY1)
End Sub

Private Sub vsDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsDetail, lnX1, lnY1)
End Sub

Private Sub vsDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub vsMoney_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    If NewCol >= vsMoney.FixedCols And NewRow >= vsMoney.FixedRows Then
        vsMoney.ForeColorSel = vsMoney.Cell(flexcpForeColor, NewRow, 0)
        Call LoadBillDetail(NewRow)
    End If
    
    On Error GoTo ErrHand
    
    Call mfrmParent.ActiveFormEnabled
    
ErrHand:
End Sub

Private Sub vsMoney_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsMoney, lnX0, lnY0)
End Sub

Private Sub vsMoney_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsMoney, lnX0, lnY0)
End Sub

Private Sub vsMoney_DblClick()
    If vsMoney.MouseRow >= vsMoney.FixedRows Then
        Call vsMoney_KeyPress(13)
    End If
End Sub

Private Sub vsMoney_GotFocus()
    vsMoney.BackColorSel = COLOR_FOCUS
End Sub

Private Sub vsMoney_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Form_KeyDown(KeyCode, Shift)
End Sub

Private Sub vsMoney_KeyPress(KeyAscii As Integer)
    
    Dim int������Դ As Integer
    Dim int��¼���� As Integer
    Dim strNO As String
    
    If KeyAscii = 13 Or KeyAscii = 32 Then
        KeyAscii = 0
        With vsMoney
            strNO = .TextMatrix(.Row, 2)
            If strNO = "" Or strNO = "[δ�Ʒ�]" Then Exit Sub
            int��¼���� = .Cell(flexcpData, .Row, 1)
        End With
        int������Դ = iPatientType
        
        frmTechnicExpense.mstrPrivs = mstrPrivs
        frmTechnicExpense.mbytInState = 1
        frmTechnicExpense.mstrInNO = strNO
        frmTechnicExpense.mint������Դ = int������Դ
        frmTechnicExpense.mint��¼���� = int��¼����
        On Error Resume Next
        frmTechnicExpense.Show 1, Me
    End If
End Sub

Private Sub vsMoney_LostFocus()
    vsMoney.BackColorSel = COLOR_LOST
End Sub

Private Sub vsDetail_GotFocus()
    vsDetail.BackColorSel = COLOR_FOCUS
End Sub

Private Sub vsDetail_LostFocus()
    vsDetail.BackColorSel = COLOR_LOST
End Sub

Private Sub InitMoneyTable()
    '----------------------------------------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------------------------------------
    
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "��������,900,1;��������,1000,1;���ݺ�,900,1;�ѱ�,900,1;Ӧ�ս��,1000,7;ʵ�ս��,1000,7;������,750,1;�Ǽ�ʱ��,1080,1;�Ǽ���,750,1;���͵���,0,1"
    arrHead = Split(strHead, ";")
    With vsMoney
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        
        .Cols = .Cols + 1
        .ExtendLastCol = True
        Call AppendRows(vsMoney, lnX0, lnY0)
        
    End With
End Sub

Private Sub InitDetailTable()
    '----------------------------------------------------------------------------------------------------
    '����:
    '----------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "���,650,1;��Ŀ,3000,1;����,1000,1;����,1000,7;Ӧ�ս��,1000,7;ʵ�ս��,1000,7;ִ�п���,1000,1"
    arrHead = Split(strHead, ";")
    With vsDetail
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        
        .Cols = .Cols + 1
        .ExtendLastCol = True
        Call AppendRows(vsDetail, lnX1, lnY1)
    End With
End Sub


Private Function LoadMoneyList(ByVal lngҽ��ID As Long, _
                                ByVal lng���ͺ� As Long, _
                                ByVal int�Ʒ�״̬ As Integer, _
                                ByVal str�ѱ� As String, _
                                ByVal int��¼���� As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ��ҽ������Ҫ���ü����ӷ���
    '˵����1.����ҽ������������ü����ӷ���,�����ÿ�����δ����
    '      2.Ŀǰ�����ݲ�֧�ֲ����˷�,�����嵥��ֻ�����ʾ
    '------------------------------------------------------------------------------------------------------------------
    
    Dim rsList As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    Dim blnMain As Boolean
    Dim blnSub As Boolean
    Dim strPre As String
    Dim lngRow As Long
    Dim curӦ�� As Currency
    Dim curʵ�� As Currency
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errH
    
    'δ�Ʒѵ�
    Dim strTmp As String
    
    If mstrSys = "LIS" Then
            
        strTmp = _
            "(Select ID From ����ҽ����¼ X,(Select " & lngҽ��ID & " as ҽ��id From dual union Select ҽ��id From ������Ŀ�ֲ� A " & _
            "Where �걾ID In (Select ID From ����걾��¼ Where ҽ��id=" & lngҽ��ID & ") " & _
                  ") Y Where Y.ҽ��id In (X.ID,X.���ID)) "
    Else
        strTmp = _
            "(Select ID from ����ҽ����¼ WHERE " & lngҽ��ID & " IN (ID,���id)) "
    End If
    
    If intִ��״̬ <> 1 Then
        strSQL = _
            "SELECT DISTINCT A.��¼����,A.NO,A.���ͺ� " & _
            "FROM ����ҽ������ A " & _
            "WHERE   A.ҽ��ID IN " & strTmp & _
                    "AND NVL(A.�Ʒ�״̬,0)=0"
        
        '����ת������
        If mblnDataMoved Then
            strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            strSQL = Replace(strSQL, "����걾��¼", "H����걾��¼")
            strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        End If
        
        Call zlDatabase.OpenRecordset(rs, strSQL, Me.Caption)
        strSQL = ""
        If rs.BOF = False Then
            Do While Not rs.EOF
                
                'δ�Ʒ�״̬,ֱ�Ӷ�ȡ�շѹ�ϵ��ʾ
                curӦ�� = 0
                curʵ�� = 0
                Call NewAdvicePrice(rs, curӦ��, curʵ��)
                
                If curʵ�� > 0 Then
                    
                    '�жϴ��շѵ��ݺ��Ƿ��շ�
                    If int��¼���� = 1 Then
                        If Not BillExistBalance(rs("NO").Value) Then
                            strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
                                    "Select 1 as ��������," & _
                                            int��¼���� & " as ��¼����," & _
                                            "0 as ���շ�," & _
                                            "'[δ�Ʒ�]' as NO," & _
                                            "'" & str�ѱ� & "' as �ѱ�," & _
                                            curӦ�� & " as Ӧ�ս��," & _
                                            curʵ�� & " as ʵ�ս��,'" & _
                                            UserInfo.���� & "' as ������," & _
                                            "Sysdate as �Ǽ�ʱ��,'" & _
                                            UserInfo.���� & "' as ����Ա,'" & rs("NO").Value & "' AS ���͵��� " & _
                                    " From Dual"
                        End If
                    Else
                    
                        strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
                            "Select 1 as ��������," & _
                                    int��¼���� & " as ��¼����," & _
                                    "0 as ���շ�," & _
                                    "'[δ�Ʒ�]' as NO," & _
                                    "'" & str�ѱ� & "' as �ѱ�," & _
                                    curӦ�� & " as Ӧ�ս��," & _
                                    curʵ�� & " as ʵ�ս��,'" & _
                                    UserInfo.���� & "' as ������," & _
                                    "Sysdate as �Ǽ�ʱ��,'" & _
                                    UserInfo.���� & "' as ����Ա,'" & rs("NO").Value & "' AS ���͵��� " & _
                            " From Dual"
                        End If
                End If
                
                rs.MoveNext
            Loop
        End If
    End If
    
    
    '�ѼƷѵ�
    strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
        " Select 1 as ��������,A.��¼����,Decode(B.��¼״̬,1,1,0) as ���շ�," & _
        " A.NO,B.�ѱ�,Sum(B.Ӧ�ս��) as Ӧ�ս��,Sum(B.ʵ�ս��) as ʵ�ս��," & _
        " B.������,B.�Ǽ�ʱ��,Nvl(B.����Ա����,B.������) as ����Ա,'' AS ���͵��� " & _
        " From ����ҽ������ A,���˷��ü�¼ B" & _
        " Where NVL(A.�Ʒ�״̬,0)=1 AND A.ҽ��ID IN " & strTmp & _
        " And A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ��ID=B.ҽ�����+0 " & _
        " And B.��¼״̬ in (0,1) " & _
        " Group by A.��¼����,B.��¼״̬,A.NO,B.�ѱ�,B.������,B.�Ǽ�ʱ��,Nvl(B.����Ա����,B.������)"
        
    '�����ò���
    strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
            " Select 2 as ��������,A.��¼����,Decode(B.��¼״̬,1,1,0) as ���շ�," & _
            " A.NO,B.�ѱ�,Sum(B.Ӧ�ս��) as Ӧ�ս��,Sum(B.ʵ�ս��) as ʵ�ս��," & _
            " B.������,B.�Ǽ�ʱ��,Nvl(B.����Ա����,B.������) as ����Ա,'' as ���͵��� " & _
            " From ����ҽ������ A,���˷��ü�¼ B" & _
            " Where A.ҽ��ID in (select id from ����ҽ����¼ where " & lngҽ��ID & " in (ID,���id)) " & _
            " And A.NO=B.NO And A.��¼����=Decode(B.��¼����,0,1,B.��¼����)" & _
            " And B.��¼״̬ in (0,1) " & _
            " Group by A.��¼����,B.��¼״̬,A.NO,B.�ѱ�,B.������,B.�Ǽ�ʱ��,Nvl(B.����Ա����,B.������)"
    
    '����ת������
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    ElseIf mblnChargeDataMoved Then
        strTmp = strSQL
        strTmp = Replace(strTmp, "���˷��ü�¼", "H���˷��ü�¼")
        strSQL = strSQL & " Union All " & strTmp
    End If
    
    strSQL = "Select * From (" & strSQL & ") Order by ��������,�Ǽ�ʱ�� Desc"
    Call zlDatabase.OpenRecordset(rsList, strSQL, Me.Caption)
    With vsMoney
        lngRow = .FixedRows
        strPre = .TextMatrix(.Row, 1) & "_" & .TextMatrix(.Row, 2)
        
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        vsDetail.Rows = vsDetail.FixedRows
        vsDetail.Rows = vsDetail.FixedRows + 1
        
        mblnCash = True
        If Not rsList.EOF Then
            .Rows = rsList.RecordCount + 1
            For i = 1 To rsList.RecordCount
            
                '�Ƿ�ȫ�����շ�,Ҫ��δ��˵ļ��ʻ��۵�,������δ�Ʒѵ�������
                If rsList!���շ� = 0 And Nvl(rsList!ʵ�ս��, 0) <> 0 And rsList!NO <> "[δ�Ʒ�]" Then mblnCash = False
                                
                .TextMatrix(i, 0) = IIF(rsList!�������� = 1, "������", "���ӷ���")
                .TextMatrix(i, 1) = IIF(rsList!��¼���� = 1, "�շѵ���" & IIF(rsList!���շ� = 1, "��", ""), "���ʵ���")
                If rsList!��¼���� = 1 And rsList!���շ� = 1 Then '���շ���ɫ��ʾ
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '����
                End If
                
                .TextMatrix(i, 2) = rsList!NO
                .TextMatrix(i, 3) = Nvl(rsList!�ѱ�)
                .TextMatrix(i, 4) = Format(Nvl(rsList!Ӧ�ս��, 0), gstrDec)
                .TextMatrix(i, 5) = Format(Nvl(rsList!ʵ�ս��, 0), gstrDec)
                .TextMatrix(i, 6) = Nvl(rsList!������)
                .TextMatrix(i, 7) = Format(rsList!�Ǽ�ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, 8) = Nvl(rsList!����Ա)
                .TextMatrix(i, 9) = Nvl(rsList!���͵���)
                '��������
                .Cell(flexcpData, i, 1) = CInt(rsList!��¼����)
                .Cell(flexcpData, i, 7) = Format(rsList!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss")
                                                
                '�Ƿ����������
                If rsList!�������� = 1 Then blnMain = True
                If rsList!�������� = 2 Then blnSub = True
                
                '��λ��ԭ��λ��
                If .TextMatrix(i, 1) & "_" & .TextMatrix(i, 2) = strPre Then lngRow = i
                rsList.MoveNext
            Next
        Else
            mblnCash = False
        End If
        If blnMain And blnSub Then
            .FrozenRows = 1
            .Select 1, 0, 1, .Cols - 1
            .CellBorder &HC00000, 0, 0, 0, 1, 0, 0
        End If
        .Row = lngRow: .Col = .FixedCols
        Call .ShowCell(.Row, .Col)
        
        .Redraw = flexRDDirect
        Call vsMoney_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    
    Call AppendRows(vsMoney, lnX0, lnY0)
        
    LoadMoneyList = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function NewAdvicePrice(ByVal rs As ADODB.Recordset, ByRef curӦ�� As Currency, ByRef curʵ�� As Currency) As Boolean
    '----------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ��ҽ���ļƼ۹�ϵ����ʱ��¼��
    '˵����Ҫ�������ĿӦ�ò��Ƕ���,Ժ��ִ��,����Ʒ�
    '----------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim dbl���� As Double
    Dim bln�������� As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
        
    On Error GoTo ErrHand
            
    '��ȡҪ���������õ�ҽ����¼(������������,��鲿λ������������)
    strSQL = _
            " Select B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,B.��������ID," & _
            " Nvl(A.��������,Sum(Nvl(C.��������,0))) as ���� " & _
            " From ����ҽ������ A,����ҽ����¼ B,����ҽ��ִ�� C" & _
            " Where NVL(A.�Ʒ�״̬,0)=0 AND A.NO=[1] " & _
                " And A.ҽ��ID=B.ID And A.���ͺ�+0=[2] " & _
                " And C.ҽ��ID(+)=A.ҽ��ID And C.���ͺ�(+)=A.���ͺ�" & _
            " Group by B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,B.��������ID,A.��������"
            
    '����ת������
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ��ִ��", "H����ҽ��ִ��")
    End If
    
    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CStr(rs("No").Value), Val(rs("���ͺ�").Value))
    For i = 1 To rsAdvice.RecordCount
        dbl���� = Nvl(rsAdvice!����, 0)
        
        '��ȡ��Ӧ���շѼ�Ŀ:ֻ��ȡ�̶�����,�Ҳ��Ǳ�۵Ķ���
        bln�������� = (rsAdvice!������� = "F" And Not IsNull(rsAdvice!���ID))
        strSQL = IIF(bln��������, "Nvl(B.�����շ���,100)/100", "1") & " as ������"
        strSQL = _
                " Select A.�շ���ĿID,A.�շ�����,B.������ĿID,D.�վݷ�Ŀ,C.���," & _
                " C.���㵥λ,C.ִ�п���,Decode(C.�Ƿ���,1,NULL,B.�ּ�) as ����," & strSQL & _
                " From �����շѹ�ϵ A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,������Ŀ D" & _
                " Where A.������ĿID=[1] " & _
                    " And A.�շ���ĿID=B.�շ�ϸĿID And A.�շ���ĿID=C.ID And B.������ĿID=D.ID" & _
                    " And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And Nvl(C.�Ƿ���,0)=0"
                    
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsAdvice!������ĿID))
        For j = 1 To rsTmp.RecordCount
        
            mrsPrice.AddNew
            
            mrsPrice!ҽ��ID = rsAdvice!ҽ��ID
            mrsPrice!��������ID = rsAdvice!��������ID
            mrsPrice!��� = rsTmp!���
            mrsPrice!�շ�ϸĿID = rsTmp!�շ���ĿID
            mrsPrice!���㵥λ = Nvl(rsTmp!���㵥λ)
            mrsPrice!�������� = IIF(bln��������, 1, 0)
            mrsPrice!ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
            mrsPrice!������ĿID = rsTmp!������ĿID
            mrsPrice!�վݷ�Ŀ = rsTmp!�վݷ�Ŀ
            mrsPrice!���� = Format(Nvl(rsTmp!����, 0), "0.00000")
            mrsPrice!���� = Format(Nvl(rsTmp!�շ�����, 0) * dbl����, "0.00000")
            mrsPrice!Ӧ�� = Format(mrsPrice!���� * mrsPrice!���� * rsTmp!������, gstrDec)
            mrsPrice!���͵��� = rs("NO").Value
            mrsPrice!���ͺ� = rs("���ͺ�").Value
            
            Select Case mstrSys
            Case "���"
                'mrsPrice!ʵ�� = Format(msgl����ۿ� * mrsPrice!Ӧ��, gstrDec)
                mrsPrice!ʵ�� = mrsPrice!Ӧ��
            Case Else
                If str�ѱ� = "" Then
                    mrsPrice!ʵ�� = mrsPrice!Ӧ��
                Else
                    mrsPrice!ʵ�� = Format(ActualMoney(str�ѱ�, mrsPrice!������ĿID, mrsPrice!Ӧ��), gstrDec)
                End If
            End Select
            
            curӦ�� = curӦ�� + zlCommFun.Nvl(mrsPrice!Ӧ��, 0)
            curʵ�� = curʵ�� + zlCommFun.Nvl(mrsPrice!ʵ��, 0)
            
            mrsPrice.Update
            rsTmp.MoveNext
        Next
        
        rsAdvice.MoveNext
        
    Next
        
    Dim sgl�ϼ� As Single
    
    If mstrSys = "���" And msgl����ۿ� > 0 Then
        If mrsPrice.RecordCount > 0 Then mrsPrice.MoveFirst
        For j = 1 To mrsPrice.RecordCount
            mrsPrice!ʵ�� = Format(msgl����ۿ� * mrsPrice!ʵ�� / curʵ��, gstrDec)
            sgl�ϼ� = sgl�ϼ� + mrsPrice!ʵ��
            mrsPrice.MoveNext
        Next
        
        If sgl�ϼ� <> msgl����ۿ� Then
            mrsPrice!ʵ�� = mrsPrice!ʵ�� + (msgl����ۿ� - sgl�ϼ�)
        End If
        
        curʵ�� = msgl����ۿ�
        
        mrsPrice.Update
    End If
    If mrsPrice.RecordCount > 0 Then mrsPrice.MoveFirst
    
    NewAdvicePrice = True
    
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
End Function

Private Function LoadAdvicePrice(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal str�ѱ� As String, ByVal strNO As String) As Boolean
    '----------------------------------------------------------------------------------------------------
    '���ܣ���ȡָ��ҽ���ļƼ۹�ϵ����ʱ��¼��
    '˵����Ҫ�������ĿӦ�ò��Ƕ���,Ժ��ִ��,����Ʒ�
    '----------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim dbl���� As Double
    Dim bln�������� As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
        
    On Error GoTo errH
            
    '��ȡҪ���������õ�ҽ����¼(������������,��鲿λ������������)
    strSQL = _
            " Select B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,B.��������ID," & _
            " Nvl(A.��������,Sum(Nvl(C.��������,0))) as ����" & _
            " From ����ҽ������ A,����ҽ����¼ B,����ҽ��ִ�� C" & _
            " Where NVL(A.�Ʒ�״̬,0)=0 AND B.���ID=[1] " & _
                " And A.ҽ��ID=B.ID And A.���ͺ�+0=[2] " & _
                " And C.ҽ��ID(+)=A.ҽ��ID And C.���ͺ�(+)=A.���ͺ�" & _
            " Group by B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,B.��������ID,A.��������"
            
    strSQL = strSQL & " Union ALL " & _
            " Select B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,B.��������ID," & _
            " Nvl(A.��������,Sum(Nvl(C.��������,0))) as ����" & _
            " From ����ҽ������ A,����ҽ����¼ B,����ҽ��ִ�� C" & _
            " Where NVL(A.�Ʒ�״̬,0)=0 AND B.ID=[1] " & _
                " And A.ҽ��ID=B.ID And A.���ͺ�+0=[2] " & _
                " And C.ҽ��ID(+)=A.ҽ��ID And C.���ͺ�(+)=A.���ͺ�" & _
                " Group by B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,B.��������ID,A.��������" & _
            " Order by ���"
    
    If mblnDataMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ��ִ��", "H����ҽ��ִ��")
    End If
    
    Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    For i = 1 To rsAdvice.RecordCount
        dbl���� = Nvl(rsAdvice!����, 0)
        
        '��ȡ��Ӧ���շѼ�Ŀ:ֻ��ȡ�̶�����,�Ҳ��Ǳ�۵Ķ���
        bln�������� = (rsAdvice!������� = "F" And Not IsNull(rsAdvice!���ID))
        strSQL = IIF(bln��������, "Nvl(B.�����շ���,100)/100", "1") & " as ������"
        strSQL = _
                " Select A.�շ���ĿID,A.�շ�����,B.������ĿID,D.�վݷ�Ŀ,C.���," & _
                " C.���㵥λ,C.ִ�п���,Decode(C.�Ƿ���,1,NULL,B.�ּ�) as ����," & strSQL & _
                " From �����շѹ�ϵ A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,������Ŀ D" & _
                " Where A.������ĿID=" & rsAdvice!������ĿID & _
                    " And A.�շ���ĿID=B.�շ�ϸĿID And A.�շ���ĿID=C.ID And B.������ĿID=D.ID" & _
                    " And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
                    " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
                    " And Nvl(C.�Ƿ���,0)=0"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        For j = 1 To rsTmp.RecordCount
        
            mrsPrice.AddNew
            
            mrsPrice!ҽ��ID = rsAdvice!ҽ��ID
            mrsPrice!��������ID = rsAdvice!��������ID
            mrsPrice!��� = rsTmp!���
            mrsPrice!�շ�ϸĿID = rsTmp!�շ���ĿID
            mrsPrice!���㵥λ = Nvl(rsTmp!���㵥λ)
            mrsPrice!�������� = IIF(bln��������, 1, 0)
            mrsPrice!ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
            mrsPrice!������ĿID = rsTmp!������ĿID
            mrsPrice!�վݷ�Ŀ = rsTmp!�վݷ�Ŀ
            mrsPrice!���� = Format(Nvl(rsTmp!����, 0), "0.00000")
            mrsPrice!���� = Format(Nvl(rsTmp!�շ�����, 0) * dbl����, "0.00000")
            mrsPrice!Ӧ�� = Format(mrsPrice!���� * mrsPrice!���� * rsTmp!������, gstrDec)
            
            Select Case mstrSys
            Case "���"
                mrsPrice!ʵ�� = Format(msgl����ۿ� * mrsPrice!Ӧ��, gstrDec)
            Case Else
                If str�ѱ� = "" Then
                    mrsPrice!ʵ�� = mrsPrice!Ӧ��
                Else
                    mrsPrice!ʵ�� = Format(ActualMoney(str�ѱ�, mrsPrice!������ĿID, mrsPrice!Ӧ��), gstrDec)
                End If
            End Select
            
            mrsPrice.Update
            rsTmp.MoveNext
        Next
        
        rsAdvice.MoveNext
        
    Next
    If mrsPrice.RecordCount > 0 Then mrsPrice.MoveFirst
    LoadAdvicePrice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
'    Set mrsPrice = Nothing
End Function

Private Function LoadBillDetail(ByVal lngRow As Long) As Boolean
    '----------------------------------------------------------------------------------------------------
    '���ܣ���ʾ������ϸ����
    '������lngRow=�����嵥��
    '----------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strNO As String, int��¼���� As Integer
    Dim lng���˿���ID As Long, int��Դ As Integer
    Dim lng��ĿID As Long, int���� As Integer
    Dim strSQL As String, i As Long
    Dim str�Ǽ�ʱ�� As String
    Dim blnҩ����λ As Boolean, strҩ����λ As String, strҩ����װ As String

    On Error GoTo errH

    If lngRow < vsMoney.FixedRows Then Exit Function

    vsDetail.Rows = vsDetail.FixedRows
    vsDetail.Rows = vsDetail.FixedRows + 1

    With vsMoney
        strNO = .TextMatrix(lngRow, 2)
        int��¼���� = Val(.Cell(flexcpData, lngRow, 1))
        lng���˿���ID = lngPatientDept
        int��Դ = iPatientType
        
        '�Ǽ�ʱ����Ϊ������ͬʱ���������δ��˵����
        str�Ǽ�ʱ�� = "To_Date('" & .Cell(flexcpData, lngRow, 7) & "','YYYY-MM-DD HH24:MI:SS')"
        
        If strNO = "" Then Exit Function
    End With

    'ҩƷ��λ
    blnҩ����λ = Val(GetSetting("ZLSOFT", "����ģ��\" & App.ProductName, "ҩƷ��λ", 0)) <> 0
    If int��Դ = 1 Then
        strҩ����λ = "���ﵥλ": strҩ����װ = "�����װ"
    Else
        strҩ����λ = "סԺ��λ": strҩ����װ = "סԺ��װ"
    End If

    If strNO = "[δ�Ʒ�]" Then
        If mrsPrice Is Nothing Then Exit Function
        With mrsPrice
            .Filter = "���͵���='" & vsMoney.TextMatrix(lngRow, 9) & "'"
            If .RecordCount > 0 Then
                .MoveFirst
                For i = 1 To .RecordCount
                    If lng��ĿID <> !�շ�ϸĿID Then int���� = i
                    strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
                        " Select " & i & " as ���," & IIF(int���� = i, "-NULL", int����) & " as �۸񸸺�," & _
                        "'" & strNO & "' as NO," & int��¼���� & " as ��¼����,1 as ��¼״̬," & _
                        "'" & !��� & "' as �շ����," & !�շ�ϸĿID & " as �շ�ϸĿID," & _
                        Get�շ�ִ�п���ID(lngPatientID, lngPageId, !���, !�շ�ϸĿID, !ִ�п���, lng���˿���ID, Nvl(!��������ID, 0), int��Դ) & " as ִ�в���ID," & _
                        !������ĿID & " as ������ĿID,1 as ����," & !���� & " as ����," & !���� & " as ��׼����," & _
                        !Ӧ�� & " as Ӧ�ս��," & !ʵ�� & " as ʵ�ս�� From Dual"
    
                    lng��ĿID = !�շ�ϸĿID
                    .MoveNext
                Next
                If strSQL = "" Then Exit Function
                strSQL = "(" & strSQL & ")"
            Else
                strSQL = "���˷��ü�¼"
            End If
            .Filter = ""
        End With
    Else
        If zlDatabase.NOMoved("���˷��ü�¼", strNO) Then
            strSQL = "H���˷��ü�¼"
        Else
            strSQL = "���˷��ü�¼"
        End If
    End If

    strSQL = "Select C.���� as ���,Nvl(F.����,B.����)||Decode(B.���,NULL,NULL,' '||B.���) as ��Ŀ," & _
        " Sum(A.��׼����" & IIF(blnҩ����λ, "*Nvl(E." & strҩ����װ & ",1)", "") & ") as ����," & _
        " Avg(Nvl(A.����,1)*A.����" & IIF(blnҩ����λ, "/Nvl(E." & strҩ����װ & ",1)", "") & ") as ����," & _
        IIF(blnҩ����λ, "Decode(E.ҩƷID,NULL,B.���㵥λ,E." & strҩ����λ & ")", "B.���㵥λ") & " as ���㵥λ," & _
        " Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��,D.���� as ִ�в���" & _
        " From " & strSQL & " A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� E,�շ���Ŀ���� F" & _
        " Where Decode(A.��¼����,0,1,A.��¼����)=" & int��¼���� & _
        " And A.NO='" & strNO & "' And A.��¼״̬ in (0,1) And A.�շ�ϸĿID=B.ID" & _
        " And A.�շ����=C.���� And A.ִ�в���ID=D.ID(+) And B.ID=E.ҩƷID(+)" & _
        " And A.�շ�ϸĿID=F.�շ�ϸĿID(+) And F.����(+)=1 And F.����(+)=" & IIF(gbln��Ʒ��, 3, 1) & _
        IIF(strSQL = "���˷��ü�¼", " And A.�Ǽ�ʱ��=" & str�Ǽ�ʱ��, "") & _
        " Group by Nvl(A.�۸񸸺�,A.���),C.����,Nvl(F.����,B.����),B.���,B.���㵥λ,D.����,E.ҩƷID,E." & strҩ����λ
        
    strSQL = strSQL & " Order by Nvl(A.�۸񸸺�,A.���)"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)

    With vsDetail
        .Redraw = flexRDNone
        If Not rsTmp.EOF Then
            .Rows = rsTmp.RecordCount + .FixedRows
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, 0) = rsTmp!���
                .TextMatrix(i, 1) = rsTmp!��Ŀ
                .TextMatrix(i, 2) = FormatEx(rsTmp!����, 5) & " " & Nvl(rsTmp!���㵥λ)
                .TextMatrix(i, 3) = Format(rsTmp!����, "0.00000")
                .TextMatrix(i, 4) = Format(rsTmp!Ӧ�ս��, gstrDec)
                .TextMatrix(i, 5) = Format(rsTmp!ʵ�ս��, gstrDec)
                .TextMatrix(i, 6) = Nvl(rsTmp!ִ�в���)
                rsTmp.MoveNext
            Next
        End If
        .Row = .FixedRows: .Col = .FixedCols
        .Redraw = flexRDDirect
    End With
    
    Call AppendRows(vsDetail, lnX1, lnY1)
    
    LoadBillDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsMoney_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    If Button = 2 And mfrmParent.mnuCharge.Visible And mfrmParent.mnuCharge.Enabled Then PopupMenu mfrmParent.mnuCharge
    
End Sub

Private Function AppendRows(ByVal objVsf As Object, ByRef objLineX As Variant, ByRef objLineY As Variant, Optional ByVal lngHideRows As Long = 0) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '����:������ؼ��Ŀ���
    '����:objVsf Ҫ�����еı��ؼ�����
    '����:���ɹ�����True,���򷵻� False
    '--------------------------------------------------------------------------------------------------------
    Dim lngTop As Long
    Dim lngLoop As Long
    Dim lngIndex As Long
    Dim lngLastRow As Long
    
    On Error GoTo ErrHand
    
    If objVsf.Rows = 0 Then Exit Function
    
    For lngLoop = objVsf.Rows - 1 To 1 Step -1
        If objVsf.RowHidden(lngLoop) = False Then
            lngLastRow = lngLoop
            Exit For
        End If
    Next
    
    lngTop = objVsf.Cell(flexcpTop, lngLastRow, 0) + objVsf.RowHeight(lngLastRow)
    
    '1.�������е���
    For lngLoop = 1 To objLineX.UBound
        objLineX(lngLoop).Visible = False
    Next
    
    For lngLoop = 1 To objLineY.UBound
        objLineY(lngLoop).Visible = False
    Next
    
    '2.���¼�����Ҫ������
    For lngLoop = 1 To objVsf.Cols - 1

        If objLineY.UBound < lngLoop Then Load objLineY(lngLoop)

        With objLineY(lngLoop)

            .ZOrder

            .X1 = objVsf.Cell(flexcpLeft, 0, lngLoop) - 15
            .X2 = .X1
            .Y1 = lngTop
            .Y2 = objVsf.Height

            .BorderColor = objVsf.GridColor

            .Visible = True
        End With

    Next

    '3.���¼�����Ҫ�ĺ���
    lngIndex = 0
    Do While (lngTop + objVsf.RowHeight(0)) < objVsf.Height

        lngIndex = lngIndex + 1
        If objLineX.UBound < lngIndex Then Load objLineX(lngIndex)

        With objLineX(lngIndex)

            .ZOrder

            .X1 = 0
            .X2 = objVsf.Width
            .Y1 = lngTop + objVsf.RowHeight(0) + 15
            .Y2 = .Y1

            .BorderColor = objVsf.GridColor

            .Visible = True

            lngTop = .Y1
        End With

    Loop
        
    AppendRows = True
    
    Exit Function
    
ErrHand:
    
End Function


