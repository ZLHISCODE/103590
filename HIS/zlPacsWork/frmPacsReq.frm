VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPACSReq 
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
   Begin VB.Frame fraSplit1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   6
      Top             =   3245
      Width           =   7110
   End
   Begin VB.Frame fraFee 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   3360
      Width           =   7935
      Begin VB.Label lblCash 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   7560
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   " ����"
         ForeColor       =   &H8000000E&
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   50
         Width           =   450
      End
   End
   Begin VB.PictureBox picFile 
      Height          =   2055
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   7875
      TabIndex        =   3
      Top             =   0
      Width           =   7935
   End
   Begin VB.Frame fraBill 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   0
      Top             =   4855
      Width           =   7110
   End
   Begin VSFlex8Ctl.VSFlexGrid vsMoney 
      Height          =   1140
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   7185
      _cx             =   12674
      _cy             =   2011
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      GridLinesFixed  =   1
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
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDetail 
      Height          =   1380
      Left            =   0
      TabIndex        =   2
      Top             =   5040
      Width           =   7200
      _cx             =   12700
      _cy             =   2434
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
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
      GridLinesFixed  =   1
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
   End
End
Attribute VB_Name = "frmPACSReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_LOST = &HFFEBD7
Private Const COLOR_FOCUS = &HFFCC99

Private mrsPrice As ADODB.Recordset 'δ�Ʒ�ҽ����������
Private objBillForm As Object
Private WithEvents frmParent As Form
Attribute frmParent.VB_VarHelpID = -1

Private pgbLoad As Object
Private AdviceID As Long, lngSendNO As Long
Private iPatientType As Integer, lngPatientID As Long, lngPatientDept As Long
Private lngPageId As Long, strCheckNo As String
Private str�Ʒ�״̬ As String, str�ѱ� As String, int��¼���� As Integer
Private intִ��״̬ As Integer, strNO As String, lng��������ID As Long
Private mstrPrivs As String
Private mblnMoved As Boolean

Public Sub zlRefresh(objParent As Object, ByVal lngAdviceID As Long, ByVal SendNO As Long, _
    ByVal strPrivs As String, Optional objpgbLoad As Object, Optional ByVal blnMoved As Boolean = False)
    
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    
    If objBillForm Is Nothing Then Exit Sub
    On Error GoTo DBError
    mblnMoved = blnMoved
    
    strSQL = _
        " Select X.��¼���� as ��������,X.��¼״̬ as ����״̬," & _
        " A.ҽ��ID,A.���ͺ�,B.���ID,B.���,B.�������,B.������ĿID,A.����ʱ�� as ʱ��,A.NO," & _
        " A.��¼����,A.ִ��״̬,A.�Ʒ�״̬,B.����ID,B.��ҳID,B.�Һŵ�,B.���˿���ID,E.���� as ����,D.����," & _
        " Decode(B.������Դ,1,D.�����,2,D.סԺ��,4,D.�����,NULL) as ��ʶ��,Nvl(F.�ѱ�,D.�ѱ�) as �ѱ�," & _
        " Decode(B.������Դ,1,'����',2,'סԺ',3,'����',4,'���') as ��Դ,C.���� as ����,A.ִ�м�,A.ִ�в���ID" & _
        " From ����ҽ������ A,����ҽ����¼ B,������ĿĿ¼ C,������Ϣ D,���ű� E,������ҳ F,���˷��ü�¼ X" & _
        " Where A.ҽ��ID=B.ID And B.������ĿID=C.ID And B.����ID=D.����ID" & _
        " And B.���˿���ID=E.ID And B.����ID=F.����ID(+) And B.��ҳID=F.��ҳID(+)" & _
        " And A.NO=X.NO(+) And A.��¼����=Decode(X.��¼����(+),0,1,X.��¼����(+))" & _
        " And X.��¼״̬(+)<>2 And X.ҽ�����(+)=A.ҽ��ID And X.���(+)=1" & _
        " And A.ҽ��ID= [1]  And A.���ͺ�= [2] " & _
        " Order by A.����ʱ�� Desc,B.����ID,B.���"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, SendNO)
   
    Set frmParent = objParent
    Set pgbLoad = objpgbLoad
    AdviceID = lngAdviceID: lngSendNO = SendNO: iPatientType = 1
    lngPatientID = 0: lngPageId = 0: strCheckNo = "": lngPatientDept = 0
    str�Ʒ�״̬ = "": str�ѱ� = "": int��¼���� = 1: mstrPrivs = strPrivs
    intִ��״̬ = 0: strNO = "": lng��������ID = 0
    If Not rsTmp.EOF Then
        iPatientType = Decode(rsTmp("��Դ"), "����", 1, "���", 1, 2)
        lngPatientID = rsTmp("����ID"): lngPageId = Nvl(rsTmp("��ҳID"), 0): strCheckNo = Nvl(rsTmp("�Һŵ�"), "")
        lngPatientDept = Nvl(rsTmp("���˿���ID"), 0)
        str�Ʒ�״̬ = GetSendMoneyState(lngAdviceID, SendNO): str�ѱ� = Nvl(rsTmp!�ѱ�): int��¼���� = Nvl(rsTmp!��¼����, 1)
        intִ��״̬ = Nvl(rsTmp!ִ��״̬, 0): strNO = Nvl(rsTmp!NO): lng��������ID = Nvl(rsTmp!ִ�в���ID, 0)
    End If
    ShowMenu
    
    If frmParent.Visible Then
        objBillForm.ShowMe AdviceID, pgbLoad
        Call LoadMoneyList(AdviceID, lngSendNO, str�Ʒ�״̬, str�ѱ�, int��¼����)
    Else
        Me.Tag = "Loading":
    End If
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlMenuClick(objMenu As Menu)
    Dim strMenu As String
    
    If objMenu.Caption Like "*(&*)*" Then
        strMenu = Split(objMenu.Caption, "(&")(0)
    Else
        strMenu = objMenu.Caption
    End If
    Select Case strMenu
        Case "����������"
                MoneyMain
        Case "�޸ĸ��ӷ���"
                MoneyModi
        Case "ɾ�����ӷ���"
                MoneyDel
        Case "�շѵ���"
                Call MoneyNewBilling(1)
        Case "���ʵ���"
                Call MoneyNewBilling(2)
        Case "��Ѻ��õǼ�"
                Call MoneyNewBilling(2, True)
    End Select
End Sub

Public Sub zlButtonClick(objButton As MSComctlLib.Button)
    Select Case objButton.Key
        Case "����"
            MoneyMain
        Case "����"
            frmParent.PopupMenu frmParent.mnuMoneyFunc(2)
        Case "�ķ�"
            MoneyModi
        Case "ɾ��"
            MoneyDel
    End Select
End Sub

Public Sub zlPrint(ByVal bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    Dim bytR As Byte, i As Long
    Dim lngRow As Long, lngCol As Long
    Dim strWidth As String
    Dim objGrid As Object
    
    On Error Resume Next
    If frmParent.lvwPati.SelectedItem Is Nothing Then Exit Sub
    
    '��ͷ
    objOut.Title.Text = "���˷����嵥"
    Set objGrid = Me.vsMoney
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True

    '����
    With frmParent.lvwPati.SelectedItem
        Set objRow = New zlTabAppRow
        objRow.Add "���ˣ�" & .SubItems(2) & " ��Դ��" & .Text & " ��ʶ�ţ�" & .SubItems(6)
        objRow.Add "���ݣ�" & .SubItems(1) & " ���ݣ�" & .SubItems(3)
        objOut.UnderAppRows.Add objRow
    End With

    '����
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow

    '����
    Set objOut.Body = objGrid

    '���
    objGrid.Redraw = False
    lngRow = objGrid.Row: lngCol = objGrid.Col

    strWidth = ""
    For i = 0 To objGrid.Cols - 1
        strWidth = strWidth & "," & objGrid.ColWidth(i)
        If i <= objGrid.FixedCols - 1 Or objGrid.ColHidden(i) Then
            objGrid.ColWidth(i) = 0
        End If
    Next

    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If

    strWidth = Mid(strWidth, 2)
    For i = 0 To objGrid.Cols - 1
        objGrid.ColWidth(i) = Split(strWidth, ",")(i)
    Next
    objGrid.Row = lngRow: objGrid.Col = lngCol
    objGrid.Redraw = True
End Sub

Private Sub MoneyMain()
    Dim rsPati As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    
    Dim lng����ID As Long, lng��ҳID As Long, lng���ͺ�, lngҽ��ID As Long
    Dim int��Դ As Integer
    Dim int���� As Integer, lng��ĿID As Long, lngִ�в���ID As Long
    Dim lng���˲���ID As Long, lng���˿���ID As Long, lng���ID As Long
    Dim arrSQL As Variant, strSQL As String, strDate As String, i As Long
    Dim int������Ŀ�� As Integer, lng���մ���ID As Long, str���ձ��� As String, curͳ���� As Currency
    Dim lng��������ID As Long, str����ҽ�� As String, int��� As Integer, strMsg As String
    
    If AdviceID = 0 Then Exit Sub
    If InStr(str�Ʒ�״̬, ",-1,") > 0 Then
        MsgBox "��ִ����Ŀ����Ʒѡ�" & vbCrLf & "�����Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
        Exit Sub
    ElseIf InStr(str�Ʒ�״̬, ",1,") > 0 Then
        MsgBox "��ִ����Ŀ���������Ѿ��Ʒѡ�" & vbCrLf & "�����Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
        Exit Sub
    End If
    If mrsPrice Is Nothing Then Exit Sub
    If mrsPrice.RecordCount = 0 Then
        MsgBox "��ִ����Ŀû�п��ԼƷѵ������á�" & vbCrLf & "�����Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
        Exit Sub
    End If
    If intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If int��¼���� = 1 Then
        If BillExistBalance(strNO) Then
            MsgBox "���� " & strNO & " �Ѿ��շѣ��������������ŵ��ݵ������á�" & vbCrLf & "�����Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    If MsgBox("ȷʵҪ���ɸ���Ŀ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
        
    Screen.MousePointer = 11
    
    lng���ͺ� = lngSendNO
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
        " Where A.����ID= [1] And A.����ID=B.����ID(+)" & _
        " And B.��ҳID(+)= [2] And A.ҽ�Ƹ��ʽ=C.����(+)"
    Set rsPati = OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��ҳID)
    
    '���ܶ��շ���ΪҩƷ����
    If int��¼���� = 1 Then
        lng���ID = ExistIOClass(8) '���ﻮ�۵�
    Else
        lng���ID = ExistIOClass(9) '����/סԺ���ʵ�
    End If
    
    '���ܷ���ʱ���Զ������˲���������,�������ֹ�����ʣ�ಿ�ݡ�
    '1.��Ϊ���ݺ���ͬ,����Ҫ�����������
    '2.����������շѻ��۵���Ҫ��֤һ�ŵ����еǼ�ʱ����ͬ(��Ȼ�շ��޷�����)
    '3.��2����������������������Ѿ��շѣ�������������������
    int��� = GetBillMax���(strNO, int��¼����, strDate)
    If int��¼���� = 2 Or strDate = "" Then
        strDate = "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    Else
        strDate = "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')"
    End If
    
    arrSQL = Array()
    With mrsPrice
        .MoveFirst
        For i = 1 To .RecordCount
            '��ȡ��Ӧ��ҽ����Ϣ
            If lngҽ��ID <> !ҽ��ID Then
                strSQL = "Select ҽ����Ч,���˿���ID,��������ID,����ҽ��,Ӥ��,ִ��Ƶ��,�Ƽ�����" & _
                    " From ����ҽ����¼ Where ID= [1] "
                Set rsAdvice = OpenSQLRecord(strSQL, Me.Caption, CLng(!ҽ��ID))
                
                '����ǰ�����Ʒ�ҽ�����Ϊ�ѼƷ�
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "ZL_����ҽ������_�Ʒ�(" & !ҽ��ID & "," & lng���ͺ� & ")"
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
            
            '�������Ҽ�������
            lng��������ID = rsAdvice!��������ID
            str����ҽ�� = rsAdvice!����ҽ��
            
            'ÿ���շ���Ŀ�Ĵ���
            If lng��ĿID <> !�շ�ϸĿID Then
                int���� = int��� '��ȡ�۸񸸺�
                lngִ�в���ID = Get�շ�ִ�п���ID(!���, !�շ�ϸĿID, !ִ�п���, Nvl(rsAdvice!���˿���ID, 0), int��Դ)
                            
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
            
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If int��Դ = 1 Then
                If int��¼���� = 1 Then
                    '�������ﻮ�۵���
                    arrSQL(UBound(arrSQL)) = _
                        "zl_���ﻮ�ۼ�¼_Insert('" & strNO & "'," & int��� & "," & lng����ID & ",NULL," & _
                        ZVal(Nvl(rsPati!�����, 0)) & ",'" & Nvl(rsPati!������) & "','" & Nvl(rsPati!����) & "'," & _
                        "'" & Nvl(rsPati!�Ա�) & "','" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�ѱ�) & "',NULL," & _
                        lng���˲���ID & "," & lng���˿���ID & "," & lng��������ID & ",'" & str����ҽ�� & "'," & _
                        "NULL," & lng��ĿID & ",'" & !��� & "','" & !���㵥λ & "',NULL,1," & !���� & "," & _
                        !�������� & "," & ZVal(lngִ�в���ID) & "," & IIf(int���� = int���, "NULL", int����) & "," & _
                        !������ĿID & ",'" & Nvl(!�վݷ�Ŀ) & "'," & !���� & "," & !Ӧ�� & "," & !ʵ�� & "," & _
                        strDate & "," & strDate & ",NULL,'" & UserInfo.���� & "'," & ZVal(lng���ID) & ",NULL," & _
                        !ҽ��ID & ",'" & Nvl(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & Nvl(rsAdvice!ҽ����Ч, 0) & "," & _
                        Nvl(rsAdvice!�Ƽ�����, 0) & ",1)"
                Else
                    '����������ʵ���
                    arrSQL(UBound(arrSQL)) = _
                        "zl_������ʼ�¼_Insert('" & strNO & "'," & int��� & "," & lng����ID & "," & _
                        ZVal(Nvl(rsPati!�����, 0)) & ",'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�Ա�) & "'," & _
                        "'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�ѱ�) & "',NULL," & ZVal(Nvl(rsAdvice!Ӥ��, 0)) & "," & _
                        lng���˲���ID & "," & lng���˿���ID & "," & lng��������ID & "," & _
                        "'" & str����ҽ�� & "',NULL," & lng��ĿID & ",'" & !��� & "'," & _
                        "'" & !���㵥λ & "',1," & !���� & "," & !�������� & "," & ZVal(lngִ�в���ID) & "," & _
                        IIf(int���� = int���, "NULL", int����) & "," & !������ĿID & ",'" & Nvl(!�վݷ�Ŀ) & "'," & !���� & "," & _
                        !Ӧ�� & "," & !ʵ�� & "," & strDate & "," & strDate & ",NULL,NULL,'" & UserInfo.��� & "'," & _
                        "'" & UserInfo.���� & "'," & ZVal(lng���ID) & ",NULL,NULL," & !ҽ��ID & "," & _
                        "'" & Nvl(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & Nvl(rsAdvice!ҽ����Ч, 0) & "," & _
                        Nvl(rsAdvice!�Ƽ�����, 0) & ")"
                End If
            Else
                '����סԺ���ʵ���
                arrSQL(UBound(arrSQL)) = _
                    "zl_סԺ���ʼ�¼_Insert('" & strNO & "'," & int��� & "," & lng����ID & "," & ZVal(lng��ҳID) & "," & _
                    ZVal(Nvl(rsPati!סԺ��, 0)) & ",'" & Nvl(rsPati!����) & "','" & Nvl(rsPati!�Ա�) & "'," & _
                    "'" & Nvl(rsPati!����) & "'," & ZVal(Nvl(rsPati!����)) & ",'" & Nvl(rsPati!�ѱ�) & "'," & _
                    lng���˲���ID & "," & lng���˿���ID & ",NULL," & ZVal(Nvl(rsAdvice!Ӥ��, 0)) & "," & _
                    lng��������ID & ",'" & str����ҽ�� & "',NULL," & lng��ĿID & ",'" & !��� & "'," & _
                    "'" & !���㵥λ & "'," & int������Ŀ�� & "," & ZVal(lng���մ���ID) & ",'" & str���ձ��� & "'," & _
                    "1," & !���� & "," & !�������� & "," & ZVal(lngִ�в���ID) & "," & _
                    IIf(int���� = int���, "NULL", int����) & "," & !������ĿID & ",'" & Nvl(!�վݷ�Ŀ) & "'," & !���� & "," & _
                    !Ӧ�� & "," & !ʵ�� & "," & curͳ���� & "," & strDate & "," & strDate & ",NULL,NULL," & _
                    "'" & UserInfo.��� & "','" & UserInfo.���� & "',NULL," & ZVal(lng���ID) & ",NULL,NULL,NULL," & _
                    !ҽ��ID & ",'" & Nvl(rsAdvice!ִ��Ƶ��) & "',NULL,NULL," & Nvl(rsAdvice!ҽ����Ч, 0) & "," & _
                    Nvl(rsAdvice!�Ƽ�����, 0) & ",NULL)"
            End If
            
            int��� = int��� + 1
            
            .MoveNext
        Next
    End With
    On Error GoTo errH
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSQL)
        Call ExecuteProc(arrSQL(i), Me.Caption)
    Next
    
    '���ύǰ����ҽ������
    If int��Դ = 2 And Not IsNull(rsPati!����) Then
        If gclsInsure.GetCapability(support�����ϴ�, , rsPati!����) And Not gclsInsure.GetCapability(support������ɺ��ϴ�, , rsPati!����) Then
            strMsg = ""
            If Not gclsInsure.TranChargeDetail(2, strNO, 2, 1, strMsg, , rsPati!����) Then
                gcnOracle.RollbackTrans
                If strMsg <> "" Then MsgBox strMsg, vbInformation, gstrSysName
                Screen.MousePointer = 0: Exit Sub
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
    Me.Tag = "Loading": Call Form_Activate
    Exit Sub
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub MoneyModi()
    Dim lng����ID As Long, lng��ҳID As Long
    Dim lng���˿���ID As Long
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim int������Դ As Integer, int��¼���� As Integer
    Dim strNO As String, bln��� As Boolean
    
    If AdviceID = 0 Then Exit Sub
    If vsMoney.TextMatrix(vsMoney.Row, 0) = "������" Then
        MsgBox "ִ����Ŀ�������ò����޸ġ������Ҫ��������ֹ����丽�ӷ��á�", vbInformation, gstrSysName
        Exit Sub
    End If
    If intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsMoney
        strNO = .TextMatrix(.Row, 2)
        If strNO = "" Or strNO = "[δ�Ʒ�]" Then Exit Sub
        int��¼���� = .Cell(flexcpData, .Row, 1)
        
        If InStr(.TextMatrix(.Row, 1), "��") > 0 Then
            MsgBox "�õ����Ѿ��շѣ��������޸ġ�", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    lngҽ��ID = AdviceID
    lng���ͺ� = lngSendNO
    lng����ID = lngPatientID
    lng��ҳID = lngPageId
    int������Դ = iPatientType
    lng���˿���ID = lngPatientDept
    
    If EditExpense(Me, 0, int��¼����, mstrPrivs, strNO, lngҽ��ID, lng���ͺ�, lng����ID, lng��ҳID, _
        int������Դ, lng��������ID, lng���˿���ID, False) Then
        'ˢ��
        Me.Tag = "Loading": Call Form_Activate
    End If
End Sub

Private Sub MoneyDel()
    Dim int������Դ As Integer, int��¼���� As Integer
    Dim strNO As String
    
    If AdviceID = 0 Then Exit Sub
    If vsMoney.TextMatrix(vsMoney.Row, 0) = "������" Then
        MsgBox "ִ����Ŀ�������ò���ɾ����", vbInformation, gstrSysName
        Exit Sub
    End If
    If intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsMoney
        strNO = .TextMatrix(.Row, 2)
        If strNO = "" Or strNO = "[δ�Ʒ�]" Then Exit Sub
        int��¼���� = .Cell(flexcpData, .Row, 1)
    
        If InStr(.TextMatrix(.Row, 1), "��") > 0 Then
            MsgBox "�õ����Ѿ��շѣ�������ɾ����", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    int������Դ = iPatientType
    
    If EditExpense(Me, 3, int��¼����, mstrPrivs, strNO, 0, 0, 0, 0, _
        int������Դ, 0, 0) Then
        'ˢ��
        Me.Tag = "Loading": Call Form_Activate
    End If
End Sub

Private Sub MoneyNewBilling(ByVal iRecordType As Integer, Optional OnlyRecord As Boolean = False)
    Dim lng����ID As Long, lng��ҳID As Long
    Dim lng���˿���ID As Long
    Dim lngҽ��ID As Long, lng���ͺ� As Long
    Dim int������Դ As Integer
    
    If AdviceID = 0 Then Exit Sub
    
    If intִ��״̬ = 1 Then
        MsgBox "��ִ����Ŀ�Ѿ�ִ����ɣ������ټ���������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    lngҽ��ID = AdviceID
    lng���ͺ� = lngSendNO
    lng����ID = lngPatientID
    lng��ҳID = lngPageId
    int������Դ = iPatientType
    lng���˿���ID = lngPatientDept
    
    If EditExpense(Me, 0, iRecordType, mstrPrivs, "", lngҽ��ID, lng���ͺ�, lng����ID, lng��ҳID, _
        int������Դ, lng��������ID, lng���˿���ID, OnlyRecord) Then
        'ˢ��
        Me.Tag = "Loading": Call Form_Activate
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If Me.Tag = "Loading" Then
        frmParent.Refresh: Me.Tag = ""
        pgbLoad.Visible = True
        objBillForm.ShowMe AdviceID, pgbLoad
        Call LoadMoneyList(AdviceID, lngSendNO, str�Ʒ�״̬, str�ѱ�, int��¼����)
        pgbLoad.Visible = False
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ShowError
    
    Set mrsPrice = Nothing
    
    Set objBillForm = getRequestForm
    
    SetWindowLong objBillForm.Hwnd, GWL_STYLE, WS_CHILD
    objBillForm.Show , Me
    SetParent objBillForm.Hwnd, picFile.Hwnd
    
    InitMoneyTable
    InitDetailTable
    Exit Sub
ShowError:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    On Error Resume Next
    With Me.fraSplit1
        .Left = 0: .Width = Me.ScaleWidth
        .Top = Me.ScaleHeight - (Me.vsDetail.Top + Me.vsDetail.Height - .Top)
    End With
    With Me.fraBill
        .Left = 0: .Width = Me.ScaleWidth
        .Top = Me.ScaleHeight - (Me.vsDetail.Top + Me.vsDetail.Height - .Top)
    End With
    
    With picFile
        .Top = 0: .Left = 0
        .Width = Me.ScaleWidth: .Height = Me.fraSplit1.Top - .Top
    End With
    With fraFee
        .Left = 0: .Top = Me.fraSplit1.Top + Me.fraSplit1.Height
        .Width = Me.ScaleWidth
        
        lblCash.Left = .Width - lblCash.Width
    End With
    With Me.vsMoney
        .Left = 0: .Top = Me.fraFee.Top + Me.fraFee.Height
        .Width = Me.ScaleWidth: .Height = Me.fraBill.Top - .Top
    End With
    With Me.vsDetail
        .Left = 0: .Top = Me.fraBill.Top + Me.fraBill.Height
        .Width = Me.ScaleWidth: .Height = Me.ScaleHeight - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload objBillForm
    Set objBillForm = Nothing
    
    Set mrsPrice = Nothing
End Sub

Private Sub fraBill_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    With fraBill
        .BackColor = RGB(0, 0, 0)
        If .Top + y - Me.vsMoney.Top < 1000 Then
            .Top = Me.vsMoney.Top + 1000
        ElseIf Me.ScaleHeight - .Top - y < 1000 Then
            .Top = Me.ScaleHeight - 1000
        Else
            .Top = .Top + y
        End If
    End With
End Sub

Private Sub fraBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    fraBill.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub fraSplit1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    With fraSplit1
        .BackColor = RGB(0, 0, 0)
        If .Top + y - Me.picFile.Top < 3000 Then
            .Top = Me.picFile.Top + 3000
        ElseIf Me.ScaleHeight - .Top - y < 2100 Then
            .Top = Me.ScaleHeight - 2100
        Else
            .Top = .Top + y
        End If
    End With
End Sub

Private Sub fraSplit1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub

    fraSplit1.BackColor = Me.BackColor
    Form_Resize
End Sub

Private Sub frmParent_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub picFile_Resize()
    Dim vRect As RECT
    
    On Error Resume Next
    If Not objBillForm Is Nothing Then
        MoveWindow objBillForm.Hwnd, 0, 0, picFile.ScaleWidth / Screen.TwipsPerPixelX, picFile.ScaleHeight / Screen.TwipsPerPixelY, 1
        Call GetWindowRect(objBillForm.Hwnd, vRect)
        SetWindowPos objBillForm.Hwnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED
    End If
End Sub

Private Sub vsMoney_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = OldRow Then Exit Sub
    If NewCol >= vsMoney.FixedCols And NewRow >= vsMoney.FixedRows Then
        vsMoney.ForeColorSel = vsMoney.Cell(flexcpForeColor, NewRow, 0)
        Call LoadBillDetail(NewRow)
    End If
End Sub

Private Sub vsMoney_DblClick()
    If vsMoney.MouseRow >= vsMoney.FixedRows Then
        Call vsMoney_KeyPress(13)
    End If
End Sub

Private Sub vsMoney_GotFocus()
    vsMoney.BackColorSel = COLOR_FOCUS
End Sub

Private Sub vsMoney_KeyPress(KeyAscii As Integer)
    Dim int������Դ As Integer, int��¼���� As Integer
    Dim strNO As String
    
    If KeyAscii = 13 Or KeyAscii = 32 Then
        KeyAscii = 0
        With vsMoney
            strNO = .TextMatrix(.Row, 2)
            If strNO = "" Or strNO = "[δ�Ʒ�]" Then Exit Sub
            int��¼���� = .Cell(flexcpData, .Row, 1)
        End With
        int������Դ = iPatientType
        
        Call EditExpense(Me, 1, int��¼����, mstrPrivs, strNO, 0, 0, 0, 0, _
            int������Դ, 0, 0, False)
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
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "��������,900,1;��������,1000,1;���ݺ�,900,1;�ѱ�,900,1;Ӧ�ս��,1000,7;ʵ�ս��,1000,7;��������,1000,1;������,750,1;�Ǽ�ʱ��,1080,1;�Ǽ���,750,1"
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
    End With
End Sub

Private Sub InitDetailTable()
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = "���,650,1;��Ŀ,2000,1;����,1000,1;����,1000,7;Ӧ�ս��,1000,7;ʵ�ս��,1000,7;ִ�п���,1000,1"
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
    End With
End Sub

Private Function LoadMoneyList(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, _
    ByVal str�Ʒ�״̬ As String, ByVal str�ѱ� As String, ByVal int��¼���� As Integer) As Boolean
'���ܣ���ȡָ��ҽ������Ҫ���ü����ӷ���
'˵����1.����ҽ������������ü����ӷ���,�����ÿ�����δ����
'      2.Ŀǰ�����ݲ�֧�ֲ����˷�,�����嵥��ֻ�����ʾ
    Dim rsTmp As New ADODB.Recordset
    Dim rsList As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim blnMain As Boolean, blnSub As Boolean
    Dim strPre As String, lngRow As Long
    Dim curӦ�� As Currency, curʵ�� As Currency
    Dim strҽ��ID As String, lngMain As Long
    Dim blnCash As Boolean  '���Ѿ��շѣ�ֻҪ���ڼ��ʼ�¼��һ��Ϊ�շѵļ�¼������Ϊδ�շѣ�
    
    On Error GoTo errH

    Set mrsPrice = Nothing
    '����δ�Ʒ�״̬,ֱ�Ӷ�ȡ�շѹ�ϵ��ʾ
    If InStr(str�Ʒ�״̬, ",0,") > 0 Then
        Call LoadAdvicePrice(lngҽ��ID, lng���ͺ�, str�ѱ�)
        If mrsPrice.RecordCount > 0 Then
            For i = 1 To mrsPrice.RecordCount
                curӦ�� = curӦ�� + Nvl(mrsPrice!Ӧ��, 0)
                curʵ�� = curʵ�� + Nvl(mrsPrice!ʵ��, 0)
                mrsPrice.MoveNext
            Next
            
            strSQL = "Select B.���� as ��������,����ҽ�� From ����ҽ����¼ A,���ű� B Where A.��������ID=B.ID And A.ID= [1] "
            If mblnMoved Then
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            End If
            Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID)
            If Not rsTmp.EOF Then
                strSQL = "Select " & _
                    " 1 as ��������," & int��¼���� & " as ��¼����,0 as ���շ�,'[δ�Ʒ�]' as NO," & _
                    "'" & str�ѱ� & "' as �ѱ�," & curӦ�� & " as Ӧ�ս��," & curʵ�� & " as ʵ�ս��," & _
                    "'" & Nvl(rsTmp!��������) & "' as ��������,'" & Nvl(rsTmp!����ҽ��) & "' as ������," & _
                    " Sysdate as �Ǽ�ʱ��,'" & UserInfo.���� & "' as ����Ա" & _
                    " From Dual"
            Else
                strSQL = ""
            End If
        End If
    End If
    
    '�����ѼƷ�״̬,Ӧ�ÿ���ֱ�Ӷ�ȡ�����ò���(ֻ��һ�ŵ���,���ܺ�����ҽ������;��ɾ�����˷�����,����ʾ)
    If InStr(str�Ʒ�״̬, ",1,") > 0 Then
        '������鲿λ������������������ϵķ���
        strҽ��ID = _
            " Select ID From ����ҽ����¼ Where ID= [1] Union All" & _
            " Select ID From ����ҽ����¼ Where ���ID= [1] And ������� IN('F','D')"
        strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
            " Select 1 as ��������,A.��¼����,Decode(B.��¼״̬,1,1,0) as ���շ�," & _
            " A.NO,B.�ѱ�,Sum(B.Ӧ�ս��) as Ӧ�ս��,Sum(B.ʵ�ս��) as ʵ�ս��," & _
            " C.���� as ��������,B.������,B.�Ǽ�ʱ��,Nvl(B.����Ա����,B.������) as ����Ա" & _
            " From ����ҽ������ A,���˷��ü�¼ B,���ű� C" & _
            " Where A.ҽ��ID IN(" & strҽ��ID & ") And A.���ͺ�= [2] " & _
            " And A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ��ID=B.ҽ�����+0" & _
            " And B.��¼״̬ IN(0,1) And B.��������ID=C.ID" & _
            " Group by A.��¼����,B.��¼״̬,A.NO,B.�ѱ�,C.����,B.������,B.�Ǽ�ʱ��,Nvl(B.����Ա����,B.������)"
    End If
    
    '�����ò���(��ɾ�����˷�����,����ʾ)
    'ҽ��ID������������,�Ը��ӷ��ö�����ͬ�����
    strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
        " Select 2 as ��������,A.��¼����,Decode(B.��¼״̬,1,1,0) as ���շ�," & _
        " A.NO,B.�ѱ�,Sum(B.Ӧ�ս��) as Ӧ�ս��,Sum(B.ʵ�ս��) as ʵ�ս��," & _
        " C.���� as ��������,B.������,B.�Ǽ�ʱ��,Nvl(B.����Ա����,B.������) as ����Ա" & _
        " From ����ҽ������ A,���˷��ü�¼ B,���ű� C" & _
        " Where A.ҽ��ID= [1]  And A.���ͺ�= [2] " & _
        " And A.NO=B.NO And A.��¼����=B.��¼���� And A.ҽ��ID=B.ҽ�����+0" & _
        " And B.��¼״̬ IN(0,1) And B.��������ID=C.ID" & _
        " Group by A.��¼����,B.��¼״̬,A.NO,B.�ѱ�,C.����,B.������,B.�Ǽ�ʱ��,Nvl(B.����Ա����,B.������)"
        
    strSQL = "Select * From (" & strSQL & ") Order by ��������,�Ǽ�ʱ�� Desc"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
    End If
    Set rsList = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    With vsMoney
        lngRow = .FixedRows
        strPre = .TextMatrix(.Row, 1) & "_" & .TextMatrix(.Row, 2)
        
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = .FixedRows + 1
        vsDetail.Rows = vsDetail.FixedRows
        vsDetail.Rows = vsDetail.FixedRows + 1
        
        blnCash = True
        If Not rsList.EOF Then
            .Rows = rsList.RecordCount + 1
            For i = 1 To rsList.RecordCount
                If rsList!���շ� = 0 And Nvl(rsList!ʵ�ս��, 0) <> 0 And rsList!NO <> "[δ�Ʒ�]" Then
                    blnCash = False 'Ҫ��δ��˵ļ��ʻ��۵�,������δ�Ʒѵ�������
                End If
                
                .TextMatrix(i, 0) = IIf(rsList!�������� = 1, "������", "���ӷ���")
                
                '�Ƿ�������շ�(��ͻ���շѵ���)
                .TextMatrix(i, 1) = IIf(rsList!��¼���� = 1, "�շѵ���" & IIf(rsList!���շ� = 1, "��", ""), "���ʵ���")
                If rsList!��¼���� = 1 And rsList!���շ� = 1 Then '���շ���ɫ��ʾ
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &HC00000 '����
                End If
                .TextMatrix(i, 2) = rsList!NO
                .TextMatrix(i, 3) = Nvl(rsList!�ѱ�)
                .TextMatrix(i, 4) = Format(Nvl(rsList!Ӧ�ս��, 0), gstrDec)
                .TextMatrix(i, 5) = Format(Nvl(rsList!ʵ�ս��, 0), gstrDec)
                .TextMatrix(i, 6) = Nvl(rsList!��������)
                .TextMatrix(i, 7) = Nvl(rsList!������)
                .TextMatrix(i, 8) = Format(rsList!�Ǽ�ʱ��, "MM-dd HH:mm")
                .TextMatrix(i, 9) = Nvl(rsList!����Ա)
                                                
                '��������
                .Cell(flexcpData, i, 1) = CInt(rsList!��¼����)
                .Cell(flexcpData, i, 8) = Format(rsList!�Ǽ�ʱ��, "yyyy-MM-dd HH:mm:ss")
                                                
                '�Ƿ����������
                If rsList!�������� = 1 Then blnMain = True
                If rsList!�������� = 2 Then
                    If blnMain Then lngMain = i - 1
                    blnSub = True
                End If
                
                '��λ��ԭ��λ��
                If .TextMatrix(i, 1) & "_" & .TextMatrix(i, 2) = strPre Then lngRow = i
                rsList.MoveNext
            Next
        Else
            blnCash = False
        End If
        If blnMain And blnSub Then
            .FrozenRows = lngMain
            .Select lngMain, 0, lngMain, .Cols - 1
            .CellBorder &HC00000, 0, 0, 0, 1, 0, 0
        End If
        .Row = lngRow: .Col = .FixedCols
        Call .ShowCell(.Row, .Col)
        
        .Redraw = flexRDDirect
        Call vsMoney_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    LoadMoneyList = True
    
    Me.lblCash.Caption = IIf(blnCash, "��", "")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadAdvicePrice(ByVal lngҽ��ID As Long, ByVal lng���ͺ� As Long, ByVal str�ѱ� As String) As Boolean
'���ܣ���ȡָ��ҽ���ļƼ۹�ϵ����ʱ��¼��
'˵����Ҫ�������ĿӦ�ò��Ƕ���,Ժ��ִ��,����Ʒ�
    Dim rsTmp As New ADODB.Recordset
    Dim rsAdvice As New ADODB.Recordset
    Dim dbl���� As Double, bln�������� As Boolean
    Dim strSQL As String, strҽ��ID As String
    Dim i As Long, j As Long
    
    Set mrsPrice = New ADODB.Recordset
    mrsPrice.Fields.Append "ҽ��ID", adBigInt
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
    
    mrsPrice.CursorLocation = adUseClient
    mrsPrice.LockType = adLockOptimistic
    mrsPrice.CursorType = adOpenStatic
    mrsPrice.Open
    
    On Error GoTo errH
            
    '��ȡҪ���������õ�ҽ����¼(������������,��鲿λ������������)
    '������鲿λ������������������ϵķ���
    strҽ��ID = _
        " Select ID From ����ҽ����¼ Where ID= [1] Union All" & _
        " Select ID From ����ҽ����¼ Where ���ID= [1] And ������� IN('F','D')"
    strSQL = _
        " Select B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID," & _
        " Nvl(A.��������,Sum(Nvl(C.��������,0))) as ����" & _
        " From ����ҽ������ A,����ҽ����¼ B,����ҽ��ִ�� C" & _
        " Where Nvl(A.�Ʒ�״̬,0)=0 And B.ID IN(" & strҽ��ID & ")" & _
        " And A.ҽ��ID=B.ID And A.���ͺ�+0= [2] " & _
        " And C.ҽ��ID(+)=A.ҽ��ID And C.���ͺ�(+)=A.���ͺ�" & _
        " Group by B.���,A.ҽ��ID,B.���ID,B.�������,B.������ĿID,A.��������" & _
        " Order by ���"
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ������", "H����ҽ������")
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "����ҽ��ִ��", "H����ҽ��ִ��")
    End If
    Set rsAdvice = OpenSQLRecord(strSQL, Me.Caption, lngҽ��ID, lng���ͺ�)
    
    For i = 1 To rsAdvice.RecordCount
        dbl���� = Nvl(rsAdvice!����, 0)
        
        '��ȡ��Ӧ���շѼ�Ŀ:ֻ��ȡ�̶�����,�Ҳ��Ǳ�۵Ķ���
        bln�������� = (rsAdvice!������� = "F" And Not IsNull(rsAdvice!���ID))
        strSQL = IIf(bln��������, "Nvl(B.�����շ���,100)/100", "1") & " as ������"
        strSQL = _
            " Select A.�շ���ĿID,A.�շ�����,B.������ĿID,D.�վݷ�Ŀ,C.���," & _
            " C.���㵥λ,C.ִ�п���,Decode(C.�Ƿ���,1,NULL,B.�ּ�) as ����," & strSQL & _
            " From �����շѹ�ϵ A,�շѼ�Ŀ B,�շ���ĿĿ¼ C,������Ŀ D" & _
            " Where A.������ĿID= [1] " & _
            " And A.�շ���ĿID=B.�շ�ϸĿID And A.�շ���ĿID=C.ID And B.������ĿID=D.ID" & _
            " And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)" & _
            " And ((Sysdate Between B.ִ������ and B.��ֹ����) or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL))" & _
            " And Nvl(A.���ж���,0)=1 And Nvl(C.�Ƿ���,0)=0"
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, CLng(rsAdvice!������ĿID))
        
        For j = 1 To rsTmp.RecordCount
            mrsPrice.AddNew
            mrsPrice!ҽ��ID = rsAdvice!ҽ��ID
            mrsPrice!��� = rsTmp!���
            mrsPrice!�շ�ϸĿID = rsTmp!�շ���ĿID
            mrsPrice!���㵥λ = Nvl(rsTmp!���㵥λ)
            mrsPrice!�������� = IIf(bln��������, 1, 0)
            mrsPrice!ִ�п��� = Nvl(rsTmp!ִ�п���, 0)
            mrsPrice!������ĿID = rsTmp!������ĿID
            mrsPrice!�վݷ�Ŀ = rsTmp!�վݷ�Ŀ
            mrsPrice!���� = Format(Nvl(rsTmp!����, 0), "0.00000")
            mrsPrice!���� = Format(Nvl(rsTmp!�շ�����, 0) * dbl����, "0.00000")
            mrsPrice!Ӧ�� = Format(mrsPrice!���� * mrsPrice!���� * rsTmp!������, gstrDec)
            If str�ѱ� = "" Then
                mrsPrice!ʵ�� = mrsPrice!Ӧ��
            Else
                mrsPrice!ʵ�� = Format(ActualMoney(str�ѱ�, mrsPrice!������ĿID, mrsPrice!Ӧ��), gstrDec)
            End If
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
    Set mrsPrice = Nothing
End Function

Private Function LoadBillDetail(ByVal lngRow As Long) As Boolean
'���ܣ���ʾ������ϸ����
'������lngRow=�����嵥��
    Dim rsTmp As New ADODB.Recordset
    Dim strNO As String, int��¼���� As Integer
    Dim lng���˿���ID As Long, int��Դ As Integer
    Dim lng��ĿID As Long, int���� As Integer
    Dim lngҽ��ID As Long, str�Ǽ�ʱ�� As String
    Dim strSQL As String, strIDs As String, strIDsSql As String, i As Long
    
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
        If strNO = "" Then Exit Function
    
        '�Ǽ�ʱ����Ϊ������ͬʱ���������δ��˵����
        str�Ǽ�ʱ�� = .Cell(flexcpData, lngRow, 8)
    End With
    
    'ҽ��ID������������,�Ը��ӷ��ö�����ͬ�����
    lngҽ��ID = AdviceID
    
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
            .MoveFirst
            For i = 1 To .RecordCount
                If lng��ĿID <> !�շ�ϸĿID Then int���� = i
                strSQL = strSQL & IIf(strSQL <> "", " Union ALL ", "") & _
                    " Select " & i & " as ���," & IIf(int���� = i, "-NULL", int����) & " as �۸񸸺�," & _
                    "'" & strNO & "' as NO," & int��¼���� & " as ��¼����,1 as ��¼״̬," & _
                    !ҽ��ID & " as ҽ�����,'" & !��� & "' as �շ����," & !�շ�ϸĿID & " as �շ�ϸĿID," & _
                    Get�շ�ִ�п���ID(!���, !�շ�ϸĿID, !ִ�п���, lng���˿���ID, int��Դ) & " as ִ�в���ID," & _
                    !������ĿID & " as ������ĿID,1 as ����," & !���� & " as ����," & !���� & " as ��׼����," & _
                    !Ӧ�� & " as Ӧ�ս��," & !ʵ�� & " as ʵ�ս��,To_Date('" & str�Ǽ�ʱ�� & "','YYYY-MM-DD HH24:MI:SS') as �Ǽ�ʱ�� From Dual"
                    
                lng��ĿID = !�շ�ϸĿID
                strIDs = strIDs & "," & !ҽ��ID
                .MoveNext
            Next
            If strSQL = "" Then Exit Function
            strSQL = "(" & strSQL & ")"
            strIDs = Mid(strIDs, 2)  'ȡ����������漰��ҽ��ID
        End With
    Else
        strSQL = "���˷��ü�¼"
        '������鲿λ������������������ϵķ���
        strIDsSql = _
            " Select ID From ����ҽ����¼ Where ID=" & lngҽ��ID & " Union All" & _
            " Select ID From ����ҽ����¼ Where ���ID=" & lngҽ��ID & " And ������� IN('F','D') "
    End If
    
    strSQL = "Select C.���� as ���,Nvl(F.����,B.����)||Decode(B.���,NULL,NULL,' '||B.���) as ��Ŀ," & _
        " Sum(A.��׼����" & IIf(blnҩ����λ, "*Nvl(E." & strҩ����װ & ",1)", "") & ") as ����," & _
        " Avg(Nvl(A.����,1)*A.����" & IIf(blnҩ����λ, "/Nvl(E." & strҩ����װ & ",1)", "") & ") as ����," & _
        IIf(blnҩ����λ, "Decode(E.ҩƷID,NULL,B.���㵥λ,E." & strҩ����λ & ")", "B.���㵥λ") & " as ���㵥λ," & _
        " Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��,D.���� as ִ�в���" & _
        " From " & strSQL & " A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,ҩƷ��� E,�շ���Ŀ���� F" & _
        " Where A.��¼����= [1] And A.��¼״̬ IN(0,1)" & _
        " And A.�շ�ϸĿID=B.ID And A.�շ����=C.���� And A.ִ�в���ID=D.ID(+)" & _
        " And B.ID=E.ҩƷID(+) And A.�շ�ϸĿID=F.�շ�ϸĿID(+)" & _
        " And F.����(+)=1 And F.����(+)=" & IIf(gbln��Ʒ��, 3, 1) & _
        " And A.NO= [3] And " & _
        IIf(strIDsSql <> "", " a.ҽ����� in (" & strIDsSql & ")", " instr([4],','||A.ҽ�����||',')> 0 ") & _
        " And A.�Ǽ�ʱ��= [2] " & _
        " Group by Nvl(A.�۸񸸺�,A.���),C.����,Nvl(F.����,B.����),B.���,B.���㵥λ,D.����,E.ҩƷID,E." & strҩ����λ & _
        " Order by Nvl(A.�۸񸸺�,A.���)"
        
    If mblnMoved Then
        strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
        strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, int��¼����, CDate(Format(str�Ǽ�ʱ��, "yyyy-MM-dd hh:mm:ss")), strNO, "," & strIDs & ",")
    
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
    
    LoadBillDetail = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowMenu()
    frmParent.mnuMoneyAdd(0).Visible = iPatientType = 1
End Sub

Private Sub vsMoney_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If Button = 2 And frmParent.mnuMoney.Visible And frmParent.mnuMoney.Enabled Then PopupMenu frmParent.mnuMoney, 2
End Sub
