VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmMulitChargeSelect 
   Caption         =   "�����շѵ���ѡ��"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11985
   Icon            =   "frmMulitChargeSelect.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11985
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox picDown 
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   570
      ScaleHeight     =   765
      ScaleWidth      =   12465
      TabIndex        =   1
      Top             =   7185
      Width           =   12465
      Begin VB.PictureBox picNoInfo 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   30
         ScaleHeight     =   465
         ScaleWidth      =   9225
         TabIndex        =   4
         Top             =   210
         Width           =   9225
         Begin VB.TextBox txtInvoiceNo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   15
            Width           =   3240
         End
         Begin VB.TextBox txtCurTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   990
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Width           =   1275
         End
         Begin VB.TextBox txtAllTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   3390
            Locked          =   -1  'True
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label lblInvoice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ʊ��Ϣ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   4875
            TabIndex        =   10
            Top             =   75
            Width           =   960
         End
         Begin VB.Label lblCurTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ǰ����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   9
            Top             =   60
            Width           =   960
         End
         Begin VB.Label lblAllTotal 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2385
            TabIndex        =   8
            Top             =   60
            Width           =   960
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9255
         TabIndex        =   3
         ToolTipText     =   "�ȼ�F2,�Ҽ���������Ϊ���۵�(��CTRL+S)"
         Top             =   195
         Width           =   1440
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ȡ��(&C)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   10770
         TabIndex        =   2
         ToolTipText     =   "�ȼ�:Esc"
         Top             =   195
         Width           =   1440
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Height          =   5835
      Left            =   -525
      TabIndex        =   0
      Top             =   900
      Width           =   9510
      _cx             =   16775
      _cy             =   10292
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   10.5
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
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMulitChargeSelect.frx":0442
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmMulitChargeSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrNOs As String
Private mlngModule As Long
Private mblnUnLoad As Boolean
Private mblnNOMoved As Boolean
Private mblnOk As Boolean
Private mstrNo As String
Private mstrShowInVoiceNo As String
Private mblnOldDelSelect As Boolean

Private Function LoadData(ByVal strNos As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-04-12 16:40:13
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strҽ����� As String, i As Long, strTemp As String
    Dim cur�ϼ� As Double, strReturnNos As String
    Dim strNoTemp As String
    
    On Error GoTo errHandle
    Screen.MousePointer = 11
    strSQL = "" & _
    " Select A.����ID,A.NO,A.��¼״̬,Nvl(A.�۸񸸺�,A.���)  as ���,A.�շ����,A.ִ�в���ID,A.��������ID, A.�շ�ϸĿID," & _
    "           A.��������,A.���㵥λ,A.ҽ�����,A.�ѱ�,A.��������," & _
    "          Avg(Nvl(A.����,1)) as ����,Avg(Nvl(A.����,0)) as ����,Sum(A.��׼����) as ����,sum(A.Ӧ�ս��) as Ӧ�ս��,sum(A.ʵ�ս��) as ʵ�ս��," & _
    "          max(Decode(A.��¼״̬,2,NULL,A.����Ա����)) as  ����Ա����,Max(decode(A.��¼״̬,2,NULL,A.�Ǽ�ʱ��)) as �Ǽ�ʱ��,max(decode(A.��¼״̬,2,NULL,A.ժҪ)) as ժҪ" & _
    " From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A,Table(f_Str2list([1])) J" & _
    " Where A.��¼����=1  And A.NO=J.Column_Value " & _
    "  Group by A.NO,A.��¼״̬,A.����ID,Nvl(A.�۸񸸺�,A.���),A.�շ���� ,A.ִ�в���ID,A.��������ID, A.�շ�ϸĿID,A.��������,A.���㵥λ,A.ҽ�����,A.�ѱ�,A.��������"
    
    strSQL = _
    " Select  A.NO,A.���,A.��������,A.�ѱ�," & _
    "        A.�շ�ϸĿID,C.���� as �����,C.���� as �����,B.����,B.����,B.���,Nvl(A.��������,B.��������) ��������," & _
        IIf(gblnҩ����λ, "Decode(X.ҩƷID,NULL,A.���㵥λ,X." & gstrҩ����λ & ")", "A.���㵥λ") & " as ���㵥λ," & _
    "       Sum(Nvl(A.����,1)*A.����" & IIf(gblnҩ����λ, "/Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ʣ������," & _
    "       Max(A.����" & IIf(gblnҩ����λ, "*Nvl(X." & gstrҩ����װ & ",1)", "") & ") as ����," & _
    "       Sum(A.Ӧ�ս��) as Ӧ�ս��,Sum(A.ʵ�ս��) as ʵ�ս��," & _
    "       D.���� as ִ�п���,E.���� as ��������,Max(A.����Ա����) as ����Ա����,Max(A.�Ǽ�ʱ��) as �Ǽ�ʱ��, " & _
    "       Max(A.ժҪ) as ժҪ,Max(A.ҽ�����) as ҽ�����" & _
    " From (" & strSQL & ") A,�շ���ĿĿ¼ B,�շ���Ŀ��� C,���ű� D,���ű� E,ҩƷ��� X" & _
    " Where A.�շ�ϸĿID=B.ID And C.����=A.�շ���� And A.�շ�ϸĿID=X.ҩƷID(+)" & _
    "       And A.ִ�в���ID=D.ID(+) And A.��������ID=E.ID(+)  " & _
    " Group by  A.NO,A.��� ,A.��������,A.�ѱ�,A.�շ�ϸĿID,C.����,C.����,B.����,B.����," & _
    "       B.���,Nvl(A.��������,B.��������),A.���㵥λ,D.����,E.����,X.ҩƷID,X." & gstrҩ����λ
    
    strSQL = "" & _
    "   Select /*+ rule */ " & _
    "        A.NO,A.���,A.��������,A.�ѱ�,A.�����,A.�����,A.����,Nvl(B.����,A.����) as ����,E1.���� as ��Ʒ��,A.���,A.��������," & _
    "       A.���㵥λ,A.ҽ����� ,A.�շ�ϸĿID,A.ʣ������,A.����,A.Ӧ�ս��,A.ʵ�ս��," & _
    "       A.ִ�п���,A.��������,A.����Ա����,A.�Ǽ�ʱ��, A.ժҪ , M.ҽ������  as ҽ������ " & _
    "   From (" & strSQL & ") A,�շ���Ŀ���� B,�շ���Ŀ���� E1,����ҽ����¼ M" & _
    "   Where       nvl(A.ʣ������,0)<>0 And A.�շ�ϸĿID=B.�շ�ϸĿID(+) And B.����(+)=1 And B.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
    "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3 " & _
    "       And A.ҽ�����=M.ID(+)  " & _
    " Order by A.NO,A.���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Replace(strNos, "'", ""))

    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        strNoTemp = ""
        For i = 1 To rsTemp.RecordCount
            If InStr(strReturnNos & ",", "," & Nvl(rsTemp!NO) & ",") = 0 Then
                strReturnNos = strReturnNos & "," & Nvl(rsTemp!NO)
            End If
            .Cell(flexcpData, i, .ColIndex("��Ŀ")) = Nvl(rsTemp!��������)
            .Cell(flexcpData, i, .ColIndex("����ID")) = Nvl(rsTemp!ҽ�����) & "," & Nvl(rsTemp!�շ�ϸĿID)
            If Val(Nvl(rsTemp!ҽ�����)) <> 0 And InStr(strҽ����� & ",", "," & Nvl(rsTemp!ҽ�����) & ",") = 0 Then
                strҽ����� = strҽ����� & "," & Nvl(rsTemp!ҽ�����)
            End If
            strTemp = ""
            If Val(Nvl(rsTemp!��������)) <> 0 Then
                rsTemp.MoveNext
                strTemp = "��"
                If rsTemp.EOF Then
                    strTemp = "��"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("��Ŀ"))) <> Nvl(rsTemp!��������) Then
                    strTemp = "��"
                End If
                rsTemp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If
    
            .RowData(i) = CLng(rsTemp!���)
            .TextMatrix(i, .ColIndex("���ݺ�")) = rsTemp!NO
            .TextMatrix(i, .ColIndex("���")) = rsTemp!�����
            .TextMatrix(i, .ColIndex("��Ŀ")) = strTemp & rsTemp!���� & IIf(IsNull(rsTemp!���), "", " " & rsTemp!���)
            .TextMatrix(i, .ColIndex("��Ʒ��")) = strTemp & Nvl(rsTemp!��Ʒ��)
            .TextMatrix(i, .ColIndex("����")) = FormatEx(Val(Nvl(rsTemp!ʣ������)), 5)
            .TextMatrix(i, .ColIndex("��λ")) = Nvl(rsTemp!���㵥λ)
            .TextMatrix(i, .ColIndex("����")) = Format(rsTemp!����, gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("Ӧ�ս��")) = Format(rsTemp!Ӧ�ս��, gstrDec)
            .TextMatrix(i, .ColIndex("ʵ�ս��")) = Format(rsTemp!ʵ�ս��, gstrDec)
            .TextMatrix(i, .ColIndex("��������")) = Nvl(rsTemp!��������)
            .TextMatrix(i, .ColIndex("ִ�п���")) = Nvl(rsTemp!ִ�п���)
            .TextMatrix(i, .ColIndex("����Ա")) = Nvl(rsTemp!����Ա����)
            .TextMatrix(i, .ColIndex("ʱ��")) = Format(rsTemp!�Ǽ�ʱ��, "MM-dd HH:mm")
            .TextMatrix(i, .ColIndex("����ID")) = 0
            .TextMatrix(i, .ColIndex("ҽ��")) = Nvl(rsTemp!ҽ������)
            .TextMatrix(i, .ColIndex("ԭʼ����")) = 0
            .TextMatrix(i, .ColIndex("׼������")) = 0
            .TextMatrix(i, .ColIndex("ҽ�����")) = Nvl(rsTemp!ҽ�����)
            If InStr(strNoTemp & ",", "," & rsTemp!NO & ",") = 0 Then
                '�����ָ���
                If strNoTemp <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                strNoTemp = strNoTemp & "," & rsTemp!NO
            End If
            cur�ϼ� = cur�ϼ� + rsTemp!ʵ�ս��
            rsTemp.MoveNext
        Next
        If .Rows <= 1 Then .Rows = 2
        .Row = .FixedRows: .Col = .ColIndex("��Ŀ")
        Call vsBill_AfterRowColChange(-1, -1, .Row, .Col)
        .SelectionMode = flexSelectionByRow
        .Redraw = flexRDBuffered
    End With
    txtAllTotal.Text = Format(cur�ϼ�, gstrDec)
    Screen.MousePointer = 0
    If strReturnNos <> "" Then strReturnNos = Mid(strReturnNos, 2)
    'û�п�ѡ���ݻ���ֻ��һ�ŵ���ʱ���˳�
    If strReturnNos = "" Or InStr(strReturnNos, ",") = 0 Then mstrNo = strReturnNos: mblnOk = True: Exit Function
    LoadData = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
End Function
Private Sub InitBillHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ͷ����Ϣ
    '����: �ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2012-09-11 09:47:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    Dim varTemp As Variant, intCol As Integer
    
    strHead = "���ݺ�,1000,1;���,720,1;��Ŀ,2800,1;��Ʒ��,2000,1;����,750,7;��λ,550,1;����,1100,7;" & _
        "Ӧ�ս��,1100,7;ʵ�ս��,1100,7;��������,1000,1;ִ�п���,1000,1;����Ա,850,1;ʱ��,1260,1;����ID,0,0;ҽ��,1560,1;" & _
        "ԭʼ����,0,4;׼������,0,4;ҽ�����,0,4"
    
    arrHead = Split(strHead, ";")
    With vsBill
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .COLS = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            varTemp = Split(arrHead(i) & ",,,", ",")
            intCol = .FixedCols + i
            .ColKey(intCol) = varTemp(0)
            .TextMatrix(.FixedRows - 1, intCol) = varTemp(0)
            If UBound(varTemp) > 0 Then
                .ColHidden(intCol) = False
                .ColWidth(intCol) = Val(varTemp(1))
                If .ColWidth(intCol) = 0 Then .ColHidden(intCol) = True
                .ColAlignment(intCol) = Val(varTemp(2))
            Else
                .ColHidden(intCol) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .COLS - 1) = 4
        .ColHidden(.ColIndex("��Ʒ��")) = gTy_System_Para.bytҩƷ������ʾ <> 2
        .FrozenCols = 2
    End With
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    mstrNo = ""
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Dim strNo As String
    With vsBill
        If .Row < 1 Then Exit Sub
        strNo = Trim(.TextMatrix(.Row, .ColIndex("���ݺ�")))
        If strNo = "" Then Exit Sub
    End With
    mstrNo = strNo
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Activate()
    picNoInfo.Visible = Not mblnOldDelSelect
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOK_Click
End Sub

Private Sub Form_Load()
    mblnUnLoad = False
    txtInvoiceNo.Text = mstrShowInVoiceNo
    Call InitBillHead
    Call RestoreWinState(Me, App.ProductName)
    If LoadData(mstrNOs) = False Then mblnUnLoad = True: Unload Me
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With picDown
        .Left = ScaleLeft
        .Top = ScaleHeight - .Height
        .Width = ScaleWidth
        
    End With
    With vsBill
        .Left = ScaleLeft + 50
        .Top = ScaleTop
        .Height = picDown.Top - .Top
        .Width = ScaleWidth - .Left * 2
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
Private Sub picDown_Resize()
    Err = 0: On Error Resume Next
    With picDown
        cmdCancel.Left = .ScaleWidth - cmdCancel.Width - 100
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 50
    End With
End Sub
Private Sub vsBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim cur�ϼ� As Currency, i As Long
    If NewRow <> OldRow Then
        With vsBill
            If .TextMatrix(NewRow, .ColIndex("���ݺ�")) <> "" Then
                For i = NewRow - 1 To .FixedRows Step -1
                    If .TextMatrix(i, .ColIndex("���ݺ�")) <> .TextMatrix(NewRow, .ColIndex("���ݺ�")) Then Exit For
                    cur�ϼ� = cur�ϼ� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                Next
                For i = NewRow To .Rows - 1
                    If .TextMatrix(i, .ColIndex("���ݺ�")) <> .TextMatrix(NewRow, .ColIndex("���ݺ�")) Then Exit For
                    cur�ϼ� = cur�ϼ� + Val(.TextMatrix(i, .ColIndex("ʵ�ս��")))
                Next
            End If
            txtCurTotal.Text = Format(cur�ϼ�, gstrDec)
        End With
    End If
End Sub
 
Public Function zlShowSelect(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strNos As String, _
    strShowInVoiceNo As String, ByRef strNo As String, _
    Optional blnOldDelSelect As Boolean) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��Դ��뵥��,ѡ������һ�ŵ���
    '���:frmMain-���õ�������
    '       strNos-���ݺ�,�ö��ŷ���:A0001,A0002
    '       strShowInVoiceNo-��ʾ�ķ�Ʊ��
    '����:strNO-����ѡ�еĵ���
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-04-12 17:41:39
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule
    mstrNOs = strNos: mblnOldDelSelect = blnOldDelSelect
    mstrShowInVoiceNo = strShowInVoiceNo
    Screen.MousePointer = 0: mblnOk = False
    Err = 0: On Error Resume Next
    Me.Show 1, frmMain
    Screen.MousePointer = 11
    strNo = mstrNo
    zlShowSelect = mblnOk
End Function

Private Sub vsBill_DblClick()
    Call cmdOK_Click
End Sub
