VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPatiSelect 
   Caption         =   "����ѡ��"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   10545
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   10545
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10545
      TabIndex        =   3
      Top             =   4260
      Width           =   10545
      Begin VB.ComboBox cboPatient 
         Height          =   300
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   112
         Width           =   2055
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   112
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   7635
         TabIndex        =   8
         Top             =   87
         Width           =   1150
      End
      Begin VB.CommandButton cmdCanc 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   9015
         TabIndex        =   7
         Top             =   87
         Width           =   1150
      End
      Begin VB.CheckBox chkState 
         Caption         =   "��Ժ(&g)"
         Height          =   195
         Index           =   2
         Left            =   3960
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   165
         Width           =   1035
      End
      Begin VB.CheckBox chkState 
         Caption         =   "Ԥ��Ժ(&P)"
         Height          =   195
         Index           =   1
         Left            =   2760
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   165
         Width           =   1155
      End
      Begin VB.CheckBox chkState 
         Caption         =   "��Ժ(&I)"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   165
         Width           =   1035
      End
   End
   Begin VB.PictureBox picVsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   2565
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4215
      ScaleWidth      =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   45
   End
   Begin MSComctlLib.ImageList imgSort 
      Left            =   7560
      Top             =   4275
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   9
      ImageHeight     =   8
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPatiSelect.frx":04DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDept 
      Height          =   4230
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   2580
      _cx             =   4551
      _cy             =   7461
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
      BackColorSel    =   12640511
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
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
      FormatString    =   $"frmPatiSelect.frx":09B4
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
      ExplorerBar     =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsPati 
      Height          =   4230
      Left            =   2625
      TabIndex        =   1
      Top             =   0
      Width           =   7905
      _cx             =   13944
      _cy             =   7461
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
      BackColorSel    =   12640511
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPatiSelect.frx":09FC
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
      ExplorerBar     =   1
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
Attribute VB_Name = "frmPatiSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mfrmParent As Form
Public mlngUnitID As Long   '��ͨ���ʺͿ��ҷ�ɢ����ʱ,���벡��ID,���������ID
Public mbytUseType As Byte  '0:��ͨ����,1-���ҷ�ɢ����,2-ҽ�����Ҽ���,3-����,4-��������
Public mstrPrivs As String

Private mstrPrivsOpt As String '���ʲ���1150ģ�����Ȩ����
Private mblnNotDo As Boolean



Private Sub cboPatient_Click()
    If Visible Then
        Call vsDept_AfterRowColChange(-1, -1, vsDept.Row, vsDept.Col)
    End If
End Sub

Private Sub cboType_Click()
    If Visible Then
        Call InitUnits(-1)
        Call vsDept_AfterRowColChange(-1, -1, vsDept.Row, vsDept.Col)
    End If
End Sub

Private Sub chkSettle_Click()
    If Visible Then
        Call vsDept_AfterRowColChange(-1, -1, vsDept.Row, vsDept.Col)
    End If
End Sub

Private Sub chkState_Click(Index As Integer)
    If Visible Then
        '����Ҫѡ��һ��,����.value��ִ�б�����
        If chkState(0).Value = 0 And chkState(1).Value = 0 And chkState(2).Value = 0 Then
            chkState(Index).Value = 1
        Else
            Call InitUnits(vsDept.RowData(vsDept.Row))
            Call vsDept_AfterRowColChange(-1, -1, vsDept.Row, vsDept.Col)
        End If
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If vsPati.Rows > 1 And vsPati.TextMatrix(1, 0) <> "" Then
        mfrmParent.txtPatient.Text = "-" & vsPati.TextMatrix(vsPati.Row, 0)
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.ScaleHeight < 2000 Or Me.ScaleWidth < 3000 Then Exit Sub
        
    vsDept.Height = Me.ScaleHeight - picBottom.Height - 100
    vsPati.Height = vsDept.Height
    picVsc.Height = vsPati.Height
    picVsc.Left = Me.ScaleLeft + vsDept.Width
    
    vsPati.Width = Me.ScaleWidth - picVsc.Left - picVsc.Width
    
    cmdCanc.Left = Me.ScaleLeft + Me.ScaleWidth - 200 - cmdCanc.Width
    cmdOK.Left = cmdCanc.Left - cmdOK.Width - 200
End Sub


Private Sub picVsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsDept.Width + X < 200 Or vsPati.Width - X < 200 Then Exit Sub
        picVsc.Left = picVsc.Left + X
        vsDept.Width = vsDept.Width + X
        vsPati.Left = vsPati.Left + X
        vsPati.Width = vsPati.Width - X
        Me.Refresh
    End If
End Sub

Private Function GetSelectCount() As Integer
    Dim i As Integer, j As Integer
    
    For i = 0 To chkState.UBound
        If chkState(i).Value = 1 Then j = j + 1
    Next
    
    GetSelectCount = j
End Function

Private Sub vsDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, j As Long, lngDepID As Long, intCNT As Integer
    Dim strSql As String, strCond As String, strUnIndex As String
    Dim blnByUnit As Boolean, strRange As String, lng�������� As Long
    Dim str��Ժ As String
    If mblnNotDo Then Exit Sub '�ֹ�������ʱ������
    
    If NewRow = OldRow Then Exit Sub
    If Not (Visible Or OldRow = -1) Then Exit Sub
    vsPati.Rows = vsPati.FixedRows
    vsPati.Rows = vsPati.FixedRows + 1
    If vsDept.RowData(vsDept.Row) = 0 Then Exit Sub
    
    lngDepID = vsDept.RowData(vsDept.Row)   '��ͨ���ʺͿ��ҷ�ɢ�����ǲ���ID
    blnByUnit = cboType.ListIndex = 0
    
    If mbytUseType = 0 Or mbytUseType = 1 Or mbytUseType = 2 Or mbytUseType = 4 Then
        If chkState(1).Value = 1 Or chkState(2).Value = 1 Then
            If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
                strCond = ""
            ElseIf InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") > 0 Then
                strCond = " And Exists(Select 1 From ������� X Where A.����ID=X.����ID And X.����=1 And X.����=2 and Nvl(X.�������,0)<>0)"
            ElseIf InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") > 0 Then
                strCond = " And Not Exists(Select 1 From ������� X Where A.����ID=X.����ID And X.����=1 And X.����=2 and Nvl(X.�������,0)<>0)"
            Else
                '���������û��,Ԥ��Ժ�ͳ�Ժѡ���ǽ��õ�
            End If
        End If
    End If
    
    '�ò�����ҳ�ĳ�Ժ����ID��������,����Ϊ�������������۲���(���ۿ��һ���û�д�λ),���Բ��ܴӴ�λ״����¼��ȥ�Ҳ�
    intCNT = GetSelectCount
    str��Ժ = " And  Exists(Select 1 From ��Ժ���� ZY Where ZY.����ID=B.����ID)"
    
    If intCNT = 1 Then
        If chkState(0).Value = 1 Then       '1.��Ժ
            strSql = " And B.��Ժ���� is NULL And B.״̬<>3  " & str��Ժ
             
        ElseIf chkState(1).Value = 1 Then   '2.Ԥ��Ժ
            strSql = " And B.��Ժ���� is NULL And B.״̬=3��" & str��Ժ & strCond
             
        Else                                '3.��Ժ
            strSql = " And B.��Ժ����>Trunc(Sysdate-" & gintOutDay & ")" & strCond
            strUnIndex = "+0"
        End If
    ElseIf intCNT = 2 Then
        If chkState(2).Value = 0 Then       '1.��Ժ��Ԥ��Ժ
            If strCond <> "" Then strCond = " And (B.״̬<>3 Or B.״̬=3" & strCond & ") ��  " & str��Ժ
            strSql = " And B.��Ժ���� is NULL" & strCond
        ElseIf chkState(1).Value = 0 Then   '2.��Ժ�ͳ�Ժ
            strSql = " And (B.��Ժ���� is NULL And B.״̬<>3   " & str��Ժ & " Or B.��Ժ����>Trunc(Sysdate-" & gintOutDay & ")" & strCond & ")"
        Else                                '3.Ԥ��Ժ�ͳ�Ժ
            strSql = " And (B.��Ժ���� is NULL And B.״̬=3 " & str��Ժ & " Or B.��Ժ����>Trunc(Sysdate-" & gintOutDay & "))" & strCond
        End If
    ElseIf intCNT = 3 Then
        If strCond <> "" Then
            strCond = " And (B.��Ժ���� is NULL And B.״̬<>3  " & str��Ժ & " Or (B.��Ժ���� is NULL And B.״̬=3   " & str��Ժ & " Or B.��Ժ����>Trunc(Sysdate-" & gintOutDay & "))" & strCond & ")"
        Else
            strCond = " And (B.��Ժ���� is NULL   " & str��Ժ & " Or B.��Ժ����>Trunc(Sysdate-" & gintOutDay & "))"
        End If
        strSql = strCond
    End If
    
    If mbytUseType = 3 Then '����
        If cboPatient.ListIndex > 0 Then '0-���ѽ��岡��
            Select Case cboPatient.ListIndex
                Case 1  '�κη���δ���岡��
                    strRange = ""
                Case 2  '���δ����Ĳ���
                    strRange = " And C.��Դ;�� = 4"
                Case 3  'סԺδ����Ĳ���
                    strRange = " And C.��Դ;�� = 2"
                Case 4  '����δ����Ĳ���
                    strRange = " And C.��Դ;�� = 1"
            End Select
            strSql = strSql & " And Exists(Select 1 From ����δ����� C Where C.����id=B.����ID And C.��ҳID=B.��ҳID" & strRange & ")"
        End If
        
        strSql = "" & _
        "Select A.����ID,A.סԺ��,nvl(B.����,A.����) as ����,A.��ǰ���� as ��λ,nvl(B.�Ա�,A.�Ա�) as �Ա�,A.����,B.�ѱ�," & _
        "       Decode(B.��Ժ����,NULL,'��','') as ��Ժ,B.��Ժ����,B.��������" & _
        " From ������Ϣ A,������ҳ B" & _
        " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID" & _
        "          And " & IIf(blnByUnit, "B.��ǰ����ID", "B.��Ժ����ID") & strUnIndex & " =[1]" & _
        "          And A.��ҳID=B.��ҳID And Nvl(B.��ҳID,0)<>0" & strSql & _
                IIf(InStr(1, mstrPrivs, ";��ͨ���˽���;") > 0, "", " And nvl(B.����,0)<>0") & _
        " Order by A.סԺ�� Desc"
    Else    '����
        If mbytUseType = 0 Or mbytUseType = 1 Or mbytUseType = 2 Then
            '���۲��˼���Ȩ��
            If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) And (InStr(mstrPrivsOpt, "סԺ���ۼ���") > 0 And gblnסԺ����) Then
                strSql = strSql & " And Nvl(B.��������,0) IN(0,1,2)"
            ElseIf InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln�������� Then
                strSql = strSql & " And Nvl(B.��������,0) IN(0,1)"
            ElseIf InStr(mstrPrivsOpt, ";סԺ���ۼ���;") > 0 And gblnסԺ���� Then
                strSql = strSql & " And Nvl(B.��������,0) IN(0,2)"
            Else
                strSql = strSql & " And Nvl(B.��������,0)=0"
            End If
        End If
        strSql = "" & _
        " Select A.����ID,A.סԺ��,nvl(B.����,A.����) as ����,B.��Ժ���� as ��λ," & _
        "       nvl(B.�Ա�,A.�Ա�) as �Ա�,A.����,B.�ѱ�," & _
        "       Decode(B.��Ժ����,NULL,'��','') as ��Ժ,B.��Ժ����,B.��������" & _
        " From ������Ϣ A,������ҳ B" & _
        " Where A.ͣ��ʱ�� is NULL And A.����ID=B.����ID And A.��ҳID=B.��ҳID  And Nvl(B.��ҳID,0)<>0 and B.��Ŀ���� is NULL " & strSql & _
        "   And " & IIf(blnByUnit, "B.��ǰ����ID", "B.��Ժ����ID") & strUnIndex & " =[1]" & _
        " Order by A.סԺ�� Desc"
    End If
    
    Screen.MousePointer = 11
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngDepID)
    
    Set vsPati.DataSource = rsTmp
    vsPati.ToolTipText = "���ҵ�:" & rsTmp.RecordCount & "λ����!"
    If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsPati.Rows = vsPati.FixedRows Then
        vsPati.Rows = vsPati.FixedRows + 1
    End If
    vsPati.AutoSize 0, vsPati.Cols - 1
    vsPati.RowHeight(0) = 320
    
    vsPati.Cell(flexcpAlignment, 0, 0, 0, vsPati.Cols - 1) = 4
    vsPati.ColAlignment(0) = 1
    vsPati.ColAlignment(1) = 1
    vsPati.ColAlignment(3) = 1
    vsPati.ColAlignment(4) = 4
    vsPati.ColAlignment(6) = 4
    If mbytUseType = 3 Then
        If gintOutDay > 0 Then vsPati.ColAlignment(7) = 1
    End If
    
    
    lng�������� = VsfGetColNum(vsPati, "��������")
    For i = 1 To vsPati.Rows - 1
        vsPati.Cell(flexcpForeColor, i, 0, i, vsPati.Cols - 1) = zlDatabase.GetPatiColor(vsPati.TextMatrix(i, lng��������))
    Next
    
    Call RestoreColSort(vsPati)
    vsPati.Row = 1
    Screen.MousePointer = 0
        
    Call vsPati_AfterRowColChange(-1, -1, vsPati.Row, vsPati.Col)
    
    If Not rsTmp.EOF Then
        If Visible Then vsPati.SetFocus
    Else
        If Visible Then vsDept.SetFocus
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsDept_AfterSort(ByVal Col As Long, Order As Integer)
    With vsDept
        .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        
        If Order Mod 2 = 1 Then
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(1).Picture
        Else
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(2).Picture
        End If
        
        Call vsDept_AfterRowColChange(-1, -1, .Row, .Col)
            
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\VSFlexGrid", .Name & "ColSort", Col & "," & Order
    End With
End Sub

Private Sub vsDept_GotFocus()
    vsDept.BackColorSel = &HC0E0FF
    vsPati.BackColorSel = &HC0C0C0
End Sub

Private Sub vsDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Call vsPati_KeyDown(KeyCode, Shift)
End Sub

Private Sub vsPati_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsPati.Cell(flexcpForeColor, NewRow, NewCol) = vbRed Then
        vsPati.ForeColorSel = vbRed
    Else
        vsPati.ForeColorSel = vsPati.Cell(flexcpForeColor, NewRow, NewCol)
    End If
End Sub

Private Sub vsPati_AfterSort(ByVal Col As Long, Order As Integer)
    With vsPati
        .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        
        If Order Mod 2 = 1 Then
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(1).Picture
        Else
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(2).Picture
        End If
        
        Call vsPati_AfterRowColChange(-1, -1, .Row, .Col)
            
        SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\VSFlexGrid", .Name & "ColSort", Col & "," & Order
    End With
End Sub

Private Sub vsPati_DblClick()
    cmdOK_Click
End Sub

Private Sub vsPati_GotFocus()
    vsDept.BackColorSel = &HC0C0C0
    vsPati.BackColorSel = &HC0E0FF
End Sub

Private Sub vsPati_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOK_Click
End Sub

Private Sub Form_Activate()
    vsPati.SetFocus
End Sub

Private Sub Form_Load()
    Dim i As Long, lngModul As Long
    
     Call RestoreWinState(Me, App.ProductName)
    
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p���ʲ���)
    mblnNotDo = False
            
    If mbytUseType = 3 Then
        lngModul = 1137
    Else
        lngModul = 1150
    End If
            
    i = Val(zlDatabase.GetPara("��ʾ��Ժ����", glngSys, lngModul, 0))
    chkState(0).Value = IIf(i > 0, 1, 0)
    i = Val(zlDatabase.GetPara("��ʾԤ��Ժ����", glngSys, lngModul, 0))
    chkState(1).Value = IIf(i > 0, 1, 0)
    i = Val(zlDatabase.GetPara("��ʾ��Ժ����", glngSys, lngModul, 0))
    chkState(2).Value = IIf(i > 0, 1, 0)
    
    If mbytUseType = 3 Then
        cboPatient.Clear
        cboPatient.AddItem "���������ѽ��岡��"
        cboPatient.AddItem "�κη���δ����Ĳ���"
        cboPatient.AddItem "������δ����Ĳ���"
        cboPatient.AddItem "סԺ����δ����Ĳ���"
        cboPatient.AddItem "�������δ����Ĳ���"
        i = Val(zlDatabase.GetPara("��ʾ���岡��", glngSys, lngModul, 0))
        If i < cboPatient.ListCount And i >= 0 Then
            cboPatient.ListIndex = i
        Else
            cboPatient.ListIndex = 0
        End If
    Else
        cboPatient.Visible = False
    End If
          
    If chkState(0).Value = 0 And chkState(1).Value = 0 And chkState(2).Value = 0 Then
        chkState(0).Value = 1
    End If
            
    '������ز����ĳ�Ժ����δ����,����ó�Ժ
    chkState(2).Enabled = gintOutDay > 0
    If Not chkState(2).Enabled Then chkState(2).Value = 0
    chkState(2).ToolTipText = "�뱾�ز������õ�����鿴n��ĳ�Ժ�����й�."
    
    If mbytUseType = 0 Or mbytUseType = 1 Or mbytUseType = 2 Or mbytUseType = 4 Then
        '����,  �Ƿ����ǿ�Ƽ���Ȩ�� ,Ԥ��Ժ�ͳ�Ժ������Ȩ��Ӱ��
        If InStr(mstrPrivsOpt, ";��Ժδ��ǿ�Ƽ���;") = 0 And InStr(mstrPrivsOpt, ";��Ժ����ǿ�Ƽ���;") = 0 Then
            chkState(1).Enabled = False: chkState(2).Enabled = False
            chkState(1).Value = 0: chkState(2).Value = 0
            chkState(0).ToolTipText = "û��[��Ժδ������ǿ�Ƽ���]Ȩ��,����ѡ��Ԥ��Ժ�ͳ�Ժ����"
        End If
    
        chkState(2).ToolTipText = "��Ȩ��[��Ժδ��(����)ǿ�Ƽ���]�й�."
    End If
    
    cboType.Clear
    cboType.AddItem "��������ʾ"
    cboType.AddItem "��������ʾ"
        
    If mbytUseType = 3 Then
        i = Val(zlDatabase.GetPara("��ʾ���˷�ʽ", glngSys, 1137, 0))
    Else
        i = Val(zlDatabase.GetPara("��ʾ���˷�ʽ", glngSys, 1150, 0))
    End If
    cboType.ListIndex = IIf(i = 0, 0, 1) '������Click�¼�
    
    Call InitUnits(mlngUnitID)
    
    Call vsDept_AfterRowColChange(-1, -1, vsDept.Row, vsDept.Col)
End Sub

Private Sub InitUnits(ByVal lngDeptID As Long)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long, lngUnitID As Long
    Dim str��Դ As String, strUnitIDs As String
    Dim blnByUnit As Boolean, blnLimitUnit As Boolean
    'ͨ��[��λ״����¼]�����ǲ�����ҳȡ���˲��������,���Ա���Դ���ȫ��ɨ��,
    '�����ĳ���һ���ȫ�Ǽ�ͥ��������,���󲻳���

    blnByUnit = cboType.ListIndex = 0
    blnLimitUnit = (mbytUseType = 0 Or mbytUseType = 1) And InStr(mstrPrivs, ";���в���;") = 0   'ҽ�����Ҽ��ʺͽ��ʲ����Ʋ��˲���
    If blnLimitUnit Then strUnitIDs = GetUserUnits

    If (InStr(mstrPrivsOpt, ";�������ۼ���;") > 0 And gbln��������) Or mbytUseType = 2 Or mbytUseType = 3 Or mbytUseType = 4 Then
        str��Դ = "1,2,3"
    Else
        str��Դ = "2,3"
    End If
    
    strSql = " And Exists (Select 'X' From ��λ״����¼ X Where " & _
            IIf(chkState(2).Value = 1, " 1=1 ", " X.����ID Is Not Null ") & _
            IIf(blnLimitUnit, " And X.����ID In (Select Column_Value From Table(Cast(f_str2list([1]) As Zltools.t_strlist))) ", "") & _
            IIf(blnByUnit, " Group by X.����ID Having X.����ID=A.ID)", " Group by X.����ID Having X.����ID=A.ID)")
    
    strSql = "Select Distinct A.ID,A.����,A.����" & _
        " From ���ű� A,��������˵�� B " & _
        " Where A.ID = B.����ID And B.��������=" & IIf(blnByUnit, "'����'", "'�ٴ�'") & _
        " And B.������� IN (" & str��Դ & ") " & strSql & _
        " And (A.����ʱ�� is NULL Or A.����ʱ��=TO_Date('3000-01-01','YYYY-MM-DD'))" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
        " Order by A.����"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strUnitIDs)
 
    
    '���������������ú�,��ʹ�ð�
    vsDept.Rows = 1
    vsDept.Rows = rsTmp.RecordCount + 1
    vsDept.TextMatrix(0, 1) = IIf(blnByUnit, "����", "����")
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            vsDept.TextMatrix(i, 0) = rsTmp!����
            vsDept.TextMatrix(i, 1) = rsTmp!����
            vsDept.RowData(i) = CLng(rsTmp!ID)
            rsTmp.MoveNext
        Next
    End If
    Call RestoreColSort(vsDept)
    
    mblnNotDo = True        '������vsDept_AfterRowColChange
    If lngDeptID <> -1 Then
        vsDept.Row = vsDept.FindRow(lngDeptID)      'colȱʡ-1��ʾ��rowdata����
        If vsDept.Row = -1 Then
            If vsDept.Rows > 1 Then vsDept.Row = 1
        End If
    Else
        If vsDept.Rows > 1 Then vsDept.Row = 1
    End If
    If vsDept.Row < 0 Then vsDept.Row = 0
    mblnNotDo = False
    
    Call vsDept.ShowCell(vsDept.Row, 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim lngModul As Long
    Dim blnHavePrivs As Boolean
    
    
    If mbytUseType = 3 Then
        lngModul = 1137
        mstrPrivs = ";" & GetPrivFunc(glngSys, lngModul) & ";"
        blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
        zlDatabase.SetPara "��ʾ���岡��", cboPatient.ListIndex, glngSys, lngModul, blnHavePrivs
    Else
        lngModul = 1150
        mstrPrivs = ";" & GetPrivFunc(glngSys, lngModul) & ";"
        blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    End If
    
    zlDatabase.SetPara "��ʾ��Ժ����", chkState(0).Value, glngSys, lngModul, blnHavePrivs
    zlDatabase.SetPara "��ʾԤ��Ժ����", chkState(1).Value, glngSys, lngModul, blnHavePrivs
    zlDatabase.SetPara "��ʾ��Ժ����", chkState(2).Value, glngSys, lngModul, blnHavePrivs
    
    
    Call SaveWinState(Me, App.ProductName)
    
    
    If mbytUseType = 3 Then
        zlDatabase.SetPara "��ʾ���˷�ʽ", cboType.ListIndex, glngSys, 1137, blnHavePrivs
    Else
        zlDatabase.SetPara "��ʾ���˷�ʽ", cboType.ListIndex, glngSys, 1150, blnHavePrivs
    End If
    mbytUseType = 0
    mlngUnitID = 0
End Sub

Private Sub vsPati_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Then
        If vsDept.RowData(vsDept.Row) = 0 Then Exit Sub
        If KeyCode = vbKeyLeft Then
            If vsDept.Row - 1 >= 1 Then vsDept.Row = vsDept.Row - 1
        ElseIf KeyCode = vbKeyRight Then
            If vsDept.Row + 1 <= vsDept.Rows - 1 Then
                vsDept.Row = vsDept.Row + 1
            End If
        End If
        Call vsDept.ShowCell(vsDept.Row, 0)
    End If
End Sub

Private Sub RestoreColSort(vsGrid As Object)
'���ܣ�������
    Dim strSort As String, i As Long
        
    With vsGrid
        Set .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 7
        If gblnMyStyle Then
            strSort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name & "\VSFlexGrid", .Name & "ColSort", "")
            If strSort <> "" Then
                .Col = Val(Split(strSort, ",")(0))
                .Sort = Val(Split(strSort, ",")(1))
                If Val(Split(strSort, ",")(1)) Mod 2 = 1 Then
                    .Cell(flexcpPicture, 0, .Col) = imgSort.ListImages(1).Picture
                Else
                    .Cell(flexcpPicture, 0, .Col) = imgSort.ListImages(2).Picture
                End If
            End If
        End If
    End With
End Sub
