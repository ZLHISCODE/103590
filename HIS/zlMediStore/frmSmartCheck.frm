VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmSmartCheck 
   Caption         =   "�̵�����ܼ��"
   ClientHeight    =   5775
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8355
   Icon            =   "frmSmartCheck.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8355
   StartUpPosition =   1  '����������
   Begin VSFlex8Ctl.VSFlexGrid vsfGrid 
      Height          =   4335
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   7815
      _cx             =   13785
      _cy             =   7646
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSmartCheck.frx":6852
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
   Begin VB.PictureBox picCondition 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   480
      Width           =   7815
      Begin VB.CheckBox chkType 
         BackColor       =   &H80000003&
         Caption         =   "�޿����δ�̵�"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   10
         ToolTipText     =   "���������̵�����ظ��̵��ҩƷ"
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtDay 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6000
         TabIndex        =   9
         Text            =   "3"
         Top             =   105
         Width           =   375
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   270
         Left            =   6375
         TabIndex        =   8
         Top             =   105
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtDay"
         BuddyDispid     =   196612
         OrigLeft        =   4200
         OrigTop         =   120
         OrigRight       =   4455
         OrigBottom      =   375
         Max             =   30
         Min             =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdCheck 
         Caption         =   "���(&C)"
         Height          =   300
         Left            =   6840
         TabIndex        =   6
         Top             =   90
         Width           =   855
      End
      Begin VB.CheckBox chkType 
         BackColor       =   &H80000003&
         Caption         =   "�̵��ظ�"
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   3
         ToolTipText     =   "���������̵�����ظ��̵��ҩƷ"
         Top             =   120
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkType 
         BackColor       =   &H80000003&
         Caption         =   "�̵���©"
         Height          =   255
         Index           =   0
         Left            =   2160
         TabIndex        =   2
         ToolTipText     =   "���������̵������©��ҩƷ"
         Top             =   120
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.Label lblDay 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "���������"
         Height          =   180
         Left            =   5160
         TabIndex        =   4
         Top             =   150
         Width           =   900
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackColor       =   &H80000003&
         Caption         =   "������ͣ�"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   900
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5415
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1235
            Text            =   "������"
            TextSave        =   "������"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12594
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgTool 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmSmartCheck.frx":68D6
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmSmartCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const mcon��ӡ As Integer = 103
Private Const mconԤ�� As Integer = 102
Private Const mcon�˳� As Integer = 191

Private mcbrToolBar As CommandBar
Private mobjPopup As CommandBar
Private mblnSuccess As Boolean '�����Ƿ����仯


Private mlng�ⷿID As Long
Private mfraPar As Form

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.id
        Case mcon��ӡ
            cbsFilePrint
        Case mconԤ��
            cbsFilePreView
        Case mcon�˳�
            Unload Me
    End Select
End Sub

Private Sub cbsFilePrint()
    '��ӡ
    vsfGrid.Redraw = flexRDNone
    subPrint 1
    vsfGrid.Redraw = flexRDDirect
    vsfGrid.Col = 0
    vsfGrid.ColSel = vsfGrid.Cols - 1
End Sub

Private Sub cbsFilePreView()
    '��ӡԤ��
    vsfGrid.Redraw = flexRDNone
    subPrint 2
    vsfGrid.Redraw = flexRDDirect
    vsfGrid.Col = 0
    vsfGrid.ColSel = vsfGrid.Cols - 1
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = "�����"
        
    objRow.Add "ʱ�䣺" & zlDataBase.Currentdate
    objRow.Add "���ţ�" & mfraPar.cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow

    objRow.Add "��ӡ��:" & UserInfo.�û�����
    objRow.Add "��ӡ����:" & Format(Sys.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = vsfGrid
    
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub


Private Sub cmdCheck_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim rsPhysic As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim strNo As String
    Dim lngҩƷid As Long
    Dim lng���� As Long
    Dim str����״̬ As String
    Dim str��ʾ��Ϣ As String
    Dim lng�̵�Sum As Long
    Dim lngҩƷSum As Long
    Dim lngҩƷSum2 As Long
    Dim str��ʾҩƷ As String
    Dim lng��ʾҩƷsum As Long
    Dim str©��ҩƷ As String
    Dim lng©��ҩƷsum As Long
    Dim bln��ɫ As Boolean
    Dim str©��ҩƷ��Ϣ As String
    
    On Error GoTo ErrHandle
    vsfGrid.Clear
    vsfGrid.rows = 1
    vsfGrid.RowHeight(-1) = 300
    
    '�����������δ�������̵��
    gstrSQL = "Select a.No, a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) ����, a.��¼״̬, a.����,b.���, a.����, b.����, b.����, Decode(a.�������, Null, Null, '����') ����״̬," & vbNewLine & _
            "       a.������, a.��������, a.�����, a.�������" & vbNewLine & _
            "From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B" & vbNewLine & _
            "Where a.��¼״̬ = 1 And a.���� = 12 And a.�������� > Sysdate - [2] And a.�ⷿid = [1] And a.ҩƷid = b.Id" & vbNewLine & _
            "Order By a.ҩƷid, ����"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "", mlng�ⷿID, Val(txtDay.Text))
    '���ظÿⷿ�������ô洢���Լ��п���ҩƷ
    gstrSQL = "Select a.*,b.����,b.����,b.���" & vbNewLine & _
            "From (Select a.Id ҩƷid, Null ����,Null ����" & vbNewLine & _
            "       From �շ���ĿĿ¼ A, ҩƷ��� B, �շ�ִ�п��� C" & vbNewLine & _
            "       Where a.Id = b.ҩƷid And a.Id = c.�շ�ϸĿid And c.ִ�п���id = [1] And Not Exists" & vbNewLine & _
            "        (Select 1" & vbNewLine & _
            "              From ҩƷ��� D" & vbNewLine & _
            "              Where a.Id = d.ҩƷid And d.�ⷿid = [1] And (ʵ������ <> 0 or ʵ�ʽ�� <> 0 or ʵ�ʲ�� <> 0))" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select Distinct a.ҩƷid, Nvl(a.����, 0) ����,�ϴ����� ����" & vbNewLine & _
            "       From ҩƷ��� A" & vbNewLine & _
            "       Where a.�ⷿid = [1] And (ʵ������ <> 0 or ʵ�ʽ�� <> 0 or ʵ�ʲ�� <> 0)) A, �շ���ĿĿ¼ B" & vbNewLine & _
            "Where a.ҩƷid = b.Id" & vbNewLine & _
            "Order By ҩƷid, ����"
    Set rsPhysic = zlDataBase.OpenSQLRecord(gstrSQL, "", mlng�ⷿID)
    
    If rsTemp.RecordCount = 0 And rsPhysic.RecordCount = 0 Then
        staThis.Panels(2).Text = "����ɣ�"
        Exit Sub
    End If
    
    '������������
    lngҩƷSum = IIf(chkType(0).Value = 0, 0, rsPhysic.RecordCount)
    lng�̵�Sum = IIf(chkType(1).Value = 0, 0, rsTemp.RecordCount)
    lngҩƷSum2 = IIf(chkType(2).Value = 0, 0, rsPhysic.RecordCount)
    
    If chkType(1).Value = 1 Then '�ظ��̵�ҩƷ
        Do While Not rsTemp.EOF
            strNo = rsTemp!NO
            lngҩƷid = rsTemp!ҩƷid
            lng���� = nvl(rsTemp!����, 0)
            str����״̬ = nvl(rsTemp!����״̬, "")
            str��ʾ��Ϣ = rsTemp!������ & "��" & rsTemp!�������� & "����" & IIf(IsNull(rsTemp!�����), "", "��" & rsTemp!����� & "��" & rsTemp!������� & "���")
            
            rsTemp.MoveNext
            
            If Not rsTemp.EOF Then
                If lngҩƷid = rsTemp!ҩƷid And lng���� = nvl(rsTemp!����, 0) Then
                
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "�̵��ظ������ݡ�" & strNo & "�� ҩƷ��[" & rsTemp!���� & "]" & rsTemp!���� & "(" & rsTemp!��� & ")" & "��" & "���ţ�" & IIf(IsNull(rsTemp!����), "��", rsTemp!����) & IIf(str����״̬ = "", "", "(" & str����״̬ & ")")
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 1) = strNo
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 2) = lngҩƷid
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 3) = rsTemp!����
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 4) = IIf(str����״̬ = "", 2, 4)
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 5) = str��ʾ��Ϣ
                    vsfGrid.rows = vsfGrid.rows + 1
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "          ���ݡ�" & rsTemp!NO & "�� ҩƷ��[" & rsTemp!���� & "]" & rsTemp!���� & "(" & rsTemp!��� & ")" & "��" & "���ţ�" & IIf(IsNull(rsTemp!����), "��", rsTemp!����) & IIf(IsNull(rsTemp!����״̬), "", "(" & rsTemp!����״̬ & ")")
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 1) = rsTemp!NO
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 2) = rsTemp!ҩƷid
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 3) = rsTemp!����
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 4) = IIf(IsNull(rsTemp!����״̬), 2, 4)
                    vsfGrid.TextMatrix(vsfGrid.rows - 1, 5) = rsTemp!������ & "��" & rsTemp!�������� & "����" & IIf(IsNull(rsTemp!�����), "", "��" & rsTemp!����� & "��" & rsTemp!������� & "���")
                    
                    If Not bln��ɫ Then vsfGrid.Cell(flexcpBackColor, vsfGrid.rows - 2, 0, vsfGrid.rows - 1, 0) = &H8000000F
                    bln��ɫ = Not bln��ɫ
                    
                    vsfGrid.rows = vsfGrid.rows + 1
                    
                End If
                
                Call zlControl.StaShowPercent(rsTemp.AbsolutePosition / (lngҩƷSum + lng�̵�Sum + lngҩƷSum2), staThis.Panels(2), frmSmartCheck)
            End If
            
        Loop
    End If
    
    If chkType(0).Value = 1 Then '�̵�©��ҩƷ
        rsPhysic.MoveFirst
        Do While Not rsPhysic.EOF
            If Not IsNull(rsPhysic!����) Then '�п�浫��ָ��ʱ�����̵���в����ڵ�ҩƷ
                rsTemp.Filter = "ҩƷid = " & rsPhysic!ҩƷid & " And ���� = " & rsPhysic!����
                
                If rsTemp.RecordCount = 0 Then
                    If str©��ҩƷ��Ϣ = "" Then
                        str©��ҩƷ��Ϣ = rsPhysic!ҩƷid & ":" & rsPhysic!����
                    Else
                        str©��ҩƷ��Ϣ = str©��ҩƷ��Ϣ & ";" & rsPhysic!ҩƷid & ":" & rsPhysic!����
                    End If
                    
                    If str©��ҩƷ = "" Then
                        lng©��ҩƷsum = lng©��ҩƷsum + 1
                        str©��ҩƷ = "[" & rsPhysic!���� & "]" & rsPhysic!���� & "(" & rsPhysic!��� & ")" & rsPhysic!����
                    Else
                        lng©��ҩƷsum = lng©��ҩƷsum + 1
                        If lng©��ҩƷsum <= 3 Then str©��ҩƷ = str©��ҩƷ & "��[" & rsPhysic!���� & "]" & rsPhysic!���� & "(" & rsPhysic!��� & ")" & rsPhysic!����
                    End If
                End If
                
            End If
            
            Call zlControl.StaShowPercent((rsPhysic.AbsolutePosition + lng�̵�Sum) / (lngҩƷSum + lng�̵�Sum + lngҩƷSum2), staThis.Panels(2), frmSmartCheck)
            
            rsPhysic.MoveNext
        Loop
        
        If lng©��ҩƷsum = 0 Then
            vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "�������̵���©��ҩƷ"
            vsfGrid.rows = vsfGrid.rows + 1
        Else
            If lng©��ҩƷsum > 0 Then
                vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "����������δ�̵��ҩƷ��:" & str©��ҩƷ & IIf(lng©��ҩƷsum > 3, "��" & lng©��ҩƷsum & "��", "")
                vsfGrid.TextMatrix(vsfGrid.rows - 1, 2) = str©��ҩƷ��Ϣ
                vsfGrid.RowHeight(vsfGrid.rows - 1) = IIf(lng©��ҩƷsum > 1, 600, 300)
                If Not bln��ɫ Then vsfGrid.Cell(flexcpBackColor, vsfGrid.rows - 1, 0, vsfGrid.rows - 1, 0) = &H8000000F
                vsfGrid.rows = vsfGrid.rows + 1
            End If
        End If
        
    End If
    
    If chkType(2).Value = 1 Then '�޿��δ�̵�
        rsPhysic.MoveFirst
        Do While Not rsPhysic.EOF
            If IsNull(rsPhysic!����) Then '���ô洢�������޿�棬ֻ���ҩƷid���������
                rsTemp.Filter = "ҩƷid = " & rsPhysic!ҩƷid
                
                If rsTemp.RecordCount = 0 Then
                    If str��ʾҩƷ = "" Then
                        lng��ʾҩƷsum = lng��ʾҩƷsum + 1
                        str��ʾҩƷ = "[" & rsPhysic!���� & "]" & rsPhysic!���� & "(" & rsPhysic!��� & ")"
                    Else
                        lng��ʾҩƷsum = lng��ʾҩƷsum + 1
                        If lng��ʾҩƷsum <= 3 Then str��ʾҩƷ = str��ʾҩƷ & "��[" & rsPhysic!���� & "]" & rsPhysic!���� & "(" & rsPhysic!��� & ")"
                    End If
                End If
            
            End If
            
            Call zlControl.StaShowPercent((rsPhysic.AbsolutePosition + lng�̵�Sum + lngҩƷSum) / (lngҩƷSum + lng�̵�Sum + lngҩƷSum2), staThis.Panels(2), frmSmartCheck)
            
            rsPhysic.MoveNext
        Loop
        
        If lng��ʾҩƷsum = 0 Then
            vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "�������޿��δ�̵�ҩƷ"
            vsfGrid.rows = vsfGrid.rows + 1
        Else
            If lng��ʾҩƷsum > 0 Then
                vsfGrid.TextMatrix(vsfGrid.rows - 1, 0) = "����������Ҳδ�̵��ҩƷ��:" & str��ʾҩƷ & IIf(lng��ʾҩƷsum > 3, "��" & lng��ʾҩƷsum & "��", "")
                vsfGrid.RowHeight(vsfGrid.rows - 1) = IIf(lng��ʾҩƷsum > 1, 600, 300)
                If Not bln��ɫ Then vsfGrid.Cell(flexcpBackColor, vsfGrid.rows - 1, 0, vsfGrid.rows - 1, 0) = &H8000000F
                vsfGrid.rows = vsfGrid.rows + 1
            End If
        End If
        
    End If
    
    staThis.Panels(2).Text = "����ɣ�"
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ShowME(ByVal lng�ⷿID As Long, ByVal fraPar As Form, ByRef blnSuccess As Boolean)
    mlng�ⷿID = lng�ⷿID
    Set mfraPar = fraPar
    
    Me.Show 1, fraPar
    
    blnSuccess = mblnSuccess
End Sub


Private Sub CmdExit_Click()
    Unload Me
End Sub


Private Sub Form_Load()
        InitComandBars
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 8595 Then Me.Width = 8595
    If Me.Height < 6345 Then Me.Height = 6345
    
    picCondition.Width = Me.ScaleWidth
    cmdCheck.Left = Me.ScaleWidth - cmdCheck.Width - 300
    vsfGrid.Move vsfGrid.Left, vsfGrid.Top, Me.ScaleWidth, Me.ScaleHeight - picCondition.Height - staThis.Height - 500
    vsfGrid.ColWidth(0) = vsfGrid.Width - 10
End Sub


Private Sub txtDay_Change()
    If txtDay.Text > 30 Then txtDay.Text = 30
    If txtDay.Text < 0 Then txtDay.Text = 1
End Sub

Private Sub txtDay_GotFocus()
    txtDay.SelStart = 0
    txtDay.SelLength = Len(txtDay.Text)
    txtDay.SelText = txtDay.Text
End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub vsfGrid_DblClick()
    Dim blnSuccess As Boolean
    
    If vsfGrid.TextMatrix(vsfGrid.Row, 1) <> "" Then
    
        MousePointer = vbHourglass
        With vsfGrid
            frmNewCheckCard.ShowCard mfraPar, .TextMatrix(.Row, 1), Val(.TextMatrix(.Row, 4)), 1, blnSuccess, Val(.TextMatrix(.Row, 2)), Val(.TextMatrix(.Row, 3))
        End With
        MousePointer = vbDefault
        
    ElseIf vsfGrid.TextMatrix(vsfGrid.Row, 2) <> "" Then
        With vsfGrid
            frmNewCheckCard.ShowCard mfraPar, .TextMatrix(.Row, 1), 9, 1, blnSuccess, , , .TextMatrix(.Row, 2)
        End With
    End If
    
    mblnSuccess = blnSuccess
    If blnSuccess Then cmdCheck_Click
End Sub

Private Sub InitComandBars()
    '��ʼ���������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim ctrCustom As CommandBarControlCustom
    Dim intCount As Integer

    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto

    Me.cbsMain.VisualTheme = xtpThemeOffice2003 + xtpThemeOfficeXP

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16

    End With

    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = imgTool.Icons
    
    
    '����������
    Set mcbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagAlignAny Or xtpFlagHideWrap
    
    With mcbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mcon��ӡ, "��ӡ")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, mconԤ��, "Ԥ��")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        
        Set cbrControlMain = .Add(xtpControlButton, mcon�˳�, "�˳�")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        
    End With

    cbsMain.Item(1).Delete
     
     '�����
    With Me.cbsMain.KeyBindings
        .Add 0, VK_ESCAPE, mcon�˳�
    End With

End Sub

Private Sub vsfGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If vsfGrid.MouseRow >= 0 Then vsfGrid.ToolTipText = vsfGrid.TextMatrix(vsfGrid.MouseRow, 5)
End Sub
