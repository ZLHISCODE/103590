VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDeliveryInvoice 
   Caption         =   "�ͻ���Ʊ����"
   ClientHeight    =   8070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10935
   Icon            =   "frmDeliveryInvoice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   10935
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picGetParams 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   240
      ScaleHeight     =   4860
      ScaleWidth      =   3735
      TabIndex        =   2
      Top             =   2280
      Width           =   3735
      Begin VB.CommandButton cmdGetData 
         Caption         =   "��ȡ����(&G)"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Frame fraParams 
         Height          =   4095
         Left            =   150
         TabIndex        =   3
         Top             =   120
         Width           =   3375
         Begin VB.ComboBox cboDrugDH 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1680
            Width           =   3015
         End
         Begin VB.ComboBox cboDrugWH 
            Height          =   300
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txtProvider 
            Height          =   270
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   2775
         End
         Begin VB.CommandButton cmdPS 
            Caption         =   "��"
            Height          =   255
            Left            =   2880
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   480
            Width           =   255
         End
         Begin VB.OptionButton optParams01 
            Caption         =   "��ȡ�����ѵ��������ⵥ"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   2160
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton optParams01 
            Caption         =   "��ȡ���ѵ��������ⵥ"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   2450
            Width           =   2655
         End
         Begin VB.Frame fraParams01 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   1250
            Left            =   480
            TabIndex        =   13
            Top             =   2720
            Width           =   2775
            Begin VB.TextBox txtParam02 
               Height          =   270
               Left            =   1320
               TabIndex        =   17
               Top             =   300
               Width           =   1335
            End
            Begin VB.OptionButton optParams02 
               Caption         =   "��Ʊ����(&D)"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   18
               Top             =   670
               Width           =   1290
            End
            Begin VB.OptionButton optParams02 
               Caption         =   "��Ʊ��(&I)"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   14
               Top             =   0
               Value           =   -1  'True
               Width           =   1170
            End
            Begin VB.TextBox txtParam01 
               Height          =   270
               Left            =   1320
               TabIndex        =   15
               Top             =   0
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker dtpParam01 
               Height          =   270
               Left            =   1320
               TabIndex        =   19
               Top             =   670
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   476
               _Version        =   393216
               Format          =   179175425
               CurrentDate     =   40290
            End
            Begin MSComCtl2.DTPicker dtpParam02 
               Height          =   270
               Left            =   1320
               TabIndex        =   21
               Top             =   980
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   476
               _Version        =   393216
               Format          =   179175425
               CurrentDate     =   40290
            End
            Begin VB.Label lblTo 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   2
               Left            =   1080
               TabIndex        =   16
               Top             =   300
               Width           =   180
            End
            Begin VB.Label lblTo 
               AutoSize        =   -1  'True
               Caption         =   "��"
               Height          =   180
               Index           =   3
               Left            =   1080
               TabIndex        =   20
               Top             =   980
               Width           =   180
            End
         End
         Begin VB.Label lblDH 
            AutoSize        =   -1  'True
            Caption         =   "ҩ��(&H)"
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   1440
            Width           =   630
         End
         Begin VB.Label lblWH 
            AutoSize        =   -1  'True
            Caption         =   "ҩ��(&W)"
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   630
         End
         Begin VB.Label lblProvider 
            AutoSize        =   -1  'True
            Caption         =   "��Ӧ��(&P)"
            Height          =   180
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   810
         End
      End
   End
   Begin VB.PictureBox picView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2280
      ScaleHeight     =   1575
      ScaleWidth      =   5100
      TabIndex        =   0
      Top             =   360
      Width           =   5100
      Begin VB.PictureBox picFunc 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   400
         Left            =   240
         ScaleHeight     =   405
         ScaleWidth      =   3975
         TabIndex        =   25
         Top             =   1100
         Width           =   3975
         Begin VB.TextBox txtIVNO 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1320
            TabIndex        =   27
            Top             =   70
            Width           =   1935
         End
         Begin VB.Label lblIVNO 
            AutoSize        =   -1  'True
            Caption         =   "���ҷ�Ʊ��(&N)"
            Height          =   180
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   1170
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfView 
         Height          =   1000
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   2655
         _cx             =   4683
         _cy             =   1764
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
         BackColorAlternate=   -2147483645
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
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
         FixedCols       =   1
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
   End
   Begin MSComctlLib.TreeView tvwProvider 
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   2143
      _Version        =   393217
      Indentation     =   529
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   24
      Top             =   7710
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   635
      SimpleText      =   $"frmDeliveryInvoice.frx":1CFA
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDeliveryInvoice.frx":1D41
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14208
            Text            =   "��ɫ��Ϊ�Ѵ���������ݣ� ��ɫ��Ϊ����ѡ������ݣ� ��ɫ��Ϊ�������ݡ�"
            TextSave        =   "��ɫ��Ϊ�Ѵ���������ݣ� ��ɫ��Ϊ����ѡ������ݣ� ��ɫ��Ϊ�������ݡ�"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin XtremeCommandBars.CommandBars cmbMain 
      Left            =   9600
      Top             =   600
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmDeliveryInvoice.frx":25D5
      Left            =   9120
      Top             =   600
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDeliveryInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytMarked As Byte     '0-δ��ȡ�������ݣ�1-�ѻ�ȡ��������

Private Sub cboDrugWH_Click()
    If cboDrugWH.ListIndex < 0 Then Exit Sub
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
    strSQL = "SELECT DISTINCT a.id, a.���� " _
           & "From ���ű� a, ҩƷ������� b " _
           & "Where a.id = b.�Է��ⷿid and b.���ڿⷿID = [1] " _
           & "Order by a.���� "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, "��ȡҩ����Ϣ", cboDrugWH.ItemData(cboDrugWH.ListIndex))
    With cboDrugDH
        .Clear
        .AddItem ""
        .ItemData(.NewIndex) = "0"
        Do While Not rsTmp.EOF
            .AddItem rsTmp!����
            .ItemData(.NewIndex) = rsTmp!Id
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End With
End Sub

Private Sub cmbMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case enm_Pop_File.FilePrintSet
            frmOutsideLinkSet.Show vbModal, Me
        Case enm_Pop_File.EditIgnore
            Call DataIgnore(vsfView)
        Case enm_Pop_File.EditProcess
            Call ProcProcess
        Case enm_Pop_File.EditCurrChoose
            SignData vsfView, 4, True
        Case enm_Pop_File.EditCurrCancel
            SignData vsfView, 4, False
        Case enm_Pop_File.EditChooChoose
            SignData vsfView, 3, True
        Case enm_Pop_File.EditChooCancel
            SignData vsfView, 3, False
        Case enm_Pop_File.EditAllChoose
            SignData vsfView, 1, True
        Case enm_Pop_File.EditAllCancel
            SignData vsfView, 0, False
        Case enm_Pop_File.ViewRefresh
            Call cmdGetData_Click
        Case enm_Pop_File.ViewFindButton
            Call FindString
        Case enm_Pop_File.ViewToolsButton
            Control.Checked = Not Control.Checked
            cmbMain(2).Visible = Control.Checked
            cmbMain.RecalcLayout
        Case enm_Pop_File.ViewToolsLabel
            Dim cbcControl As CommandBarControl
            Control.Checked = Not Control.Checked
            For Each cbcControl In Me.cmbMain(2).Controls
                cbcControl.Style = IIf(cbcControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
            Next
            cmbMain.RecalcLayout
        Case enm_Pop_File.ViewToolsIcon
            Control.Checked = Not Control.Checked
            cmbMain.Options.LargeIcons = Not Me.cmbMain.Options.LargeIcons
            cmbMain.RecalcLayout
        Case enm_Pop_File.ViewStatebar
            Control.Checked = Not Control.Checked
            stbThis.Visible = Not stbThis.Visible
            cmbMain.RecalcLayout
        Case enm_Pop_File.FileExit
            Unload Me
    End Select
End Sub

Private Sub cmbMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cmdGetData_Click()
    Dim strDB As String, strServer As String, strUser As String, strPWD As String
    Dim strSQL As String, strProvider As String
    Dim isConn As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim dtEnd As Date
    Dim str�ⷿID As String

    '�������
'    If Trim(txtProvider.Text) = "" Then
'        MsgBox "��¼�롰��Ӧ�̡���Ϣ��", vbInformation, GSTR_MESSAGE
'        txtProvider.SetFocus
'        Exit Sub
'    End If
    If cboDrugWH.ListIndex < 0 Then
        MsgBox "��ѡ��ҩ�⡱��Ϣ��", vbInformation, GSTR_MESSAGE
        cboDrugWH.SetFocus
        Exit Sub
    End If

    If optParams01(1).Value Then
        If optParams02(0).Value Then
            If Len(Trim(txtParam01.Text)) = 0 Or Len(Trim(txtParam02.Text)) = 0 Then
                MsgBox "������Ҫ��ȡ[��Ʊ��]��ʼ����������Ϣ��", vbInformation, GSTR_MESSAGE
                txtParam01.SetFocus
                Exit Sub
            End If
        Else
            If Len(Trim(dtpParam01.Value)) = 0 Or Len(Trim(dtpParam02.Value)) = 0 Then
                MsgBox "������Ҫ��ȡ[��Ʊ����]��ʼ����������Ϣ��", vbInformation, GSTR_MESSAGE
                dtpParam01.SetFocus
                Exit Sub
            End If
            If IsDate(dtpParam01.Value) = False Or IsDate(dtpParam02.Value) = False Then
                MsgBox "���������[��Ʊ����]��", vbInformation, GSTR_MESSAGE
                dtpParam01.SetFocus
                Exit Sub
            End If
        End If
    End If

'��ȡ�ⲿ����
'step1 �����ⲿ���ݿ�
    strDB = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="DBNAME", Default:="")
    strServer = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="SERVER", Default:="")
    strUser = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="USER", Default:="")
    strPWD = GetSetting(appName:="ZLSOFT", Section:=GSTR_REGEDIT_PATH, Key:="PASSWORD", Default:="")
    strPWD = StringEnDeCodecn(strPWD, 68)
    'Ĭ��MSSQL��ʽ����
    isConn = MSSQLServerOpen(strServer, strDB, strUser, strPWD)
    
    If isConn = False Then
        MsgBox "���ӷ�����ʧ�ܣ��������м����ݿ�����ӣ�", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

'step2 ��ȡ���ݼ�
    Screen.MousePointer = vbHourglass

    strProvider = Trim(txtProvider.Text)
    If cboDrugDH.ListIndex < 0 Then
        str�ⷿID = cboDrugWH.ItemData(cboDrugWH.ListIndex) & "|" & cboDrugWH.ItemData(cboDrugWH.ListIndex)
    Else
        str�ⷿID = cboDrugWH.ItemData(cboDrugWH.ListIndex) & "|" & IIf(cboDrugDH.ItemData(cboDrugDH.ListIndex) = 0, cboDrugWH.ItemData(cboDrugWH.ListIndex), cboDrugDH.ItemData(cboDrugDH.ListIndex))
    End If
    
    On Error GoTo ErrHand
    strSQL = "select saler_code ��Ӧ��ID,saler_name ��Ӧ��,medical_code ҩƷID,medical_manu ������, plan_code �ƻ�����" _
           & "  ,produce_code ����,produce_date ��������,avail_date Ч��,medical_amt ��Ʊ����,b.in_sum PDA��������" _
           & "  ,his_checkQTY ����������, medical_amt - isnull(his_checkqty,0) ��������, purvey_price ������" _
           & "  ,invoice_code ��Ʊ��,invoice_date ��Ʊ����,pay_sum ��Ʊ���,his_check_status imported,detail_id, " _
           & "  (select tu.cName from dbo.t_User tu where tu.cUserName in " _
           & "     (select top 1 convert(varchar,c.iuserid) from t_InStore c " _
           & "      where c.idetail_id = a.detail_id and c.isaler_code = a.saler_code)) as ������ " _
           & "from WCMS_DOWN_INVOICE a left join " _
           & "(select idetail_id, isaler_code, sum(convert(decimal,iNum)) In_Sum from t_InStore group by idetail_id,isaler_code) b" _
           & " on a.detail_id=b.idetail_id and a.saler_code=b.isaler_code " _
           & "where (a.storagcode='" & str�ⷿID & "' or a.storagcode+'|'+a.storagcode='" & str�ⷿID & "') "
    If optParams01(0).Value Then
        strSQL = strSQL & " and isnull(his_check_status,'')<>'1' "
        If optParams02(0).Value Then
            If Trim(txtParam01.Text) <> "" And Trim(txtParam02.Text) <> "" Then
                strSQL = strSQL & " and invoice_code between '" & txtParam01.Text & "' and '" & txtParam02.Text & "' "
            End If
        Else
            strSQL = strSQL & " and cast(invoice_date as smalldatetime) between '" & Format(dtpParam01.Value, "yyyy-mm-dd hh:mm:ss") & "'" _
                   & " and '" & Format(dtpParam02.Value, "yyyy-mm-dd 23:59:59") & "'"
        End If
    ElseIf optParams01(1).Value Then
        If optParams02(0).Value Then    '��Ʊ��
            strSQL = strSQL & " and invoice_code between '" & txtParam01.Text & "' and '" & txtParam02.Text & "'"
        Else                            '��Ʊ����
            strSQL = strSQL & " and cast(invoice_date as smalldatetime) between '" & Format(dtpParam01.Value, "yyyy-mm-dd hh:mm:ss") & "'" _
                   & " and '" & Format(dtpParam02.Value, "yyyy-mm-dd 23:59:59") & "'"
        End If
    End If
    
    '��Ӧ������
    If strProvider <> "" Then
        strSQL = strSQL & " and saler_name like '%" & strProvider & "%'"
    End If
    
    strSQL = strSQL & " order by invoice_code,medical_code "
    rsTmp.Open strSQL, gcnOutside, adOpenStatic, adLockReadOnly

    If rsTmp.RecordCount <= 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "�ⲿ���ݿ�����ʱ�����ݿɻ�ȡ��", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If

'step3 װ������
    On Error GoTo 0
    mbytMarked = 0
    DataLoading vsfView, rsTmp, 1, IIf(optParams01(0).Value, 0, 1)
    mbytMarked = IIf(optParams01(0).Value, 0, 1)
   
    Err = 0: On Error Resume Next
    RefreshTVWProvider tvwProvider, vsfView
    If Err <> 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "װ�ع�Ӧ����Ϣʱ�쳣��", vbInformation, GSTR_MESSAGE
        Err = 0: On Error GoTo 0
        Exit Sub
    End If
    Err = 0: On Error GoTo 0
    
    '����ⷿ��Ϣ
    With cmbMain.FindControl(, enm_Pop_File.ImportControl)
        If cboDrugWH.Text <> "" Then
            .Text = cboDrugWH.Text
        Else
            .Text = ""
        End If
    End With
    lblWH.Tag = cboDrugWH.ItemData(cboDrugWH.ListIndex)
    If cboDrugDH.ListIndex < 0 Then
        lblDH.Tag = "0"
    Else
        lblDH.Tag = cboDrugDH.ItemData(cboDrugDH.ListIndex)
    End If
    
    Screen.MousePointer = vbDefault
    'MsgBox "��ȡ������ɣ�", vbInformation, GSTR_MESSAGE
    Exit Sub

ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox "��ȡ�ⲿ���ݴ���", vbInformation, GSTR_MESSAGE
End Sub

Private Sub cmdPS_Click()
    ProviderSelecter Me, txtProvider, True
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
        Case 1: Item.Handle = picView.hwnd
        Case 2: Item.Handle = tvwProvider.hwnd
        Case 3: Item.Handle = picGetParams.hwnd
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()
    Call GetUserNameInfo
    InitCommandBars cmbMain
    Call InitDKPMain
    Call InitToolBar
    Call SetMenu
    InitVSF vsfView, True
    dtpParam01.Value = Date - 7
    dtpParam02.Value = Date
    optParams02_Click 0
    Call SetMedicalWH
End Sub

Private Sub InitDKPMain()
'��ʼ��dkpMain
    Dim pneMain As Pane, pneProvider As Pane, pneGetParams As Pane ', pneFind As Pane
    With dkpMain
        Set pneMain = .CreatePane(1, Me.ScaleHeight, 0, DockRightOf)
        pneMain.Options = PaneNoCloseable + PaneNoHideable + PaneNoFloatable
        pneMain.Title = "����������"
        
        Set pneProvider = .CreatePane(2, 230, 400, DockLeftOf)
        pneProvider.Options = PaneNoCloseable + PaneNoFloatable '+ PaneNoHideable
        pneProvider.Title = "��Ӧ���б�"
        pneProvider.MinTrackSize.Width = 230
        pneProvider.MinTrackSize.Height = 50
        
        Set pneGetParams = .CreatePane(3, 245, 320, DockBottomOf, pneProvider)
        pneGetParams.Options = PaneNoCloseable + PaneNoFloatable
        pneGetParams.Title = "��������"
        pneGetParams.MinTrackSize.Height = 320
        pneGetParams.MaxTrackSize.Height = 320
        pneGetParams.MinTrackSize.Width = 245
        
        .VisualTheme = ThemeOffice2003
        .Options.HideClient = True
        If Not cmbMain Is Nothing Then .SetCommandBars cmbMain
    End With
    
End Sub

Private Sub SetMenu()
    Dim cbcControl As CommandBarControl, cbcControlParent As CommandBarControl
    Dim cbpMenuBar As CommandBarPopup
    
    cmbMain.ActiveMenuBar.Title = "�˵�"
    cmbMain.ActiveMenuBar.EnableDocking xtpFlagAlignTop
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.File, "�ļ�(&F)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.File
    With cbpMenuBar.CommandBar.Controls
        'Set cbcControl = .Add(xtpControlButton, arrMenuBars(1).Id, arrMenuBars(1).Caption & arrMenuBars(1).HotKey)
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FilePrintSet, "�������ݿ�����(&S)")
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FileExit, "�˳�(&X)")
        cbcControl.BeginGroup = True        '����Ϊһ��Ŀ�ʼ
    End With
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.Edit, "�༭(&E)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.Edit
    With cbpMenuBar.CommandBar.Controls
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditIgnore, "����(&I)")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditProcess, "���ݴ���(&P)")
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditCurrChoose, "��ǰ��Ӧ�̴�")
        cbcControl.BeginGroup = True
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditCurrCancel, "��ǰ��Ӧ��ȡ��")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditChooChoose, "ѡ�д�")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditChooCancel, "ѡ��ȡ��")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditAllChoose, "ȫ����")
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditAllCancel, "ȫ��ȡ��")
    End With
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.View, "�鿴(&V)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.View
    With cbpMenuBar.CommandBar.Controls
        Set cbcControlParent = .Add(xtpControlPopup, enm_Pop_File.ViewTools, "������(&T)")
        Set cbcControl = cbcControlParent.CommandBar.Controls.Add(xtpControlButton, enm_Pop_File.ViewToolsButton, "��׼��ť(&S)", -1, False)
        cbcControl.Checked = True
        Set cbcControl = cbcControlParent.CommandBar.Controls.Add(xtpControlButton, enm_Pop_File.ViewToolsLabel, "�ı���ǩ(&T)", -1, False)
        cbcControl.Checked = True
        Set cbcControl = cbcControlParent.CommandBar.Controls.Add(xtpControlButton, enm_Pop_File.ViewToolsIcon, "��ͼ��(&B)", -1, False)
        cbcControl.Checked = True
        
        Set cbcControlParent = .Add(xtpControlButton, enm_Pop_File.ViewStatebar, "״̬��(&S)")
        cbcControlParent.Checked = True
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.ViewRefresh, "ˢ��(&R)")
        cbcControl.ShortcutText = "F5"
        cbcControl.BeginGroup = True
    End With
    
    Set cbpMenuBar = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enm_Pop_File.Help, "����(&H)", -1, False)
    cbpMenuBar.Id = enm_Pop_File.Help
    With cbpMenuBar.CommandBar.Controls
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.HelpHelp, "��������(&H)")
        Set cbcControl = .Add(xtpControlPopup, enm_Pop_File.HelpWeb, "&WEB�ϵ�����")
        cbcControl.CommandBar.Controls.Add xtpControlButton, enm_Pop_File.HelpWebhome, "������ҳ(&H)", -1, False
        cbcControl.CommandBar.Controls.Add xtpControlButton, enm_Pop_File.HelpWebBBS, "������̳(&F)", -1, False
        cbcControl.CommandBar.Controls.Add xtpControlButton, enm_Pop_File.HelpWebFeelback, "���ͷ���(&M)", -1, False
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.HelpAbout, "����(&A)��")
        cbcControl.BeginGroup = True
    End With
    
    '�����
    With cmbMain.KeyBindings
'        .Add FCONTROL, Asc("X"), conMenu_File_Exit
        .Add 0, VK_F5, enm_Pop_File.ViewRefresh
        .Add 0, VK_F1, enm_Pop_File.HelpHelp
    End With
    
    For Each cbcControl In cbpMenuBar.Controls
        cbcControl.Style = xtpButtonIconAndCaption
    Next

End Sub

Private Sub InitToolBar()
    Dim cbcControl As CommandBarControl
    Dim cbrToolBar As CommandBar

    Set cbrToolBar = cmbMain.Add("������", xtpBarTop)
    'cbrToolBar.ShowTextBelowIcons = True
    cbrToolBar.EnableDocking xtpFlagAlignAny + xtpFlagStretched
    With cbrToolBar.Controls
        'Set cbcControl = .Add(xtpControlButton, arrMenuBars(1).Id, arrMenuBars(1).Caption)
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FilePrintSet, "����")
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditIgnore, "����")
        cbcControl.BeginGroup = True
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.EditProcess, "����")
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.ViewRefresh, "ˢ��")
        cbcControl.BeginGroup = True
        
        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.FileExit, "�˳�")
        cbcControl.BeginGroup = True
    End With
    For Each cbcControl In cbrToolBar.Controls
        If cbcControl.Type = xtpControlButton Then
            cbcControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With cbrToolBar.Controls
        Set cbcControl = .Add(xtpControlLabel, enm_Pop_File.ImportTitle, "����ҩ�⣺")
        cbcControl.BeginGroup = True
        cbcControl.Flags = xtpFlagRightAlign
        Set cbcControl = .Add(xtpControlEdit, enm_Pop_File.ImportControl, "")
        cbcControl.Flags = xtpFlagRightAlign
        cbcControl.Width = 200
        cbcControl.Enabled = False
    End With

'    Set cbrToolBar = cmbMain.Add("ҩ��", xtpBarTop)
'    cbrToolBar.EnableDocking xtpFlagAlignAny
'    With cbrToolBar.Controls
'        Set cbcControl = .Add(xtpControlLabel, enm_Pop_File.ImportTitle, "����ҩ�⣺")
'        cbcControl.Flags = xtpFlagRightAlign
'        Set cbcControl = .Add(xtpControlEdit, enm_Pop_File.ImportControl, "")
'        cbcControl.Width = 200
'        cbcControl.Enabled = False
'
''        Set cbcControl = .Add(xtpControlLabel, enm_Pop_File.ViewFindTitle, "����(��Ʊ��)��")
''        cbcControl.BeginGroup = True
''        Set cbcControl = .Add(xtpControlEdit, enm_Pop_File.ViewFindEdit, "")
''        cbcControl.Width = 120
''        Set cbcControl = .Add(xtpControlButton, enm_Pop_File.ViewFindButton, "")
'    End With
    
End Sub

Private Sub optParams02_Click(Index As Integer)
    Dim lngBackColor As Long
    On Error Resume Next
    If Index = 0 Then
        txtParam01.Enabled = True
        txtParam02.Enabled = True
        txtParam01.BackColor = vbWhite
        txtParam02.BackColor = vbWhite
        dtpParam01.Enabled = False
        dtpParam02.Enabled = False
        txtParam01.SetFocus
    Else
        txtParam01.Enabled = False
        txtParam02.Enabled = False
        txtParam01.BackColor = &H80000004
        txtParam02.BackColor = &H80000004
        dtpParam01.Enabled = True
        dtpParam02.Enabled = True
        dtpParam01.SetFocus
    End If
End Sub

Private Sub picGetParams_Resize()
    fraParams.Width = IIf(picGetParams.Width > 300, picGetParams.Width - 300, 0)
    txtProvider.Width = IIf(picGetParams.Width > 700 + cmdPS.Width, picGetParams.Width - 700 - cmdPS.Width, 0)
    cmdPS.Left = IIf(txtProvider.Width > 0, txtProvider.Left + txtProvider.Width + 20, 0)
    fraParams01.Width = IIf(picGetParams.Width > fraParams01.Left + 500, picGetParams.Width - fraParams01.Left - 500, 0)
    cboDrugWH.Width = IIf(picGetParams.Width > 650, picGetParams.Width - 650, 0)
'    txtParam01.Width = picGetParams.Width - 2090
'    txtParam02.Width = txtParam01.Width
'    dtpParam01.Width = txtParam01.Width
'    dtpParam02.Width = txtParam01.Width
End Sub

Private Sub picView_Resize()
    With picFunc
        .Top = picView.Height - picFunc.Height
        .Left = 0
        .Width = picView.Width
    End With
    With vsfView
        .Top = 0
        .Left = 0
        .Width = picView.Width
        If picView.Height > picFunc.Height Then
            .Height = picView.Height - picFunc.Height
        Else
            .Height = picView.Height
        End If
    End With
End Sub

Private Sub SetMedicalWH()
'����ҩ��combobox��Ϣ��ͬHIS�����û�Ҫ��HIS�Ĳ���Ȩ��һ����
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Dim i, j As Integer
    
    'ҩ����Ϣ
    strSQL = "SELECT DISTINCT a.id, a.���� " _
           & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
           & "Where  a.id = c.����id and c.�������� = b.����" _
           & "  and Instr('HIJ',b.����,1) > 0 AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
           & "  and a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1]) " _
           & "Order by a.���� "
    Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, glngUserID)
    
    cboDrugWH.Clear
    For i = 0 To rsTmp.RecordCount - 1
        cboDrugWH.AddItem rsTmp!����
        cboDrugWH.ItemData(i) = rsTmp!Id
        rsTmp.MoveNext
    Next
    rsTmp.Close

'    'ҩ����Ϣ
'    strSQL = "SELECT DISTINCT a.id, a.���� " _
'           & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
'           & "Where  a.id = c.����id and c.�������� = b.����" _
'           & "  and Instr('LMN',b.����,1) > 0 AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'" _
'           & "  and a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1]) " _
'           & "Order by a.���� "
'    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gintUserID)
'
'    cboDrugDH.Clear
'    cboDrugDH.AddItem "", 0: cboDrugDH.ItemData(0) = 0
'    For i = 1 To rsTmp.RecordCount
'        cboDrugDH.AddItem rsTmp!����
'        cboDrugDH.ItemData(i) = rsTmp!Id
'        rsTmp.MoveNext
'    Next
'    rsTmp.Close
        
End Sub

Private Sub tvwProvider_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer, intCounter As Integer
    Dim bytState As Byte
    'Check״̬��ʾ
    vsfView.Redraw = flexRDNone
    If Node.Key = "Root" Then
        For i = 2 To tvwProvider.Nodes.Count
            tvwProvider.Nodes(i).Checked = Node.Checked
        Next
    Else
        For i = 2 To tvwProvider.Nodes.Count
            If i = 2 Then
                If tvwProvider.Nodes(i).Checked Then
                    bytState = 2
                Else
                    bytState = 1
                End If
            Else
                If (bytState = 1 And tvwProvider.Nodes(i).Checked) Or (bytState = 2 And tvwProvider.Nodes(i).Checked = False) Then
                    bytState = 0
                    Exit For
                End If
            End If
        Next
        Select Case bytState
            Case 1: tvwProvider.Nodes(1).Checked = False
            Case 2: tvwProvider.Nodes(1).Checked = True
            Case Else: tvwProvider.Nodes(1).Checked = 0
        End Select
    End If
    '����VSFView����ɵļ�¼
    If Node.Key = "Root" Then
        For i = 1 To vsfView.Rows - 1
            vsfView.RowHidden(i) = Not Node.Checked
        Next
    Else
        For i = 1 To vsfView.Rows - 1
            If Node.Tag = -1 Then   '��������
                If vsfView.TextMatrix(i, vsfView.ColIndex("imported")) = "0,0" Then
                    vsfView.RowHidden(i) = Not Node.Checked
                End If
            ElseIf Node.Tag = Val(vsfView.TextMatrix(i, vsfView.ColIndex("providerid"))) Then
                If vsfView.TextMatrix(i, vsfView.ColIndex("imported")) <> "0,0" Then
                    vsfView.RowHidden(i) = Not Node.Checked
                End If
            End If
            
        Next
    End If
    '��д���
    intCounter = 1
    For i = 1 To vsfView.Rows - 1
        If vsfView.RowHidden(i) = False Then
            vsfView.TextMatrix(i, 1) = intCounter
            intCounter = intCounter + 1
        End If
    Next
    vsfView.Redraw = flexRDBuffered
End Sub

Private Sub txtIVNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call FindString
    End If
End Sub

Private Sub txtProvider_GotFocus()
    txtProvider.SelStart = 0: txtProvider.SelLength = Len(txtProvider.Text)
End Sub

Private Sub txtProvider_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call ProviderSelecter(Me, txtProvider, False)
    End If
End Sub

Private Sub vsfView_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfView
        'ѡ������޸�
        If Col = .ColIndex("choose") Then
            'If Mid(.TextMatrix(Row, .ColIndex("imported")), 3, 1) = "1"  Then
            If .TextMatrix(Row, .ColIndex("imported")) = "0,1" And Val(.TextMatrix(Row, .ColIndex("qty"))) > 0 Then
                Cancel = False
            Else
                Cancel = True
            End If
        '������������ʱ��״̬��"��Ʊ����������������"�����޸�����
        ElseIf Col = .ColIndex("qty") And Mid(.TextMatrix(Row, .ColIndex("imported")), 3, 1) = "1" Then '.TextMatrix(Row, .ColIndex("provider")) = "��Ʊ����������������" Then
            If CheckProvider(.TextMatrix(Row, .ColIndex("providerid"))) = "" Then
                .TextMatrix(Row, .ColIndex("provider")) = "��Ӧ��ID��"
                .TextMatrix(Row, .ColIndex("imported")) = "0,0"
                .TextMatrix(Row, .ColIndex("choose")) = 0
                .Cell(flexcpForeColor, Row, 3, Row, .ColIndex("mess")) = vbRed
            ElseIf Val(.TextMatrix(Row, .ColIndex("qty"))) <= Val(.TextMatrix(Row, .ColIndex("ivqty"))) - Val(.TextMatrix(Row, .ColIndex("chkqty"))) Then 'And Val(.TextMatrix(Row, .ColIndex("qty"))) > 0 Then
                .TextMatrix(Row, .ColIndex("provider")) = CheckProvider(.TextMatrix(Row, .ColIndex("providerid")))
                .TextMatrix(Row, .ColIndex("imported")) = "0,1"
                If Val(.TextMatrix(Row, .ColIndex("qty"))) > 0 Then
                    .TextMatrix(Row, .ColIndex("choose")) = 1
                Else
                    .TextMatrix(Row, .ColIndex("choose")) = 0
                End If
                .Cell(flexcpForeColor, Row, 3, Row, .ColIndex("mess")) = vbBlack
            ElseIf Val(.TextMatrix(Row, .ColIndex("qty"))) > Val(.TextMatrix(Row, .ColIndex("ivqty"))) - Val(.TextMatrix(Row, .ColIndex("chkqty"))) Then
                .TextMatrix(Row, .ColIndex("imported")) = "0,1"
                .TextMatrix(Row, .ColIndex("choose")) = 0
                .Cell(flexcpForeColor, Row, 3, Row, .ColIndex("mess")) = vbBlack
            Else
                .TextMatrix(Row, .ColIndex("imported")) = "0,0"
                .TextMatrix(Row, .ColIndex("choose")) = 0
                .Cell(flexcpForeColor, Row, 3, Row, .ColIndex("mess")) = vbRed
            End If
            If Val(.TextMatrix(Row, .ColIndex("pdaqty"))) > 0 Or Val(.TextMatrix(Row, .ColIndex("ivqty"))) <= Val(.TextMatrix(Row, .ColIndex("chkqty"))) Then
                Cancel = True
            Else
                Cancel = False
            End If
        Else: Cancel = True
        End If
    End With
End Sub

Private Sub vsfView_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col < 3 Then Cancel = True
End Sub

Private Sub vsfView_EnterCell()
    With vsfView
        '������ɫ
        .ForeColorSel = .Cell(flexcpForeColor, .Row, 3)
    End With
End Sub

Private Sub vsfView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopupMenu As CommandBarPopup
    Dim cbcControl As CommandBarControl
    
    If vsfView.Rows <= 1 Then Exit Sub
    
    If Button = vbRightButton Then
        Set objPopupMenu = cmbMain.ActiveMenuBar.FindControl(, enm_Pop_File.Edit)
        If Not objPopupMenu Is Nothing Then
            '����Ҫ���صĲ˵���
            For Each cbcControl In objPopupMenu.CommandBar.Controls
                If cbcControl.Id = enm_Pop_File.EditProcess Then
                    cbcControl.Visible = False
                    Exit For
                End If
            Next
            objPopupMenu.CommandBar.ShowPopup
            '�ָ�
            If Not cbcControl Is Nothing Then
                cbcControl.Visible = True
            End If
        End If
    End If
End Sub

Private Sub vsfView_RowColChange()
    '��ǰ��¼�ü�ͷָʾ
    vsfView.Cell(flexcpText, 0, 0, vsfView.Rows - 1, 0) = ""
    If vsfView.Row > 0 Then
        vsfView.Cell(flexcpFontName, , 0) = "Marlett"
        vsfView.TextMatrix(vsfView.Row, 0) = 4
    End If
End Sub

Private Sub ProcProcess()
    Dim strTmp As String
    Dim cboWH As CommandBarComboBox
    
    If vsfView.Rows <= 1 Or CheckRecord(vsfView) = False Then
        MsgBox "�����ݿ��Դ������Ȼ�ȡ���ݣ�", vbInformation, GSTR_MESSAGE
        Exit Sub
    End If
    
    '�ⲿ���ݿ��Ƿ�����
    On Error GoTo ExitSub
    If gcnOutside.State = adStateClosed Then gcnOutside.Open
    On Error GoTo 0

    '�������ݿ�
    If MsgBox("��ȷ��Ҫ������", vbInformation Or vbYesNo Or vbDefaultButton2, GSTR_MESSAGE) = vbNo Then Exit Sub
    
    Call ProcImport(cboWH)
    
    Exit Sub
    
ExitSub:
    MsgBox "�ⲿ���ݿ�����ʧ��!", vbCritical
    Exit Sub
End Sub

Private Sub ProcImport(ByVal cboWH As CommandBarComboBox)
    '��ⵥ�������ݴ���
    Dim strInsert As String, strTmp As String, strProviderID As String
    Dim strNO As String, strInDate As String, strIVNO As String
    Dim i As Integer, intCounter As Integer, intMaxXQ As Integer
    Dim lngPackageQTY As Long
    Dim dblFactQTY As Double
    Dim bytLotPrice As Byte
    Dim numAddRate As Double, numCurPrice As Double, numTmp As Double, numCost As Double
    Dim rsTmp As New ADODB.Recordset, rsSign As New ADODB.Recordset, rsSort As New ADODB.Recordset
    Dim strSQL As String
    Dim lngMedicalID As Long
    Dim intReturn As Integer, intRows As Integer
    Dim strMess As String
    Dim bytErrNo As Byte
    Dim strҩ�� As String
    
    
    '�����ݼ�����
    With rsSort
        On Error GoTo ErrProc
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Fields.Append "ID", adDouble, , adFldIsNullable
        .Fields.Append "��ҩ��λID", adDouble, , adFldIsNullable
        .Fields.Append "����", adVarChar, 60, adFldIsNullable
        .Fields.Append "����", adVarChar, 20, adFldIsNullable
        .Fields.Append "Ч��", adDate, 20, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, , adFldIsNullable
        .Fields.Append "�ɱ���", adDouble, , adFldIsNullable
        .Fields.Append "��Ʊ��", adVarChar, 200, adFldIsNullable
        .Fields.Append "��Ʊ����", adDate, , adFldIsNullable
        .Fields.Append "��Ʊ���", adDouble, , adFldIsNullable
        .Fields.Append "��������", adDate, , adFldIsNullable
        .Fields.Append "�ƻ�����", adVarChar, 50, adFldIsNullable
        .Fields.Append "detail_id", adVarChar, 40, adFldIsNullable
        .Open
        For i = 1 To vsfView.Rows - 1
            If Val(vsfView.TextMatrix(i, vsfView.ColIndex("choose"))) = 0 Or vsfView.RowHidden(i) = True Then GoTo Continue
            If Val(vsfView.TextMatrix(i, vsfView.ColIndex("ivqty"))) - Val(vsfView.TextMatrix(i, vsfView.ColIndex("chkqty"))) < Val(vsfView.TextMatrix(i, vsfView.ColIndex("qty"))) _
                And Val(vsfView.TextMatrix(i, vsfView.ColIndex("choose"))) <> 0 Then GoTo Continue
            'If CheckProvider(vsfView.TextMatrix(i, vsfView.ColIndex("providerid"))) = "" Then GoTo Continue
            '�������
'            If Val(vsfView.TextMatrix(i, vsfView.ColIndex("id"))) = 0 Then
'                MsgBox "��ZLHISҩƷID����Ӧ������" & i & "�У�", vbInformation, GSTR_MESSAGE
'                .Close
'                Exit Sub
'            End If
'            If Val(vsfView.TextMatrix(i, vsfView.ColIndex("providerid"))) = 0 Then
'                MsgBox "��ZLHISҩƷ��Ӧ��ID����Ӧ������" & i & "�У�", vbInformation, GSTR_MESSAGE
'                .Close
'                Exit Sub
'            End If
            .AddNew
            !Id = vsfView.TextMatrix(i, vsfView.ColIndex("id"))
            !��ҩ��λid = vsfView.TextMatrix(i, vsfView.ColIndex("providerid"))
            !���� = vsfView.TextMatrix(i, vsfView.ColIndex("producer"))
            !���� = vsfView.TextMatrix(i, vsfView.ColIndex("lot_no"))
            !Ч�� = vsfView.TextMatrix(i, vsfView.ColIndex("avail_date"))
            !ʵ������ = vsfView.TextMatrix(i, vsfView.ColIndex("qty"))
            !�ɱ��� = vsfView.TextMatrix(i, vsfView.ColIndex("price"))
            !��Ʊ�� = vsfView.TextMatrix(i, vsfView.ColIndex("invoice"))
            !��Ʊ���� = vsfView.TextMatrix(i, vsfView.ColIndex("idate"))
            !��Ʊ��� = vsfView.TextMatrix(i, vsfView.ColIndex("iamount"))
            !�������� = vsfView.TextMatrix(i, vsfView.ColIndex("pdate"))
            !�ƻ����� = vsfView.TextMatrix(i, vsfView.ColIndex("plan_code"))
            !detail_id = vsfView.TextMatrix(i, vsfView.ColIndex("detail_id"))
            .Update
Continue:
        Next
        .Sort = "��ҩ��λID,��Ʊ��,ID"
        If .RecordCount > 0 Then .MoveFirst
    
'        'ע��: ��ҽ��Ժ��zl9comlib.dll�汾��9.36.0.120 (2010-6-11)
        strInDate = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        
        strҩ�� = Trim(cboDrugDH.Text)
'        If strҩ�� = "" Then
'            strҩ�� = cmbMain.FindControl(, enm_Pop_File.ImportControl).Text
'        End If
        
        gcnOracle.BeginTrans
        gcnOutside.BeginTrans
        On Error GoTo ErrHand
        
        Do While Not .EOF
            'ȡ��ⵥ�ݺ�(NO)
            If strProviderID <> IIf(IsNull(!��ҩ��λid), -99, !��ҩ��λid) Then
                'ע��: ��ҽ��Ժ��zl9comlib.dll�汾��9.36.0.120 (2010-6-11)
                strNO = gobjComLib.zlDatabase.GetNextNo(21, lblWH.Tag)
                intCounter = 1
            End If
            'lngMedicalID = !Id
            strSQL = "Select A.���Ч��, A.ҩ���װ, A.�Ƿ���, round(1 / (1 - B.ָ������� / 100) - 1, 5) �ӳ���, c.�ּ� " _
                   & "From ҩƷĿ¼ A, ҩƷ��� B, �շѼ�Ŀ c " _
                   & "Where A.ҩƷid = B.ҩƷid and a.ҩƷid=c.�շ�ϸĿid and A.ҩƷid = [1] " _
                   & "  and (c.��ֹ���� Is Null Or Sysdate Between c.ִ������ And Nvl(c.��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))"
            Set rsTmp = gobjComLib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, !Id)
            
            '�������ݴ���
            If rsTmp.EOF Then
                intMaxXQ = 0            '���Ч��
                lngPackageQTY = 0       'ҩ���װ��
                bytLotPrice = 0         '����
                numAddRate = 0          '�ӳ���
                numCurPrice = 0         '�ּ�
            Else
                lngPackageQTY = rsTmp!ҩ���װ
                bytLotPrice = rsTmp!�Ƿ���
                numAddRate = rsTmp!�ӳ���
                numCurPrice = rsTmp!�ּ�
                '����SQL��
                strInsert = "zl_ҩƷ�⹺_INSERT("
                'NO
                strInsert = strInsert & "'" & strNO & "'"
                '���
                strInsert = strInsert & "," & intCounter
                '�ⷿID(ҩ��ID)
                strInsert = strInsert & "," & lblWH.Tag
                '�Է�����ID(ҩ��ID)
                strInsert = strInsert & "," & IIf(Val(lblDH.Tag) <= 0, "Null", lblDH.Tag)
                '��ҩ��λID
                strInsert = strInsert & "," & IIf(IsNull(!��ҩ��λid), "", !��ҩ��λid)
                'ҩƷID
                strInsert = strInsert & "," & IIf(IsNull(!Id), "", !Id)
                '����
                strInsert = strInsert & ",'" & IIf(IsNull(!����), "", !����) & "'"
                '����
                strInsert = strInsert & ",'" & IIf(IsNull(!����), "", !����) & "'"
                'Ч��
                'strTmp = Format(DateAdd("M", intMaxXQ, vsfView.TextMatrix(i, vsfView.ColIndex("pdate"))), "yyyy-mm-dd")
                'If gbytЧ�� = 1 Then '��Ч��
                '    strTmp = Format(DateAdd("D", -1, CDate(strTmp)), "yyyy-mm-dd")
                'End If
                'strInsert = strInsert & "," & IIf(strTmp = "", "Null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                strTmp = IIf(IsNull(!Ч��), "", !Ч��)
                strInsert = strInsert & "," & IIf(strTmp = "", "Null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                'ʵ������
                dblFactQTY = IIf(IsNull(!ʵ������), 0, !ʵ������) * lngPackageQTY
                strInsert = strInsert & "," & Round(dblFactQTY, 5)
                '�ɱ���
                numCost = IIf(IsNull(!�ɱ���), 0, !�ɱ���) / lngPackageQTY
                strInsert = strInsert & "," & numCost
                '�ɱ����
                strInsert = strInsert & "," & Round(dblFactQTY * numCost, 5)
                '����
                strInsert = strInsert & ",100"
                '���ۼ�
                numTmp = IIf(bytLotPrice = 1, numCost * (1 + numAddRate), numCurPrice)
                strInsert = strInsert & "," & numTmp
                '���۽��
                strInsert = strInsert & "," & Round(dblFactQTY * numTmp, 5)  'dblFactQTY���������ʵ��������numTmp���ۼ�
                '���
                strInsert = strInsert & "," & Round(dblFactQTY * numTmp, 5) - Round(dblFactQTY * numCost, 5)
                'ժҪ
                strInsert = strInsert & IIf(strҩ�� = "", ",Null", ",'��(" & strҩ�� & ")���졣'")
                '������
                strInsert = strInsert & ",'" & gstrUserNameNew & "'"
                '��Ʊ��
                strIVNO = IIf(IsNull(!��Ʊ��), "", !��Ʊ��)
                strInsert = strInsert & ",'" & strIVNO & "'"
                '��Ʊ����
                strTmp = IIf(IsNull(!��Ʊ����), "", !��Ʊ����)
                strInsert = strInsert & "," & IIf(strTmp = "", "Null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '��Ʊ���
                If Trim(strIVNO) <> "" And strTmp <> "" Then
                    'strInsert = strInsert & "," & IIf(IsNull(!��Ʊ���), 0, !��Ʊ���)
                    strInsert = strInsert & "," & Round(dblFactQTY * numTmp, 5)
                Else
                    strInsert = strInsert & ",Null"
                End If
                '��������
                strInsert = strInsert & ",to_date('" & strInDate & "','yyyy-mm-dd HH24:MI:SS')"
                '���
                strInsert = strInsert & ",Null"
                '��Ʒ�ϸ�֤
                strInsert = strInsert & ",Null"
                '�˲���
                strInsert = strInsert & ",Null"
                '�˲�����
                strInsert = strInsert & ",Null"
                '����
                strInsert = strInsert & ",Null"
                '�Ƿ��˻�
                strInsert = strInsert & ",1"
                '��������
                strTmp = !��������
                strInsert = strInsert & "," & IIf(strTmp = "", "Null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '��׼�ĺ�
                strInsert = strInsert & ",Null"
                '�������
                strInsert = strInsert & ",Null"
                '����
                strInsert = strInsert & ",Null"
                '�ӳ���
                strInsert = strInsert & "," & numAddRate
                strInsert = strInsert & ")"
                '�������Ӳ���
                bytErrNo = 1
                gobjComLib.zlDatabase.ExecuteProcedure strInsert, Me.Caption

'Step2 ��־�ѵ��봦��
                bytErrNo = 2
                strSQL = "declare @return int, @mess varchar(200) " & Chr(13)
                strSQL = strSQL & "execute sj_updInvoiceStatus_pro " _
                       & "'" & IIf(IsNull(!��ҩ��λid), "", !��ҩ��λid) _
                       & "','" & IIf(IsNull(!��Ʊ��), "", !��Ʊ��) _
                       & "','" & IIf(IsNull(!detail_id), "", !detail_id) _
                       & "'," & IIf(IsNull(!ʵ������), 0, !ʵ������) _
                       & ",@return output, @mess output " & Chr(13)
                strSQL = strSQL & "select @return return_, @mess mess"
                rsSign.Open strSQL, gcnOutside
                bytErrNo = 3
                If rsSign.EOF Then
                    intReturn = 0
                    strMess = "����м����ݱ�ʧ�ܣ�"
                Else
                    intReturn = rsSign!return_
                    strMess = rsSign!mess
                End If
                rsSign.Close

'�޸�ZLHIS�ƻ�����ִ������
                bytErrNo = 4
                If IIf(IsNull(!�ƻ�����), "", Trim(!�ƻ�����)) <> "" Then
                    strSQL = "Zl_ҩƷ�ƻ�����_�޸�ִ������('" _
                           & IIf(IsNull(!�ƻ�����), "", !�ƻ�����) & "', '" _
                           & IIf(IsNull(!Id), "", !Id) & "," & dblFactQTY & "')"
                    gobjComLib.zlDatabase.ExecuteProcedure strSQL, Me.Caption & "-�޸�ִ������"
                End If

'Step3 ��ɴ���
                numTmp = 0: dblFactQTY = 0: strTmp = ""
                bytErrNo = 5
                For i = 1 To vsfView.Rows - 1
                    If vsfView.TextMatrix(i, vsfView.ColIndex("detail_id")) = !detail_id Then
                        vsfView.TextMatrix(i, vsfView.ColIndex("mess")) = strMess
                        If intReturn = 1 Then
                            vsfView.TextMatrix(i, vsfView.ColIndex("mess")) = "OK"
                            intCounter = intCounter + 1
                        End If
                    End If
                Next
                
            End If
            rsTmp.Close
            
            strProviderID = IIf(IsNull(!��ҩ��λid), -99, !��ҩ��λid)      '����
            .MoveNext
        Loop
    End With

    '�ύ����
    gcnOracle.CommitTrans
    gcnOutside.CommitTrans
    
    'ˢ��VsfView
    For i = vsfView.Rows - 1 To 1 Step -1
        If vsfView.TextMatrix(i, vsfView.ColIndex("mess")) = "OK" Then
            vsfView.RemoveItem i
        End If
    Next
    If rsSort.State = adStateOpen Then rsSort.Close
    
    Exit Sub

ErrProc:
    'Call ErrCenter
    MsgBox "��������ʱ����", vbInformation, GSTR_MESSAGE
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    gcnOutside.RollbackTrans
    'Call ErrCenter
    MsgBox Err.Description & vbNewLine & "����ţ�" & bytErrNo
End Sub

Private Sub FindString()
    Dim i As Integer
    
    If Trim(txtIVNO.Text) <> "" And vsfView.Rows > 1 Then
        '���ҷ�Ʊ��
        With vsfView
            For i = 1 To .Rows - 1
                If UCase(.TextMatrix(i, .ColIndex("invoice"))) = UCase(Trim(txtIVNO.Text)) And .RowHidden(i) = False Then
                    .Row = i
                    .TopRow = i
                    .SetFocus
                    Exit Sub
                End If
            Next
        End With
        MsgBox "δ�ҵ���¼��ķ�Ʊ�ţ�", , GSTR_MESSAGE
    End If
    
'    Dim cbeFind As CommandBarEdit
'    Set cbeFind = cmbMain.FindControl(, enm_Pop_File.ViewFindEdit)
'
'    If cbeFind Is Nothing Then Exit Sub
'
'    If Trim(cbeFind.Text) <> "" And vsfView.Rows > 1 Then
'        '���ҷ�Ʊ��
'        Dim i As Integer
'        With vsfView
'            For i = 1 To .Rows - 1
'                If UCase(.TextMatrix(i, .ColIndex("invoice"))) = UCase(Trim(cbeFind.Text)) And .RowHidden(i) = False Then
'                    .Row = i
'                    .TopRow = i
'                    .SetFocus
'                    Exit Sub
'                End If
'            Next
'        End With
'        MsgBox "δ�ҵ���¼��ķ�Ʊ�ţ�", , GSTR_MESSAGE
'    End If
End Sub

Private Sub SignData(ByVal vsfVal As VSFlexGrid, ByVal bytVal As Byte, ByVal blnVal As Boolean)
'0: ȫ��ȡ��; 1:ȫ��ѡ��; 2: ѡ��ȡ��; 3:ѡ�д�; 4:��Ӧ��
    Dim i As Integer
    Dim strTmp As String
    
    If vsfVal.Rows < 2 Then Exit Sub
    
    With vsfVal
        strTmp = .TextMatrix(.Row, .ColIndex("provider"))
        'ע��: SelectedRowsҪ��Ч��SelectMode��ҪΪ flexSelectionListBox
        For i = 1 To .Rows - 1
            Select Case bytVal
                Case 0, 1
                    'vsfView.TextMatrix(i, 2) = IIf(blnVal And Mid(vsfView.TextMatrix(i, vsfView.ColIndex("imported")), 3, 1) = "1", "1", "0")
                    .TextMatrix(i, 2) = IIf(blnVal And .TextMatrix(i, .ColIndex("imported")) = "0,1", "1", "0")
                Case 2, 3
                    If .IsSelected(i) = True Then
                        .TextMatrix(i, 2) = IIf(blnVal And .TextMatrix(i, .ColIndex("imported")) = "0,1", "1", "0")
                    End If
                Case 4
                    If .TextMatrix(i, .ColIndex("provider")) = strTmp Then
                        .TextMatrix(i, 2) = IIf(blnVal And .TextMatrix(i, .ColIndex("imported")) = "0,1", "1", "0")
                    End If
            End Select
        Next
    End With
End Sub

Private Sub DataIgnore(ByVal vsfVal As VSFlexGrid)
'���Դ�������ƽ̨���ݿ�����Ѿ�������
    Dim i As Integer
    Dim strSQL As String
    Dim rsSign As New ADODB.Recordset
    
    If vsfVal.Rows < 2 Then Exit Sub
    
    If MsgBox("��ȷ�������ݺ��Բ�����", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    
    With vsfVal
        On Error GoTo errHandle
        If .SelectedRows > 0 Then gcnOutside.BeginTrans
        
        For i = 1 To .Rows - 1
            If .IsSelected(i) And Left(.TextMatrix(i, .ColIndex("imported")), 1) <> "1" Then
                '��־�ѵ��봦��
                strSQL = "declare @return int, @mess varchar(200) " & Chr(13)
'                strSQL = strSQL & "execute sj_updInvoiceStatus_pro '" & .TextMatrix(i, .ColIndex("detail_id")) & "', @return output, @mess output " & Chr(13)
                strSQL = strSQL & "execute sj_updInvoiceStatus_pro " _
                       & "'" & .TextMatrix(i, .ColIndex("providerid")) _
                       & "','" & .TextMatrix(i, .ColIndex("invoice")) _
                       & "','" & .TextMatrix(i, .ColIndex("detail_id")) _
                       & "',@return output, @mess output " & Chr(13)
                strSQL = strSQL & "select @return return_, @mess mess"
                rsSign.Open strSQL, gcnOutside
'                If rsSign.EOF Then
'                    intReturn = 0
'                    strMess = ""
'                Else
'                    intReturn = rsSign!return_
'                    strMess = rsSign!mess
'                End If
                rsSign.Close
            End If
        Next
        
        If .SelectedRows > 0 Then
            gcnOutside.CommitTrans
            Call cmdGetData_Click
        End If
    End With
    Exit Sub
    
errHandle:
    gcnOutside.RollbackTrans
    Call gobjComLib.ErrCenter
    Call gobjComLib.SaveErrLog
End Sub


