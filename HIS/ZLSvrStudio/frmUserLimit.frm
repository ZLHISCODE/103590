VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmUserLimit 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   10575
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   16305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmUserLimit.frx":0000
   ScaleHeight     =   10575
   ScaleWidth      =   16305
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctContainer 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9855
      Left            =   0
      ScaleHeight     =   9855
      ScaleWidth      =   14055
      TabIndex        =   10
      Top             =   480
      Width           =   14055
      Begin VB.PictureBox pctOpt 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3495
         ScaleWidth      =   13935
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   6000
         Width           =   13935
         Begin VB.PictureBox pctOption 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2850
            Left            =   8280
            ScaleHeight     =   2820
            ScaleWidth      =   5505
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   360
            Width           =   5535
            Begin VB.CommandButton cmdCancel 
               Caption         =   "ȡ��(&C)"
               Height          =   350
               Left            =   3900
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   2400
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.CommandButton cmdSave 
               Caption         =   "ȷ��(&O)"
               Height          =   350
               Left            =   3900
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   2040
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   3
               Left            =   2640
               MaxLength       =   3
               TabIndex        =   5
               Tag             =   "IP��ַ"
               Top             =   638
               Width           =   315
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   2
               Left            =   2160
               MaxLength       =   3
               TabIndex        =   4
               Tag             =   "IP��ַ"
               Top             =   638
               Width           =   315
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   0
               Left            =   1080
               MaxLength       =   3
               TabIndex        =   2
               Tag             =   "IP��ַ"
               Top             =   638
               Width           =   315
            End
            Begin VB.TextBox txtIP 
               Height          =   300
               Index           =   2
               Left            =   3420
               MaxLength       =   3
               TabIndex        =   6
               Tag             =   "IP"
               Top             =   615
               Width           =   390
            End
            Begin VB.TextBox txtUser 
               Height          =   350
               Left            =   1020
               TabIndex        =   1
               Top             =   120
               Width           =   2415
            End
            Begin VB.CommandButton cmdStop 
               Caption         =   "ͣ��(&S)"
               Height          =   350
               Left            =   3900
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   1560
               Width           =   1455
            End
            Begin VB.CommandButton cmdEdit 
               Caption         =   "�޸�����(&M)"
               Height          =   350
               Left            =   3900
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   600
               Width           =   1455
            End
            Begin VB.CommandButton cmdAdd 
               Caption         =   "�������(&A)"
               Height          =   350
               Left            =   3900
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   120
               Width           =   1455
            End
            Begin VB.CommandButton cmdDel 
               Caption         =   "ɾ������(&D)"
               Height          =   350
               Left            =   3900
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1080
               Width           =   1455
            End
            Begin VB.CommandButton cmdMore 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   6.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   3420
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   120
               Width           =   375
            End
            Begin VB.TextBox txtbeforeIp 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   225
               IMEMode         =   3  'DISABLE
               Index           =   1
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   3
               Tag             =   "IP��ַ"
               Top             =   638
               Width           =   315
            End
            Begin VB.TextBox txtDesc 
               Height          =   495
               Left            =   1020
               MaxLength       =   99
               MultiLine       =   -1  'True
               TabIndex        =   9
               Top             =   2040
               Width           =   2715
            End
            Begin MSComCtl2.DTPicker dtpStart 
               Height          =   345
               Left            =   1020
               TabIndex        =   7
               Top             =   1080
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/MM/dd HH:mm"
               DateIsNull      =   -1  'True
               Format          =   93978627
               CurrentDate     =   43024
            End
            Begin MSComCtl2.DTPicker dtpEnd 
               Height          =   345
               Left            =   1020
               TabIndex        =   8
               Top             =   1560
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/MM/dd HH:mm"
               DateIsNull      =   -1  'True
               Format          =   93978627
               CurrentDate     =   43024
            End
            Begin VB.TextBox txtIP 
               Enabled         =   0   'False
               Height          =   300
               Index           =   1
               Left            =   1020
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   20
               TabStop         =   0   'False
               Tag             =   "IP"
               Text            =   "    ��    ��    ��"
               Top             =   600
               Width           =   1965
            End
            Begin VB.Label lblIP 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "IP��"
               Height          =   180
               Left            =   480
               TabIndex        =   30
               Top             =   690
               Width           =   360
            End
            Begin VB.Label lblUser 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "�û���"
               Height          =   180
               Left            =   300
               TabIndex        =   29
               Top             =   210
               Width           =   540
            End
            Begin VB.Label lblEdit 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "-"
               Height          =   180
               Index           =   11
               Left            =   3180
               TabIndex        =   28
               Top             =   675
               Width           =   90
            End
            Begin VB.Label lblStart 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "��ʼʱ��"
               Height          =   180
               Left            =   120
               TabIndex        =   27
               Top             =   1125
               Width           =   720
            End
            Begin VB.Label lblEnd 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "����ʱ��"
               Height          =   180
               Left            =   120
               TabIndex        =   26
               Top             =   1635
               Width           =   720
            End
            Begin VB.Label lblDesc 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               Caption         =   "˵��"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   480
               TabIndex        =   25
               Top             =   2040
               Width           =   360
            End
         End
         Begin VB.TextBox txtTip 
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2895
            Left            =   1560
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   31
            TabStop         =   0   'False
            Text            =   "frmUserLimit.frx":803A
            Top             =   480
            Width           =   7095
         End
         Begin VB.TextBox txtStop 
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   480
            Width           =   105
         End
         Begin VB.Image imgIcon 
            Appearance      =   0  'Flat
            Height          =   1035
            Left            =   480
            Picture         =   "frmUserLimit.frx":81F1
            Stretch         =   -1  'True
            Top             =   120
            Width           =   915
         End
      End
      Begin VB.PictureBox pctLimit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6015
         Left            =   0
         ScaleHeight     =   6015
         ScaleWidth      =   12975
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   12975
         Begin VB.CheckBox chkAll 
            BackColor       =   &H00FFFFFF&
            Caption         =   "ֻ��ʾ�����������û��Ĺ���"
            Height          =   255
            Left            =   9600
            TabIndex        =   13
            Top             =   120
            Width           =   2655
         End
         Begin VB.TextBox txtFind 
            ForeColor       =   &H80000010&
            Height          =   350
            Left            =   960
            TabIndex        =   12
            Text            =   "�����û����������󰴻س����ж�λ"
            Top             =   80
            Width           =   3135
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfLimit 
            Height          =   5415
            Left            =   120
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   480
            Width           =   12135
            _cx             =   21405
            _cy             =   9551
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
            BackColorBkg    =   16777215
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
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
            Begin MSComctlLib.ImageList img16 
               Left            =   9360
               Top             =   1200
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   2
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmUserLimit.frx":9999
                     Key             =   "unCheck"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmUserLimit.frx":9F33
                     Key             =   "Check"
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Label lblFind 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "����"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   165
            Width           =   360
         End
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�û���IP����"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmUserLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsUsers As ADODB.Recordset  '�����û��ļ�¼��,���ڼ���û��������Ժͼ��ز��Ҵ���
Private Const conCol = "ѡ��,250,1;�û���,1200,1;����,1200,1;��ʼIP,1200,1;����IP,1200,1;��ʼʱ��,1200,1;����ʱ��,1200,1;״̬,500,1;˵��,1200,1"
Private mstrIDs As String     '���浱ǰѡ�е�ID,�����޸�/ɾ��
Private mdtStart As Date
Private mdtEnd As Date
Private Enum Color
    tipColor = &H80000010
    txtColor = &H80000012
End Enum
Private mblnKeypress As Boolean

Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���

End Function

Private Sub chkAll_Click()
    LoadLimit
End Sub

Private Sub ChangeCmdVisiable(ByVal blnIsAdd)
    '�޸İ�ť�ɼ���
    cmdAdd.Visible = Not blnIsAdd
    cmdDel.Visible = Not blnIsAdd
    cmdEdit.Visible = Not blnIsAdd
    cmdStop.Visible = Not blnIsAdd
    cmdSave.Visible = blnIsAdd
    cmdCancel.Visible = blnIsAdd
    
    'ѡ���
    If blnIsAdd Then
        txtUser.Text = ""
        txtbeforeIp(0).Text = ""
        txtbeforeIp(1).Text = ""
        txtbeforeIp(2).Text = ""
        txtbeforeIp(3).Text = ""
        txtIP(2).Text = ""
        dtpStart.Value = ""
        dtpEnd.Value = ""
        txtDesc.Text = ""
    Else
        With vsfLimit
            vsfLimit_AfterRowColChange 0, 0, .Row, .Col
        End With
    End If
    
    '����������
    cmdMore.Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
    txtUser.Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(0).Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(1).Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(2).Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
    txtbeforeIp(3).Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
    txtIP(2).Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
    dtpStart.Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
    dtpEnd.Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
    txtDesc.Enabled = Val(vsfLimit.Row) > 0 Or cmdSave.Visible
End Sub

Private Sub cmdAdd_Click()
    ChangeCmdVisiable True
End Sub

Private Sub cmdCancel_Click()
    ChangeCmdVisiable False
End Sub
Private Sub cmdSave_Click()
    Dim strTmp As String, i As Integer
    Dim strID As String, strStartIP As String, strEndIp As String
    Dim strStartTime As String, strEndTime As String
    Dim strDesc As String
    Dim varIDs As Variant
    
    On Error GoTo errh
    'У������
    Call GetDataFromCard(strID, strStartIP, strEndIp, strStartTime, strEndTime, strDesc)
    
    If mrsUsers Is Nothing Then
        Set mrsUsers = LoadUsers
    End If
    
    strTmp = CheckExist("�û���", strID, mrsUsers)
    If strTmp <> "" Then
        MsgBox "�����û���:" & strTmp & "������,������������롣", , "��ʾ"
        Exit Sub
    End If
    
    strTmp = ValidateTxt
    If strTmp <> "" Then
        frmMDIMain.stbThis.Panels(2).Text = strTmp
        Exit Sub
    End If
    
    '�ύ����
    Screen.MousePointer = vbArrowHourglass
    gcnOracle.BeginTrans
    If Len(strID) < 2000 Then
        Call ExecuteProcedure("Zl_Zlloginlimit_Edit(1,Null," & strStartTime & "," & strEndTime & ",'" & strStartIP & "','" & strEndIp & "',1,'" & strDesc & "','" & strID & "')", Me.Caption)
    Else
        varIDs = TranStr2Var(strID, ",", 2000)
        For i = 0 To UBound(varIDs)
            Call ExecuteProcedure("Zl_Zlloginlimit_Edit(1,Null," & strStartTime & "," & strEndTime & ",'" & strStartIP & "','" & strEndIp & "',1,'" & strDesc & "','" & varIDs(i) & "')", Me.Caption)
        Next
    End If
    gcnOracle.CommitTrans
    Screen.MousePointer = vbDefault
    
    With vsfLimit
        .Redraw = flexRDNone
        Call LoadLimit
    End With
    frmMDIMain.stbThis.Panels(2).Text = "��ӹ���ɹ���"
    Exit Sub
errh:
    frmMDIMain.stbThis.Panels(2).Text = ""
    Screen.MousePointer = vbDefault
    
    If InStr(1, UCase(err.Description), "ORA") Then '���ݿ����,�ַ����ϳ�,������ʾ,ͬʱ��������
        MsgBox "��ӹ���ʧ�ܡ�ԭ��" & vbNewLine & err.Description
        gcnOracle.RollbackTrans
    Else
        frmMDIMain.stbThis.Panels(2).Text = "��ӹ���ʧ�ܡ�ԭ��" & vbNewLine & err.Description
    End If
End Sub

Private Sub cmdDel_Click()
    Dim varIDs As Variant, i As Integer, intSRow As Integer
    
    On Error GoTo errh
    mstrIDs = GetSelectData
    
    '�ַ�������С��2000��,ֱ�ӽ���ɾ��,����2000��,���в�ֺ����ɾ��
    Screen.MousePointer = vbArrowHourglass
    gcnOracle.BeginTrans
    If Len(mstrIDs) < 2000 Then
        Call ExecuteProcedure("Zl_Zlloginlimit_Delete('" & mstrIDs & "')", Me.Caption)
    Else
        varIDs = TranStr2Var(mstrIDs, ",", 2000)
        For i = 0 To UBound(varIDs)
            Call ExecuteProcedure("Zl_Zlloginlimit_Delete('" & varIDs(i) & "')", Me.Caption)
        Next
    End If
    gcnOracle.CommitTrans
    Screen.MousePointer = vbDefault
    
    '�ڱ����ɾ������,�������»�ȡ����
    With vsfLimit
        intSRow = .Row
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - .FixedRows
            If i > .Rows - .FixedRows Or .Rows = .FixedRows Then
                Exit For
            End If
            If InstrEx(mstrIDs, .RowData(i)) Then
                .RemoveItem (i)
                i = i - 1
            End If
        Next
        .Redraw = flexRDDirect
        
        If intSRow > .Rows - .FixedRows Then
            .Select .Rows - .FixedRows, 0
        Else
            .Select intSRow, 0
        End If
        .TopRow = .Row
    End With
    
    mstrIDs = GetSelectData
    frmMDIMain.stbThis.Panels(2).Text = "ɾ������ɹ���"
    Exit Sub
errh:
    frmMDIMain.stbThis.Panels(2).Text = ""
    Screen.MousePointer = vbDefault
    
    If InStr(1, UCase(err.Description), "ORA") Then '���ݿ����,�ַ����ϳ�,������ʾ,ͬʱ��������
        MsgBox "��ӹ���ʧ�ܡ�ԭ��" & vbNewLine & err.Description
        frmMDIMain.stbThis.Panels(2).Text = ""
        gcnOracle.RollbackTrans
    Else
        frmMDIMain.stbThis.Panels(2).Text = "��ӹ���ʧ�ܡ�ԭ��" & vbNewLine & err.Description
    End If
End Sub

Private Sub cmdEdit_Click()
    EditLimit
End Sub

Private Sub cmdMore_Click()
    Dim strUsers As String
    Dim p As PointAPI
    Dim rstmp As ADODB.Recordset
    Dim strTmp() As String, i As Integer
    
    p.X = (pctOption.Left + cmdMore.Left + cmdMore.Width - FindUserWidth) / Screen.TwipsPerPixelX
    p.Y = (pctOpt.Top + pctContainer.Top - FindUserHeight) / Screen.TwipsPerPixelY
    ClientToScreen Me.hwnd, p
    
    If mrsUsers Is Nothing Then
        Set mrsUsers = LoadUsers
    End If
    
    strUsers = frmFindUser.ShowMe(Me, mrsUsers, Trim(txtUser.Text), p.X * Screen.TwipsPerPixelX, p.Y * Screen.TwipsPerPixelY)
    txtUser.Text = strUsers

End Sub

Private Sub CmdStop_Click()
    EditLimit ("Stop")
End Sub

Private Sub dtpEnd_Change()
    If IsNull(dtpEnd.Value) Then
        dtpStart.Value = Null
    Else
        mdtEnd = dtpEnd.Value
        dtpStart.Value = mdtStart
    End If
End Sub

Private Sub dtpStart_Change()
    If IsNull(dtpStart.Value) Then
        dtpEnd.Value = Null
    Else
        mdtStart = dtpStart.Value
        dtpEnd.Value = mdtEnd
    End If
End Sub

Private Sub Form_Load()
    Call InitTable(vsfLimit, conCol)
    Call LoadLimit
    Call ChangeCmdVisiable(False)
    
    '��ʼ�����ѡ���
    With vsfLimit
        .ColSort(-1) = flexSortCustom
        .ColSort(0) = flexSortNone
        .ColDataType(0) = flexDTBoolean
        .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        .Cell(flexcpText, 0, 0) = ""
        .Cell(flexcpPictureAlignment, 0, 0) = flexPicAlignCenterCenter
        .Editable = flexEDKbdMouse
    End With
    
    mdtStart = Now: mdtEnd = Now
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    pctContainer.Width = Me.ScaleWidth
    pctContainer.Height = Me.ScaleHeight - pctContainer.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsUsers = Nothing
End Sub

Private Sub pctContainer_Resize()
    
    On Error Resume Next
    
    LockWindowUpdate Me.hwnd
    pctLimit.Top = 0: pctLimit.Left = 0
    pctLimit.Width = pctContainer.ScaleWidth
    pctLimit.Height = pctContainer.ScaleHeight - pctOpt.Height
    
    pctOpt.Top = pctLimit.Height
    pctOpt.Left = 0
    pctOpt.Width = pctContainer.ScaleWidth
    LockWindowUpdate 0
End Sub

Private Sub pctLimit_Resize()
    On Error Resume Next
    
    vsfLimit.Width = pctLimit.ScaleWidth - 240
    vsfLimit.Height = pctLimit.ScaleHeight - vsfLimit.Top - 30
    
    lblFind.Left = vsfLimit.Left
    txtFind.Left = lblFind.Left + lblFind.Width + 45
    
    chkAll.Left = vsfLimit.Left + vsfLimit.Width - chkAll.Width
End Sub


Private Sub pctOpt_Resize()
    On Error Resume Next

    pctOption.Left = vsfLimit.Width + vsfLimit.Left - pctOption.Width
End Sub


Private Sub txtbeforeIp_Change(Index As Integer)
    Dim lngLineNo As Long '�к�
    Dim lngColNo  As Long '�к�
    err = 0
    On Error Resume Next
    
    Call GetCursorPos(Me.txtbeforeIp(Index).hwnd, lngLineNo, lngColNo)
    If lngColNo > 3 Then
        If Index < 3 Then
            If txtbeforeIp(Index + 1).Enabled Then txtbeforeIp(Index + 1).SetFocus
        End If
    End If
End Sub

Private Sub txtbeforeIp_GotFocus(Index As Integer)
    txtbeforeIp(Index).SelStart = 0
    txtbeforeIp(Index).SelLength = Len(txtbeforeIp(Index).Text)
End Sub

Private Sub txtbeforeIp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim lngLineNo As Long '�к�
    Dim lngColNo  As Long '�к�
    err = 0
    Call GetCursorPos(Me.txtbeforeIp(Index).hwnd, lngLineNo, lngColNo)
    
    Select Case KeyCode
    Case 37    '<-
        If Index > 0 Then
        If lngColNo > 1 Then Exit Sub
            If txtbeforeIp(Index - 1).Enabled Then
                txtbeforeIp(Index - 1).SelStart = Len(txtbeforeIp(Index - 1))
                txtbeforeIp(Index - 1).SetFocus
            End If
        End If
    Case 39    '->
        If Index < 3 Then
            If lngColNo <= Len(txtbeforeIp(Index)) Then Exit Sub
            If txtbeforeIp(Index + 1).Enabled Then
                txtbeforeIp(Index + 1).SelStart = 0
                txtbeforeIp(Index + 1).SetFocus
            End If
        End If
    Case 8     'BACKSPACE
        If Index > 0 Then
            If lngColNo > 1 Then Exit Sub
            If txtbeforeIp(Index - 1).Enabled Then
                txtbeforeIp(Index - 1).SelStart = Len(txtbeforeIp(Index - 1))
                txtbeforeIp(Index - 1).SetFocus
            End If
        End If
    End Select
    
    If InStr(1, "1234567890", Chr(KeyCode)) = 0 Then
        KeyCode = 0
    End If
    
End Sub

Private Sub txtbeforeIp_KeyPress(Index As Integer, KeyAscii As Integer)
    err = 0
    On Error Resume Next
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        If KeyAscii <> 13 Then
            If KeyAscii <> 8 Then
                If KeyAscii = Asc(".") Then
                    If Index < 3 And Index >= 0 And Trim(txtbeforeIp(Index)) <> "" Then
                        If txtbeforeIp(Index + 1).Enabled Then txtbeforeIp(Index + 1).SetFocus
                    End If
                End If
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Trim(txtFind.Text) = "" Then
            '���������ݰ��»س�,��ˢ��
            LoadLimit
        Else
            Call GetRowPos(vsfLimit, txtFind.Text, "�û���,����")
        End If
    End If
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text = "�����û����������󰴻س����ж�λ" Then
        txtFind.Text = ""
        txtFind.ForeColor = txtColor
    End If
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "�����û����������󰴻س����ж�λ"
        txtFind.ForeColor = tipColor
    End If
End Sub

Private Sub txtIp_KeyPress(Index As Integer, KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
End Sub

Private Sub txtUser_LostFocus()
    Call txtUser_KeyPress(13)
End Sub

Private Sub txtUser_Validate(Cancel As Boolean)
     If mblnKeypress Then
        mblnKeypress = False
    Else
        Call txtUser_KeyPress(13)
    End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    Dim strTmp As String, intPos As Integer
    
    If KeyAscii = 13 Then    '���»س�
        strTmp = Trim(txtUser.Text)
        intPos = InStrRev(strTmp, ",")
        strTmp = UCase(Mid(strTmp, intPos + 1))
        If strTmp = "" Then Exit Sub
        strTmp = Left(Trim(txtUser.Text), intPos) & FindUser(strTmp)
        
        txtUser.Text = strTmp
        txtUser.SelStart = Len(strTmp)
    End If
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    mblnKeypress = True
End Sub


Private Sub LoadLimit(Optional ByVal strFind As String)
    '����:��������
    Dim strSQL As String, rstmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo errh
    
    strSQL = "Select a.Id, a.�û���, c.����, a.��ʼip, a.����ip, a.��ʼʱ��, a.����ʱ��, Decode(a.״̬, 1, '������', 'δ����') ״̬, a.˵��" & vbNewLine & _
                    "From Zlloginlimit A, �ϻ���Ա�� B, ��Ա�� C" & vbNewLine & _
                    "Where a.�û��� = b.�û���(+) And b.��Աid = c.Id(+)" & IIf(chkAll.Value = 1, " And A.�û��� is null ", "") & strFind
    Set rstmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, "LoadLimit")
    
    With vsfLimit
        
        If rstmp.RecordCount = 0 Then
             .Rows = .FixedRows
            Exit Sub
        End If

        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Rows = rstmp.RecordCount + .FixedRows
        
        i = .FixedRows
        Do While Not rstmp.EOF
            .RowData(i) = "" & rstmp!id
            .TextMatrix(i, 0) = "0"
            .TextMatrix(i, .ColIndex("�û���")) = rstmp!�û��� & ""
            .TextMatrix(i, .ColIndex("����")) = rstmp!���� & ""
            .TextMatrix(i, .ColIndex("��ʼIP")) = rstmp!��ʼIP & ""
            .TextMatrix(i, .ColIndex("����IP")) = rstmp!����IP & ""
            .TextMatrix(i, .ColIndex("��ʼʱ��")) = rstmp!��ʼʱ�� & ""
            .TextMatrix(i, .ColIndex("����ʱ��")) = rstmp!����ʱ�� & ""
            .TextMatrix(i, .ColIndex("״̬")) = rstmp!״̬ & ""
            .TextMatrix(i, .ColIndex("˵��")) = rstmp!˵�� & ""
            i = i + 1: rstmp.MoveNext
        Loop
        
        .AutoResize = True: .AutoSize 0, .Cols - 2
        
        .Redraw = flexRDDirect
        If .Rows > .FixedRows Then
            .Select .FixedRows, 0
        End If
    End With
    
    Exit Sub
errh:
    MsgBox err.Description
    If 0 = 1 Then
        Resume
    End If
End Sub



Private Sub vsfLimit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Then Cancel = True
End Sub

Private Sub vsfLimit_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer
    
    With vsfLimit
        If .Rows = .FixedRows Then Exit Sub
        If Col = 0 Then
            If .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture Then
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "-1"
                Next
            Else
                .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
                For i = .FixedRows To .Rows - .FixedRows
                    .TextMatrix(i, 0) = "0"
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfLimit_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer, blnAllSelectd As Boolean
    
    blnAllSelectd = True
    With vsfLimit
        If .Redraw = flexRDNone Then Exit Sub
        
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 0) = "0" Then
                blnAllSelectd = False
                Exit For
            End If
        Next

        
        If blnAllSelectd Then
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("Check").Picture
        Else
            .Cell(flexcpPicture, 0, 0) = img16.ListImages("unCheck").Picture
        End If

    End With
End Sub

Private Sub vsfLimit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp() As String
    
    With vsfLimit
        If .Redraw = flexRDNone Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        txtUser.Text = .TextMatrix(NewRow, .ColIndex("�û���"))
        txtDesc.Text = .TextMatrix(NewRow, .ColIndex("˵��"))
        cmdStop.Caption = IIf(.TextMatrix(NewRow, .ColIndex("״̬")) = "������", "ͣ��", "����")
        
        If .TextMatrix(NewRow, .ColIndex("��ʼIP")) <> "" Then
            strTmp = Split(.TextMatrix(NewRow, .ColIndex("��ʼIP")), ".")
            txtbeforeIp(0).Text = strTmp(0)
            txtbeforeIp(1).Text = strTmp(1)
            txtbeforeIp(2).Text = strTmp(2)
            txtbeforeIp(3).Text = strTmp(3)
            txtIP(2).Text = Split(.TextMatrix(NewRow, .ColIndex("����IP")), ".")(3)
        Else
            txtbeforeIp(0).Text = ""
            txtbeforeIp(1).Text = ""
            txtbeforeIp(2).Text = ""
            txtbeforeIp(3).Text = ""
            txtIP(2).Text = ""
        End If
        
        If .TextMatrix(NewRow, .ColIndex("��ʼʱ��")) <> "" Then
            dtpStart.Value = Format(.TextMatrix(NewRow, .ColIndex("��ʼʱ��")), "yyyy/MM/dd HH:mm")
            dtpEnd.Value = Format(.TextMatrix(NewRow, .ColIndex("����ʱ��")), "yyyy/MM/dd HH:mm")
        Else
            dtpStart.Value = ""
            dtpEnd.Value = ""
        End If
        
    End With
End Sub


Private Function ValidateTxt() As String
'����:У�������Ƿ�Ϸ�,������벻�Ϸ�,���ش�����Ϣ.
    Dim strStartIP As String, strEndIp As String
    Dim strErr As String
    
    If txtIP(2).Text = "" And txtbeforeIp(0).Text = "" And txtUser.Text = "" And IsNull(dtpStart.Value) Then
        ValidateTxt = "����һ����������Ϊ�գ����������롣"
        Exit Function
    End If
    
    strStartIP = txtbeforeIp(0) & "." & txtbeforeIp(1) & "." & txtbeforeIp(2) & "." & txtbeforeIp(3)
    strEndIp = txtbeforeIp(0) & "." & txtbeforeIp(1) & "." & txtbeforeIp(2) & "." & txtIP(2)
    Call CheckIpValidate(strStartIP, strEndIp, strErr)
    If strErr <> "" Then
        ValidateTxt = strErr
        Exit Function
    End If
    
    If (IsNull(dtpStart.Value) And Not IsNull(dtpEnd.Value)) Or (Not IsNull(dtpStart.Value) And IsNull(dtpEnd.Value)) Then
        ValidateTxt = "��ʼ/����ʱ���������������롣"
        Exit Function
    End If
    
    If Not IsNull(dtpStart.Value) Then
        If dtpStart.Value > dtpEnd.Value Then
            ValidateTxt = "��ʼʱ�䲻�ܴ��ڽ���ʱ�䣬���������롣"
            Exit Function
        End If
    End If
    
End Function

Private Function GetSelectData() As String
'����:��ȡ��ѡ��ѡ��������,���ض�ӦID,���򷵻ؿ�ֵ
    Dim i As Integer, strIDs As String
    
    With vsfLimit
        If .Rows = .FixedRows Then Exit Function
        
        '����Ƿ���ѡ������
        For i = .FixedRows To .Rows - .FixedRows
            If .TextMatrix(i, 0) = "-1" Then
                If strIDs = "" Then
                    strIDs = .RowData(i)
                Else
                    strIDs = strIDs & "," & .RowData(i)
                End If
            End If
        Next
        
        If strIDs = "" Then
            '��Ϊ��ֵ,��ȡ��ǰѡ��������
            GetSelectData = .RowData(.Row)
        Else
            GetSelectData = strIDs
        End If
    End With
End Function

Private Sub EditLimit(Optional ByVal strStop As String)
'����:�޸Ĺ���
    Dim strTmp As String, i As Integer
    Dim strID As String, strStartIP As String, strEndIp As String
    Dim strStartTime As String, strEndTime As String
    Dim strDesc As String, strUser As String
    Dim varIDs As Variant
    
    On Error GoTo errh
    'У������
    Call GetDataFromCard(strUser, strStartIP, strEndIp, strStartTime, strEndTime, strDesc)
    strID = vsfLimit.RowData(vsfLimit.Row)
    strTmp = ValidateTxt
    If strTmp <> "" Then
        frmMDIMain.stbThis.Panels(2).Text = strTmp
        Exit Sub
    End If
    
    '�ύ����
    If strStop = "" Then
        '˵��û�д���ͣ�ò���,��������
        strStop = IIf(vsfLimit.TextMatrix(vsfLimit.Row, vsfLimit.ColIndex("״̬")) = "������", 1, 0)
    Else
        strStop = IIf(vsfLimit.TextMatrix(vsfLimit.Row, vsfLimit.ColIndex("״̬")) = "������", 0, 1)
    End If
    Call ExecuteProcedure("Zl_Zlloginlimit_Edit(2,'" & strID & "'," & strStartTime & "," & strEndTime & ",'" & strStartIP & "','" & strEndIp & "','" & strStop & "','" & strDesc & "','" & strUser & "')", Me.Caption)
    
    cmdStop.Caption = IIf(strStop = 0, "����", "ͣ��")
    With vsfLimit
        .Redraw = flexRDNone
        Call LoadLimit
    End With
    frmMDIMain.stbThis.Panels(2).Text = "�޸Ĺ���ɹ���"
    Exit Sub
errh:
    frmMDIMain.stbThis.Panels(2).Text = ""
    
    If InStr(1, UCase(err.Description), "ORA") > 0 Then '���ݿ����,�ַ����ϳ�,������ʾ
        MsgBox "�޸Ĺ���ʧ�ܡ�ԭ��" & vbNewLine & err.Description
    Else
        frmMDIMain.stbThis.Panels(2).Text = "�޸Ĺ���ʧ�ܡ�ԭ��" & vbNewLine & err.Description
    End If
End Sub

Private Sub GetDataFromCard(ByRef strUser As String, ByRef strStartIP As String, ByRef strEndIp As String, _
                                                        ByRef strStartTime As String, ByRef strEndTime As String, ByRef strDesc As String)
'����:�ӿ�Ƭ���л�ȡ����
    
    strUser = Trim(txtUser.Text)
    
    If txtbeforeIp(0).Text = "" Then
        strStartIP = "": strEndIp = ""
    Else
        strStartIP = txtbeforeIp(0).Text & "." & txtbeforeIp(1).Text & "." & txtbeforeIp(2).Text & "." & txtbeforeIp(3).Text
        strEndIp = txtbeforeIp(0).Text & "." & txtbeforeIp(1).Text & "." & txtbeforeIp(2).Text & "." & IIf(txtIP(2).Text = "", txtbeforeIp(3).Text, txtIP(2).Text)
    End If
    
    strStartTime = IIf(IsNull(dtpStart.Value), "''", "To_Date('" & dtpStart.Value & "','yyyy/mm/dd hh24:mi:ss')")
    strEndTime = IIf(IsNull(dtpEnd.Value), "''", "To_Date('" & dtpEnd.Value & "','yyyy/mm/dd hh24:mi:ss')")
    strDesc = txtDesc.Text
    
End Sub


Private Sub GetCursorPos(ByVal hwnd5 As Long, LineNo As Long, ColNo As Long)
    Dim i As Long, j As Long
    Dim lParam As Long, wParam As Long
    Dim K As Long
    
    i = SendMessage(hwnd5, EM_GETSEL, wParam, lParam)
    j = i / 2 ^ 16 'ȡ��Ŀǰ�������λ��ǰ�ж��ٸ�Byte
    LineNo = SendMessage(hwnd5, EM_LINEFROMCHAR, j, 0) 'ȡ�ù��ǰ���ж�����
    LineNo = LineNo + 1
    K = SendMessage(hwnd5, EM_LINEINDEX, -1, 0)
    'ȡ��Ŀǰ���������ǰ���ж��ٸ�Byte
    ColNo = j - K + 1
End Sub



