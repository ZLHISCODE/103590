VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���Ӱ��¼-�༭"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7380
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   7380
   StartUpPosition =   1  '����������
   Begin VB.Frame fraEdit 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      Begin VB.Frame fraSplit1 
         Height          =   30
         Left            =   0
         TabIndex        =   24
         Top             =   4080
         Width           =   6855
      End
      Begin VB.CommandButton cmdIn 
         Caption         =   "��"
         Height          =   290
         Left            =   3225
         TabIndex        =   21
         Top             =   2300
         Width           =   255
      End
      Begin VB.ComboBox cboType 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1082
         Width           =   1935
      End
      Begin VB.ComboBox cboHoldType 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2736
         Width           =   1935
      End
      Begin VB.ComboBox cboDept 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtPer 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1560
         TabIndex        =   2
         Top             =   661
         Width           =   1935
      End
      Begin VB.TextBox txtHold 
         Height          =   300
         Left            =   1560
         TabIndex        =   4
         Top             =   2315
         Width           =   1935
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   4440
         TabIndex        =   6
         Top             =   4320
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   5760
         TabIndex        =   7
         Top             =   4320
         Width           =   1100
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   285
         Left            =   1560
         TabIndex        =   13
         Top             =   1500
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   243662851
         CurrentDate     =   401769
         MaxDate         =   402133
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   1905
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   243662851
         CurrentDate     =   401769
         MaxDate         =   402133
      End
      Begin MSComCtl2.DTPicker dtpHoldBegin 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Top             =   3150
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   243662851
         CurrentDate     =   401769
         MaxDate         =   402133
      End
      Begin MSComCtl2.DTPicker dtpHoldEnd 
         Height          =   285
         Left            =   1560
         TabIndex        =   18
         Top             =   3570
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   243662851
         CurrentDate     =   401769
         MaxDate         =   402133
      End
      Begin MSComctlLib.TreeView tvwDoc 
         Height          =   1575
         Left            =   1080
         TabIndex        =   22
         Top             =   4200
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   2778
         _Version        =   393217
         Indentation     =   353
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgList"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfPaiType 
         Height          =   3375
         Left            =   3840
         TabIndex        =   25
         Top             =   480
         Width           =   3075
         _cx             =   5433
         _cy             =   5953
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
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   11
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmEdit.frx":6852
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
         Editable        =   2
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
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�Զ��������²���"
         Height          =   180
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label lblInEnd 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�Ӱ����ʱ��"
         Height          =   180
         Left            =   360
         TabIndex        =   20
         Top             =   3600
         Width           =   1080
      End
      Begin VB.Label lblInBegin 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�Ӱ࿪ʼʱ��"
         Height          =   180
         Left            =   360
         TabIndex        =   19
         Top             =   3180
         Width           =   1080
      End
      Begin VB.Label lblEnd 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�������ʱ��"
         Height          =   180
         Left            =   360
         TabIndex        =   16
         Top             =   1950
         Width           =   1080
      End
      Begin VB.Label lblBegin 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "���࿪ʼʱ��"
         Height          =   180
         Left            =   360
         TabIndex        =   15
         Top             =   1545
         Width           =   1080
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   1080
         TabIndex        =   12
         Top             =   315
         Width           =   360
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�Ӱ���"
         Height          =   180
         Index           =   3
         Left            =   720
         TabIndex        =   11
         Top             =   2775
         Width           =   720
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�Ӱ�ҽ��"
         Height          =   180
         Index           =   4
         Left            =   720
         TabIndex        =   10
         Top             =   2370
         Width           =   720
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "������"
         Height          =   180
         Index           =   6
         Left            =   720
         TabIndex        =   9
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "����ҽ��"
         Height          =   180
         Index           =   7
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   720
      End
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   8760
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":69C1
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":6F5B
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":74F5
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":DD57
            Key             =   "Up"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":145B9
            Key             =   "Down"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":1AE1B
            Key             =   "add"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":2167D
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEdit.frx":2208F
            Key             =   "Dept"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngId As Long
Private mrsTime As ADODB.Recordset
Private mstrDeptId As String
Private mblnOk As Boolean
Private mrsDoc As ADODB.Recordset
Private mrsPati As ADODB.Recordset
Private mbytType As Byte '0-������1-�޸�
Private mstrDept As String '���Ҵ���
Private mstrOutPer As String '����������
Private mstrOutTime As String '����ʱ�䷶Χ
Private mstrInPer As String '�Ӱ�������
Private mstrInTime As String '�Ӱ�ʱ�䷶Χ

Public Function ShowMe(ByVal bytType As Byte, Optional ByVal lngId As Long, Optional strDept As String, Optional strOutPer As String, Optional strOutTime As String, _
                Optional strInPer As String, Optional strInTime As String) As Boolean
'bytType:0-�������Ӱ��¼��1-�޸Ľ��Ӱ��¼
'strDept��ʽ-������ѡ���������|���������1|���������2...
    mstrDeptId = ""
    mlngId = lngId
    mbytType = bytType
    mstrDept = strDept
    mstrOutPer = strOutPer
    mstrOutTime = strOutTime
    mstrInPer = strInPer
    mstrInTime = strInTime
    
    Me.Show 1
    ShowMe = mblnOk
End Function

Private Sub SetBasic()
    Dim strTemp As String
    Dim varTemp As Variant, varData As Variant
    Dim i As Long, lngTemp As Long
    Dim rsTemp As ADODB.Recordset
    
    Select Case mbytType
        Case 0
            Me.Caption = "���Ӱ��¼-����"
            Me.Width = fraEdit.Width + 100
            Me.Height = fraEdit.Height + 350
            fraEdit.Visible = True
            fraEdit.Move 0, 0
            Set rsTemp = GetPatientType
            With vsfPaiType
                .Rows = 1
                .Rows = rsTemp.RecordCount + 1
                Do While Not rsTemp.EOF
                    .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("���")) = rsTemp!���
                    .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("��������")) = rsTemp!����
                    .TextMatrix(rsTemp.AbsolutePosition, .ColIndex("��ȡSQL")) = rsTemp!��ȡSQL & ""
                    rsTemp.MoveNext
                Loop
                For i = 0 To .Rows - 1
                    .Cell(flexcpChecked, i, 0) = flexChecked
                Next
                If .Rows > 11 Then
                    .ColWidth(.ColIndex("��������")) = 1800
                Else
                    .ColWidth(.ColIndex("��������")) = 2055
                End If
            End With
            Call vsfPaiType_AfterRowColChange(1, 1, 0, 1)
        Case 1
            Me.Caption = "���Ӱ��¼-�޸�"
            cboDept.Enabled = False
            lblType.Visible = False
            vsfPaiType.Visible = False
            cboType.Enabled = False
            Me.Width = fraEdit.Width - vsfPaiType.Width
            Me.Height = fraEdit.Height + 350
            cmdCancel.Left = dtpHoldEnd.Left + dtpHoldEnd.Width - cmdCancel.Width
            cmdOK.Left = cmdCancel.Left - 200 - cmdOK.Width
            fraEdit.Visible = True
            fraEdit.Move 0, 0
    End Select
End Sub

Private Sub cboDept_Click()
    Dim strDeptID As Long
    Dim rsTemp As ADODB.Recordset
    
    strDeptID = cboDept.ItemData(cboDept.ListIndex)
    Set rsTemp = GetShiftType(2, strDeptID)
    Set mrsTime = GetShiftType(1, strDeptID)
    cboType.Clear
    cboHoldType.Clear
    Do While Not rsTemp.EOF
        cboType.AddItem rsTemp!�������
        cboHoldType.AddItem rsTemp!�������
        rsTemp.MoveNext
    Loop
End Sub

Private Sub cboHoldType_Change()
    If cboHoldType.Text = "" Then
        dtpHoldBegin.Value = "3000/1/1"
        dtpHoldEnd.Value = "3000/1/1"
    End If
End Sub

Private Sub cboHoldType_Click()
    Dim objDate As Date
    
    objDate = Format(IIf(cboType.Text = "", zlDatabase.Currentdate, dtpEnd.Value), "yyyy-mm-dd")
    mrsTime.Filter = "�������='" & cboHoldType.Text & "'"
    If mrsTime.RecordCount = 1 Then
        dtpHoldBegin.Value = objDate & " " & mrsTime!��ʼʱ��
        dtpHoldEnd.Value = IIf(mrsTime!��ʼʱ�� >= mrsTime!����ʱ��, objDate + 1, objDate) & " " & mrsTime!����ʱ��
    End If
End Sub

Private Sub cboType_Change()
    If cboType = "" Then
        dtpBegin.Enabled = False
        dtpBegin.Value = "3000/1/1"
        dtpEnd.Value = "3000/1/1"
    Else
        dtpBegin.Enabled = True
    End If
End Sub

Private Sub cboType_Click()
'������ѡ����Զ���ʾ���࿪ʼʱ��ͽ������ʱ�䣬���࿪ʼʱ��ɵ���
    Dim objDate As Date
    
    objDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    mrsTime.Filter = "�������='" & cboType.Text & "'"
    If mrsTime.RecordCount = 1 Then
        If mrsTime!��ʼʱ�� >= mrsTime!����ʱ�� Then
            dtpBegin.Value = objDate - 1 & " " & mrsTime!��ʼʱ��
            dtpEnd.Value = objDate & " " & mrsTime!����ʱ��
        Else
            dtpBegin.Value = objDate & " " & mrsTime!��ʼʱ��
            dtpEnd.Value = objDate & " " & mrsTime!����ʱ��
        End If
        If mrsTime!��ʼʱ�� = mrsTime!����ʱ�� Then
            cboHoldType.Text = cboType.Text
            Call cboHoldType_Click
        Else
            mrsTime.Filter = "��ʼʱ��='" & mrsTime!����ʱ�� & "'"
            If mrsTime.RecordCount > 0 Then
                cboHoldType.Text = mrsTime!�������
                Call cboHoldType_Click
            End If
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdIn_Click()

    mrsDoc.Filter = ""
    If mrsDoc.RecordCount = 0 Then Exit Sub
    tvwDoc.Visible = True
    tvwDoc.SetFocus
End Sub

Private Sub ShowDoc()
'����ҽ����Ϣ������
    Dim strDept As String
    Dim objNode As Object
    
    On Error GoTo errH
    If mbytType = 0 Then
        gstrSQL = "Select b.����id, c.����, a.Id,a.���, a.����,a.���� From ��Ա�� a, ������Ա b, ���ű� c" & vbNewLine & _
            "Where a.Id = b.��Աid And b.ȱʡ = 1 And b.����id In(Select * From Table(f_str2list([1]))) And b.����id = c.Id" & vbNewLine & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) Order By ����id, Id"
        Set mrsDoc = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrDeptId)
    Else
        gstrSQL = "Select b.����id, c.����, a.Id, a.���, a.����,a.���� " & vbNewLine & _
            "From ��Ա�� a, ������Ա b, ���ű� c" & vbNewLine & _
            "Where a.Id = b.��Աid And b.ȱʡ = 1 And" & vbNewLine & _
            "      b.����id In (Select ����id From �ٴ����� Where �������� In (Select �������� From �ٴ����� Where ����id =[1])) And b.����id = c.Id" & vbNewLine & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) Order By ����id, Id"
        Set mrsDoc = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrDeptId))
    End If
    If mrsDoc.RecordCount = 0 Then Exit Sub
    
    With tvwDoc
        .Nodes.Clear
        Do While Not mrsDoc.EOF
            If strDept <> mrsDoc!���� Then
                '���ź���Ա��id�����ظ����ʹؼ�����id������һ��
                Set objNode = .Nodes.Add(, , "K" & mrsDoc!����id & mrsDoc!����, mrsDoc!����, "Dept")
                Set objNode = .Nodes.Add("K" & mrsDoc!����id & mrsDoc!����, tvwChild, "K" & mrsDoc!id, mrsDoc!����, "Person")
                strDept = mrsDoc!����
            Else
                Set objNode = .Nodes.Add("K" & mrsDoc!����id & mrsDoc!����, tvwChild, "K" & mrsDoc!id, mrsDoc!����, "Person")
            End If
            mrsDoc.MoveNext
        Loop
    End With
    tvwDoc.Left = txtHold.Left
    tvwDoc.Top = txtHold.Top + txtHold.Height
    tvwDoc.ZOrder 0
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub cmdOK_Click()
    Dim arrTemp As Variant, arrSQL As Variant
    Dim i As Long, lngId As Long
    Dim blnBegin As Boolean
        
    If CheckRecordData = False Then Exit Sub
    gstrSQL = ""
    arrTemp = Array()
    arrSQL = Array()
    If mbytType = 0 Then
        '����������һ��ҽ��ֵ��������
        If cboDept.Text = "���п���" Then
            For i = 1 To cboDept.ListCount - 1
                ReDim Preserve arrTemp(UBound(arrTemp) + 1)
                arrTemp(UBound(arrTemp)) = cboDept.ItemData(i) & ",'" & txtPer.Text & "','" & cboType.Text & "'," & _
                zlStr.To_Date(dtpBegin.Value) & "," & zlStr.To_Date(dtpEnd.Value) & ",'" & _
                txtHold.Text & "','" & cboHoldType.Text & "'," & _
                zlStr.To_Date(dtpHoldBegin.Value) & "," & zlStr.To_Date(dtpHoldEnd.Value)
            Next
        Else
            ReDim Preserve arrTemp(UBound(arrTemp) + 1)
            arrTemp(UBound(arrTemp)) = cboDept.ItemData(cboDept.ListIndex) & ",'" & txtPer.Text & "','" & cboType.Text & "'," & _
            zlStr.To_Date(dtpBegin.Value) & "," & zlStr.To_Date(dtpEnd.Value) & ",'" & _
            txtHold.Text & "','" & cboHoldType.Text & "'," & _
            zlStr.To_Date(dtpHoldBegin.Value) & "," & zlStr.To_Date(dtpHoldEnd.Value)
        End If
        Set mrsPati = GetTimeRangePati(dtpBegin.Value, dtpEnd.Value, mstrDeptId)
        For i = LBound(arrTemp) To UBound(arrTemp)
            lngId = GetNextId("ҽ�����Ӱ��¼", "��¼ID")
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ��¼_Edit(0," & lngId & "," & arrTemp(i) & ",'" & grsUserInfo!���� & "')"
            Call SavePatiData(arrSQL, lngId, Mid(arrTemp(i), 1, InStr(arrTemp(i), ",") - 1))
        Next
    Else '�޸�
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ��¼_Edit(1," & mlngId & "," & Val(mstrDeptId) & ",'" & txtPer.Text & "','" & cboType.Text & "'," & _
            zlStr.To_Date(dtpBegin.Value) & "," & zlStr.To_Date(dtpEnd.Value) & ",'" & _
            txtHold.Text & "','" & cboHoldType.Text & "'," & _
            zlStr.To_Date(dtpHoldBegin.Value) & "," & zlStr.To_Date(dtpHoldEnd.Value) & ",'" & grsUserInfo!���� & "')"
    End If
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    blnBegin = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans
    mblnOk = True
    Unload Me
    Exit Sub
ErrHand:
    If blnBegin Then gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub SavePatiData(arrSQL As Variant, ByVal lngRecordId As Long, ByVal lngDeptId As Long)
'���ݹ�ѡ���潻�Ӱ����������Լ���������
    Dim rsTemp As ADODB.Recordset
    Dim strType As String, strTypes As String, strPsiId As String, str��� As String, str���� As String
    Dim i As Long, lng���� As Long, lng��Ժ As Long, lng������ As Long
    
    '������¼ʱ�Զ�������ܱ�����
    Set rsTemp = GetPatiType
    Do While Not rsTemp.EOF
        mrsPati.Filter = "��Ժ����id=" & lngDeptId & " And ����='" & rsTemp!��� & "'"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ����_Insert(" & lngRecordId & "," & rsTemp!˳�� & ",'" & rsTemp!��� & "'," & mrsPati.RecordCount & ")"
        rsTemp.MoveNext
    Loop
    
    'סԺ������
    gstrSQL = "Select Count(*) ���� From ��Ժ���� Where ����id =[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDeptId)
    lng������ = rsTemp!����
    If DateDiff("s", dtpEnd.Value, zlDatabase.Currentdate) <= 0 Then
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ����_Insert(" & lngRecordId & ",99,'סԺ��'," & rsTemp!���� & ")"
    Else
        gstrSQL = "Select count(*) ���� " & vbNewLine & _
            "From ������ҳ a" & vbNewLine & _
            "Where a.��Ժ���� > " & zlStr.To_Date(dtpEnd.Value) & " And" & vbNewLine & _
            "      a.��Ժ���� <=sysdate And a.��Ժ���� Is Null and a.��Ժ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDeptId)
        lng���� = rsTemp!����
        gstrSQL = "Select count(*) ����" & vbNewLine & _
            "From ������ҳ a" & vbNewLine & _
            "Where a.��Ժ���� > " & zlStr.To_Date(dtpEnd.Value) & " And" & vbNewLine & _
            "      a.��Ժ���� <=sysdate and a.��Ժ����id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngDeptId)
        lng��Ժ = rsTemp!����
        lng������ = lng������ - lng���� + lng��Ժ
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ����_Insert(" & lngRecordId & ",99,'סԺ������'," & IIf(lng������ > 0, lng������, 0) & ")"
    End If
    
    mrsPati.Filter = ""
    If mrsPati.RecordCount > 0 Then
        Set rsTemp = zlDatabase.CopyNewRec(mrsPati)
        With vsfPaiType
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("ѡ��")) = flexChecked Then
                    mrsPati.Filter = "��Ժ����id=" & lngDeptId & " And ����='" & .TextMatrix(i, .ColIndex("���")) & "'"
                    Do While Not mrsPati.EOF
                        'һ������������ڶ������ͣ��������������ַ���ƴ����(���)
                        'һ������ֻ�ܼ���һ�������У����ձ��˳�����е�
                        strType = mrsPati!����
                        rsTemp.Filter = "��Ժ����id=" & lngDeptId & " And ����id=" & mrsPati!����ID & " And ����<>" & "'" & .TextMatrix(i, .ColIndex("���")) & "'"
                        Do While Not rsTemp.EOF
                            strType = strType & "," & rsTemp!����
                            rsTemp.MoveNext
                        Loop
                        If InStr(strPsiId & ",", "," & mrsPati!����ID & ",") = 0 Then
                            strPsiId = strPsiId & "," & mrsPati!����ID
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = "Zl_ҽ�����Ӱ�����_Edit(0,0," & lngRecordId & ",0,'" & strType & "'," & mrsPati!����ID & "," & NVL(mrsPati!��ҳID, 0) & ",'" & _
                                mrsPati!���� & "','" & mrsPati!�Ա� & "','" & mrsPati!���� & "','" & mrsPati!���� & "" & "'," & NVL(mrsPati!��ʶ��, "Null") & _
                                "," & zlStr.To_Date(mrsPati!��Ժʱ��) & ",'" & _
                                mrsPati!��Ժ��ʽ & "')"
                        End If
                        mrsPati.MoveNext
                    Loop
                End If
            Next
        End With
    End If
End Sub

Private Function CheckRecordData() As Boolean
'���Ӱ��¼���ݵļ��
    
    If cboType.Text = "" Then MsgBox "�����β���Ϊ�գ���ѡ��", vbExclamation, Me.Caption: Call zlControl.ControlSetFocus(txtPer): Exit Function
    If txtHold.Text = "" Then MsgBox "�Ӱ�ҽ������Ϊ�գ�����д��", vbExclamation, Me.Caption: Call zlControl.ControlSetFocus(txtHold): Exit Function
    If cboHoldType.Text = "" Then MsgBox "�Ӱ��β���Ϊ�գ���ѡ��", vbExclamation, Me.Caption: Call zlControl.ControlSetFocus(cboHoldType): Exit Function
    If dtpEnd.Value <> dtpHoldBegin.Value Then MsgBox "�������ʱ����Ӱ࿪ʼʱ�䲻һ�£�����!", vbExclamation, Me.Caption: Call zlControl.ControlSetFocus(cboHoldType): Exit Function
    mrsDoc.Filter = "����='" & txtHold.Text & "'"
    If mrsDoc.RecordCount = 0 Then
        MsgBox "�Ӱ�ҽ�������ڵ�ǰ�������ң�������ѡ��", vbExclamation, Me.Caption
        Exit Function
    End If
    CheckRecordData = True
End Function

Private Sub dtpBegin_CloseUp()
    
    mrsTime.Filter = "�������='" & cboType.Text & "'"
    If mrsTime.RecordCount = 1 Then
        If mrsTime!��ʼʱ�� >= mrsTime!����ʱ�� Then
            dtpEnd.Value = Format(dtpBegin.Value + 1, "yyyy-mm-dd") & " " & mrsTime!����ʱ��
        Else
            dtpEnd.Value = Format(dtpBegin.Value, "yyyy-mm-dd") & " " & mrsTime!����ʱ��
        End If
        Call cboHoldType_Click
    End If
End Sub

Private Sub dtpEnd_CloseUp()
    Call cboHoldType_Click
End Sub

Private Sub Form_Load()
    Dim varTemp As Variant, varData As Variant
    Dim i As Long

    Call SetBasic
    varTemp = Split(mstrDept, "|")
    Select Case mbytType
        Case 0
            '����ʱ������������Ŀ���һ��
            For i = 1 To UBound(varTemp)
                varData = Split(varTemp(i), ",")
                cboDept.AddItem varData(0)
                cboDept.ItemData(cboDept.NewIndex) = varData(1)
                mstrDeptId = IIf(mstrDeptId = "", "", mstrDeptId & ",") & varData(1)
            Next
            cboDept.ListIndex = IIf(varTemp(0) < 0, 0, varTemp(0))
        Case 1
            cboDept.AddItem varTemp(0)
            cboDept.ListIndex = 0
            mstrDeptId = varTemp(1)
    End Select
    '�����ˡ������Ρ�����ʱ��
    txtPer.Text = mstrOutPer
    varTemp = Split(mstrOutTime, "|")
    If UBound(varTemp) > 0 Then
        cboType.AddItem varTemp(0)
        cboType.ListIndex = 0
        dtpBegin.Value = varTemp(1)
        dtpEnd.Value = varTemp(2)
    End If
    '�Ӱ��ˡ��Ӱ��Ρ��Ӱ�ʱ��
    txtHold.Text = mstrInPer
    varTemp = Split(mstrInTime, "|")
    If UBound(varTemp) > 0 Then
        cboHoldType.AddItem varTemp(0)
        cboHoldType.ListIndex = 0
        dtpHoldBegin.Value = varTemp(1)
        dtpHoldEnd.Value = varTemp(2)
    End If
    Call ShowDoc
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsTime = Nothing
    Set mrsDoc = Nothing
    Set mrsPati = Nothing
End Sub

Private Sub tvwDoc_DblClick()
    
    If Not tvwDoc.SelectedItem.Parent Is Nothing Then
        txtHold.Text = tvwDoc.SelectedItem.Text
        Call tvwDoc_LostFocus
    End If
End Sub

Private Sub tvwDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call tvwDoc_LostFocus
    End If
End Sub

Private Sub tvwDoc_LostFocus()
    tvwDoc.Visible = False
End Sub

Private Sub txtHold_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strTemp As String
        
    If KeyCode = vbKeyReturn Then
        strTemp = UCase(txtHold.Text)
        mrsDoc.MoveFirst
        Do While Not mrsDoc.EOF
            If InStr(mrsDoc!����, strTemp) > 0 Or InStr(mrsDoc!����, strTemp) > 0 Then
                txtHold.Text = mrsDoc!����
                Exit Do
            End If
            mrsDoc.MoveNext
        Loop
    End If
End Sub

Private Sub txtHold_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("%") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txtHold_Validate(Cancel As Boolean)

    '��ȡ��Ҳ�ᴥ������¼������ڱ����ʱ����
'    If txtHold.Text = "" Then Exit Sub
'    mrsDoc.Filter = "����='" & txtHold.Text & "'"
'    If mrsDoc.RecordCount = 0 Then
'        MsgBox "�Ӱ�ҽ�������ڵ�ǰ�������ң�������ѡ��"
'    End If
End Sub

Private Sub vsfPaiType_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    With vsfPaiType
        If Col = 0 Then
            'ʵ��ȫѡ��ȫ���Ĺ���
            If Row = 0 Then
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    .Cell(flexcpChecked, 0, 0) = flexChecked
                    For i = .FixedRows To .Rows - .FixedRows
                        .Cell(flexcpChecked, i, 0) = flexChecked
                    Next
                Else
                    .Cell(flexcpChecked, 0, 0) = flexUnchecked
                    For i = .FixedRows To .Rows - .FixedRows
                        .Cell(flexcpChecked, i, 0) = flexUnchecked
                    Next
                End If
            Else
                If .Cell(flexcpChecked, 0, 0) = flexChecked Then
                    .Cell(flexcpChecked, 0, 0) = flexUnchecked
                End If
                For i = .FixedRows To .Rows - .FixedRows
                    If .Cell(flexcpChecked, i, 0) = flexUnchecked Then: Exit For
                    If i = .Rows - .FixedRows Then
                        .Cell(flexcpChecked, 0, 0) = flexChecked
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub vsfPaiType_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Or NewRow < 1 Then Exit Sub
    With vsfPaiType
        If NewRow = 1 Then
            .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
            .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Down").Picture
        Else
            If NewRow = .Rows - 1 Then
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = ""
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Up").Picture
            Else
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Up").Picture
                .Cell(flexcpPicture, NewRow, .ColIndex("����")) = imgList.ListImages("Down").Picture
            End If
        End If
        If OldRow < .Rows Then
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
            .Cell(flexcpPicture, OldRow, .ColIndex("����")) = ""
        End If
    End With
End Sub

Private Sub vsfPaiType_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPaiType
        If Col <> .ColIndex("ѡ��") Then Cancel = True
    End With
End Sub

Private Sub vsfPaiType_Click()
    Dim lngCheck As Long, lngNum As Long, lngRow As Long
    Dim strPati As String, strName As String, strSQL As String
    
    With vsfPaiType
        If .Row < 1 Then Exit Sub
        If .Col = .ColIndex("����") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                lngRow = .Row - 1
            End If
        ElseIf .Col = .ColIndex("����") Then
            If Not .Cell(flexcpPicture, .Row, .ColIndex("����")) Is Nothing Then
                lngRow = .Row + 1
            End If
        End If
        If lngRow = 0 Then Exit Sub
        lngCheck = .Cell(flexcpChecked, .Row, .ColIndex("ѡ��"))
        strPati = .TextMatrix(.Row, .ColIndex("��������"))
        strName = .TextMatrix(.Row, .ColIndex("���"))
        strSQL = .TextMatrix(.Row, .ColIndex("��ȡSQL"))
        
        .Cell(flexcpChecked, .Row, .ColIndex("ѡ��")) = .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��"))
        .TextMatrix(.Row, .ColIndex("��������")) = .TextMatrix(lngRow, .ColIndex("��������"))
        .TextMatrix(.Row, .ColIndex("���")) = .TextMatrix(lngRow, .ColIndex("���"))
        .TextMatrix(.Row, .ColIndex("��ȡSQL")) = .TextMatrix(lngRow, .ColIndex("��ȡSQL"))
        .Cell(flexcpChecked, lngRow, .ColIndex("ѡ��")) = lngCheck
        .TextMatrix(lngRow, .ColIndex("��������")) = strPati
        .TextMatrix(lngRow, .ColIndex("���")) = strName
        .TextMatrix(lngRow, .ColIndex("��ȡSQL")) = strSQL
        .Row = lngRow
        .ShowCell lngRow, 1
    End With
    
End Sub


