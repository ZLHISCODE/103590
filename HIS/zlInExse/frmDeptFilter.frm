VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeptFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   ControlBox      =   0   'False
   Icon            =   "frmDeptFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdDef 
      Caption         =   "ȱʡ(&D)"
      Height          =   350
      Left            =   5940
      TabIndex        =   26
      Top             =   1755
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5940
      TabIndex        =   24
      Top             =   810
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5940
      TabIndex        =   23
      Top             =   390
      Width           =   1100
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frmDeptFilter.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpB"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dtpE"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "�շ���Ŀ(&1)"
      TabPicture(1)   =   "frmDeptFilter.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl������Ŀ(0)"
      Tab(1).Control(1)=   "ListFeeItem(0)"
      Tab(1).Control(2)=   "tlbOpt(0)"
      Tab(1).Control(3)=   "txtInput(0)"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "������Ŀ(&2)"
      TabPicture(2)   =   "frmDeptFilter.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl������Ŀ(1)"
      Tab(2).Control(1)=   "tlbOpt(1)"
      Tab(2).Control(2)=   "txtInput(1)"
      Tab(2).Control(3)=   "ListFeeItem(1)"
      Tab(2).ControlCount=   4
      Begin VB.ListBox ListFeeItem 
         Height          =   1740
         Index           =   1
         Left            =   -73680
         Style           =   1  'Checkbox
         TabIndex        =   22
         ToolTipText     =   "Ctrl+Aȫѡ,Ctrl+Cȫ��,���һ����δѡ���ʾ������"
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   21
         ToolTipText     =   "���ƥ��100���������"
         Top             =   480
         Width           =   2160
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   -73680
         MaxLength       =   40
         TabIndex        =   18
         ToolTipText     =   "���ƥ��100���������"
         Top             =   480
         Width           =   2160
      End
      Begin MSComctlLib.Toolbar tlbOpt 
         Height          =   600
         Index           =   0
         Left            =   -74760
         TabIndex        =   27
         Top             =   1080
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1058
         ButtonWidth     =   1614
         ButtonHeight    =   1058
         Style           =   1
         ImageList       =   "ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�Ƴ�(&M)"
               Key             =   "Delete"
               Object.ToolTipText     =   "�Ƴ���ǰѡ����б���"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���(&T)"
               Key             =   "Clear"
               Object.ToolTipText     =   "����б���Ŀ"
               ImageKey        =   "Clear"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����(&S)"
               Key             =   "Save"
               Object.ToolTipText     =   "����ѡ����б���Ŀ"
               ImageKey        =   "Save"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbOpt 
         Height          =   600
         Index           =   1
         Left            =   -74760
         TabIndex        =   28
         Top             =   1080
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1058
         ButtonWidth     =   1614
         ButtonHeight    =   1058
         Style           =   1
         ImageList       =   "ils16"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�Ƴ�(&M)"
               Key             =   "Delete"
               Object.ToolTipText     =   "�Ƴ���ǰѡ����б���"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���(&T)"
               Key             =   "Clear"
               Object.ToolTipText     =   "����б���Ŀ"
               ImageKey        =   "Clear"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����(&S)"
               Key             =   "Save"
               Object.ToolTipText     =   "����ѡ����б���Ŀ"
               ImageKey        =   "Save"
            EndProperty
         EndProperty
      End
      Begin VB.ListBox ListFeeItem 
         Height          =   1740
         Index           =   0
         Left            =   -73680
         Style           =   1  'Checkbox
         TabIndex        =   19
         ToolTipText     =   "Ctrl+Aȫѡ,Ctrl+Cȫ��,���һ����δѡ���ʾ������"
         Top             =   960
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   300
         Left            =   3525
         TabIndex        =   1
         Top             =   600
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   107216899
         CurrentDate     =   37068
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   300
         Left            =   1080
         TabIndex        =   0
         Top             =   600
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   107216899
         CurrentDate     =   37068
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   2580
         Left            =   105
         TabIndex        =   29
         Top             =   885
         Width           =   5520
         Begin VB.TextBox txtHospitalNO 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   975
            MaxLength       =   18
            TabIndex        =   14
            Top             =   1755
            Width           =   2025
         End
         Begin VB.TextBox txtName 
            Height          =   300
            IMEMode         =   1  'ON
            Left            =   960
            MaxLength       =   100
            TabIndex        =   16
            Top             =   2130
            Width           =   4515
         End
         Begin VB.TextBox txtNoEnd 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3420
            MaxLength       =   8
            TabIndex        =   12
            Top             =   1380
            Width           =   2055
         End
         Begin VB.TextBox txtNOBegin 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   975
            MaxLength       =   8
            TabIndex        =   11
            Top             =   1380
            Width           =   2055
         End
         Begin VB.ComboBox cbo����Ա 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   990
            Width           =   2055
         End
         Begin VB.CheckBox chkType 
            Caption         =   "���ʵ���"
            ForeColor       =   &H000000C0&
            Height          =   255
            Index           =   1
            Left            =   4260
            TabIndex        =   5
            Top             =   165
            Width           =   1095
         End
         Begin VB.CheckBox chkType 
            Caption         =   "���ʵ���"
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   3120
            TabIndex        =   4
            Top             =   165
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.CheckBox chkBill 
            Caption         =   "��ͨ����"
            Height          =   195
            Index           =   0
            Left            =   3120
            TabIndex        =   6
            Top             =   630
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chkBill 
            Caption         =   "��������"
            Height          =   195
            Index           =   2
            Left            =   3120
            TabIndex        =   8
            Top             =   1035
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chkBill 
            Caption         =   "�Զ�����"
            Height          =   210
            Index           =   1
            Left            =   4260
            TabIndex        =   7
            Top             =   615
            Width           =   1020
         End
         Begin VB.CheckBox chkBill 
            Caption         =   "��������"
            Height          =   210
            Index           =   3
            Left            =   4245
            TabIndex        =   9
            Top             =   1020
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   975
            TabIndex        =   3
            Top             =   570
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   107216899
            CurrentDate     =   36588
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   975
            TabIndex        =   2
            Top             =   150
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   529
            _Version        =   393216
            CalendarTitleBackColor=   -2147483647
            CalendarTitleForeColor=   -2147483634
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   107216899
            CurrentDate     =   36588
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��"
            Height          =   180
            Left            =   360
            TabIndex        =   13
            Top             =   1815
            Width           =   540
         End
         Begin VB.Label lblName 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Left            =   480
            TabIndex        =   15
            Top             =   2205
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ݺ�"
            Height          =   180
            Left            =   360
            TabIndex        =   34
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   3120
            TabIndex        =   33
            Top             =   1440
            Width           =   180
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Left            =   180
            TabIndex        =   32
            Top             =   630
            Width           =   720
         End
         Begin VB.Label lblDate 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼʱ��"
            Height          =   180
            Left            =   180
            TabIndex        =   31
            Top             =   210
            Width           =   720
         End
         Begin VB.Label lbl����Ա 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����Ա"
            Height          =   180
            Left            =   360
            TabIndex        =   30
            Top             =   1050
            Width           =   540
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3240
         TabIndex        =   36
         Top             =   660
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   300
         TabIndex        =   35
         Top             =   660
         Width           =   720
      End
      Begin VB.Label lbl������Ŀ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ(&R)"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   20
         Top             =   540
         Width           =   990
      End
      Begin VB.Label lbl������Ŀ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շ���Ŀ(&F)"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   17
         Top             =   540
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   6240
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptFilter.frx":0060
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptFilter.frx":03FA
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptFilter.frx":0794
            Key             =   "Clear"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDeptFilter.frx":0B2E
            Key             =   "Delete"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDeptFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrPrivs As String
Public mstrFilter As String
Public mlngDeptID As Long, mlngUnitID As Long
Public mblnDateMoved As Boolean '��ǰ��ѡ�����������Ƿ��ں����ݱ���
'��������
Public mstrFeeItems As String '�շ���ĿID��
Public mstrIncomeItems As String '������ĿID��
Private mintTab As Integer
Private Enum chkTypes
    ���ʵ��� = 0
    ���ʵ��� = 1
End Enum
Public mlngPrePatient As Long
Private mblnKeyReturn As Boolean
Private mblnNotClick As Boolean
Private mblnUnChange  As Boolean
Private mrsInfo As ADODB.Recordset
Private mblnOlnyBJYB As Boolean
Private mblnSeekName As Boolean '�Ƿ�ģ������,��ʱ��֧��ģ������,�ȶ��� �Ժ�ֵʹ��
 
Private Sub chkBill_Click(Index As Integer)
    Dim i As Integer, j As Integer
    
    j = 0
    For i = 0 To chkBill.UBound
        If chkBill(i).Value = 0 Then j = j + 1
    Next
    If j = i Then
        If Index = chkBills.�Զ����� And Not (frmManageBilling.tbs.SelectedItem.Key = "Auditing") Then
            '���۽����Զ�����
            chkBill(chkBills.��ͨ����).Value = 1
        Else
            chkBill(Index).Value = 1  '����i�Ǽ���1��
        End If
    End If
    
End Sub

Private Sub chkType_Click(Index As Integer)
    If chkType(0).Value = 0 And chkType(1).Value = 0 Then
        chkType((Index + 1) Mod 2).Value = 1
    End If
End Sub

Private Sub cbo����Ա_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlcontrol.CboMatchIndex(cbo����Ա.hWnd, KeyAscii)
        If lngIdx = -1 And cbo����Ա.ListCount > 0 Then lngIdx = 0
        cbo����Ա.ListIndex = lngIdx
    End If
End Sub


Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub



Private Sub cmdOK_Click()
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "�������ݺŲ���С�ڿ�ʼ���ݺţ�", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    
    Call MakeFilter
    
    gblnOK = True
    Hide
End Sub

Private Sub dtpB_Change()
    On Error Resume Next
    If dtpB.Value <= dtpBegin.MaxDate Then
        dtpBegin.Value = Format(dtpB.Value, "yyyy-MM-dd 00:00:00")
    End If
End Sub

Private Sub dtpE_Change()
    On Error Resume Next
    dtpB.MaxDate = dtpE.Value
    If dtpE.Value <= dtpEnd.MaxDate Then
        dtpBegin.Value = Format(dtpB.Value, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = Format(dtpE.Value, "yyyy-MM-dd 23:59:59")
    End If
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    mblnSeekName = True '֧������ģ������,����Ժ���ʲô��������,�����ٽ��е���
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If mintTab = 1 Or mintTab = 2 Then txtInput(mintTab - 1).SetFocus
    ElseIf KeyCode = vbKeyReturn And Not (mintTab = 1 Or mintTab = 2) Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Shift = 2 Then
        If mintTab = 1 Or mintTab = 2 Then
            Dim i As Integer, Index As Integer
            
            Index = mintTab - 1
            If UCase(Chr(KeyCode)) = "A" Then
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    ListFeeItem(Index).Selected(i) = True
                Next
            ElseIf UCase(Chr(KeyCode)) = "C" Then
                For i = 0 To ListFeeItem(Index).ListCount - 1
                    ListFeeItem(Index).Selected(i) = False
                Next
            End If
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim Curdate As Date, i As Long, Index As Integer
    Dim strListFeeItem As String
    Dim arrItem As Variant
    
    gblnOK = False
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    
    mstrFeeItems = ""
    mstrIncomeItems = ""
    '���ó�ʼֵ
    Curdate = zlDatabase.Currentdate
    dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = dtpBegin.MaxDate
    
    dtpB.MaxDate = Format(Curdate + 7, "yyyy-MM-dd 23:59:59")
    dtpB.Value = Format(DateAdd("m", -1, Curdate), "yyyy-MM-dd 00:00:00")
    dtpE.Value = dtpB.MaxDate
    
    Call GetOperator
    
    Call SSTab1_Click(0)
        
    
    If InStr(1, mstrPrivs, ";��ϸ��Ŀ����;") = 0 Then
        SSTab1.TabVisible(1) = False
        SSTab1.TabVisible(2) = False
    Else
        For Index = 0 To 1
            strListFeeItem = ""
            ListFeeItem(Index).Clear
            
            Call GetRegisterItem(g˽��ģ��, Me.Name & "\" & ListFeeItem(0).Name, IIf(Index = 0, "�շ���Ŀ�б�", "������Ŀ�б�"), strListFeeItem)
            If strListFeeItem <> "" Then
                arrItem = Split(strListFeeItem, ";")
                
                For i = 0 To UBound(arrItem)
                    ListFeeItem(Index).AddItem Split(arrItem(i), ",")(0)
                    ListFeeItem(Index).ItemData(ListFeeItem(Index).NewIndex) = Val(Split(arrItem(i), ",")(1))
                    ListFeeItem(Index).Selected(ListFeeItem(Index).NewIndex) = IIf(Val(Split(arrItem(i), ",")(2)) = 1, True, False)
                Next
            End If
        Next
    End If
End Sub

Private Sub ListFeeItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If ListFeeItem(Index).ListIndex >= 0 Then
            ListFeeItem(Index).RemoveItem ListFeeItem(Index).ListIndex
        End If
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
    Select Case SSTab1.Caption
        Case "����(&0)"
           mintTab = 0
        Case "�շ���Ŀ(&1)"
            mintTab = 1
        Case "������Ŀ(&2)"
            mintTab = 2
    End Select
    
End Sub


Private Sub tlbOpt_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Delete"
            If ListFeeItem(Index).ListIndex >= 0 Then
                Call ListFeeItem(Index).RemoveItem(ListFeeItem(Index).ListIndex)
            End If
        Case "Clear"
            ListFeeItem(Index).Clear
        Case "Save"
            Dim strTmp As String, i As Long
            With ListFeeItem(Index)
                For i = 0 To .ListCount - 1
                    strTmp = strTmp & ";" & .List(i) & "," & .ItemData(i) & "," & IIf(.Selected(i), 1, 0)
                Next
            End With
            strTmp = Mid(strTmp, 2)
            Call SaveRegisterItem(g˽��ģ��, Me.Name & "\" & ListFeeItem(0).Name, IIf(Index = 0, "�շ���Ŀ�б�", "������Ŀ�б�"), strTmp)
    End Select
End Sub

Private Sub txtInput_GotFocus(Index As Integer)
    Call zlcontrol.TxtSelAll(txtInput(Index))
End Sub

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
Dim strSql As String, strInput As String, strMatch As String, strIF As String
    Dim rsTmp As ADODB.Recordset
    Dim blnCancel As Boolean, i As Long
    Dim vRect As RECT
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        strInput = UCase(Trim(txtInput(Index).Text))
        If strInput = "" Then Exit Sub
        strMatch = IIf(Len(strInput) < 3, "", gstrLike)
        
        If Index = 0 Then
        '�շ���Ŀ
            If zlCommFun.IsNumOrChar(strInput) Then
                strIF = " And (A.���� like [1] Or B.���� like [1] And B.���� in(3," & gbytCode + 1 & "))"
            Else
                strIF = " And B.���� like [1]"
            End If
            strSql = "Select Distinct A.ID, A.����, B.���� ,A.���, A.����, A.���㵥λ " & _
                  " From �շ���ĿĿ¼ A,�շ���Ŀ���� B Where A.id=B.�շ�ϸĿID " & strIF & _
                  " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                  " And rownum<101 Order by ����"
        Else
        '������Ŀ
            If zlCommFun.IsNumOrChar(strInput) Then
                If IsNumeric(strInput) Then
                    strIF = " And ���� like [1]"
                Else
                    strIF = " And ���� like [1]"
                End If
            Else
                strIF = " And ���� like [1]"
            End If
            
            strSql = "Select ID, ����, ���� From ������Ŀ Where ĩ��=1 " & strIF & _
                " And rownum<101 Order by ����"
        End If
        
        On Error GoTo errH
        vRect = zlcontrol.GetControlRect(txtInput(Index).hWnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "��Ŀѡ��", 1, "", "��ѡ��", False, False, True, vRect.Left, vRect.Top, txtInput(Index).Height, blnCancel, False, True, strMatch & strInput & "%")
        If Not rsTmp Is Nothing Then
            With ListFeeItem(Index)
                For i = 0 To .ListCount - 1
                    If .ItemData(i) = rsTmp!ID Then
                        txtInput(Index).SetFocus
                        txtInput(Index).SelStart = 0
                        txtInput(Index).SelLength = Len(txtInput(Index).Text)
                        Exit Sub
                    End If
                Next
                If .ListCount < 100 Then
                    If Index = 0 Then
                        .AddItem rsTmp!���� & "-" & rsTmp!���� & "(" & rsTmp!��� & ")"
                    Else
                        .AddItem rsTmp!���� & "-" & rsTmp!����
                    End If
                    .ItemData(.NewIndex) = rsTmp!ID
                    .Selected(.NewIndex) = True
                Else
                    MsgBox "�������ܿ���,������Ŀ���ֻ�������100��!", vbInformation, gstrSysName
                End If
            End With
        End If
        
        txtInput(Index).SetFocus
        txtInput(Index).SelStart = 0
        txtInput(Index).SelLength = Len(txtInput(Index).Text)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Public Function GetOperator() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, i As Long
    
    On Error GoTo errH
    
    '����Ա
    cbo����Ա.Clear
    cbo����Ա.AddItem "���в���Ա"
    cbo����Ա.ListIndex = 0
    
    If mlngDeptID = 0 Then
        cbo����Ա.ListIndex = 0
        strSql = "Select Distinct A.ID,A.���,A.����,A.����" & _
            " From ��Ա�� A Where (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by A.����"
    Else
        strSql = "Select Distinct A.ID,A.���,A.����,A.����" & _
            " From ��Ա�� A,������Ա C" & _
            " Where A.ID=C.��ԱID And C.����ID IN([1],[2]) And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by A.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngDeptID, mlngUnitID)
    
    For i = 1 To rsTmp.RecordCount
        cbo����Ա.AddItem rsTmp!���� & "-" & rsTmp!����
        cbo����Ա.ItemData(cbo����Ա.NewIndex) = rsTmp!ID
        If rsTmp!ID = UserInfo.ID Then cbo����Ա.ListIndex = cbo����Ա.NewIndex
        rsTmp.MoveNext
    Next
    GetOperator = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    mlngDeptID = 0
    mlngUnitID = 0
    mstrPrivs = ""
End Sub
Private Sub txtName_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlcontrol.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlcontrol.TxtCheckKeyPress txtNOBegin, KeyAscii, m�ı�ʽ
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 14)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 14)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlcontrol.TxtSelAll txtNoEnd
End Sub

Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlcontrol.TxtCheckKeyPress txtNoEnd, KeyAscii, m�ı�ʽ
End Sub

Private Sub MakeFilter()
    Dim bln��ͨ���� As Boolean
    Dim i As Long, Index As Integer
    Dim strIDs As String
    
    mstrFilter = " And �Ǽ�ʱ�� Between [1] And [2]"
    
    mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
    
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And NO=[3]"
    End If
    
    If cbo����Ա.ListIndex <> -1 Then
        If cbo����Ա.ItemData(cbo����Ա.ListIndex) <> 0 Then
            mstrFilter = mstrFilter & " And ����Ա����||''=[5]"
        End If
    End If
    
    
    '�Զ�����
    bln��ͨ���� = chkBill(chkBills.��ͨ����).Value = 1 Or chkBill(chkBills.��������).Value = 1 Or chkBill(chkBills.��������).Value = 1
    If chkBill(chkBills.�Զ�����).Value = 1 And bln��ͨ���� Then
        mstrFilter = mstrFilter & " And ��¼���� IN(2,3)"
    ElseIf chkBill(chkBills.�Զ�����).Value = 0 And bln��ͨ���� Then
        mstrFilter = mstrFilter & " And ��¼����=2"
    ElseIf chkBill(chkBills.�Զ�����).Value = 1 And Not bln��ͨ���� Then
        mstrFilter = mstrFilter & " And ��¼����=3"
    End If
    
    '���ʻ�����
    If chkType(chkTypes.���ʵ���).Value = 1 And chkType(chkTypes.���ʵ���).Value = 1 Then
        mstrFilter = mstrFilter & " And ��¼״̬ IN(1,2,3)"
    ElseIf chkType(chkTypes.���ʵ���).Value = 1 Then
        mstrFilter = mstrFilter & " And ��¼״̬ IN(1,3)"
    ElseIf chkType(chkTypes.���ʵ���).Value = 1 Then
        mstrFilter = mstrFilter & " And ��¼״̬=2"
    End If
    
    If InStr(1, mstrPrivs, ";��ϸ��Ŀ����;") > 0 Then
        For Index = 0 To 1
            strIDs = ""
            For i = 0 To ListFeeItem(Index).ListCount - 1
                If ListFeeItem(Index).Selected(i) Then
                    strIDs = strIDs & "," & ListFeeItem(Index).ItemData(i)
                End If
            Next
            If strIDs <> "" Then
                strIDs = Mid(strIDs, 2)
                If Index = 0 Then
                    mstrFeeItems = strIDs
                    mstrFilter = mstrFilter & " And Instr(','||[8]||',',','||�շ�ϸĿID||',')>0"
                Else
                    mstrIncomeItems = strIDs
                    mstrFilter = mstrFilter & " And Instr(','||[9]||',',','||������ĿID||',')>0"
                End If
            End If
        Next
    End If
    
    'ҽ�����ж�����������
End Sub

            
Private Sub txtHospitalNO_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Sub txtName_GotFocus()
    zlcontrol.TxtSelAll txtName
    zlCommFun.OpenIme True
End Sub

Private Sub txtHospitalNO_GotFocus()
    zlcontrol.TxtSelAll txtHospitalNO
    zlCommFun.OpenIme False
End Sub

