VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBatchChangeNumNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�����������Ź���"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10740
   Icon            =   "frmBatchChangeNumNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10740
   StartUpPosition =   3  '����ȱʡ
   Begin VSFlex8Ctl.VSFlexGrid vsf�Һ���Ϣ 
      Height          =   2805
      Left            =   60
      TabIndex        =   5
      Top             =   2355
      Width           =   10635
      _cx             =   18759
      _cy             =   4948
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
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Frame fra���� 
      Height          =   2175
      Left            =   9090
      TabIndex        =   2
      Top             =   150
      Width           =   1620
      Begin VB.CommandButton btnȡ�� 
         Caption         =   "ȡ��"
         Height          =   350
         Left            =   255
         TabIndex        =   4
         Top             =   1065
         Width           =   1100
      End
      Begin VB.CommandButton btnȷ�� 
         Caption         =   "ȷ��"
         Height          =   350
         Left            =   255
         TabIndex        =   3
         Top             =   315
         Width           =   1100
      End
   End
   Begin VB.Frame fra�¹ҺŰ��� 
      Caption         =   "�¹ҺŰ�����Ϣ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2235
      Left            =   4605
      TabIndex        =   1
      Top             =   105
      Width           =   4410
      Begin VB.CommandButton cmd���Ű� 
         Caption         =   "P"
         Height          =   375
         Left            =   1635
         TabIndex        =   43
         Top             =   315
         Width           =   405
      End
      Begin VB.TextBox txtNewFilter 
         Height          =   375
         Left            =   675
         TabIndex        =   41
         Top             =   315
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   0
         TabIndex        =   40
         Top             =   750
         Width           =   4350
      End
      Begin VB.TextBox txt�ºű� 
         Height          =   300
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "�ű�"
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txt�¿��� 
         Height          =   300
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   32
         Text            =   "����"
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txt��ҽ�� 
         Height          =   300
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   "ҽ��"
         Top             =   1350
         Width           =   1380
      End
      Begin VB.TextBox txt�º��� 
         Height          =   300
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   30
         Text            =   "����"
         Top             =   1335
         Width           =   1380
      End
      Begin VB.TextBox txt���޺� 
         Height          =   300
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "�޺�"
         Top             =   1740
         Width           =   1380
      End
      Begin VB.TextBox txt����Լ 
         Height          =   300
         Left            =   2835
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "��Լ"
         Top             =   1725
         Width           =   1365
      End
      Begin VB.TextBox txt�°���ID 
         Height          =   285
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   27
         Text            =   "����ID"
         Top             =   855
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txt����Ŀ 
         Height          =   285
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "��Ŀ"
         Top             =   885
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label Label12 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   42
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label11 
         Caption         =   "�ű�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   39
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label10 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2205
         TabIndex        =   38
         Top             =   975
         Width           =   480
      End
      Begin VB.Label Label9 
         Caption         =   "ҽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   37
         Top             =   1380
         Width           =   465
      End
      Begin VB.Label Label8 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2220
         TabIndex        =   36
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label7 
         Caption         =   "�޺�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   35
         Top             =   1785
         Width           =   465
      End
      Begin VB.Label Label6 
         Caption         =   "��Լ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2235
         TabIndex        =   34
         Top             =   1770
         Width           =   480
      End
   End
   Begin VB.Frame fraԭ�ҺŰ��� 
      Caption         =   "ԭ�ҺŰ�����Ϣ"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2235
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   4410
      Begin VB.TextBox txt��Ŀ 
         Height          =   285
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "��Ŀ"
         Top             =   660
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txt����ID 
         Height          =   285
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   24
         Text            =   "����ID"
         Top             =   675
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.CommandButton cmdԭ�Ű� 
         Caption         =   "P"
         Height          =   375
         Left            =   3885
         TabIndex        =   23
         Top             =   285
         Width           =   405
      End
      Begin VB.TextBox txt��Լ 
         Height          =   300
         Left            =   2775
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "��Լ"
         Top             =   1755
         Width           =   1380
      End
      Begin VB.TextBox txt�޺� 
         Height          =   300
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "�޺�"
         Top             =   1755
         Width           =   1380
      End
      Begin VB.Frame Frame1 
         Height          =   30
         Left            =   30
         TabIndex        =   18
         Top             =   750
         Width           =   4350
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "����"
         Top             =   1350
         Width           =   1380
      End
      Begin VB.TextBox txtҽ�� 
         Height          =   300
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "ҽ��"
         Top             =   1350
         Width           =   1380
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "����"
         Top             =   945
         Width           =   1380
      End
      Begin VB.TextBox txt�ű� 
         Height          =   300
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "�ű�"
         Top             =   960
         Width           =   1380
      End
      Begin VB.TextBox txtFilter 
         Height          =   375
         Left            =   2910
         TabIndex        =   8
         Top             =   315
         Width           =   1380
      End
      Begin MSComCtl2.DTPicker dtpԤԼ���� 
         Height          =   345
         Left            =   990
         TabIndex        =   6
         Top             =   315
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   609
         _Version        =   393216
         Format          =   103940097
         CurrentDate     =   41128
      End
      Begin VB.Label Label5 
         Caption         =   "��Լ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2190
         TabIndex        =   21
         Top             =   1770
         Width           =   480
      End
      Begin VB.Label Label4 
         Caption         =   "�޺�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   165
         TabIndex        =   19
         Top             =   1785
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         TabIndex        =   16
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "ҽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   14
         Top             =   1410
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2145
         TabIndex        =   12
         Top             =   990
         Width           =   480
      End
      Begin VB.Label lbl�ű� 
         Caption         =   "�ű�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   10
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label lbl���� 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2430
         TabIndex        =   9
         Top             =   390
         Width           =   915
      End
      Begin VB.Label lblԤԼ���� 
         Caption         =   "ԤԼ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   7
         Top             =   390
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmBatchChangeNumNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mcbrControl As CommandBarControl, mcbrMenuBar As CommandBarPopup, mcbrToolBar As CommandBar
'��������
Private Const ID_PANE_Դ�ű���Ϣ = 1
Private Const ID_PANE_Ŀ��ű���Ϣ = 2
'��������
Private Enum mEnum_��ѯ����
     �軻���Ű� = 1
     �ɻ����Ű� = 2
End Enum
'A.��ʶ,A.����,A.����,A.����,A.��Ŀ,A.ҽ������,A.�޺���,A.��Լ��,A.��Լ��,A.�ѽ�����,A.ʱ���,A.����,A.����,A.��ſ���

Private mrs�ҺŰ��� As Recordset
Private mrs�ɻ����Ű� As Recordset
Private mrs�Һ���Ϣ As Recordset

Private Sub btnȡ��_Click()
    Unload Me
End Sub

Private Sub btnȷ��_Click()
    If CheckValid = False Then Exit Sub
    '��������
    If SaveData = True Then
        MsgBox "���ųɹ���", vbInformation, gstrSysName
        vsf�Һ���Ϣ.Clear 1
    Else
        MsgBox "����ʧ�ܣ�", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdԭ�Ű�_Click()
    Call Show�ҺŰ�����Ϣ(mEnum_��ѯ����.�軻���Ű�)
    If txt����ID.Text <> "" Then
        Call Show�ҺŰ�����Ϣ(mEnum_��ѯ����.�ɻ����Ű�)
    Else
        Call Clear�°�����Ϣ
    End If
    
    Call Show�Һ���Ϣ
End Sub

Private Sub cmd���Ű�_Click()
        Call Show�ҺŰ�����Ϣ(mEnum_��ѯ����.�ɻ����Ű�)
End Sub

Private Sub DTPԤԼ����_Change()
     'Call Show�ҺŰ�����Ϣ(mEnum_��ѯ����.�軻���Ű�)
End Sub

Private Sub Form_Load()
    RestoreWinState Me, App.ProductName
    Call Clearԭ������Ϣ
    Call Clear�°�����Ϣ
    Call InitԤԼʱ��
    Call SetHeader
    Call DTPԤԼ����_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrs�ҺŰ��� = Nothing
    Set mrs�Һ���Ϣ = Nothing
    SaveWinState Me, App.ProductName
End Sub

Private Sub Show�ҺŰ�����Ϣ(enum��ѯ���� As mEnum_��ѯ����)
    '----------------------------------------------------------------------------------------------
    '����:��ʾ�ҺŰ�����Ϣ
    '����:
    '����:����
    '����:2012/7/30
    '----------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strSQLBody As String
    Dim strSQL��ǰʱ�� As String
    Dim strSQL��������ʱ�� As String
    Dim strSQL�ɻ���ʱ�� As String
    Dim strSQL�ɻ����Ű� As String
    Dim strSQLWhere As String
    Dim datʱ�� As Date
    Dim strʱ�� As String
    
    On Error GoTo errHanl
    
    datʱ�� = dtpԤԼ����.Value
    strʱ�� = "to_date('" & datʱ�� & "','yyyy-MM-dd HH24:mi:ss')"
    
    strSQLBody = "Select a.Id, b.����, b.����, c.���� As ����, b.����id, d.���� As ��Ŀ, Nvl(a.����ҽ������, a.ҽ������) As ҽ������, a.ҽ��id, a.�޺���, a.��Լ��, a.�ѹ���, a.��Լ��," & vbNewLine & _
                "       a.�����ѽ��� As �ѽ�����, a.�ϰ�ʱ�� As ʱ���, Decode(b.�Ƿ񽨲���, 0, '', '��') As ����," & vbNewLine & _
                "       Decode(A.���﷽ʽ, 1, 'ָ��', 2, '��̬', 3, 'ƽ��', '') As ����, Decode(a.�Ƿ���ſ���, 0, '', '��') As ��ſ���" & vbNewLine & _
                "From �ٴ������¼ A, �ٴ������Դ B, ���ű� C, �շ���ĿĿ¼ D" & vbNewLine & _
                "Where a.��Դid = b.Id And a.��Ŀid = d.Id And b.����id = c.Id And (c.վ�� Is Null Or c.վ�� = [1]) And a.�������� = [2]"

    
    'ģ������
    '���ҷ�Χ:��Ŀ,ҽ������,����,����
    If Trim(txtFilter.Text) <> "" And enum��ѯ���� = �軻���Ű� Then
        strSQLWhere = " Where " & _
        "     ��Ŀ Like '%" & Trim(txtFilter.Text) & "%'" & _
        "  Or ҽ������ Like '%" & Trim(txtFilter.Text) & "%'" & _
        "  Or ���� Like '%" & Trim(txtFilter.Text) & "%'" & _
        "  Or ���� Like '%" & Trim(txtFilter.Text) & "%'"
    End If
    
    If enum��ѯ���� = �ɻ����Ű� Then
        strSQLBody = strSQLBody & " And (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��)) "
    End If
    
    strSQL = "" & _
    "   Select ID As ID,����,����,����,����ID,��Ŀ,ҽ������,ҽ��ID,�޺���,��Լ��,�ѹ���,��Լ��,�ѽ�����,ʱ���,����,����,��ſ��� From(" & strSQLBody & ") " & strSQLWhere & " Order By ����"
    
    If enum��ѯ���� = �軻���Ű� Then
        Set mrs�ҺŰ��� = zlDatabase.ShowSQLSelect(frmBatchChangeNumNew, strSQL, 0, "ԭ�ҺŰ�����Ϣ", False, "", "", False, False, False, txtFilter.Left, txtFilter.Top + txtFilter.Height, 1000, True, False, False, gstrNodeNo, datʱ��)
        If mrs�ҺŰ��� Is Nothing Then
           Call Clearԭ������Ϣ
        Else
           Call Setԭ������Ϣ
        End If
    End If
    
    
    If enum��ѯ���� = �ɻ����Ű� Then
        If mrs�ҺŰ��� Is Nothing Then Exit Sub
        '��ȡ��ͬʱ�ε��Ű���ϢSQL
        strSQL�ɻ���ʱ�� = "Select a.Id, b.����, b.����, c.���� As ����, b.����id, d.���� As ��Ŀ, Nvl(a.����ҽ������, a.ҽ������) As ҽ��, a.ҽ��id, a.�޺���, a.��Լ��, a.�ѹ���, a.��Լ��," & vbNewLine & _
                        "       a.�����ѽ��� As �ѽ�����, a.�ϰ�ʱ�� As ʱ���, Decode(b.�Ƿ񽨲���, 0, '', '��') As ����," & vbNewLine & _
                        "       Decode(a.���﷽ʽ, 1, 'ָ��', 2, '��̬', 3, 'ƽ��', '') As ����, Decode(a.�Ƿ���ſ���, 0, '', '��') As ��ſ���" & vbNewLine & _
                        "From �ٴ������¼ A, �ٴ������Դ B, ���ű� C, �շ���ĿĿ¼ D" & vbNewLine & _
                        "Where a.��Դid = b.Id And a.��Ŀid = d.Id And b.����id = c.Id And (c.վ�� Is Null Or c.վ�� = [1]) And a.�������� = [2] And" & vbNewLine & _
                        "      (a.��ʼʱ�� < Nvl(a.ͣ�￪ʼʱ��, a.��ֹʱ��) Or a.��ֹʱ�� > Nvl(a.ͣ����ֹʱ��, a.��ʼʱ��))"

        
        '�����ɻ����޶�����
        '�޺�����ͬ,��Լ����ͬ,������ͬ,������ͬ,�շ���Ŀ��ͬ,��Լ�����ѽ�����Ϊ0����ʾ���Ű�û�б�ԤԼ��
        strSQLWhere = "" & _
        "   And  Nvl(A.�޺���,'0') >= '" & Val(IIf(Not mrs�ҺŰ��� Is Nothing, Nvl(mrs�ҺŰ���!�ѹ���, "0"), "0")) & "'" & _
        "   And  Nvl(A.��Լ��,'0') >= '" & Val(IIf(Not mrs�ҺŰ��� Is Nothing, Nvl(mrs�ҺŰ���!��Լ��, "0"), "0")) & "'" & _
        "   And  Nvl(A.����,'��') = '" & IIf(Trim(txt����.Text) <> "", Trim(txt����.Text), "��") & "'" & _
        "   And  Nvl(A.����,'��') = '" & IIf(Trim(txt����.Text) <> "", Trim(txt����.Text), "��") & "'" & _
        "   And  Nvl(A.��Ŀ,'��') = '" & IIf(Trim(txt��Ŀ.Text) <> "", Trim(txt��Ŀ.Text), "��") & "'" & _
        "   And  Nvl(A.��Լ��,0) = 0 And Nvl(A.�ѽ�����,0) = 0 And Nvl(A.�ѹ���,0) = 0 " & _
        "   " & IIf(Not mrs�ҺŰ��� Is Nothing, "And A.���� <> '" & mrs�ҺŰ���!���� & "'", "")
        'ģ������
        '���ҷ�Χ:��Ŀ,ҽ������,����,����
        If Trim(txtNewFilter.Text) <> "" Then
            strSQLWhere = strSQLWhere & " And (" & _
            "     ��Ŀ Like '%" & Trim(txtNewFilter.Text) & "%'" & _
            "  Or ҽ������ Like '%" & Trim(txtNewFilter.Text) & "%'" & _
            "  Or ���� Like '%" & Trim(txtNewFilter.Text) & "%'" & _
            "  Or ���� Like '%" & Trim(txtNewFilter.Text) & "%'" & _
            "   )"
        End If

        '��ȡ�ɻ��ŵ��Ű�SQL
        strSQL�ɻ����Ű� = "" & _
        "   Select A.ID,A.����,A.����,A.����,A.����ID,A.��Ŀ,A.ҽ������,A.ҽ��ID,A.�޺���,A.��Լ��,A.��Լ��,A.�ѽ�����,A.ʱ���,A.����,A.����,A.��ſ��� From(" & strSQL & ") A,(" & strSQL�ɻ���ʱ�� & ") B " & _
        "   Where A.ID=B.ID(+)" & _
        "   " & strSQLWhere & _
        "   Order By ����"
               
        '��ѯ�ɻ����Ű���Ϣ
        
        Set mrs�ɻ����Ű� = zlDatabase.ShowSQLSelect(frmBatchChangeNumNew, strSQL�ɻ����Ű�, 0, "�¹ҺŰ�����Ϣ", False, "", "", False, False, False, txtFilter.Left, txtFilter.Top + txtFilter.Height, 1000, True, False, False, gstrNodeNo, datʱ��)
        
        If mrs�ɻ����Ű� Is Nothing Then
            Call Clear�°�����Ϣ
        Else
            If mrs�ɻ����Ű�.RecordCount = 0 Then
                Call Clear�°�����Ϣ
            Else
                Call Set�°�����Ϣ(mrs�ɻ����Ű�)
            End If
        End If
    End If
    Exit Sub
errHanl:
    MsgBox Err.Description
End Sub

Private Sub Show�Һ���Ϣ()
    '----------------------------------------------------------------------------------------------
    '����:��ʾ�Һ���Ϣ
    '����:
    '����:����
    '����:2012/7/30
    '----------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim strʱ�� As String
    
    If Not mrs�ҺŰ��� Is Nothing And Not mrs�ɻ����Ű� Is Nothing Then
        If Val(Nvl(mrs�ҺŰ���!�ѹ���)) <> 0 Then
            If MsgBox("ѡ�񻻺ŵļ�¼�����Ѿ�����ļ�¼,�Ƿ���������������ţ�", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Clearԭ������Ϣ
                Clear�°�����Ϣ
                vsf�Һ���Ϣ.Clear 1
                Exit Sub
            End If
        End If
    End If
    
    strʱ�� = "to_date('" & dtpԤԼ����.Value & "','yyyy-MM-dd HH24:mi:ss')"
    
    '�����:56164
    strSQL = "" & _
    "      Select A.No,A.���� As ���,A.����ʱ�� As ԤԼʱ��,A.����,A.�Ա�,A.����,B.���֤��,B.��ϵ�˵绰,Decode(A.��¼����,1,'�ѽ���',2,'��ԤԼ','��Ԥ��') As ״̬" & _
    "      From ���˹Һż�¼ A,������Ϣ B " & _
    "      Where A.��¼״̬=1 " & _
    "      And A.����ID=B.����ID(+) " & _
    "      And A.�����¼ID=" & Val(txt�ű�.Tag) & _
    ""
'    "      Union ALL" & _
'    "      Select A.No,A.���� As ���,A.����ʱ�� As ԤԼʱ��,A.����,A.�Ա�,A.����,B.���֤��,B.��ϵ�˵绰,Decode(A.��¼����,1,'�ѽ���',2,'��ԤԼ','��Ԥ��') As ״̬" & _
'    "      From ���˹Һż�¼ A,������Ϣ B " & _
'    "      Where A.����ʱ�� Between Trunc(" & strʱ�� & ") And Trunc(" & strʱ�� & ") +1 -1/24/60/60" & _
'    "      And A.��¼״̬=1 And Nvl(ԤԼ,0) = 1" & _
'    "      And A.����ID=B.����ID(+) " & _
'    "      And Nvl(�ű�,'�ű�')= '" & IsNothing(Trim(txt�ű�.Text), "�ű�") & "'"
    Set mrs�Һ���Ϣ = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    '���ùҺ���Ϣ�б�
    Set vsf�Һ���Ϣ.DataSource = mrs�Һ���Ϣ
    '������ͷ��Ϣ
    SetHeader
End Sub
Private Sub SetHeader()
    Dim str�ҺŰ��� As String
    Dim str�Һ���Ϣ As String

    'str�ҺŰ��� = "��ʶ,1,850|����,1,850|����,1,850|����,1,1700|��Ŀ,1,1800|ҽ������,1,1000|��Լ��,1,650|��Լ��,1,650|�ѽ�����,1,850|ʱ���,1,850|����,1,650|����,1,650|��ſ���,1,650"
    
    str�Һ���Ϣ = "No,4,0|���,4,650|ԤԼʱ��,4,2000|����,4,1000|�Ա�,4,650|����,4,650|���֤��,4,2000|��ϵ�˵绰,4,2000|״̬,4,650"
    
    '���ùҺŰ���
    SetVSGrid vsf�Һ���Ϣ, str�Һ���Ϣ
End Sub


Private Sub SetVSGrid(vsGrid As VSFlexGrid, str���� As String)
    '----------------------------------------------------------------------------------------------
    '����:����FlexGrid����
    '����:
    '����:����
    '����:2012/8/3
    '----------------------------------------------------------------------------------------------
    Dim strArr() As String
    Dim lngCols As Long
    Dim i As Long
    
    strArr = Split(str����, "|")
    lngCols = UBound(strArr) + 1
    
    'vsGrid.Clear
    'vsGrid.Rows = 2
    vsGrid.Cols = lngCols
    
    With vsGrid
        .ColWidthMin = 0
        .Redraw = False
         For i = 0 To UBound(strArr)
            .ColKey(i) = Split(strArr(i), ",")(0)
            .TextMatrix(0, i) = Split(strArr(i), ",")(0)
            .ColAlignment(i) = Split(strArr(i), ",")(1)
            .ColWidth(i) = Split(strArr(i), ",")(2)
        Next
        .RowHeight(0) = 320
        .ExtendLastCol = True
        .Redraw = True
    End With
End Sub


Private Sub InitԤԼʱ��()
    '----------------------------------------------------------------------------------------------
    '����:��ʼ��ԤԼʱ��ؼ�
    '����:
    '����:����
    '����:2012/8/6
    '----------------------------------------------------------------------------------------------
    Dim Curdate As Date
    
    Curdate = zlDatabase.Currentdate
    
    dtpԤԼ����.Value = Format(Curdate + 1, "yyyy-MM-dd ")
    dtpԤԼ����.MinDate = Format(Curdate + 1, "yyyy-MM-dd ")
End Sub

Private Sub Clearԭ������Ϣ()
     txt����ID.Text = ""
     txt�ű�.Text = ""
     txt�ű�.Tag = ""
     txt����.Text = ""
     txtҽ��.Text = ""
     txt����.Text = ""
     txt�޺�.Text = ""
     txt��Լ.Text = ""
     txt��Ŀ.Text = ""
End Sub

Private Sub Setԭ������Ϣ()
     txt����ID.Text = Nvl(mrs�ҺŰ���!ID, "")
     txt�ű�.Text = Nvl(mrs�ҺŰ���!����, "")
     txt�ű�.Tag = Nvl(mrs�ҺŰ���!ID, "")
     txt����.Text = Nvl(mrs�ҺŰ���!����, "")
     txtҽ��.Text = Nvl(mrs�ҺŰ���!ҽ������, "")
     txt����.Text = Nvl(mrs�ҺŰ���!����, "")
     txt�޺�.Text = Nvl(mrs�ҺŰ���!�޺���, "")
     txt��Լ.Text = Nvl(mrs�ҺŰ���!��Լ��, "")
     txt��Ŀ.Text = Nvl(mrs�ҺŰ���!��Ŀ, "")
End Sub

Private Sub Clear�°�����Ϣ()
     txt�°���ID.Text = ""
     txt�ºű�.Text = ""
     txt�ºű�.Tag = ""
     txt�¿���.Text = ""
     txt��ҽ��.Text = ""
     txt�º���.Text = ""
     txt���޺�.Text = ""
     txt����Լ.Text = ""
     txt����Ŀ.Text = ""
End Sub

Private Sub Set�°�����Ϣ(rs�°��� As Recordset)
     txt�°���ID.Text = Nvl(rs�°���!ID, "")
     txt�ºű�.Text = Nvl(rs�°���!����, "")
     txt�ºű�.Tag = Nvl(rs�°���!ID, "")
     txt�¿���.Text = Nvl(rs�°���!����, "")
     txt��ҽ��.Text = Nvl(rs�°���!ҽ������, "")
     txt�º���.Text = Nvl(rs�°���!����, "")
     txt���޺�.Text = Nvl(rs�°���!�޺���, "")
     txt����Լ.Text = Nvl(rs�°���!��Լ��, "")
     txt����Ŀ.Text = Nvl(rs�°���!��Ŀ, "")
End Sub
Public Function Nvl(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
    '-----------------------------------------------------------------------------------
    '����:ȡĳ�ֶε�ֵ
    '����:rsObj          �������ֶ�
    '     varValue       ��rsObjΪNULLֵʱ��ȡ��ֵ
    '����:�����Ϊ��ֵ,����ԭ����ֵ,���Ϊ��ֵ,�򷵻�ָ����varValueֵ
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        Nvl = varValue
    Else
        Nvl = rsObj
    End If
End Function

Public Function IsNothing(varValue As Variant, Optional varDefalt As Variant = "") As String
    '-----------------------------------------------------------------------------------
    '����:�жϱ����Ƿ�Ϊ��,Ϊ�շ���Ĭ��ֵ
    '����:objValue   ��Ҫ�жϵĶ���
    '     strDefalt  Ĭ��ֵ
    '����:�����Ϊ��ֵ,����ԭ����ֵ,���Ϊ��ֵ,�򷵻�ָ����strDefaltֵ
    '-----------------------------------------------------------------------------------
    Dim var����ֵ As Variant
    
    var����ֵ = IIf(Trim(varValue) <> "", varValue, varDefalt)
    IsNothing = var����ֵ
End Function

Public Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------
    '����:���滻������
    '����:
    '����:
    '����:����
    '����:2012-08-22
    '-----------------------------------------------------------------------------------
    Dim strNos As String
    Dim strSQL As String

    On Error GoTo Errhand
    
    mrs�Һ���Ϣ.MoveFirst
    '��ȡ���ݺ�
    While mrs�Һ���Ϣ.EOF = False
        strNos = strNos & "|" & Nvl(mrs�Һ���Ϣ!NO, "")
        mrs�Һ���Ϣ.MoveNext
    Wend
    
    If strNos <> "" Then
        strNos = Mid(strNos, 2)
    End If
    
    strSQL = "zl_���˹Һż�¼_��������("
    '���ݺ�     Nos_In varchar
    strSQL = strSQL & "'" & strNos & "',"
    '�ºű�     �ºű�_In ���˹Һż�¼.�ű�%Type
    strSQL = strSQL & "'" & mrs�ɻ����Ű�!���� & "',"
    '��ҽ������ ��ҽ������_In �ҺŰ���.ҽ������%Type
    strSQL = strSQL & "'" & Nvl(mrs�ɻ����Ű�!ҽ������, Null) & "',"
    '��ҽ��ID   ��ҽ��ID_In �ҺŰ���.ҽ��ID%Type
    strSQL = strSQL & "'" & Nvl(mrs�ɻ����Ű�!ҽ��ID, Null) & "',"
    '�¿���ID   �¿���ID_In �ҺŰ���.����ID%Type
    strSQL = strSQL & "'" & Nvl(mrs�ɻ����Ű�!����ID, Null) & "',"
    'ԭҽ������ ԭҽ������_In �ҺŰ���.ҽ������%Type
    strSQL = strSQL & "'" & Nvl(mrs�ҺŰ���!ҽ������, Null) & "',"
    'ԭҽ��ID   ԭҽ��ID_In �ҺŰ���.ҽ��ID%Type
    strSQL = strSQL & "'" & Nvl(mrs�ҺŰ���!ҽ��ID, Null) & "',"
    'ԭ�ű�     ԭ�ű�_In   ���˹Һż�¼.�ű�%Type
    strSQL = strSQL & "'" & mrs�ҺŰ���!���� & "',"
    '����Ա���� ����Ա����_In �Һ����״̬.����Ա����%Type
    strSQL = strSQL & "'" & IIf(UserInfo.���� = "", Null, UserInfo.����) & "',"
    strSQL = strSQL & "" & Val(mrs�ҺŰ���!ID) & ","
    strSQL = strSQL & "" & Val(mrs�ɻ����Ű�!ID) & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    SaveData = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Public Function CheckValid() As Boolean
    '-----------------------------------------------------------------------------------
    '����:���滻������
    '����:
    '����:�ɹ�����True,ʧ�ܷ���False
    '����:����
    '����:2012-08-22
    '-----------------------------------------------------------------------------------

    If mrs�ҺŰ��� Is Nothing Then
        MsgBox "����û��ѡ����Ҫ���л��Ų����ĺű�,���ܽ��л��Ų���!", vbInformation, gstrSysName
        CheckValid = False
        Exit Function
    End If
    If mrs�ɻ����Ű� Is Nothing Then
        MsgBox "����û��ѡ����Ҫ�µĺű�,���ܽ��л��Ų���!", vbInformation, gstrSysName
        CheckValid = False
        Exit Function
    End If
    If mrs�Һ���Ϣ Is Nothing Then
        MsgBox "�úű���û�кű��ҳ�����Ҫ���л��Ų���!", vbInformation, gstrSysName
        CheckValid = False
        Exit Function
    ElseIf mrs�Һ���Ϣ.RecordCount <= 0 Then
        MsgBox "�úű���û�кű��ҳ�����Ҫ���л��Ų���!", vbInformation, gstrSysName
        CheckValid = False
        Exit Function
    End If
    
    CheckValid = True
End Function
