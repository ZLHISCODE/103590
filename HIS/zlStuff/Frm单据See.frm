VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form Frm����See 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ݲ�ѯ"
   ClientHeight    =   6690
   ClientLeft      =   120
   ClientTop       =   2190
   ClientWidth     =   11730
   Icon            =   "Frm����See.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Cmdȡ�� 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   10515
      TabIndex        =   5
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   5715
      Left            =   0
      ScaleHeight     =   5655
      ScaleWidth      =   11655
      TabIndex        =   9
      Top             =   0
      Width           =   11715
      Begin ZL9BillEdit.BillEdit msf 
         Height          =   2895
         Left            =   240
         TabIndex        =   46
         Top             =   1320
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5106
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox Txt������ 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   4725
         Width           =   1005
      End
      Begin VB.TextBox Txt��ҩ�� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   10155
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   4725
         Width           =   1005
      End
      Begin VB.TextBox Txt�Ƿ��� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5745
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   4725
         Width           =   1005
      End
      Begin VB.ComboBox Cbo���� 
         Height          =   300
         Left            =   7140
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   765
         Width           =   1515
      End
      Begin VB.ComboBox Cbo�Ա� 
         Height          =   300
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   795
         Width           =   915
      End
      Begin VB.TextBox Txt���� 
         Height          =   285
         Left            =   5490
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "20"
         Top             =   795
         Width           =   435
      End
      Begin VB.TextBox TxtסԺ�� 
         Height          =   270
         Left            =   975
         TabIndex        =   29
         Top             =   435
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox Txt���� 
         Height          =   270
         Left            =   975
         TabIndex        =   28
         Top             =   825
         Width           =   1365
      End
      Begin VB.ComboBox Cbo���� 
         Height          =   300
         Left            =   9000
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   780
         Width           =   2430
      End
      Begin VB.TextBox Txt��Ʊ�� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5565
         MaxLength       =   8
         TabIndex        =   2
         Top             =   780
         Width           =   1035
      End
      Begin VB.TextBox TxtMoney 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   210
         TabIndex        =   24
         Top             =   4350
         Width           =   11235
      End
      Begin VB.ComboBox Cbo������ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   420
         Width           =   1635
      End
      Begin VB.TextBox TxtNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox Txt�������� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   2820
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   5160
         Width           =   1395
      End
      Begin VB.TextBox Txt������� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   10050
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   5160
         Width           =   1395
      End
      Begin VB.TextBox Txt����� 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   7950
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   5160
         Width           =   1005
      End
      Begin VB.TextBox Txt������ 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   780
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   5160
         Width           =   1005
      End
      Begin VB.TextBox TxtժҪ 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   765
         TabIndex        =   4
         Top             =   4740
         Width           =   10665
      End
      Begin VB.CommandButton Cmd��ҩ��λ 
         Caption         =   "��"
         Height          =   285
         Left            =   11190
         TabIndex        =   20
         Top             =   780
         Width           =   255
      End
      Begin VB.TextBox Txt��ҩ��λ 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   9000
         MaxLength       =   50
         TabIndex        =   3
         Top             =   780
         Width           =   2205
      End
      Begin VB.ComboBox Cbo�ⷿ 
         Enabled         =   0   'False
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   810
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker Dtp���� 
         Height          =   285
         Left            =   9750
         TabIndex        =   30
         Top             =   795
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   129761283
         CurrentDate     =   36471
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   165
         TabIndex        =   45
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl�Ƿ��� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƿ���"
         Height          =   180
         Left            =   5145
         TabIndex        =   44
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl��ҩ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��"
         Height          =   180
         Left            =   9570
         TabIndex        =   43
         Top             =   4800
         Width           =   540
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   6705
         TabIndex        =   39
         Top             =   825
         Width           =   360
      End
      Begin VB.Label Lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   225
         TabIndex        =   38
         Top             =   870
         Width           =   720
      End
      Begin VB.Label Lbl�Ա� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2790
         TabIndex        =   37
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5010
         TabIndex        =   36
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   9300
         TabIndex        =   35
         Top             =   855
         Width           =   360
      End
      Begin VB.Label LblסԺ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ס Ժ ��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   34
         Top             =   495
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Lbl���� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   8565
         TabIndex        =   27
         Top             =   855
         Width           =   360
      End
      Begin VB.Label Lbl��Ʊ�� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ʊ��"
         Height          =   180
         Left            =   4920
         TabIndex        =   25
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   225
         TabIndex        =   23
         Top             =   480
         Width           =   720
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO"
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
         Left            =   9675
         TabIndex        =   21
         Top             =   330
         Width           =   240
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   9240
         TabIndex        =   17
         Top             =   5220
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   7365
         TabIndex        =   16
         Top             =   5220
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2040
         TabIndex        =   15
         Top             =   5220
         Width           =   720
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   180
         TabIndex        =   14
         Top             =   5220
         Width           =   540
      End
      Begin VB.Label LblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժ  Ҫ"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Top             =   4815
         Width           =   540
      End
      Begin VB.Label Lbl��ҩ��λ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ��λ"
         Height          =   180
         Left            =   8205
         TabIndex        =   12
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Lbl�ⷿ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��    ��"
         Height          =   180
         Left            =   210
         TabIndex        =   11
         Top             =   885
         Width           =   720
      End
      Begin VB.Label Lbl���� 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "���ݲ�ѯ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   315
         TabIndex        =   10
         Top             =   150
         Width           =   11535
      End
   End
   Begin MSComctlLib.StatusBar Sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   6324
      Width           =   11736
      _ExtentX        =   20690
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "Frm����See.frx":030A
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15372
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
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
End
Attribute VB_Name = "Frm����See"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strNo As String               '�������:���ݺ�
Public int��¼״̬ As Integer        '�������:��¼״̬
Public byt���� As Byte               '�������:���ݱ�־:24-�շѴ�����25-���ʵ�������26-���ʱ���
Private blnFirst As Boolean
Private UnitLevel As Integer '��λ����

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Private Sub RefreshHead()
    '-----------------------------------------------------
    '--����:ˢ�±�ͷ�ṹ
    '--����:byt����
    '          24-�շѴ���
    '          25-���ʵ�����
    '          26-���ʱ���
    '--����:
    '-----------------------------------------------------
    With msf
        Select Case byt����
            Case 24, 25             '����
                
                Me.Caption = "���Ĵ�����"
                Me.Lbl���� = "���Ĵ�����"
               
                .Cols = 10
                .TextMatrix(0, 0) = "���������"
                .TextMatrix(0, 1) = "���"
                .TextMatrix(0, 2) = "����"
                .TextMatrix(0, 3) = "��λ"
                .TextMatrix(0, 4) = "����"
                .TextMatrix(0, 5) = "����"
                .TextMatrix(0, 6) = "�ۼ�"
                .TextMatrix(0, 7) = "���"
                .TextMatrix(0, 8) = "�ɱ����"
                .TextMatrix(0, 9) = "���"
                
                .ColWidth(0) = 2500
                .ColWidth(1) = 800
                .ColWidth(2) = 1000
                .ColWidth(3) = 500
                .ColWidth(4) = 1100
                .ColWidth(5) = 800
                .ColWidth(6) = 1000
                .ColWidth(7) = 800
                .ColWidth(8) = 800
                .ColWidth(9) = 800
                
                .ColData(0) = 1
                .ColData(1) = 5
                .ColData(2) = 5
                .ColData(3) = 5
                .ColData(4) = 4
                .ColData(5) = 5
                .ColData(6) = 5
                .ColData(7) = 5
                .ColData(8) = 5
            Case 26                     '��ҩ��
            
                Me.Caption = "���ʴ�����"
                Me.Lbl���� = "���ʴ�����"
               
                .Cols = 11
                .TextMatrix(0, 0) = "����"
                .TextMatrix(0, 1) = "��������"
                .TextMatrix(0, 2) = "���������"
                .TextMatrix(0, 3) = "���"
                .TextMatrix(0, 4) = "����"
                .TextMatrix(0, 5) = "��λ"
                .TextMatrix(0, 6) = "����"
                .TextMatrix(0, 7) = "�ۼ�"
                .TextMatrix(0, 8) = "���"
                .TextMatrix(0, 9) = "�ɱ����"
                .TextMatrix(0, 10) = "���"
                
                .ColWidth(0) = 1000
                .ColWidth(1) = 1000
                .ColWidth(2) = 2500
                .ColWidth(3) = 800
                .ColWidth(4) = 1000
                .ColWidth(5) = 500
                .ColWidth(6) = 1100
                .ColWidth(7) = 1000
                .ColWidth(8) = 800
                .ColWidth(9) = 800
                .ColWidth(10) = 800
                
                .ColData(0) = 4
                .ColData(1) = 4
                .ColData(2) = 1
                .ColData(3) = 5
                .ColData(4) = 5
                .ColData(5) = 5
                .ColData(6) = 4
                .ColData(7) = 5
                .ColData(8) = 5
                .ColData(9) = 5
                .ColData(10) = 0
                .PrimaryCol = 1
                .LocateCol = 9
        
        End Select
    End With
End Sub

Private Sub RefreshHeadData()
    '-----------------------------------------------------
    '--����:ˢ�±���������������
    '--����:
    '     byt����:
    '          24-�շѴ���
    '          25-���ʵ�����
    '          26-���ʱ���
    '    strNO
    '--����:
    '-----------------------------------------------------
    Dim RsHead As New ADODB.Recordset
    Dim strSql���� As String
    
    On Error GoTo errHandle
    Select Case byt����
        Case 24, 25          '����
            If byt���� = 24 Then
                gstrSQL = " Select A.NO,B.���� as ��������,G.����,G.�Ա�,G.����,0 As סԺ��,'' ����,A.ժҪ,A.������," & _
                " To_Char(A.��������,'yyyy-MM-dd') as ��������,A.�����,To_Char(A.�������,'yyyy-MM-dd') as �������," & _
                " G.����ID,0 ��ҳID,A.�Է�����ID as ��������ID,G.����Ա���� as �Ʒ���,A.����,A.ID,G.������ " & _
                          " From ҩƷ�շ���¼ A,���ű� B,������ü�¼ G " & _
                          " Where A.�Է�����id=B.id(+) And A.����=[2] And A.NO=G.NO(+) And A.����ID=G.ID And A.NO=[1] And Rownum < 2 "
            Else
                gstrSQL = " Select A.NO,B.���� as ��������,G.����,G.�Ա�,G.����,C.סԺ��,G.����,A.ժҪ,A.������," & _
                " To_Char(A.��������,'yyyy-MM-dd') as ��������,A.�����,To_Char(A.�������,'yyyy-MM-dd') as �������," & _
                " G.����ID,G.��ҳID,A.�Է�����ID as ��������ID,G.����Ա���� as �Ʒ���,A.����,A.ID,G.������ " & _
                          " From ҩƷ�շ���¼ A,���ű� B,������ü�¼ G,������Ϣ C " & _
                          " Where A.�Է�����id=B.id(+) And A.����=[2] And C.����id=G.����id And A.����ID=G.ID And A.NO=G.NO(+) And A.NO=[1] And Rownum < 2 "
                strSql���� = Replace(gstrSQL, "G.����", "'' ����")
                strSql���� = Replace(strSql����, "G.��ҳID", "0 ��ҳID")
                gstrSQL = strSql���� & " Union All " & Replace(gstrSQL, "������ü�¼", "סԺ���ü�¼")
            End If
        Case 26
            gstrSQL = " Select A.NO,A.ժҪ,A.������,To_Char(A.��������,'yyyy-MM-dd') as ��������,A.�����,To_Char(A.�������,'yyyy-MM-dd') as �������,G.ִ�в���ID,A.�Է�����ID as ��������ID,G.����Ա���� as �Ʒ���,A.����,A.ID" & _
                    " From ҩƷ�շ���¼ A,���ű� B,סԺ���ü�¼ G " & _
                    " Where A.�Է�����id=B.id(+) And A.����=[2] And A.NO=G.NO(+) And A.����ID=G.ID And A.NO=[1] And Rownum < 2 "
    End Select
    
    Set RsHead = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, byt����)
    
    With RsHead
        If .RecordCount = 0 Then Exit Sub
        '���Ʊ�ͷ����
        Select Case byt����
            Case 24, 25      '����
                Me.TxtNo = !NO
                Me.Txt������ = !������
                Me.Txt�Ƿ��� = IIf(IsNull(!�Ʒ���), " ", !�Ʒ���)
                Me.Dtp���� = !��������
                Lbl����.Caption = "��������"
                
                Me.Txt���� = IIf(IsNull(!����), " ", !����)
                Me.Cbo����.AddItem IIf(IsNull(!��������), "", !��������)
                Me.Cbo����.ListIndex = Me.Cbo����.NewIndex
                
                Me.Cbo�Ա�.AddItem IIf(IsNull(!�Ա�), "", !�Ա�)
                Cbo�Ա�.ListIndex = Cbo�Ա�.NewIndex
                
                Me.Txt���� = IIf(IsNull(!����), 20, !����)
                Me.TxtסԺ�� = IIf(IsNull(!סԺ��), " ", !סԺ��)
                Me.Txt��ҩ�� = IIf(IsNull(!�����), " ", !�����)
           
            Case 26                     '��ҩ��
                Me.TxtNo = !NO
                Me.Txt������ = !������
                Me.Txt�Ƿ��� = IIf(IsNull(!�Ʒ���), " ", !�Ʒ���)
                Me.Dtp���� = !��������
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshBodyData()
    '-----------------------------------------------------
    '--����:ˢ�±�����������
    '--����:
    '     byt����:
    '          24-�շѴ���
    '          25-���ʵ�����
    '          26-���ʱ���
    '    strNO
    '--����:
    '-----------------------------------------------------
    Dim RsBody As New ADODB.Recordset
    Dim intRow As Long
    Dim ii As Long
    On Error GoTo errHandle
    Select Case byt����
        Case 24        '�շѴ���
            gstrSQL = "SELECT DISTINCT '��'||F.����||'��'||NVL(E.����,F.����) AS ������Ϣ,F.���,F.����,F.���㵥λ AS ��λ," & _
                " B.����ϵ��,A.ʵ������ AS ����,A.���ۼ� AS ����," & _
                " A.���۽��,A.����,A.�ɱ����,A.���,A.ҩƷID,A.�ⷿID,A.���ϵ��,A.���,A.ID,B.���Ч��,B.ָ�����ۼ�,B.ָ�������" & _
                " FROM ҩƷ�շ���¼ A,�������� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F " & _
                " WHERE A.ҩƷID=B.����ID AND B.����ID=F.ID " & _
                " AND B.����ID=E.�շ�ϸĿID(+) AND E.����(+)=3 AND E.����(+)=1" & _
                " AND A.��¼״̬=[3] AND A.����=[2] AND A.NO =[1]" & _
                " ORDER BY A.���"
        Case 25      '���ʵ�����
            gstrSQL = "" & _
                "   SELECT DISTINCT '��'||F.����||'��'||NVL(E.����,F.����) AS ������Ϣ,F.���,F.����,F.���㵥λ AS ��λ," & _
                "       b.����ϵ��,A.ʵ������ AS ����,A.���ۼ� AS ����," & _
                "       A.���۽��,A.����,A.�ɱ����,A.���,A.ҩƷID,A.�ⷿID,A.���ϵ��,A.���,A.ID,B.���Ч��,B.ָ�����ۼ�,B.ָ�������" & _
                " FROM ҩƷ�շ���¼ A,�������� B,�շ���Ŀ���� E,�շ���ĿĿ¼ F " & _
                " WHERE A.ҩƷID=B.����ID AND B.����ID=F.ID " & _
                " AND B.����ID=E.�շ�ϸĿID(+) AND E.����(+)=3 AND E.����(+)=1" & _
                " AND A.����=[2] AND A.��¼״̬=[3] AND A.NO =[1]" & _
                " ORDER BY A.���"
        Case 26        '���ʱ���
            gstrSQL = "" & _
                "   SELECT DISTINCT P.���� ����,G.����,G.�Ա�,G.����,G.����ID,G.��ҳID,D.סԺ��,G.����," & _
                "       '��'||F.����||'��'||NVL(C.����,F.����) AS ������Ϣ,F.���,F.����," & _
                "       B.����ϵ��,A.ʵ������ AS ����,F.���㵥λ AS ��λ,A.���ۼ�  AS ����," & _
                "       A.���۽��,A.����,A.�ɱ����,A.���,A.ҩƷID,A.�ⷿID,A.���ϵ��,A.���,A.ID,B.���Ч��,B.ָ�����ۼ�,B.ָ�������" & _
                " FROM ҩƷ�շ���¼ A,�������� B,סԺ���ü�¼ G,������Ϣ D,�շ���Ŀ���� C,�շ���ĿĿ¼ F,���ű� P " & _
                " WHERE A.NO=G.NO(+) AND A.����ID=G.ID AND A.�Է�����ID=P.ID AND A.ҩƷID=B.����ID AND B.����ID=F.ID AND G.����ID=D.����ID " & _
                " AND B.����ID=C.�շ�ϸĿID(+) AND C.����(+)=3 AND C.����(+)=1 " & _
                " AND A.����=[2] AND A.��¼״̬=[3] AND A.NO =[1]" & _
                " ORDER BY A.���"
    End Select
    Set RsBody = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strNo, byt����, int��¼״̬)
    
    With RsBody
       If .RecordCount <> 0 Then msf.Rows = .RecordCount + 1
        
        '���Ʊ�ͷ����
        Select Case byt����
            Case 24, 25          '����
                    For intRow = 1 To .RecordCount
                        msf.TextMatrix(intRow, 0) = !������Ϣ
                        msf.TextMatrix(intRow, 1) = IIf(IsNull(!���), "", !���)
                        msf.TextMatrix(intRow, 2) = IIf(IsNull(!����), "", !����)
                        msf.TextMatrix(intRow, 3) = !��λ
                        msf.TextMatrix(intRow, 4) = Format(!����, mFMT.FM_����)
                        msf.TextMatrix(intRow, 5) = Format(!����, mFMT.FM_����)
                        msf.TextMatrix(intRow, 6) = Format(!����, mFMT.FM_���ۼ�)
                        msf.TextMatrix(intRow, 7) = Format(!���۽��, mFMT.FM_���)
                        msf.TextMatrix(intRow, 8) = Format(!�ɱ����, mFMT.FM_���)
                        msf.TextMatrix(intRow, 9) = Format(!���, mFMT.FM_���)
                        msf.RowData(intRow) = !���
                        .MoveNext
                    Next
            Case 26         '��ҩ
                    For intRow = 1 To .RecordCount
                        msf.TextMatrix(intRow, 0) = !����
                        msf.TextMatrix(intRow, 1) = !����
                        msf.TextMatrix(intRow, 2) = !������Ϣ
                        msf.TextMatrix(intRow, 3) = IIf(IsNull(!���), "", !���)
                        msf.TextMatrix(intRow, 4) = IIf(IsNull(!����), "", !����)
                        msf.TextMatrix(intRow, 5) = !��λ
                        msf.TextMatrix(intRow, 6) = Format(!����, mFMT.FM_����)
                        msf.TextMatrix(intRow, 7) = Format(!����, mFMT.FM_���ۼ�)
                        msf.TextMatrix(intRow, 8) = Format(!���۽��, mFMT.FM_���)
                        msf.TextMatrix(intRow, 9) = Format(!�ɱ����, mFMT.FM_���)
                        msf.TextMatrix(intRow, 10) = Format(!���, mFMT.FM_���)
                        msf.RowData(intRow) = !���
                        .MoveNext
                    Next
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If Not blnFirst Then Exit Sub
    SetEV
    blnFirst = False
    Me.TxtNo = strNo
    'ˢ�±�ͷ
    RefreshHead
    'ˢ�±�ͷ����
    RefreshHeadData
    'ˢ�±�������
    
    RefreshBodyData
    '��ʾ�ϼƽ��
    SumDataMSf
    LockCons

End Sub

Private Sub Form_Load()
    blnFirst = True
  
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(0, g_�ɱ���)
        .FM_��� = GetFmtString(0, g_���)
        .FM_���ۼ� = GetFmtString(0, g_�ۼ�)
        .FM_���� = GetFmtString(0, g_����)
    End With
    
    
    SetEV
    RestoreWinState Me, App.ProductName, Me.Caption
End Sub

Private Function LockCons()
    Me.Cbo������.Enabled = False
    Me.Cbo�ⷿ.Enabled = False
    Me.Txt��Ʊ��.Enabled = False
    Me.Txt��ҩ��λ.Enabled = False
    Me.Cmd��ҩ��λ.Enabled = False
    Me.msf.Active = False
    Me.Cbo����.Enabled = False
    Me.TxtժҪ.Enabled = False
    Me.Txt������.Enabled = False
    Me.Txt�����.Enabled = False
    Me.Txt��������.Enabled = False
    Me.Txt�������.Enabled = False
    Me.Txt����.Enabled = False
    Me.Txt����.Enabled = False
    Me.Cbo����.Enabled = False
    Me.Dtp����.Enabled = False
    Me.TxtסԺ��.Enabled = False
    Me.Cbo�Ա�.Enabled = False
    Me.Txt��ҩ��.Enabled = False
    Me.Txt�Ƿ���.Enabled = False
    Me.Txt������.Enabled = False
End Function

Private Sub SumDataMSf()
    '-------------------------------------------------------------
    '--����:�Ը����ݽ��н�����
    '--����:
    '       byt����:
    '          24-�շѴ���
    '          25-���ʵ�����
    '          26-���ʱ���
    '-- ����:
    '------------------------------------------------------------
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Long
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0: TxtMoney = ""
    
    Select Case byt����
    Case 24, 25                      ' ֱ�����봦��
        For intLop = 1 To msf.Rows - 1
            curTotal = curTotal + Val(msf.TextMatrix(intLop, 7))
            Cur���ʽ�� = Cur���ʽ�� + Val(msf.TextMatrix(intLop, 8))
            Cur���ʲ�� = Cur���ʲ�� + Val(msf.TextMatrix(intLop, 9))
        Next
        TxtMoney = "���ϼƣ�" & Format(curTotal, mFMT.FM_���) & Space(10) & "���ʽ��ϼƣ�" & Format(Cur���ʽ��, mFMT.FM_���) & Space(10) & "���ʲ�ۺϼƣ�" & Format(Cur���ʲ��, mFMT.FM_���)
    Case 26                          ' ֱ�������ҩ��
        For intLop = 1 To msf.Rows - 1
            curTotal = curTotal + Val(msf.TextMatrix(intLop, 7))
            Cur���ʽ�� = Cur���ʽ�� + Val(msf.TextMatrix(intLop, 8))
            Cur���ʲ�� = Cur���ʲ�� + Val(msf.TextMatrix(intLop, 9))
        Next
        TxtMoney = "���ϼƣ�" & Format(curTotal, mFMT.FM_���) & Space(10) & "���ʽ��ϼƣ�" & Format(Cur���ʽ��, mFMT.FM_���) & Space(10) & "���ʲ�ۺϼƣ�" & Format(Cur���ʲ��, mFMT.FM_���)
    End Select
End Sub

Private Function SetEV()
    '-------------------------------------------------------------
    '--����:���ÿؼ���Visible����
    '--����:
    '       byt����:
    '          24-�շѴ���
    '          25-���ʵ�����
    '          26-���ʱ���
    '-- ����:
    '------------------------------------------------------------
    Me.Lbl������.Visible = True
    Me.Cbo������.Visible = True
    Me.Lbl�ⷿ.Caption = "��    ��"
    
    Me.Lbl�ⷿ.Visible = True
    Me.Cbo�ⷿ.Visible = True
    Me.Lbl��Ʊ��.Visible = True
    Me.Txt��Ʊ��.Visible = True
    Me.Lbl��ҩ��λ.Visible = True
    Me.Txt��ҩ��λ.Visible = True
    Me.Cmd��ҩ��λ.Visible = True
    Me.Lbl����.Visible = True
    Me.Cbo����.Visible = True
    Me.LblժҪ.Visible = True
    Me.TxtժҪ.Visible = True
    Me.Lbl������.Visible = True
    Me.Txt������.Visible = True
    Me.Lbl�����.Visible = True
    Me.Txt�����.Visible = True
    Me.Lbl��������.Visible = True
    Me.Txt��������.Visible = True
    Me.Lbl�����.Visible = True
    Me.Lbl�������.Visible = True
    Me.Txt�������.Visible = True
    Me.Lbl��������.Visible = True
    Me.Txt����.Visible = True
    Me.Lbl����.Visible = True
    Me.Txt����.Visible = True
    Me.Lbl����.Visible = True
    Me.Cbo����.Visible = True
    Me.Lbl����.Visible = True
    Me.Dtp����.Visible = True
    Me.LblסԺ��.Visible = True
    Me.TxtסԺ��.Visible = True
    Me.Lbl�Ա�.Visible = True
    Me.Cbo�Ա�.Visible = True
    Me.Lbl��ҩ��.Visible = True
    Me.Txt��ҩ��.Visible = True
    Me.Lbl�Ƿ���.Visible = True
    Me.Txt�Ƿ���.Visible = True
    Me.Lbl������.Visible = True
    Me.Txt������.Visible = True
        
    Select Case byt����
    
    Case 24, 25                                     'ֱ�����봦��
        Me.Lbl������.Visible = False
        Me.Cbo������.Visible = False
        Me.Lbl�ⷿ.Visible = False
        Me.Cbo�ⷿ.Visible = False
        Me.Lbl��Ʊ��.Visible = False
        Me.Txt��Ʊ��.Visible = False
        Me.Lbl��ҩ��λ.Visible = False
        Me.Txt��ҩ��λ.Visible = False
        Me.Cmd��ҩ��λ.Visible = False
        Me.Lbl����.Visible = False
        Me.Cbo����.Visible = False
        Me.LblժҪ.Visible = False
        Me.TxtժҪ.Visible = False
        Me.Lbl������.Visible = False
        Me.Txt������.Visible = False
        Me.Lbl�����.Visible = False
        Me.Txt�����.Visible = False
        Me.Lbl��������.Visible = False
        Me.Txt��������.Visible = False
        Me.Lbl�������.Visible = False
        Me.Lbl�����.Visible = False
        Me.Txt�������.Visible = False
    Case 26                                        'ֱ�������ҩ��
        Me.Lbl�ⷿ.Caption = "��    ��"
        Me.Lbl�ⷿ.Visible = False
        Me.Cbo�ⷿ.Visible = False
        Me.Lbl������.Visible = False
        Me.Cbo������.Visible = False
        Me.Lbl��Ʊ��.Visible = False
        Me.Txt��Ʊ��.Visible = False
        Me.Lbl��ҩ��λ.Visible = False
        Me.Txt��ҩ��λ.Visible = False
        Me.Cmd��ҩ��λ.Visible = False
        Me.Lbl����.Visible = False
        Me.Cbo����.Visible = False
        Me.LblժҪ.Visible = False
        Me.TxtժҪ.Visible = False
        Me.Lbl������.Visible = False
        Me.Txt������.Visible = False
        Me.Lbl�����.Visible = False
        Me.Txt�����.Visible = False
        Me.Lbl��������.Visible = False
        Me.Txt��������.Visible = False
        Me.Lbl�����.Visible = False
        Me.Lbl�������.Visible = False
        Me.Txt�������.Visible = False
        Me.Lbl��������.Visible = False
        Me.Txt����.Visible = False
        Me.Lbl����.Visible = False
        Me.Txt����.Visible = False
        Me.Lbl����.Visible = False
        Me.Cbo����.Visible = False
        Me.LblסԺ��.Visible = False
        Me.TxtסԺ��.Visible = False
        Me.Lbl�Ա�.Visible = False
        Me.Cbo�Ա�.Visible = False
    
    End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, Me.Caption
End Sub
