VERSION 5.00
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDiseaseQuery 
   Caption         =   "��Ⱦ�����Խ����ѯ����"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16185
   Icon            =   "frmDiseaseQuery.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   16185
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPatiList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4305
      Left            =   840
      ScaleHeight     =   4305
      ScaleWidth      =   14940
      TabIndex        =   20
      Top             =   3600
      Width           =   14940
      Begin VB.PictureBox picPati 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         Height          =   2190
         Index           =   0
         Left            =   240
         Picture         =   "frmDiseaseQuery.frx":6852
         ScaleHeight     =   2190
         ScaleWidth      =   1800
         TabIndex        =   22
         Top             =   120
         Width           =   1800
         Begin VB.Label lblSource 
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   480
            TabIndex        =   29
            Top             =   120
            Width           =   855
         End
         Begin VB.Image imgMark 
            Height          =   300
            Index           =   0
            Left            =   130
            Picture         =   "frmDiseaseQuery.frx":A219
            Stretch         =   -1  'True
            Top             =   110
            Width           =   300
         End
         Begin VB.Label lblName 
            BackColor       =   &H00C0C000&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   345
            TabIndex        =   27
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblSex 
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   26
            Top             =   1080
            Width           =   375
         End
         Begin VB.Label lblAge 
            BackStyle       =   0  'Transparent
            Caption         =   "25��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   1080
            TabIndex        =   25
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblDisease 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "��˲�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   24
            Top             =   1440
            Width           =   1400
         End
         Begin VB.Label lblTime 
            BackStyle       =   0  'Transparent
            Caption         =   "2015/01/01 00:00"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   23
            Top             =   1815
            Width           =   1575
         End
      End
      Begin VB.VScrollBar HScr 
         Height          =   5295
         LargeChange     =   10
         Left            =   14280
         Max             =   100
         SmallChange     =   5
         TabIndex        =   21
         Top             =   -120
         Width           =   255
      End
   End
   Begin VB.Frame fraHead 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   16935
      Begin VB.Frame fraLeft 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Χ����"
         Height          =   2295
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4455
         Begin VB.CommandButton cmdFind 
            Caption         =   "����"
            Enabled         =   0   'False
            Height          =   350
            Left            =   2880
            TabIndex        =   13
            Top             =   1755
            Width           =   1200
         End
         Begin VB.ComboBox cboDate 
            Height          =   300
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   870
            Width           =   3045
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   14
            Top             =   1800
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   237764611
            CurrentDate     =   40256
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   960
            TabIndex        =   15
            Top             =   1320
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   237764611
            CurrentDate     =   40256
         End
         Begin VB.Label lblRegistDept 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   1080
            TabIndex        =   18
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label lblDept 
            BackStyle       =   0  'Transparent
            Caption         =   "�Ǽǿ���"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblDate 
            BackStyle       =   0  'Transparent
            Caption         =   "�Ǽ�ʱ��"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   915
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   960
            X2              =   4200
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.Frame fraPatiInfo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "������Ϣ"
         Height          =   2295
         Left            =   4680
         TabIndex        =   1
         Top             =   240
         Width           =   12015
         Begin VB.Label lalPatiDisease 
            BackStyle       =   0  'Transparent
            Caption         =   "�ν��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3840
            TabIndex        =   10
            Top             =   1680
            Width           =   9135
         End
         Begin VB.Label lalPatiDept 
            BackStyle       =   0  'Transparent
            Caption         =   "�����ڿ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   7800
            TabIndex        =   9
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lalPatiNo 
            BackStyle       =   0  'Transparent
            Caption         =   "201512021234"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   3360
            TabIndex        =   8
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lalPatiInfoDisease 
            BackStyle       =   0  'Transparent
            Caption         =   "���Ƽ�����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2520
            TabIndex        =   7
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lalPatiInfoNo 
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ�ţ�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2520
            TabIndex        =   6
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label lblPatiDept 
            BackStyle       =   0  'Transparent
            Caption         =   "���ң�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   6960
            TabIndex        =   5
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label lblPatiAge 
            BackStyle       =   0  'Transparent
            Caption         =   "30��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7200
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblPatiSex 
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5160
            TabIndex        =   3
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblPatiName 
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2520
            TabIndex        =   2
            Top             =   360
            Width           =   1815
         End
         Begin VB.Image imgPati 
            Height          =   1755
            Left            =   240
            Picture         =   "frmDiseaseQuery.frx":3EEF1
            Stretch         =   -1  'True
            Top             =   360
            Width           =   1650
         End
      End
      Begin VB.Label lblNote 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��˫��ѡ��һ����Ŀ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Height          =   360
      Left            =   480
      TabIndex        =   19
      Top             =   8280
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Enabled         =   0   'False
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDiseaseQuery.frx":3F8E1
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17463
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Picture         =   "frmDiseaseQuery.frx":40175
            Text            =   "-��Ⱦ��"
            TextSave        =   "-��Ⱦ��"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Picture         =   "frmDiseaseQuery.frx":40C91
            Text            =   "-�Ǵ�Ⱦ��"
            TextSave        =   "-�Ǵ�Ⱦ��"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   2
            Text            =   "��д"
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   25
            Text            =   "����"
            TextSave        =   "����"
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
   Begin VB.Image imgState 
      Height          =   300
      Index           =   0
      Left            =   6000
      Picture         =   "frmDiseaseQuery.frx":417AD
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgState 
      Height          =   300
      Index           =   1
      Left            =   6480
      Picture         =   "frmDiseaseQuery.frx":76485
      Stretch         =   -1  'True
      Top             =   3720
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgPatiPhoto 
      Height          =   1185
      Left            =   7200
      Picture         =   "frmDiseaseQuery.frx":AB15D
      Top             =   3600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Image imgCardBack 
      Height          =   2190
      Index           =   1
      Left            =   3360
      Picture         =   "frmDiseaseQuery.frx":AE56A
      Top             =   3600
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image imgCardBack 
      Height          =   2190
      Index           =   0
      Left            =   1200
      Picture         =   "frmDiseaseQuery.frx":B1F31
      Top             =   3600
      Visible         =   0   'False
      Width           =   1800
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmDiseaseQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng�Ǽǿ���ID As Long
Private rsPati As ADODB.Recordset
Private mdblScaleHeight  As Double
Private mlngSelIndex As Long        'ѡ��ķ�����
Private mIntWindMode As Integer     '0- �����Ĵ���ģʽ ��1- ҽ��վ���õĴ���ģʽ
Private mlngID As Long              'ѡ��ķ�����
Private mlngCount As Long           '������������
Private mlngPageCount As Long       'һҳ�Ŀ�Ƭ����
Private mlngRowCount As Long        'һ�еĿ�Ƭ����
Private mlngColCount As Long        'һ�еĿ�Ƭ����
Private mlngCardCount As Long       '��Ƭ����
Private mlngReportCount As Long     'ʵ����ʾ�ķ���������
Private mIntCboIndex As Integer     'ѡ��ĵǼ�ʱ�������
Private mlngOldY As Long
Private mblnRefresh As Boolean      '�Ƿ�ˢ��������
Private mobjBarPopup As CommandBar  '�Ҽ��˵�

Public Function ShowDiseaseQuery(ByVal var�Ǽǿ��� As Variant) As Long
    If TypeName(var�Ǽǿ���) = "String" Then    '����
        mlng�Ǽǿ���ID = GetDeptID(var�Ǽǿ���)
    ElseIf IsNumeric(var�Ǽǿ���) Then
        mlng�Ǽǿ���ID = Val(var�Ǽǿ���)
    Else
        mlng�Ǽǿ���ID = 0
    End If
    mIntWindMode = 0
    Me.Show 1
    ShowDiseaseQuery = Val(rsPati.RecordCount)
End Function

Public Function ShowPatiDis(ByVal rsDis As ADODB.Recordset, ByRef frmParent As Object) As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If rsDis Is Nothing Then Exit Function
    If rsDis.RecordCount > 0 Then
        mIntWindMode = 1
        mlngID = 0
        Set rsPati = rsDis
        mlngCount = rsPati.RecordCount
        Call AdjustCardsPosition
        stbThis.Panels(2).Text = "һ��" & CStr(mlngCount) & "�ŷ�������"
        Me.Show 1, frmParent
        ShowPatiDis = mlngID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadPatiCard(ByVal intIndex As Integer)
    If intIndex = 0 Then
        Call SetPicVisible(0, True)
        Exit Sub
    End If
    
    Load picPati(intIndex)
    With picPati(intIndex)
        .Width = picPati(0).Width
        .Height = picPati(0).Height
        .Picture = Nothing
        .Visible = True
        .ZOrder 1
    End With

    Load lblName(intIndex)
    Set lblName(intIndex).Container = picPati(intIndex)
    lblName(intIndex).Visible = True
    lblName(intIndex).FontSize = lblName(0).FontSize
    lblName(intIndex).Left = lblName(0).Left
    lblName(intIndex).Top = lblName(0).Top
    lblName(intIndex).Width = lblName(0).Width
    lblName(intIndex).Height = lblName(0).Height
    lblName(intIndex).Caption = ""
    lblName(intIndex).ZOrder 0
    
    Load lblAge(intIndex)
    Set lblAge(intIndex).Container = picPati(intIndex)
    lblAge(intIndex).Visible = True
    lblAge(intIndex).FontSize = lblAge(0).FontSize
    lblAge(intIndex).Left = lblAge(0).Left
    lblAge(intIndex).Top = lblAge(0).Top
    lblAge(intIndex).Width = lblAge(0).Width
    lblAge(intIndex).Height = lblAge(0).Height
    lblAge(intIndex).Caption = ""
    lblAge(intIndex).ZOrder 0
    
    Load lblSex(intIndex)
    Set lblSex(intIndex).Container = picPati(intIndex)
    lblSex(intIndex).Visible = True
    lblSex(intIndex).FontSize = lblSex(0).FontSize
    lblSex(intIndex).Left = lblSex(0).Left
    lblSex(intIndex).Top = lblSex(0).Top
    lblSex(intIndex).Width = lblSex(0).Width
    lblSex(intIndex).Height = lblSex(0).Height
    lblSex(intIndex).Caption = ""
    lblSex(intIndex).ZOrder 0
    
    Load lblDisease(intIndex)
    Set lblDisease(intIndex).Container = picPati(intIndex)
    lblDisease(intIndex).Visible = True
    lblDisease(intIndex).FontSize = lblDisease(0).FontSize
    lblDisease(intIndex).Left = lblDisease(0).Left
    lblDisease(intIndex).Top = lblDisease(0).Top
    lblDisease(intIndex).Width = lblDisease(0).Width
    lblDisease(intIndex).Height = lblDisease(0).Height
    lblDisease(intIndex).Caption = ""
    lblDisease(intIndex).ZOrder 0
    
    Load lblTime(intIndex)
    Set lblTime(intIndex).Container = picPati(intIndex)
    lblTime(intIndex).Visible = True
    lblTime(intIndex).FontSize = lblTime(0).FontSize
    lblTime(intIndex).Left = lblTime(0).Left
    lblTime(intIndex).Top = lblTime(0).Top
    lblTime(intIndex).Width = lblTime(0).Width
    lblTime(intIndex).Height = lblTime(0).Height
    lblTime(intIndex).Caption = ""
    lblTime(intIndex).ZOrder 0
    
    Load lblSource(intIndex)
    Set lblSource(intIndex).Container = picPati(intIndex)
    lblSource(intIndex).Visible = True
    lblSource(intIndex).FontSize = lblSource(0).FontSize
    lblSource(intIndex).Left = lblSource(0).Left
    lblSource(intIndex).Top = lblSource(0).Top
    lblSource(intIndex).Width = lblSource(0).Width
    lblSource(intIndex).Height = lblSource(0).Height
    lblSource(intIndex).Caption = ""
    lblSource(intIndex).ZOrder 0
    
    Load imgMark(intIndex)
    Set imgMark(intIndex).Container = picPati(intIndex)
    imgMark(intIndex).Visible = True
    imgMark(intIndex).Left = imgMark(0).Left
    imgMark(intIndex).Top = imgMark(0).Top
    imgMark(intIndex).Width = imgMark(0).Width
    imgMark(intIndex).Height = imgMark(0).Height
    imgMark(intIndex).ZOrder 0
End Sub

Private Sub LoadPati(ByRef rsPati As ADODB.Recordset)
    Dim strSQL As String
On Error GoTo errH
    strSQL = "Select a.Id, a.��Դ, a.����id, a.����, a.�Ա�, a.����, e.���� As ����, a.��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.��¼״̬, f.���� As �ͼ����, a.�걾����, a.�������," & vbNewLine & _
            "       a.��Ⱦ������ As ���Ƽ���, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, a.����ʱ��, a.�������˵��" & vbNewLine & _
            "From (Select a.Id, '����' As ��Դ, a.����id, b.����, b.�Ա�, b.����, b.����� As ��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id, a.��¼״̬, a.�걾����, a.�������," & vbNewLine & _
            "              a.��Ⱦ������, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, a.����ʱ��, a.�������˵��, b.ִ�в���id As ����id" & vbNewLine & _
            "       From �������Լ�¼ A, ���˹Һż�¼ B" & vbNewLine & _
            "       Where a.����id = b.����id And a.�Һŵ� = b.No And a.�Ǽǿ���id = [1] And a.�Ǽ�ʱ�� Between [2] And [3]" & vbNewLine & _
            "       Union All" & vbNewLine & _
            "       Select a.Id, 'סԺ' As ��Դ, a.����id, c.����, c.�Ա�, c.����, c.סԺ�� As ��ʶ��, a.�ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id, a.��¼״̬, a.�걾����, a.�������," & vbNewLine & _
            "              a.��Ⱦ������, a.�Ǽ���, a.�Ǽ�ʱ��, a.������, a.����ʱ��, a.�������˵��, c.��Ժ����id As ����id" & vbNewLine & _
            "       From �������Լ�¼ A, ������ҳ C" & vbNewLine & _
            "       Where a.����id = c.����id And a.��ҳid = c.��ҳid And a.�Ǽǿ���id =[1] And a.�Ǽ�ʱ�� Between [2] And [3]) A, ���ű� E, ���ű� F" & vbNewLine & _
            "Where a.�ͼ����id = f.Id(+) And a.����id = e.Id(+) order by a.�Ǽ�ʱ�� desc"


    Set rsPati = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�Ǽǿ���ID, CDate(dtpTime(0).Value), CDate(dtpTime(1).Value + 1 - 1 / 24 / 60 / 60))
    mlngCount = rsPati.RecordCount
    mblnRefresh = True
    If rsPati.RecordCount > 0 Then
        Call AdjustCardsPosition
    Else
        Call UnloadControls(False)
    End If
    stbThis.Panels(2).Text = "һ��" & CStr(mlngCount) & "�ŷ�������"
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetPicVisible(ByVal Index As Long, ByVal blnVisible As Boolean)
    lblName(Index).Visible = blnVisible
    lblAge(Index).Visible = blnVisible
    lblSex(Index).Visible = blnVisible
    lblDisease(Index).Visible = blnVisible
    lblTime(Index).Visible = blnVisible
    lblSource(Index).Visible = blnVisible
    picPati(Index).Visible = blnVisible
End Sub

Private Sub AdjustCardsPosition(Optional ByVal lngY As Long = 0)
    Dim lngRowCount As Long
    Dim lngColCount As Long
    Dim lngX As Long, lngStart As Long
    Dim lngShowed As Long
    Dim blnAdjust As Boolean
    Dim i As Long
   
    blnAdjust = (lngY = 0)
    lngX = 50
    'ÿһ���ж��ٸ�
    lngRowCount = Abs((picPatiList.Width - HScr.Width - 50) / (picPati(0).Width + 15) - 0.5)
    If lngRowCount < 1 Then lngRowCount = 1
    lngColCount = Abs(picPatiList.Height / picPati(0).Height + 1)
    mlngPageCount = lngColCount * lngRowCount
    
    If Not mblnRefresh And mlngRowCount = lngRowCount And mlngColCount = lngColCount And lngY = mlngOldY Then
        Exit Sub
    End If
    mlngRowCount = lngRowCount
    mlngColCount = lngColCount
    mlngOldY = lngY
    mblnRefresh = False
    
    Call gobjComlib.zlControl.FormLock(Me.hwnd)
    '���ؿ�Ƭ
    If mlngCardCount < mlngPageCount Then
        For i = mlngCardCount + 1 To mlngPageCount
             Call LoadPatiCard(i)
        Next
        mlngCardCount = mlngPageCount
    End If
    '����������֮������λ��
    
    If lngY <> 0 Then
        lngStart = CLng((-1 * lngY) / picPati(0).Height - 0.5) * lngRowCount
        If lngStart < 0 Then lngStart = 0
        lngY = lngY + CLng((-1 * lngY) / picPati(0).Height - 0.5) * picPati(0).Height
    End If
    
    '���ؿ�Ƭ�������Ϣ
    Call LoadPatiCardInfo(lngStart)
    
    '���ÿ�Ƭ�Ŀɼ���
    For i = 0 To mlngReportCount - 1
        Call SetPicVisible(i, True)
    Next
    If mlngReportCount - 1 < mlngCardCount Then
        For i = mlngReportCount To mlngCardCount
            Call SetPicVisible(i, False)
        Next
    End If
    
    '����ÿ�ſ�Ƭ��λ��
    If mlngPageCount > 0 Then
        For i = 0 To mlngPageCount
            If i Mod (lngRowCount) = 0 And i <> 0 Then
                lngX = 50
                lngY = lngY + picPati(0).Height
            End If
            picPati(i).Left = lngX
            picPati(i).Top = lngY
            lngX = lngX + picPati(0).Width
        Next
    End If
    mdblScaleHeight = picPati(0).Height * IIf(mlngCount / lngRowCount > CLng(mlngCount / lngRowCount), CLng(mlngCount / lngRowCount + 0.5), CLng(mlngCount / lngRowCount))
    If blnAdjust Then
        HScr.Value = 0
        HScr.Visible = mdblScaleHeight > picPatiList.Height
    End If
    
    Call gobjComlib.zlControl.FormLock(0)
End Sub


Private Sub UnloadControls(ByVal blnUnload As Boolean)
    Dim j As Long
    For j = picPati.Count - 1 To 1 Step -1
        If blnUnload Then
            Unload lblName(j)
            Unload lblAge(j)
            Unload lblSex(j)
            Unload lblDisease(j)
            Unload lblTime(j)
            Unload lblSource(j)
            Unload imgMark(j)
            Unload picPati(j)
        Else
            Call SetPicVisible(j, False)
        End If
    Next
    Call SetPicVisible(0, False)
    mlngSelIndex = -1
    lblPatiName.Caption = "����"
    lblPatiSex.Caption = "�Ա�"
    lblPatiAge.Caption = "����"
    lalPatiNo.Caption = ""
    lalPatiDept.Caption = ""
    lalPatiDisease.Caption = ""
    imgPati.Picture = imgPatiPhoto.Picture
End Sub

Private Sub cboDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Call LoadPati(rsPati)
    End If
    KeyAscii = 0
End Sub

Private Sub LoadPatiCardInfo(ByVal lngStart As Long)
    Dim i As Long
    If rsPati.RecordCount > 0 Then
        Call rsPati.Move(lngStart, adBookmarkFirst)
        Do While Not rsPati.EOF
            If i >= mlngPageCount Then Exit Do
            picPati(i).Tag = rsPati!ID
            picPati(i).Picture = imgCardBack(1).Picture
            If Val(rsPati!��¼״̬ & "") = 1 Then
                imgMark(i).Visible = False
                imgMark(i).Tag = "δ����"
            ElseIf Val(rsPati!��¼״̬ & "") = 2 Then
                imgMark(i).Visible = True
                imgMark(i).Picture = imgState(1).Picture
                imgMark(i).Tag = "�Ѵ���Ϊ��Ⱦ��"
            ElseIf Val(rsPati!��¼״̬ & "") = 3 Then
                imgMark(i).Visible = True
                imgMark(i).Picture = imgState(0).Picture
                imgMark(i).Tag = "�Ѵ���Ϊ�Ǵ�Ⱦ��"
            ElseIf Val(rsPati!��¼״̬ & "") = 4 Then
                imgMark(i).Visible = False
                imgMark(i).Tag = "ת�ƴ�����"
            End If
            lblName(i).Caption = rsPati!���� & ""
            lblName(i).Tag = rsPati!������ & ""
            lblAge(i).Caption = rsPati!���� & ""
            lblSex(i).Caption = rsPati!�Ա� & ""
            lblDisease(i).Caption = rsPati!���Ƽ��� & ""
            If IsDate(rsPati!�Ǽ�ʱ�� & "") Then
                 lblTime(i).Caption = Format(rsPati!�Ǽ�ʱ�� & "", "yyyy-mm-dd HH:MM")
            End If
            If mIntWindMode = 0 Then
                lblSource(i).Caption = rsPati!��Դ & ""
            ElseIf mIntWindMode = 1 Then
                lblSource(i).Caption = rsPati!�Ǽǿ��� & ""
            End If
            rsPati.MoveNext
            i = i + 1
        Loop
    End If
    mlngReportCount = i
End Sub


Private Sub cboDate_Click()
    Dim curDate As Date
    
    If mIntCboIndex = cboDate.ListIndex And cboDate.ListIndex <> 5 Then Exit Sub
    mIntCboIndex = cboDate.ListIndex
    
    dtpTime(0).Enabled = (cboDate.ListIndex = cboDate.ListCount - 1)
    dtpTime(1).Enabled = (cboDate.ListIndex = cboDate.ListCount - 1)
    
    curDate = gobjComlib.zlDatabase.Currentdate
    dtpTime(0).MaxDate = curDate
    dtpTime(1).MaxDate = curDate
    
    Select Case cboDate.ListIndex
    Case 0 '����
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 1 '�������
        dtpTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 2 '�������
        dtpTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 3 '���һ��
        dtpTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 4 '���һ��
        dtpTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
    Case 5 'ָ  ��
        If Me.Visible Then
            dtpTime(0).SetFocus
            cmdFind.Enabled = True
        End If
    End Select
    
    If cboDate.ListIndex <> 5 Then cmdFind.Enabled = False
    
    If cboDate.ListIndex <> cboDate.ListCount - 1 Then
        If Me.Visible Then
            Call gobjComlib.ZLCommFun.PressKey(vbKeyReturn)
        End If
    End If
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Integer
    Dim objControl As CommandBarControl
        
    Select Case Control.ID
        Case conMenu_File_Modify
            If mlngSelIndex >= 0 Then
                If CLng(picPati(mlngSelIndex).Tag) > 0 Then
                    Call ModifyDiseaseRegist(CLng(picPati(mlngSelIndex).Tag))
                End If
            End If
        Case conMenu_File_Delete
            If mlngSelIndex >= 0 Then
                If CLng(picPati(mlngSelIndex).Tag) > 0 Then
                    Call DeleteDiseaseRegist(CLng(picPati(mlngSelIndex).Tag))
                End If
            End If
        Case conMenu_View_DiseaseRegist
            If mlngSelIndex >= 0 Then
                If CLng(picPati(mlngSelIndex).Tag) > 0 Then
                    Call frmDiseaseRegist.ShowDiseaseRegist(Me, 2, CLng(picPati(mlngSelIndex).Tag))
                End If
            End If
        Case conMenu_View_ToolBar_Button '������
            For i = 2 To cbsMain.Count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Text '��ť����
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case conMenu_View_ToolBar_Size '��ͼ��
            Me.cbsMain.Options.LargeIcons = Not Me.cbsMain.Options.LargeIcons
            Me.cbsMain.RecalcLayout
        Case conMenu_View_StatusBar '״̬��
            Me.stbThis.Visible = Not Me.stbThis.Visible
            Me.cbsMain.RecalcLayout
            cbsMain_Resize
        Case conMenu_View_Refresh
            Call LoadPati(rsPati)
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call gobjComlib.zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '������̳
            Call gobjComlib.zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call gobjComlib.zlMailTo(Me.hwnd)
        Case conMenu_File_Exit '�˳�
            Unload Me
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim blnStatesBar As Boolean
    
    blnStatesBar = stbThis.Visible
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    fraLeft.Visible = True
    fraPatiInfo.Visible = True
    lblNote.Visible = False
        
    With fraHead
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = 2655
    End With
  
    If mIntWindMode = 1 Then
        fraLeft.Visible = False
        fraPatiInfo.Visible = False
        lblNote.Visible = True
        fraHead.Height = 400
    End If
  
    With picPatiList
        .Top = fraHead.Top + fraHead.Height + 100
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = lngBottom - .Top - stbThis.Height
        If blnStatesBar Then
          .Height = lngBottom - .Top - stbThis.Height
        Else
          .Height = lngBottom - .Top
        End If
    End With
    
    If blnStatesBar Then
        With stbThis
             .Top = picPatiList.Top + picPatiList.Height
            .Left = lngLeft
            .Width = lngRight - lngLeft
        End With
    End If
    Me.Refresh
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Me.Visible = False Then Exit Sub
    Err = 0: On Error Resume Next
    Select Case Control.ID
        Case conMenu_File_Modify
            Control.Enabled = (mlngSelIndex >= 0)
            If Control.Enabled Then Control.Enabled = (imgMark(mlngSelIndex).Tag = "δ����")
            If Control.Enabled Then Control.Enabled = (CLng(picPati(mlngSelIndex).Tag) > 0)
        Case conMenu_File_Delete
            Control.Enabled = (mlngSelIndex >= 0)
            If Control.Enabled Then Control.Enabled = (imgMark(mlngSelIndex).Tag = "δ����")
            If Control.Enabled Then Control.Enabled = (CLng(picPati(mlngSelIndex).Tag) > 0)
        Case conMenu_View_DiseaseRegist
            Control.Enabled = (mlngSelIndex >= 0)
            If Control.Enabled Then Control.Enabled = (CLng(picPati(mlngSelIndex).Tag) > 0)
        Case conMenu_View_ToolBar_Button '������
            If cbsMain.Count >= 2 Then
            Control.Checked = Me.cbsMain(2).Visible
            End If
        Case conMenu_View_ToolBar_Text 'ͼ������
            If cbsMain.Count >= 2 Then
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
            End If
        Case conMenu_View_ToolBar_Size '��ͼ��
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar '״̬��
            Control.Checked = Me.stbThis.Visible
    End Select
End Sub

Private Sub cmdFind_Click()
    Call LoadPati(rsPati)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = HScr.Value
    lngMin = HScr.Min
    lngMax = HScr.Max
    
    If KeyCode = vbKeyPageDown Then '��
        If Between(lngCur + (lngMax - lngMin) / 100, lngMin, lngMax) Then
            HScr.Value = lngCur + (lngMax - lngMin) / 100
        Else
            HScr.Value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '��
        If Between(lngCur - (lngMax - lngMin) / 100, lngMin, lngMax) Then
            HScr.Value = lngCur - (lngMax - lngMin) / 100
        Else
            HScr.Value = lngMin
        End If
    End If
End Sub

Private Sub Form_Activate()
    glngPreHWnd = GetWindowLong(Me.hwnd, GWL_WNDPROC)
    SetWindowLong Me.hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
    SetWindowLong Me.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub Form_Load()
    mlngReportCount = 0
    mIntCboIndex = 0
    mlngRowCount = 0
    mlngColCount = 0
    mlngCardCount = 0
    If mIntWindMode = 0 Then
        cbsMain.ActiveMenuBar.Visible = True
        Call InitCommandBar
        
        cboDate.AddItem "��    ��"
        cboDate.AddItem "�������"
        cboDate.AddItem "�������"
        cboDate.AddItem "���һ��"
        cboDate.AddItem "���һ��"
        cboDate.AddItem "[ָ  ��]"
        cboDate.ListIndex = 3
        
        mlngSelIndex = -1
        Call GetRegistDept
        Call LoadPati(rsPati)
        lblPatiName.Caption = "����"
        lblPatiSex.Caption = "�Ա�"
        lblPatiAge.Caption = "����"
        lalPatiNo.Caption = ""
        lalPatiDept.Caption = ""
        lalPatiDisease.Caption = ""
        Me.BorderStyle = 2
        Me.Caption = "��Ⱦ�����Խ����ѯ����"
        lblNote.Visible = False
        stbThis.Visible = True
         '����ָ�
        Call gobjComlib.RestoreWinState(Me, App.ProductName)
    ElseIf mIntWindMode = 1 Then
        Me.BorderStyle = 3
        Me.Caption = "���Խ��ѡ��"
        lblNote.Visible = True

        cbsMain.ActiveMenuBar.Visible = False
        stbThis.Visible = False
        If mlngCount = 2 Then
            Me.Width = 4030
        Else
            Me.Width = 5850
        End If
        
        Me.Height = 3200
    End If
End Sub

Private Sub Form_Resize()
    Call cbsMain_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call UnloadControls(True)
    If mIntWindMode = 0 Then Call gobjComlib.SaveWinState(Me, App.ProductName)
End Sub

Private Sub HScr_Change()
    Dim lngMove As Long
    Dim lngY As Long
    If Not HScr.Visible Then Exit Sub
    '���㵥������
    lngMove = CLng((mdblScaleHeight - picPatiList.Height) / 100)

    If lngMove < 0 Then lngMove = 0
    lngY = -1 * HScr.Value * lngMove
    If lngY >= 0 And lngY < 100 Then lngY = 0
    Call AdjustCardsPosition(lngY)
End Sub

Private Sub lblAge_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub lblAge_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblAge(Index).Left + X, lblAge(Index).Top + Y)
End Sub

Private Sub lblDisease_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    gobjComlib.ZLCommFun.ShowTipInfo picPati(Index).hwnd, "���ƴ�Ⱦ����" & lblDisease(Index).Caption
End Sub

Private Sub lblDisease_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub lblDisease_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblDisease(Index).Left + X, lblDisease(Index).Top + Y)
End Sub

Private Sub lblName_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblName(Index).Left + X, lblName(Index).Top + Y)
End Sub

Private Sub lblName_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    gobjComlib.ZLCommFun.ShowTipInfo picPati(Index).hwnd, "������" & lblName(Index).Caption
End Sub

Private Sub lblName_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub lblSource_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mIntWindMode = 1 Then
        gobjComlib.ZLCommFun.ShowTipInfo picPati(Index).hwnd, "�����ң�" & lblSource(Index).Caption
    ElseIf mIntWindMode = 0 Then
        Call picPati_MouseMove(Index, Button, Shift, X, Y)
    End If
End Sub

Private Sub lblSource_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub picPati_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    gobjComlib.ZLCommFun.ShowTipInfo picPati(Index).hwnd, "״̬��" & imgMark(Index).Tag
End Sub

Private Sub lblSex_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblSex(Index).Left + X, lblSex(Index).Top + Y)
End Sub

Private Sub lblSex_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub lblSource_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblSource(Index).Left + X, lblSource(Index).Top + Y)
End Sub

Private Sub lblTime_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseDown(Index, Button, Shift, lblTime(Index).Left + X, lblTime(Index).Top + Y)
End Sub

Private Sub lblTime_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picPati_MouseUp(Index, Button, Shift, X, Y)
End Sub

Private Sub lblAge_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblDisease_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblName_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblSex_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblSource_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblTime_DblClick(Index As Integer)
    Call picPati_DblClick(Index)
End Sub

Private Sub lblTime_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    gobjComlib.ZLCommFun.ShowTipInfo picPati(Index).hwnd, "�Ǽ�ʱ�䣺" & lblTime(Index).Caption
End Sub

Private Sub picPati_DblClick(Index As Integer)
      '�鿴������
    Dim lngID As Long
    
    If mlngSelIndex < 0 Then Exit Sub
    lngID = CLng(picPati(mlngSelIndex).Tag)

    If mIntWindMode = 0 Then
        If lngID > 0 Then
            Call frmDiseaseRegist.ShowDiseaseRegist(Me, 2, lngID)
        End If
    ElseIf mIntWindMode = 1 Then
        mlngID = lngID
        Unload Me
    End If
End Sub

Private Sub picPati_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index <> mlngSelIndex Then
        If mlngSelIndex >= 0 Then
            If lblName(mlngSelIndex).Tag <> "" Then
                picPati(mlngSelIndex).Picture = imgCardBack(1).Picture
            Else
                picPati(mlngSelIndex).Picture = imgCardBack(1).Picture
            End If
        End If
        mlngSelIndex = Index
        If lblName(mlngSelIndex).Tag <> "" Then
            picPati(mlngSelIndex).Picture = imgCardBack(0).Picture
        Else
            picPati(mlngSelIndex).Picture = imgCardBack(0).Picture
        End If
        
        Call AdjustPatiInfo(mlngSelIndex)
    End If
End Sub

Private Sub AdjustPatiInfo(ByVal Index As Long)
    If Index < 0 Then Exit Sub
    
    If mIntWindMode = 1 Then Exit Sub

    rsPati.Move Index, adBookmarkFirst
    lblPatiName.Caption = rsPati!���� & ""
    lblPatiSex.Caption = rsPati!�Ա� & ""
    lblPatiAge.Caption = rsPati!���� & ""
    If rsPati!��Դ & "" = "סԺ" Then
        lalPatiInfoNo.Caption = "סԺ�ţ�"
    Else
        lalPatiInfoNo.Caption = "����ţ�"
    End If
    lalPatiNo.Caption = rsPati!��ʶ�� & ""
    lalPatiDept.Caption = rsPati!���� & ""
    lalPatiDisease.Caption = rsPati!���Ƽ��� & ""
    
    Call ReadPatPricture(Val(rsPati!����ID), imgPati)
End Sub

Private Sub picPati_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'���ܣ��Ҽ��˵�
    If Button = 2 Then
        If Not mobjBarPopup Is Nothing Then
            mobjBarPopup.ShowPopup
        End If
    End If
End Sub

Private Sub picPatiList_Resize()
On Error Resume Next
    HScr.Move picPatiList.ScaleWidth - HScr.Width, 0, HScr.Width, picPatiList.ScaleHeight
    If Me.Visible Then Call AdjustCardsPosition
End Sub

Public Sub ReadPatPricture(ByVal lng����ID As Long, ByRef imgPatient As Image)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ������Ƭ
    '������lng����ID=��ȡָ�����˵���Ƭ
    '           imgPatient=��Ƭ����λ��
    '           strFile=��Ƭ�ı���·��
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFile As String
    On Error GoTo ErrHand
    strFile = ""
    strFile = gobjComlib.sys.Readlob(glngSys, 27, lng����ID, strFile)
    If strFile <> "" Then
        imgPatient.Picture = Nothing
        imgPatient.Picture = LoadPicture(strFile)
        Kill strFile
    Else
        imgPatient.Picture = imgPatiPhoto.Picture
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub InitCommandBar()
'���ܣ������ڲ˵����岿��
'˵����
'1.���й��еĲ˵��Ͱ�ť�����У���Ϊ�Ӵ��崦��˵��Ļ�׼
'2.�����������������ҵ��Ĳ�ͬ�����ܲ�ͬ
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim strFunName As String
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
    End With
    cbsMain.EnableCustomization False
    cbsMain.ActiveMenuBar.Visible = True
    Set cbsMain.Icons = gobjComlib.ZLCommFun.GetPubIcons
    
    '�˵�����
    '-----------------------------------------------------
    '�����Ҽ��˵�
    Set mobjBarPopup = cbsMain.Add("Popup", xtpBarPopup)
    With mobjBarPopup.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_DiseaseRegist, "�鿴������")
        Set objControl = .Add(xtpControlButton, conMenu_File_Modify, "�޸�")
        Set objControl = .Add(xtpControlButton, conMenu_File_Delete, "ɾ��")
    End With
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup           '��xtpControlPopup���͵�����ID�����¸�ֵ
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Modify, "�޸�(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Delete, "ɾ��(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
        objControl.BeginGroup = True
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_View_ToolBar, "������(&T)")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_View_ToolBar_Button, "��׼��ť(&S)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Text, "�ı���ǩ(&T)", -1, False
            .Add xtpControlButton, conMenu_View_ToolBar_Size, "��ͼ��(&B)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_View_StatusBar, "״̬��(&S)")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_View_DiseaseRegist, "�鿴������(&D)")
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)"): objControl.BeginGroup = True
    End With

    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�" & gstrProductName)
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, gstrProductName & "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, gstrProductName & "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
    End With
    
    '����������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False                   '�������ϵ������Ҽ�ʱ���������ò˵�
    objBar.ShowTextBelowIcons = False                   '�������еİ�ť������ʾ��ͼ���Ҳ�
    objBar.EnableDocking xtpFlagHideWrap                '��������Ȳ���ʱҲ������
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_DiseaseRegist, "�鿴������")
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_File_Modify, "�޸�")
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_File_Delete, "ɾ��")
        objControl.Style = xtpButtonIconAndCaption
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        objControl.Style = xtpButtonIconAndCaption
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        objControl.Style = xtpButtonIconAndCaption
        objControl.IconId = 191
        objControl.BeginGroup = True
    End With
    
    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
    End With
End Sub


Private Sub GetRegistDept()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    '��ȡ�Ǽǿ���
    strSQL = "Select a.Id,a.���� From ���ű� A Where ID = [1] And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null)"
    Set rsTmp = gobjComlib.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng�Ǽǿ���ID)
    
    If rsTmp.RecordCount > 0 Then
        lblRegistDept.Caption = rsTmp!���� & ""
    Else
        lblRegistDept.Caption = ""
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DeleteDiseaseRegist(ByVal lngID As Long)
'����: ɾ�����Խ��������
'����: lngID - ������ID
    Dim strSQL As String
    Dim intErrCode As Integer
    Dim strMsg As String
    
    On Error GoTo errH
  
    If CheckOperateState(lngID, intErrCode) Then
        If MsgBox("ɾ��֮�󲻿ɻָ���ȷ��Ҫɾ���÷�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            strSQL = "Zl_�������Լ�¼_Delete(" & lngID & ")"
            Call gobjComlib.zlDatabase.ExecuteProcedure(strSQL, "ɾ�����Խ��������")
            Call LoadPati(rsPati)
        Else
            Exit Sub
        End If
    Else
        If intErrCode = 1 Then
            strMsg = "û�в�ѯ���÷�������"
        ElseIf intErrCode = 2 Then
            strMsg = "������ɾ�����˵ķ�������"
        ElseIf intErrCode = 3 Then
            strMsg = "ҽ���Ѿ������˸÷�����������ɾ����"
        End If
        MsgBox strMsg, vbInformation, gstrSysName
    End If
     
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ModifyDiseaseRegist(ByVal lngID As Long)
'����: �޸����Խ��������
'����: lngID - ������ID
    Dim strSQL As String
    Dim intErrCode As Integer
    Dim strMsg As String
    
    On Error GoTo errH
  
    If CheckOperateState(lngID, intErrCode) Then
        Call frmDiseaseRegist.ShowDiseaseRegist(Me, 3, lngID)
    Else
        If intErrCode = 1 Then
            strMsg = "û�в�ѯ���÷�������"
        ElseIf intErrCode = 2 Then
            strMsg = "�������޸����˵ķ�������"
        ElseIf intErrCode = 3 Then
            strMsg = "ҽ���Ѿ������˸÷������������޸ġ�"
        End If
        MsgBox strMsg, vbInformation, gstrSysName
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
