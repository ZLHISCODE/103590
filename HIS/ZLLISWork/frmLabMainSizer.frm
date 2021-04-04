VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmLabMainSizer 
   BackColor       =   &H00FDD6C6&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   3495
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   1965
      _Version        =   589884
      _ExtentX        =   3466
      _ExtentY        =   6165
      _StockProps     =   0
      ShowHeader      =   0   'False
   End
   Begin XtremeSuiteControls.ShortcutCaption ShortCaption 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2385
      _Version        =   589884
      _ExtentX        =   4207
      _ExtentY        =   503
      _StockProps     =   6
      Caption         =   "ɸѡ����"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.01
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   1
      Alignment       =   1
   End
End
Attribute VB_Name = "frmLabMainSizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mstrName As String
Private Const con_������ɸѡ_������ As String = "���ﲡ��;סԺ����;�����걾;����걾;δ��걾;��첡��;����ҽ��;�����걾;�ʿر걾;�����ͨ��;���δͨ��;δ����;������;�������ͨ��;�������δͨ��"

Private Const con_������ɸѡ_������ As String = "���ﲡ��;סԺ����;��첡��"

Private Const con_frmLisStationWrite As String = "�鿴����;�鿴ԭʼ���;�鿴�ϴν��;�鿴��־;�鿴��λ;�鿴�ο�;�鿴ø��;������ʾ;������˱�ʶ"
Private Enum mCol
    ѡ�� = 0
    ����
End Enum

Private Sub showData()
    Dim intLoop As Integer
    Dim lngLoop As Long
    Dim astrName() As String
    Dim Record As ReportRecord
    Dim strCheck As Boolean
    
    With Me.rptList.Columns
        
        rptList.AllowColumnRemove = False
        rptList.ShowItemsInGroups = False
        Me.rptList.Records.DeleteAll
        With rptList.PaintManager
            
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ����Ŀ..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptList.SetImageList Imglist
        Set Column = .Add(mCol.ѡ��, "", 30, False)
        Column.Icon = 0
        Set Column = .Add(mCol.����, "����", 120, False)
    End With
    
    Select Case mstrName
        Case "������"
            astrName = Split(con_������ɸѡ_������, ";")
        Case "������"
            astrName = Split(con_������ɸѡ_������, ";")
        Case "frmLisStationWrite"
            astrName = Split(con_frmLisStationWrite, ";")
    End Select
    
'    astrName = Split(mstrName, ";")
    
    For lngLoop = 0 To UBound(astrName)
        Set Record = Me.rptList.Records.Add
        For intLoop = 0 To Me.rptList.Columns.Count - 1
            Record.AddItem ""
        Next
        Record(mCol.ѡ��).HasCheckbox = True
        Record(mCol.����).Value = astrName(lngLoop)
        strCheck = zlDatabase.GetPara(mstrName & "_" & astrName(lngLoop), 100, 1208)
        Record(mCol.ѡ��).Checked = strCheck
    Next
    Me.rptList.Populate

End Sub

Public Sub ShowME(Objfrm As Object, strName As String, blnShow As Boolean)
    mstrName = strName
    If blnShow = True Then
        Unload Me
    Else
        showData
        Me.Show modal, Objfrm
    End If
End Sub

Private Sub Form_Resize()
    Me.ShortCaption.Top = 50
    Me.ShortCaption.Left = 50
    Me.ShortCaption.Width = Me.ScaleWidth - 100
    Me.rptList.Top = Me.ShortCaption.Top + Me.ShortCaption.Height
    Me.rptList.Left = 50
    Me.rptList.Width = Me.ScaleWidth - 100
    Me.rptList.Height = Me.ScaleHeight - (Me.ShortCaption.Top + Me.ShortCaption.Height) - 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim intLoop As Integer
    '��������б�
    
    For intLoop = 0 To Me.rptList.Rows.Count
        zlDatabase.SetPara mstrName & "_" & Me.rptList.Rows(intLoop).Record(mCol.����).Value, _
            Me.rptList.Rows(intLoop).Record(mCol.ѡ��).Checked, 100, 1208
    Next
    mstrName = ""
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Row.Record(mCol.ѡ��).Checked = Not Row.Record(mCol.ѡ��).Checked
    Me.rptList.Populate
End Sub



