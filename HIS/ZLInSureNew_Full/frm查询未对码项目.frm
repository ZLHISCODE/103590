VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm��ѯδ������Ŀ 
   Caption         =   "��ѯδ������Ŀ"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   Icon            =   "frm��ѯδ������Ŀ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8280
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdExcel 
      Caption         =   "���&EXCEL"
      Height          =   350
      Left            =   150
      TabIndex        =   2
      Top             =   4920
      Width           =   1100
   End
   Begin VB.CommandButton cmd�˳� 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6900
      TabIndex        =   1
      Top             =   4920
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   8440
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frm��ѯδ������Ŀ.frx":0E42
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm��ѯδ������Ŀ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer

Public Sub ShowME(ByVal objParent As Object, ByVal intinsure As Integer)
    mintInsure = intinsure
    Me.Show 1, objParent
End Sub

Private Sub cmdExcel_Click()
    '�����EXCEL
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    Dim bytStyle As Byte
    
    intRow = mshList.Row
    bytStyle = 3
    
    '��ͷ
    objOut.Title.Text = "δ������Ŀ�嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.COL = 0: mshList.ColSel = mshList.Cols - 1
    mshList.Redraw = True
End Sub

Private Sub cmd�˳�_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    '��ʾָ��ҽ������δ�������Ŀ
    '��ʽ���շ�ϸĿID|�շ����|��Ŀ����|��Ŀ����|����|���|����ʱ��
    Dim i As Integer, j As Integer
    Dim rsTemp As New ADODB.Recordset
    
    With mshList
        .Cols = 6
        .TextMatrix(0, 0) = "�շ�ϸĿID"
        .TextMatrix(0, 1) = "�շ����"
        .TextMatrix(0, 2) = "��Ŀ����"
        .TextMatrix(0, 3) = "��Ŀ����"
        .TextMatrix(0, 4) = "���"
        .TextMatrix(0, 5) = "����ʱ��"
        
        j = .Cols - 1
        For i = 0 To j
            .ColAlignmentFixed(i) = 4
        Next
        .ColAlignment(2) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColWidth(0) = 0
        .ColWidth(1) = 810
        .ColWidth(2) = 1050
        .ColWidth(3) = 2130
        .ColWidth(4) = 2820
    End With
    
    '��ȡ����δ������Ŀ
    gstrSQL = "Select C.ID As �շ�ϸĿID,B.��� As �շ����,C.���� As ��Ŀ����,C.���� As ��Ŀ����,DECODE(C.���,'��','',C.���) AS ���,C.����ʱ�� " & _
             " From  " & _
             " (Select ID As �շ�ϸĿID " & _
             " From �շ�ϸĿ " & _
             " Minus  " & _
             " Select �շ�ϸĿID " & _
             " From ����֧����Ŀ " & _
             " Where ����=[1]) A,�շ���� B,�շ�ϸĿ C " & _
             " Where A.�շ�ϸĿID=C.Id And B.����=C.��� " & _
             " And (C.����ʱ�� Is NULL Or to_char(C.����ʱ��,'yyyy-MM-dd')='3000-01-01')" & _
             " Order By C.���,C.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����δ������Ŀ", mintInsure)
    If rsTemp.RecordCount = 0 Then Exit Sub
    Set mshList.DataSource = rsTemp
    mshList.ColWidth(0) = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With cmd�˳�
        .Left = Me.ScaleWidth - .Width - 150
        .Top = Me.ScaleHeight - .Height - 150
    End With
    cmdExcel.Top = cmd�˳�.Top
    
    With mshList
        .Height = cmd�˳�.Top - 150
        .Width = Me.ScaleWidth
    End With
End Sub
