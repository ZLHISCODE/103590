VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm��ѯδ�ϴ�������ϸ 
   Caption         =   "δ�ϴ��Ĵ�����ϸ�嵥������"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "frm��ѯδ�ϴ�������ϸ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   8910
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
      Left            =   7620
      TabIndex        =   0
      Top             =   4920
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   4785
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
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
      MouseIcon       =   "frm��ѯδ�ϴ�������ϸ.frx":0E42
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm��ѯδ�ϴ�������ϸ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mintInsure As Integer

Public Function ShowME(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intinsure As Integer) As Boolean
    '���������δ�ϴ��Ĵ�����ϸ�򷵻ؼ٣�ͬʱ��ʾ������Ա�����
    On Error Resume Next
    mblnOK = False
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mintInsure = intinsure
    Me.Show 1
    ShowME = mblnOK
End Function

Private Sub cmdExcel_Click()
    '�����EXCEL
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    Dim bytStyle As Byte
    
    intRow = mshList.Row
    bytStyle = 3
    
    '��ͷ
    objOut.Title.Text = "δ�ϴ��Ĵ�����ϸ�嵥-" & mlng����ID
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
    Dim rsTemp As New ADODB.Recordset
    On Error Resume Next
    
    '��ȡָ������ָ��סԺ������δ�ϴ���ϸ
    gstrSQL = "Select DECODE(A.��¼����,3,'�Զ�����','����') AS ����,DECODE(A.��¼״̬,2,'����','����') AS ����,A.No,A.���,E.����," & _
             " trim(to_char(Nvl(A.����,0)*Nvl(A.����,1),'90009990.00')) As ����,trim(to_char(A.��׼����,'90009990.00')) AS ��׼����," & _
             " trim(to_char(A.ʵ�ս��,'90009990.00')) AS ʵ�ս��,F.��Ŀ���� AS ҽ������" & _
             " From סԺ���ü�¼ A,������Ϣ B,������ҳ C,�����ʻ� D,�շ�ϸĿ E,����֧����Ŀ F " & _
             " Where A.����ID=B.����ID And B.����ID=C.����ID And A.��ҳID=C.��ҳID And A.����ID=D.����ID And D.����=" & mintInsure & _
             " And Nvl(��¼״̬,0)<>0 And Nvl(���ӱ�־,0)<>9 And Nvl(ʵ�ս��,0)<>0 And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(���ʷ���,0)=1 " & _
             " And (Nvl(A.�����־,0)<>1 And Nvl(A.�����־,0)<>4)" & _
             " And A.�շ�ϸĿID=E.Id And E.ID=F.�շ�ϸĿID(+) And F.����(+)=" & mintInsure & _
             " And A.����ID=[1] And A.��ҳID=[2]" & _
             " Order By A.�Ǽ�ʱ��,No,���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ������ָ��סԺ������δ�ϴ���ϸ", mlng����ID, mlng��ҳID)
    If rsTemp.RecordCount = 0 Then
        mblnOK = True
        Unload Me
        Exit Sub
    End If
    
    Set mshList.DataSource = rsTemp
    With mshList
        .ColWidth(0) = 660
        .ColWidth(1) = 495
        .ColWidth(2) = 810
        .ColWidth(3) = 495
        .ColWidth(4) = 2070
        .ColWidth(5) = 1035
        .ColWidth(6) = 1035
        .ColWidth(7) = 990
        .ColWidth(8) = 1200
        .ColAlignment(4) = 1
        .ColAlignment(5) = 7
        .ColAlignment(6) = 7
        .ColAlignment(7) = 7
        .ColAlignment(8) = 1
    End With
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
