VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSquareSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��ѡ��ָ���Ŀ����"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5415
   Icon            =   "frmSquareSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3870
      TabIndex        =   1
      Top             =   5880
      Width           =   1290
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2490
      TabIndex        =   0
      Top             =   5895
      Width           =   1290
   End
   Begin MSComctlLib.ImageList imgLvw 
      Left            =   3960
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSquareSelect.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw���� 
      Height          =   5745
      Left            =   0
      TabIndex        =   2
      Top             =   30
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   10134
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgLvw"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "����"
         Object.Width           =   5292
      EndProperty
   End
End
Attribute VB_Name = "frmSquareSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnStart As Boolean, mlng�����ID As Long, mbln���ѿ� As Boolean
Private mblnOk As Boolean
Private mstrSelect As String
Private mblnUnLoad As Boolean
Private mobjOneDataObject As clsOneCardDataObject
Private Sub cmdȡ��_Click()
    mbln���ѿ� = False: mlng�����ID = 0: Unload Me
End Sub
Private Sub cmdȷ��_Click()
    If lvw����.SelectedItem Is Nothing Then Exit Sub
    mlng�����ID = Val(Mid(lvw����.SelectedItem.Key, 2)):
    mbln���ѿ� = Left(lvw����.SelectedItem.Key, 1) = "X": mblnOk = True
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
    If Not mblnStart Then Unload Me: Exit Sub
    If lvw����.ListItems.count = 1 Then cmdȷ��_Click
End Sub
Private Sub Form_Load()
    Dim lvwItem As ListItem, objCard As clsCard, I As Long
    Dim objCardBrush As clsBrushSequareCard
    Dim objCardInterface As Object
    Dim rsTemp As ADODB.Recordset
    Dim objYLCards As Cards
    Dim objYlCardObjs As Cards
    '59760
    If mobjOneDataObject.zlGetCards_YL(objYLCards) = False Then Exit Sub
    If mobjOneDataObject.zlGetYLCardObjs(objYlCardObjs) = False Then Exit Sub
    
    mblnUnLoad = False
    lvw����.ListItems.Clear
    With lvw����
        For I = 1 To objYlCardObjs.count
            '����:48005
            If Not (objYlCardObjs(I).���ѿ� And objYlCardObjs(I).���ƿ�) Or (objYlCardObjs(I).���ƿ� And objYlCardObjs(I).�ӿڳ����� <> "") Then
                If Not (objYlCardObjs(I).���ѿ� = False And (objYlCardObjs(I).���ƿ� Or Not objYlCardObjs(I).�Ƿ�����ʻ�)) Or (objYlCardObjs(I).���ƿ� And objYlCardObjs(I).�ӿڳ����� <> "") Then
                    Set lvwItem = .ListItems.Add(, IIf(objYlCardObjs(I).���ѿ�, "X", "K") & objYlCardObjs(I).�ӿ����, objYlCardObjs(I).�ӿڱ���, , 1)
                    lvwItem.SubItems(1) = objYlCardObjs(I).����
                End If
            End If
        Next
        If .ListItems.count = 0 Then
            MsgBox "�����ڵ������ӿ�,����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
            mblnUnLoad = True
            Exit Sub
        End If
        .ListItems(1).Selected = True
    End With
    mblnStart = True
End Sub

Public Function zlShowSelect(ByVal frmMain As Object, ByVal objOneDataObject As clsOneCardDataObject, ByRef lng�����ID As Long, Optional ByRef bln���ѿ� As Boolean) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�ѡ����㿨�ӿ�
    '��Σ�frmMain-�����������
    '���Σ�lng�����ID-ѡ��Ŀ����ID
    '          bln���ѿ�-�Ƿ����ѿ�
    '���أ�ѡ��ɹ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-06-18 11:23:19
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim objYlCardObjs As Cards
    
    Set mobjOneDataObject = objOneDataObject
    '59760
    If mobjOneDataObject.zlGetYLCardObjs(objYlCardObjs) = False Then Exit Function
    mblnOk = False
    If objYlCardObjs.count = 1 Then
        lng�����ID = Val(objYlCardObjs(1).�ӿ����)
        zlShowSelect = True
        Exit Function
    End If
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    lng�����ID = mlng�����ID
    bln���ѿ� = mbln���ѿ�
    zlShowSelect = mblnOk
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mobjOneDataObject Is Nothing Then Set mobjOneDataObject = Nothing
End Sub

Private Sub lvw����_DblClick()
    Call cmdȷ��_Click
End Sub
Private Sub lvw����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Call cmdȷ��_Click
End Sub


