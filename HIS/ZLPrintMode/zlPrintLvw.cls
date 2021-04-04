VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlPrintLvw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarTitle As zlTabTitle
Private mvarUnderAppItems As zlTabAppRow
Private mvarBelowAppItems As zlTabAppRow
Private mvarBody As zlLvwBody

Private mvarAppFont As StdFont
Private mvarAppColor As Long
Private mvarGridLines As Byte
Private mvarGridColor As Long

Private mvarEmptyUp As Single '���ظ���
Private mvarEmptyLeft As Single '���ظ���
Private mvarEmptyDown As Single '���ظ���
Private mvarRowSpaceRate As Single '���ظ���
'��������ֵ�ľֲ�����
Private mvarFixRow As Integer '�ֲ�����
Private mvarFixCol As Integer '�ֲ�����
'��������ֵ�ľֲ�����
Private mvarEmptyRight As Single '�ֲ�����
'��������ֵ�ľֲ�����
Private mvarPageHeader As Single '�ֲ�����
Private mvarPageFooter As Single '�ֲ�����
'��������ֵ�ľֲ�����
Private mvarHeader As String '�ֲ�����
Private mvarFooter As String '�ֲ�����
Private mvarPrivileged As Boolean '�ֲ�����

Public Property Get Privileged() As Boolean
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.Privileged
    Privileged = mvarPrivileged
End Property

Public Property Let Privileged(ByVal vData As Boolean)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.Privileged = True
    mvarPrivileged = vData
End Property

Public Property Let Footer(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Footer = 5
    mvarFooter = vData
End Property


Public Property Get Footer() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Footer
    Footer = mvarFooter
End Property



Public Property Let Header(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Header = 5
    mvarHeader = vData
End Property


Public Property Get Header() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Header
    Header = mvarHeader
End Property



Public Property Let PageFooter(ByVal vData As Single)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageFooter = 5
    mvarPageFooter = vData
End Property


Public Property Get PageFooter() As Single
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageFooter
    PageFooter = mvarPageFooter
End Property



Public Property Let PageHeader(ByVal vData As Single)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageHeader = 5
    mvarPageHeader = vData
End Property


Public Property Get PageHeader() As Single
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageHeader
    PageHeader = mvarPageHeader
End Property



Public Property Let EmptyRight(ByVal vData As Single)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.EmptyRight = 5
    mvarEmptyRight = vData
End Property


Public Property Get EmptyRight() As Single
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.EmptyRight
    EmptyRight = mvarEmptyRight
End Property



Public Property Let FixCol(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FixCol = 5
    mvarFixCol = vData
End Property


Public Property Get FixCol() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FixCol
    FixCol = mvarFixCol
End Property



Public Property Let FixRow(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FixRow = 5
    mvarFixRow = vData
End Property


Public Property Get FixRow() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FixRow
    FixRow = mvarFixRow
End Property




Public Property Get AppFont() As StdFont
    Set AppFont = mvarAppFont
End Property

Public Property Set AppFont(vData As StdFont)
    Set mvarAppFont = vData
End Property

Public Property Let AppColor(ByVal vData As Long)
    mvarAppColor = vData
End Property

Public Property Get AppColor() As Long
    AppColor = mvarAppColor
End Property

Public Property Let GridLines(ByVal vData As Byte)
    mvarGridLines = vData
End Property

Public Property Get GridLines() As Byte
    GridLines = mvarGridLines
End Property

Public Property Let GridColor(ByVal vData As Long)
    mvarGridColor = vData
End Property

Public Property Get GridColor() As Long
    GridColor = mvarGridColor
End Property


Public Property Let RowSpaceRate(ByVal vData As Single)
    mvarRowSpaceRate = vData
End Property

Public Property Get RowSpaceRate() As Single
    RowSpaceRate = mvarRowSpaceRate
End Property


Public Property Let EmptyDown(ByVal vData As Single)
    mvarEmptyDown = vData
End Property


Public Property Get EmptyDown() As Single
    EmptyDown = mvarEmptyDown
End Property


Public Property Let EmptyLeft(ByVal vData As Single)
    mvarEmptyLeft = vData
End Property


Public Property Get EmptyLeft() As Single
    EmptyLeft = mvarEmptyLeft
End Property


Public Property Let EmptyUp(ByVal vData As Single)
    mvarEmptyUp = vData
End Property


Public Property Get EmptyUp() As Single
    EmptyUp = mvarEmptyUp
End Property


Public Property Get Body() As zlLvwBody
    Set Body = mvarBody
End Property


Public Property Set Body(vData As zlLvwBody)
    Set mvarBody = vData
End Property

Public Property Get BelowAppItems() As zlTabAppRow
    If mvarBelowAppItems Is Nothing Then
        Set mvarBelowAppItems = New zlTabAppRow
    End If
    Set BelowAppItems = mvarBelowAppItems
End Property


Public Property Set BelowAppItems(vData As zlTabAppRow)
    Set mvarBelowAppItems = vData
End Property

Public Property Get UnderAppItems() As zlTabAppRow
    If mvarUnderAppItems Is Nothing Then
        Set mvarUnderAppItems = New zlTabAppRow
    End If
    Set UnderAppItems = mvarUnderAppItems
End Property


Public Property Set UnderAppItems(vData As zlTabAppRow)
    Set mvarUnderAppItems = vData
End Property


Public Property Get Title() As zlTabTitle
    Set Title = mvarTitle
End Property


Public Property Set Title(vData As zlTabTitle)
    Set mvarTitle = vData
End Property


Private Sub Class_Initialize()
    '������ zlPrintLvw ��ʱ������ mTitle ����
    Set mvarTitle = New zlTabTitle
    '������ zlPrintLvw ��ʱ������ mBody ����
    Set mvarBody = New zlLvwBody
    '������ zlPrintLvw ��ʱ������ mAppFont����
    Set mvarAppFont = New StdFont
    mvarAppFont.Name = "����"
    mvarAppFont.Size = 9
    mvarAppFont.Bold = False
    mvarAppFont.Italic = False
    '������ zlPrintLvw ��ʱ������ mUnderAppItems����
    Set mvarUnderAppItems = New zlTabAppRow
    '������ zlPrintLvw ��ʱ������ mBelowAppItems����
    Set mvarBelowAppItems = New zlTabAppRow
    '������ zlPrintLvw ��ʱ�������������
    mvarAppColor = 0
    mvarGridLines = 1
    mvarGridColor = 0
    mvarEmptyUp = 25
    mvarEmptyDown = 25
    mvarEmptyLeft = 19
    mvarEmptyRight = 19
    mvarPageHeader = 13
    mvarPageFooter = 13
    mvarRowSpaceRate = 1.1
    mvarFooter = ";;"
    mvarHeader = ";;"
    mvarPrivileged = False
End Sub


Private Sub Class_Terminate()
    Set mvarBody = Nothing
    Set mvarBelowAppItems = Nothing
    Set mvarUnderAppItems = Nothing
    Set mvarTitle = Nothing
    Set mvarAppFont = Nothing
End Sub
