VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserDatePickerDayMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'��������ֵ�ľֲ�����
Private mvarBackColor As OLE_COLOR '�ֲ�����
Private mvarForeColor As OLE_COLOR '�ֲ�����
Private mvarIsHoliday As Boolean '�ֲ�����
Private mvarIsWorkFromHoliday As Boolean '�ֲ�����
Private mvarFontBold As Boolean '�ֲ�����
Public Property Let FontBold(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FontBold = 5
    mvarFontBold = vData
End Property


Public Property Get FontBold() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FontBold
    FontBold = mvarFontBold
End Property


Public Property Let IsWorkFromHoliday(ByVal vData As Boolean)
Attribute IsWorkFromHoliday.VB_Description = "�Ƿ��ǻ����ա�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.IsWorkFromHoliday = 5
    mvarIsWorkFromHoliday = vData
End Property


Public Property Get IsWorkFromHoliday() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.IsWorkFromHoliday
    IsWorkFromHoliday = mvarIsWorkFromHoliday
End Property


Public Property Let IsHoliday(ByVal vData As Boolean)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.IsHoliday = 5
    mvarIsHoliday = vData
End Property


Public Property Get IsHoliday() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.IsHoliday
    IsHoliday = mvarIsHoliday
End Property


Public Property Let ForeColor(ByVal vData As OLE_COLOR)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ForeColor = vbRed
    mvarForeColor = vData
End Property


Public Property Get ForeColor() As OLE_COLOR
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ForeColor
    ForeColor = mvarForeColor
End Property


Public Property Let BackColor(ByVal vData As OLE_COLOR)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.BackColor = vbRed
    mvarBackColor = vData
End Property


Public Property Get BackColor() As OLE_COLOR
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.BackColor
    BackColor = mvarBackColor
End Property


