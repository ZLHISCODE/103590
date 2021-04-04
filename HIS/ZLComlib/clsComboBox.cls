VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'##############################################################################
'#功能：ComboBox相关的处理、判断等公共函数
'命名规则：函数过程名称为通俗易懂的英文即可，不用加Cbo关键字
'上级类：clsControl
'##############################################################################

Public Sub SetListWidth(ByVal cboControl As Variant, ByVal lngWidth As Long)
'功能：设置Combo控件下拉列表的宽度
'此处的宽度是批下拉列表的宽度，并且是以TWIP为单位
    If IsObject(cboControl) Then
        Call gobjComLib.zlControl.CboSetWidth(cboControl.hWnd, lngWidth)
    Else
        Call gobjComLib.zlControl.CboSetWidth(cboControl, lngWidth)
    End If
End Sub

Public Sub SetListWidthAuto(cboControl As Object, Optional ByVal dblMaxRate As Double)
'功能：自动根据Combo控件下拉列表内容设置下拉列表宽度
'参数：cboControl=自动设置下拉列表的控件
'          intMaxRate=上限倍率，自动设置后的下拉列表的宽度范围为cboControl.Width<= 下拉列表宽度<=cboControl.Width*intMaxRate
'                               <2表示不限宽度
'说明：该函数需在加载ComboBox下拉列表项后调用
    Dim i As Long
    Dim lngTmp As Long, lngMaxLen As Long
    Dim lngWidth As Long
    Dim lngOneChar As Long
    
    For i = 0 To cboControl.ListCount - 1
        lngTmp = gobjComLib.zlStr.ActualLen(cboControl.List(i))
        If lngTmp > lngMaxLen Then
            lngMaxLen = lngTmp
        End If
    Next
    If lngMaxLen <> 0 Then
        If dblMaxRate < 1 Then dblMaxRate = 0
        lngOneChar = gobjComLib.zlControl.OneCharWidth(cboControl.Font)
        lngWidth = CLng((lngMaxLen + 4) * lngOneChar)   '计算宽度,加4长度作为滚动条与边界的值
        If lngWidth < cboControl.Width Then '低于最小宽度，则为最小宽度
            lngWidth = cboControl.Width
        ElseIf dblMaxRate <> 0 And lngWidth > cboControl.Width * dblMaxRate Then '超过设定的最大倍率，则为最大宽度
            lngWidth = CLng(cboControl.Width * dblMaxRate)
        End If
    Else
        lngWidth = cboControl.Width '默认与控件宽度一致
    End If
    Call SetListWidth(cboControl.hWnd, lngWidth)
End Sub

Public Sub SetText(cboControl As Object, ByVal strText As String, Optional ByVal blnAfter As Boolean = True, Optional strSplit As String = "-")
'功能：根据文本串更新Combo控件的当前值
'参数：cboControl  准备设置的ComboBox控件
'      strText     输入的文本串
'      blnAfter    表示在分隔符之前或之后取值。如果没有分隔符，则取之后
'      strSplit    分隔符，通常为-
    Call gobjComLib.zlControl.CboSetText(cboControl, strText, blnAfter, strSplit)
End Sub

Public Sub SeekIndex(objCbo As Object, varFind As Variant, Optional Keep As Boolean, Optional ByVal blnOnlyExact As Boolean)
'功能：由字符串在ComboBox中查找索引
'参数：Keep=如果未匹配，是否保持原索引
'          blnOnlyExact:只精确匹配，该参数只针对字符
    Dim i As Integer
    i = FindIndex(objCbo, varFind, blnOnlyExact)
    If i <> -1 Then
        objCbo.ListIndex = i
    ElseIf Not Keep Then
        objCbo.ListIndex = -1
    End If
End Sub

Public Sub SetIndex(ByVal hWnd_combo As Long, ByVal lngIndex As Long)
'功能：设置Combo控件的Index值
'为一个Combo控件选择列表项，但又不触发其Click事件
    Call gobjComLib.zlControl.CboSetIndex(hWnd_combo, lngIndex)
End Sub

Public Sub SetListHeight(cboControl As Object, ByVal lngHeight As Long)
'功能：设置Combo控件的下拉列表的高度
'此处的宽度是批下拉列表的高度，并且是以TWIP为单位
    Call gobjComLib.zlControl.CboSetHeight(cboControl, lngHeight)
End Sub

Public Function MatchIndex(ByVal lnghwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：lngHwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：匹配的索引(含不匹配的索引)
'说明：只返回索引,但未设置索引。请将该函数在KeyPress事件中调用。
     MatchIndex = gobjComLib.zlControl.CboMatchIndex(lnghwnd, KeyAscii, sngInterval)
End Function

Public Function Locate(ByVal cboObj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
'建议弃用，使用Cbo.SeekIndex代替
'blnItem:True-表示根据ItemData的值定位下拉框;False-表示根据文本的内容定位下拉框
    Locate = gobjComLib.zlControl.CboLocate(cboObj, strValue, blnItem)
End Function

Public Sub LoadFromList(ByRef arrCbo As Variant, ByVal varList As Variant, Optional ByVal intDefault As Integer = -1, Optional ByVal strSplit As String = ",")
'功能：将指定数据装入指定ComboBox
'参数：arrList=List String数组或者以逗号分割的字符串
'      arrCbo=一个ComboBox对象，或存放ComboBox对象的数组
'      intDefaut=缺省索引
'     strSplit=传入字符串时，项目分隔符
    Dim i As Long, j As Long
    Dim arrTmp As Variant
    Dim arrCboTmp As Variant
    If TypeName(arrCbo) = "Variant()" Then
        arrCboTmp = arrCbo
    Else
        arrCboTmp = Array(arrCbo)
    End If
    
    If TypeName(varList) = "String" Then
        arrTmp = Split(varList, strSplit)
        For i = 0 To UBound(arrCboTmp)
            arrCboTmp(i).Clear
            For j = 0 To UBound(arrTmp)
                arrCboTmp(i).AddItem arrTmp(j)
            Next
            arrCboTmp(i).ListIndex = intDefault '缺省为未选中
        Next
    Else
        For i = 0 To UBound(arrCboTmp)
             arrCboTmp(i).Clear
            For j = 0 To UBound(varList)
                arrCboTmp(i).AddItem varList(j)
            Next
            arrCboTmp(i).ListIndex = intDefault '缺省为未选中
        Next
    End If
End Sub

Public Function ListWidth(ByVal hWnd_combo As Long) As Long
'功能： 取得 Combo 下拉的宽度,以TWIP为单位
    Dim lRetVal As Long
    lRetVal = SendMessage(hWnd_combo, &H15F, 0, 0)
    If lRetVal <> -1 Then
        ListWidth = lRetVal * Screen.TwipsPerPixelX
    Else
        ListWidth = 0
    End If
End Function

Public Function ItemExists(objControl As Object, ByVal strList As String, ByVal lngData As Long, Optional ByVal intCheckType As Integer = 2) As Boolean
'功能：判断指定项目在列表中是否已经存在
'说明：相同项目指Text及ItemData都相同
'          intCheckType=2-检查项目的Text及ItemData是否都相同,0-仅检查项目的Text是否相同，1-仅检查项目的ItemData是否相同
    Dim i As Long
    Dim blnFind As Boolean
    
    For i = 0 To objControl.ListCount - 1
        blnFind = True
        If intCheckType <> 1 Then
            If objControl.List(i) = strList Then
                blnFind = False
            End If
        End If
        If intCheckType <> 0 And blnFind Then
            If objControl.ItemData(i) = lngData Then
                blnFind = False
            End If
        End If
        If blnFind Then
            ItemExists = True: Exit For
        End If
    Next
End Function

Public Function FindIndex(objCbo As Object, varData As Variant, Optional ByVal blnOnlyExact As Boolean) As Long
'功能：由ItemData或Text查找ComboBox的索引值
'blnOnlyExact:只精确匹配，该参数只针对字符
    Dim strType As String, i As Integer
    
    FindIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '先精确查找
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    FindIndex = i: Exit Function
                ElseIf gobjComLib.zlStr.NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    FindIndex = i: Exit Function
                End If
            Next
            If Not blnOnlyExact Then
                '再模糊查找
                For i = 0 To objCbo.ListCount - 1
                    If InStr(objCbo.List(i), varData) > 0 And varData <> "" Then
                        FindIndex = i: Exit Function
                    End If
                Next
            End If
        End If
    Else
        If varData <> 0 Then
            For i = 0 To objCbo.ListCount - 1
                If objCbo.ItemData(i) = varData Then
                    FindIndex = i: Exit Function
                End If
            Next
        End If
    End If
End Function

Public Sub AppendText(cboControl As Object, KeyAscii As Integer)
'功能：对ComboBox实现输入过程中自动完成的功能
'说明：在Combox.KeyPress事件中调用
    Call gobjComLib.zlControl.CboAppendText(cboControl, KeyAscii)
End Sub


Public Sub AddData(objSource As Object, ByVal rsDataSource As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'功能:装载数据入指定的组合下拉框或网格中的下拉框中
'参数:objSource   要装载记录集的下拉框控件
'     rsDataSource     记录集数据,只能有两个数据项,第一个是下拉数据,第二个是对应的关键字,如id
'     blnClear    装载时是否清楚原有的下拉数据,缺省为True
'返回:
    Call gobjComLib.zlControl.CboAddData(objSource, rsDataSource, blnClear)
End Sub


