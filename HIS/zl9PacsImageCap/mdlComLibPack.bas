Attribute VB_Name = "mdlComLibPack"
Option Explicit

Public gobjComLib As Object 'zl9ComLib.clsComLib


Public Sub zlCL_CloseWindow()
'�ر�ʹ�ô���
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.CloseWindows
End Sub

Public Function zlCL_GetCacheDir() As String
'��ȡ����Ŀ¼
    zlCL_GetCacheDir = IIf(Len(App.Path) > 3, App.Path & "\TmpImage\", App.Path & "TmpImage\")
End Function


Public Function zlCL_GetResourceDir() As String
'��ȡ��ԴĿ¼
    zlCL_GetResourceDir = IIf(Len(App.Path) > 3, App.Path & "\..\�����ļ�\", App.Path & "..\�����ļ�\")
End Function

Public Function zlCL_GetPara(ByVal varPara As Variant, _
                        Optional ByVal lngSys As Long, _
                        Optional ByVal lngModual As Long, _
                        Optional ByVal strDefault As String, _
                        Optional ByVal arrControl As Variant, _
                        Optional ByVal blnSetup As Boolean, _
                        Optional intType As Integer) As String
'��ȡ��������
    zlCL_GetPara = ""
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_GetPara = gobjComLib.zlDataBase.GetPara(varPara, lngSys, lngModual, strDefault, arrControl, blnSetup, intType)
End Function


Public Sub zlCL_ShowFlash(Optional strNote As String, Optional frmParent As Object)
'Flash������ʾ
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlCommFun.ShowFlash(strNote, frmParent)
End Sub

Public Sub zlCL_StopFlash()
'��ֹFlash������ʾ
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlCommFun.StopFlash
End Sub


Public Sub zlCL_ExecuteProcedure(strSQL As String, ByVal strFormCaption As String)
'ִ�й���
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlDataBase.ExecuteProcedure(strSQL, strFormCaption)
End Sub

Public Function zlCL_GetDBObj() As Object
'��ȡ���ݿ�ִ�ж���
'���ڲ��ܶ�zldatabase��executedreProcedure���з�װ�����ֻ���ɶ��󵥶�ִ��
    Set zlCL_GetDBObj = Nothing
    If gobjComLib Is Nothing Then Exit Function
    
    Set zlCL_GetDBObj = gobjComLib.zlDataBase
End Function

Public Function zlCL_GetNodeNo() As String
'��ȡվ����
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_GetNodeNo = gobjComLib.gstrNodeNo
End Function

Public Function zlCL_GetPubIcons() As Object
'��ȡ����ͼ�����
    Set zlCL_GetPubIcons = Nothing
    If gobjComLib Is Nothing Then Exit Function
    
    Set zlCL_GetPubIcons = frmPubIcons.imgPublic.Icons
End Function

Public Function zlCL_Currentdate() As Date
'�����ݿ�������л�ȡ��ǰʱ��
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_Currentdate = gobjComLib.zlDataBase.Currentdate
End Function

Public Function zlCL_GetPrivFunc(lngSys As Long, lngProgId As Long) As String
'��ȡ����Ȩ��
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_GetPrivFunc = gobjComLib.GetPrivFunc(lngSys, lngProgId)
End Function

Public Sub zlCL_CboSetIndex(ByVal hWnd_combo As Long, ByVal lngIndex As Long)
'����combobox��listindex
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlControl.CboSetIndex(hWnd_combo, lngIndex)
End Sub


Public Sub zlCL_PressKey(bytKey As Byte)
'ģ�ⰴ������
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.zlCommFun.PressKey(bytKey)
End Sub


Public Function zlCL_SetPara(ByVal varPara As Variant, _
                            ByVal strValue As String, _
                            Optional ByVal lngSys As Long, _
                            Optional ByVal lngModual As Long, _
                            Optional ByVal blnSetup As Boolean = True) As Boolean
'�����������
    zlCL_SetPara = False
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_SetPara = gobjComLib.zlDataBase.SetPara(varPara, strValue, lngSys, lngModual, blnSetup)
End Function


Public Function zlCL_RestoreWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'�ָ�����״̬
    zlCL_RestoreWinState = False
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_RestoreWinState = gobjComLib.RestoreWinState(objForm, strProjectName, strUserDef)
End Function


Public Function zlCL_SaveWinState(objForm As Object, Optional ByVal strProjectName As String, Optional ByVal strUserDef As String) As Boolean
'���洰��״̬
    zlCL_SaveWinState = False
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_SaveWinState = gobjComLib.SaveWinState(objForm, strProjectName, strUserDef)
End Function


Public Function zlCl_RegCheck(Optional blnTemp As Boolean) As String
'ע����
    zlCl_RegCheck = ""
    If gobjComLib Is Nothing Then Exit Function
    
    zlCl_RegCheck = gobjComLib.zlRegCheck(blnTemp)
End Function

Public Function zlCL_GetRegInfo(strItem As String, Optional blnTemp As Boolean, Optional intBits As Integer) As String
'��ȡע����Ϣ
    zlCL_GetRegInfo = ""
    If gobjComLib Is Nothing Then Exit Function
    
    zlCL_GetRegInfo = gobjComLib.zlRegInfo(strItem, blnTemp, intBits)
End Function


Public Sub zlCL_InitCommon(cnMain As ADODB.Connection)
'��ʼ�����Ӷ���
    If gobjComLib Is Nothing Then Exit Sub
    
    Call gobjComLib.InitCommon(cnMain)
    
    Call gobjComLib.SetDbUser(cnMain.Properties.Item("User ID").value)
End Sub





