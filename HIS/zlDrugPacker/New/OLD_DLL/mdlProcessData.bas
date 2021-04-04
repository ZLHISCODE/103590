Attribute VB_Name = "mdlProcessData"
Option Explicit

Public Sub ProcDrugInfo(ByVal strDrugType As String, ByVal strLinkName As String)
'功能：获取HIS药品基本信息
'参数：
'  strDrugType：剂型串
'  strLinkName：连接名称
    
    '实例clsConnect
    
    '读HIS数据
    
    '按连接类型不同，分别格式化要上传的数据
    
    '调用mdlDrugPacker.DrugInfo
    
    
    Exit Sub

errHandle:
    If gobjComLib.ErrCenter = 1 Then Resume
    gstrMessage = Err.Description
End Sub



Public Sub SetUpload(ByVal bytType As Byte, ByVal varKey As Variant)
'功能：获取HIS相关上传信息
'参数：
'   bytType：
'       1: 门诊处方上传 (配药)
'       2: 门诊发药通知 (发药)
'       3: 住院药品医嘱上传 (配、发药)
'   varKey：
'       当bytType=1时，varKey表示“单据;库房ID;NO”；
'       格式：“单据;库房ID;NO[|单据;库房ID;NO][|...]”
'       当bytType=2时，同bytType=1
'       当bytType=3时，varKey表示药品收发ID；
'       格式：“药品收发ID[|药品收发ID][|...]”

    '读HIS数据
    
    '记录行确定要上传设备对象
    
    '格式化要上传的数据
    
    '调用mdlDrugPacker.Dispense、Dispensing


End Sub

