VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStuffService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private Const mstrUsername = "admin"
Private Const mstrPassword = "admin"
Private Const mstrClient_id = "70e99255-2e6c-4354-be3b-5ca25cc50814"
Private Const mstrPID = "f327422f-3efa-47a0-9942-dd02d946832c"
Private Const mstrBusinessURL = "https://192.168.2.62:8800"
Private mobjHttpRequest  As WinHttp.WinHttpRequest

Private mobjJson As clsJson
Private mdtSecurity As Date
Private mobjBusinessAsks As clsBusinessAsks '业务清单集
Private mstraccess_token As String
Private Const M_CONNET_TYPE = 1     '0-webApi;1-Oracle的伙
Public Enum gDataType
     BusinessDataType_Num = 0
     BusinessDataType_String = 1
     BusinessDataType_Date = 2
End Enum
 
Private Function GetReturnData(ByVal strBusinessAskName As String) As clsBusinessParsing
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定业务请求的返回数据集
    '入参:strBusinessAskName-请求的业务名称
    '出参:
    '返回:成功,返回clsBusinessDatas类型,否则返回Nothing
    '编制:刘兴洪
    '日期:2017-05-24 16:02:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objDatas As clsBusinessDatas
    Dim objdata As clsBusinessData
    Dim objBusinessParsing As clsBusinessParsing
    
    
    Set objBusinessParsing = New clsBusinessParsing
    
    
    Set objDatas = New clsBusinessDatas
    Select Case UCase(strBusinessAskName)
    Case UCase("zl_StuffSvr_GetStockCheck")   '获取卫材库存检查方式"
        objDatas.列表接点 = "item_list"
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "库房ID"
            .节点名称 = "warehouse_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "返回库房ID"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "检查方式"
            .节点名称 = "check_type"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "检查方式：0-不检查，1-检查提示名，2-检查禁止"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objBusinessParsing.objReturnRecord = objDatas
        Set GetReturnData = objBusinessParsing
    
    Case UCase("zl_StuffSvr_GetStock")   '获取指定卫材在指定库房的可用库存数"
        objDatas.列表接点 = ""
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "可用库存"
            .节点名称 = "stock"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "可用库存"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objBusinessParsing.objReturnData = objDatas
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_GetStockBatch")   '批量获取多个卫材库存及价格信息"
        objDatas.列表接点 = "item_list"
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "卫材ID"
            .节点名称 = "stuff_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "收费项目目录.ID"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "可用数量"
            .节点名称 = "stock"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "药品库存.可用数量"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "零售价"
            .节点名称 = "price"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "零售价(返回价格时才有此项)"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnRecord = objDatas
        Set GetReturnData = objBusinessParsing
        
    Case UCase("zl_StuffSvr_GetWareHouseWindow")   '根据库房获取所涉及的发料窗口"
   
        objDatas.列表接点 = "window_list"
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "库房ID"
            .节点名称 = "warehouse_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "部门表.ID"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                
                
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "发药窗口"
            .节点名称 = "warehouse_windows"
            .数据类型 = BusinessDataType_String
            .长度 = 4000
            .数据说明 = "库房ID对应的发药窗口"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnRecord = objDatas
        Set GetReturnData = objBusinessParsing
    
    Case UCase("zl_StuffSvr_GetSendWindows")   '根据库房窗口的分配规则获取指定库房的发药窗口"
       objDatas.列表接点 = "item_list"
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "库房ID"
            .节点名称 = "warehouse_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "部门表.ID"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "发药窗口"
            .节点名称 = "warehouse_window"
            .数据类型 = BusinessDataType_String
            .长度 = 50
            .数据说明 = "库房ID对应的发药窗口"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objBusinessParsing.objReturnRecord = objDatas
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_BatchGetPrice")   '批量获取卫生卫材售价(实价)"
  
        objDatas.列表接点 = "item_list"
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "卫材ID"
            .节点名称 = "stuff_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "收费项目目录.ID"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "零售价"
            .节点名称 = "price"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "卫材ID对应的价格"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnRecord = objDatas
        Set GetReturnData = objBusinessParsing
    
    Case UCase("zl_StuffSvr_GetPrice")   '获取指定卫材的售价?成本价"
    
        objDatas.列表接点 = ""
 
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "零售价"
            .节点名称 = "price"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "卫材ID对应的价格"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "成本价"
            .节点名称 = "price_cost"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "卫材ID对应的成本价"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "剩余数量"
            .节点名称 = "quantity_remain"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "剩余数量，等于0表示数量足够，大于0则表示数量不够 "
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objBusinessParsing.objReturnData = objDatas
        Set GetReturnData = objBusinessParsing
    
    Case UCase("zl_StuffSvr_GetStuffIDbyBarcode")   '根据输入的条码获取卫生卫材信息(卫材，批次)"
   
        objDatas.列表接点 = "item_list"
 
    
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "卫材ID"
            .节点名称 = "stuff_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "收费项目目录.ID"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "批次"
            .节点名称 = "batch"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "药品库存.批次"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
          
        
        Set objBusinessParsing.objReturnRecord = objDatas
        Set GetReturnData = objBusinessParsing
    
    
    Case UCase("zl_StuffSvr_CheckStoreLimit")   '检查指定卫材在指定库房的库存是否低于储备下限"
        objDatas.列表接点 = ""
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "储备状态"
            .节点名称 = "below_limit_lower"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "储备状态:1-低于储备下限，0-大于等于储备下限"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnData = objDatas
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_GetCargoSpace")   '获取指定卫材的指定库房的货位信息"
        objDatas.列表接点 = ""
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "库房货位"
            .节点名称 = "cargospace"
            .数据类型 = BusinessDataType_String
            .长度 = 50
            .数据说明 = "库房货位"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnData = objDatas
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_CheckExpiryDate")   '检查卫生卫材的灭菌效期是否过期"
    
        objDatas.列表接点 = ""
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "是否存在过期"
            .节点名称 = "exist_expiried"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "是否存在过期:0-不存在已过期项目，1-存在已过期项目"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        objDatas.列表接点 = ""
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "最小灭菌效期"
            .节点名称 = "min_expirydate"
            .数据类型 = BusinessDataType_String
            .长度 = 20
            .数据说明 = "最小灭菌效期，exist_expiried=1时返回，格式：yyyy-mm-dd"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnData = objDatas
        Set GetReturnData = objBusinessParsing
    
    
    
    Case UCase("zl_StuffSvr_UpdateChargeTag")   '更新收费状态(已收，未收，异常）"
        
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_AdjustData")   '门诊费用转住院时调整卫材关联数据"
        
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_GetBatchNumber")   '获取卫材批号信息"
        objDatas.列表接点 = ""
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "批号"
            .节点名称 = "batch_number"
            .数据类型 = BusinessDataType_String
            .长度 = 100
            .数据说明 = "批号"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnData = objDatas
        Set GetReturnData = objBusinessParsing
    
    Case UCase("zl_StuffSvr_CheckContainStuff")   '检查单据是否含备货卫材"
        objDatas.列表接点 = ""
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "是否含备货卫材"
            .节点名称 = "contain_stuff"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "是否含备货卫材:0-含有备货材料,1-不含备货材料"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnData = objDatas
        Set GetReturnData = objBusinessParsing
    
    Case UCase("zl_StuffSvr_GetStuffPutInBills")   '获取备货卫材的入库单"
    
        objDatas.列表接点 = "item_list"
        
 
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "卫材入库单id"
            .节点名称 = "bill_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "卫材入库单id"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "卫材入库单号"
            .节点名称 = "bill_no"
            .数据类型 = BusinessDataType_String
            .长度 = 50
            .数据说明 = "卫材入库单号"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "序号"
            .节点名称 = "serial_num"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "单据序号"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
               
               
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "产地"
            .节点名称 = "origin"
            .数据类型 = BusinessDataType_String
            .长度 = 200
            .数据说明 = "产地"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "供应商"
            .节点名称 = "provider"
            .数据类型 = BusinessDataType_String
            .长度 = 200
            .数据说明 = "供应商"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "批号"
            .节点名称 = "batch_number"
            .数据类型 = BusinessDataType_String
            .长度 = 100
            .数据说明 = "批号"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                                
                                
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "生产日期"
            .节点名称 = "production_date"
            .数据类型 = BusinessDataType_String
            .长度 = 20
            .数据说明 = "生产日期"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "效期"
            .节点名称 = "expiry_date"
            .数据类型 = BusinessDataType_String
            .长度 = 20
            .数据说明 = "效期"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                       
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "灭菌效期"
            .节点名称 = "sterilization_expiry_date"
            .数据类型 = BusinessDataType_String
            .长度 = 20
            .数据说明 = "灭菌效期"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
 
        
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "入库数量"
            .节点名称 = "putin_quantity"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "入库数量"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "入库零售价"
            .节点名称 = "putin_price"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "入库零售价"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "入库零售金额"
            .节点名称 = "putin_money"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "入库零售金额"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "审核日期"
            .节点名称 = "audit_time"
            .数据类型 = BusinessDataType_String
            .长度 = 20
            .数据说明 = "审核日期"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "卫材ID"
            .节点名称 = "stuff_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "卫材ID"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "批次"
            .节点名称 = "batch"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "批次"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "商品条码"
            .节点名称 = "barcode_goods"
            .数据类型 = BusinessDataType_String
            .长度 = 100
            .数据说明 = "商品条码"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                        
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "内部条码"
            .节点名称 = "barcode_inside"
            .数据类型 = BusinessDataType_String
            .长度 = 100
            .数据说明 = "内部条码"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                                                
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "可用库存"
            .节点名称 = "stock"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "可用库存"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "费用ID"
            .节点名称 = "stuffdtl_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "处方明细id"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnRecord = objDatas
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_RecipeAffirm")    '卫材处方记帐确认或卫材处方收费确认
    
         Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_GetAllowReturnNum")   '获取卫材或卫生卫材的原始数量?发药数量?未发数量及准退费"
    
    objDatas.列表接点 = "item_list"
        
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "NO"
            .节点名称 = "stuff_no"
            .数据类型 = BusinessDataType_String
            .长度 = 50
            .数据说明 = "门诊费用记录.NO"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "序号"
            .节点名称 = "serial_num"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "门诊费用记录.序号"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
               
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "费用ID"
            .节点名称 = "stuffdtl_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "费用ID "
            .是否可选 = True
        End With
        objDatas.AddItem objdata
                              
                              
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "卫材ID"
            .节点名称 = "stuff_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "收费项目目录.ID"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "原数量"
            .节点名称 = "original_num"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "原数量"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "已发数量"
            .节点名称 = "sended_num"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "已发数量"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "未发数量"
            .节点名称 = "notsend_num"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "未发数量"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "准退数量"
            .节点名称 = "allowreturn_num"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "准退数量"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "商品条码"
            .节点名称 = "barcode_goods"
            .数据类型 = BusinessDataType_String
            .数据说明 = "商品条码"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "内部条码"
            .节点名称 = "barcode_inside"
            .数据类型 = BusinessDataType_String
            .数据说明 = "内部条码"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objBusinessParsing.objReturnRecord = objDatas
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_AutoSendStuff")   '")   '")   '收费或记帐后，自动发料
        
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_AutoBatchSendStuff")   '卫生自动批量发料"
        
    Case UCase("zl_StuffSvr_AutoReturnStuff")   '针对卫生卫材进行自动退料操作"
        
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_NewRecipeBill")   '")   '")   '主要是在记帐（含划价），")   '收费(含划价)后产生新的处方或药嘱记录
        
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_DelRecipeBill")   '卫材处方销帐(或退费)"
        
        Set GetReturnData = objBusinessParsing
    Case UCase("zl_StuffSvr_GetRefuseSendList")   '获取拒发料清单"
        Dim objTemp As New clsBusinessDatas
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "是否存在"
            .节点名称 = "isExsit"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "是否存在:1-存在拒发;0-不存在拒发"
            .是否可选 = True
        End With
        objTemp.AddItem objdata
        Set objBusinessParsing.objReturnRecord = objTemp
        
        objDatas.列表接点 = "item_list"
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "NO"
            .节点名称 = "stuff_no"
            .数据类型 = BusinessDataType_String
            .长度 = 50
            .数据说明 = "门诊费用记录.NO"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "费用ID"
            .节点名称 = "stuffdtl_id"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "门诊费用记录.ID"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnRecord = objDatas
        Set GetReturnData = objBusinessParsing
    
    Case UCase("Zl_StuffSvr_Pati_Check_Execute")   '根据费用ID或单据获取未发卫材数据"
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "是否存在"
            .节点名称 = "isexist"
            .数据类型 = BusinessDataType_Num
            .数据说明 = "是否存在: 1-存在;0-不存在"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objdata = New clsBusinessData
        With objdata
            .数据名称 = "未发料信息"
            .节点名称 = "stuff_notsend_infor"
            .数据类型 = BusinessDataType_Num
            .长度 = 4000
            .数据说明 = "未发料信息,isexist=1时返回"
            .是否可选 = True
        End With
        objDatas.AddItem objdata
        Set objBusinessParsing.objReturnData = objDatas
        Set GetReturnData = objBusinessParsing
        
    Case UCase("zl_StuffSvr_UpdatePatiPageid")   '更新病人信息"
         Set GetReturnData = objBusinessParsing
    Case Else
        Set GetReturnData = objBusinessParsing
    End Select
End Function


Private Sub InitBusinessData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化业务数据
    '编制:刘兴洪
    '日期:2017-05-23 15:54:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objBusinessAsk As clsBusinessAsk
    
    Set mobjBusinessAsks = New clsBusinessAsks
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetStockCheck"
        .功能 = "1.获取卫材库存检查方式"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetStock"
        .功能 = "2.获取指定卫材在指定库房的可用库存数"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetStockBatch"
        .功能 = "3.批量获取多个卫材库存及价格信息"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetWareHouseWindow"
        .功能 = "4.根据库房获取所涉及的发料窗口"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetSendWindows"
        .功能 = "5.根据库房窗口的分配规则获取指定库房的发药窗口"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_BatchGetPrice"
        .功能 = "6.批量获取卫生卫材售价(实价)"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetPrice"
        .功能 = "7.获取指定卫材的售价、成本价"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetStuffIDbyBarcode"
        .功能 = "8.根据输入的条码获取卫生卫材信息(卫材，批次)"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_CheckStoreLimit"
        .功能 = "9.检查指定卫材在指定库房的库存是否低于储备下限"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetCargoSpace"
        .功能 = "10.获取指定卫材的指定库房的货位信息"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_CheckExpiryDate"
        .功能 = "11.检查卫生卫材的灭菌效期是否过期"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_UpdateChargeTag"
        .功能 = "12.更新收费状态(已收，未收，异常）"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_AdjustData"
        .功能 = "13.门诊费用转住院时调整卫材关联数据"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetBatchNumber"
        .功能 = "14.获取卫材批号信息"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_CheckContainStuff"
        .功能 = "15.检查单据是否含备货卫材"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetStuffPutInBills"
        .功能 = "16.获取备货卫材的入库单"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_RecipeAffirm"
        .功能 = "17.卫材处方记帐确认或卫材处方收费确认"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetAllowReturnNum"
        .功能 = "18.获取卫材或卫生卫材的原始数量、发药数量、未发数量及准退费"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_AutoSendStuff"
        .功能 = "19.收费或记帐后，自动发料"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_AutoBatchSendStuff"
        .功能 = "20.卫生自动批量发料"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_AutoReturnStuff"
        .功能 = "21.针对卫生卫材进行自动退料操作"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_NewRecipeBill"
        .功能 = "22.主要是在记帐（含划价）， 收费(含划价)后产生新的处方或药嘱记录"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_DelRecipeBill"
        .功能 = "23.卫材处方销帐(或退费)"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_GetRefuseSendList"
        .功能 = "24.获取拒发料清单"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "Zl_StuffSvr_Pati_Check_Execute"
        .功能 = "25.根据费用ID或单据获取未发卫材数据"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)
    
    Set objBusinessAsk = New clsBusinessAsk
    With objBusinessAsk
        .发布ID = ""
        .过程名 = "zl_StuffSvr_UpdatePatiPageid"
        .功能 = "26更新病人信息"
    End With
    Call mobjBusinessAsks.AddItem(objBusinessAsk, objBusinessAsk.过程名)


End Sub

Private Function ValiedSecurityCertificate(Optional ByVal strCurDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:验证证书是否在有效时间范围内
    '入参:strCurDate-当前数据库时间,格式"yyyy-mm-dd hh24:mi:ss"
    '返回:合法成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-05-23 15:34:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDate As String
    
    On Error GoTo errHandle
    strDate = strCurDate
    If strDate = "" Then strDate = Date
    If mdtSecurity >= CDate(strDate) Then ValiedSecurityCertificate = True: Exit Function
    '取过期时间，下一次调用的时候如果在有效期内，则不需要去调用token服务
    strDate = Mid$(mobjJson.GetValue("json['.expires']", False), 5, 21)
    If strDate = "" Then Exit Function
    mdtSecurity = CDate(strDate)
    ValiedSecurityCertificate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function OpenConnect(Optional cnoracle As ADODB.Connection, Optional strDBUser As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:连接到新药品系统
    '入参:cnoracle-oracle连接对象
    '     strDBUser-DB用户
    '出参:
    '返回:连接成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 09:27:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInput As String, strToken As String
    Dim strSecurityCertificate     As String  '安全证书串
   
    On Error GoTo errHandle
    
    '1.目前只支持调用oracle过程
    Set gcnOracle = cnoracle: gstrDBUser = strDBUser
    OpenConnect = Not gcnOracle Is Nothing
    Exit Function
    
    
     '调用webservice方法
    strInput = "grant_type=password&username=" & mstrUsername & "&password=" & mstrPassword & "&client_id=" & mstrClient_id
    
    '跳过安全证书提示
    mobjHttpRequest.Option(4) = 13056
    mobjHttpRequest.Open "POST", mstrBusinessURL & "/token", False '默认是POST方式
    mobjHttpRequest.SetRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    mobjHttpRequest.SetRequestHeader "Content-Length", Len(strInput)
    mobjHttpRequest.Send strInput
    strSecurityCertificate = mobjHttpRequest.ResponseText
    
    If mobjJson.OpenJson(strSecurityCertificate, False) = False Then Exit Function
    mstraccess_token = mobjJson.GetValue("access_token")
    If ValiedSecurityCertificate = False Then Exit Function
    
    OpenConnect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetBusinessObj(ByVal strAskName As String) As clsBusinessAsk
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据名字获取业务对象
    '入参:strAskName-业务请求名称（现以过程名为准)
    '出参:
    '返回:成功,返回clsBusinessAsk对象,否则返回Nothing
    '编制:刘兴洪
    '日期:2017-05-23 16:55:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set GetBusinessObj = mobjBusinessAsks(strAskName)
    If Err <> 0 Then Err = 0: Set GetBusinessObj = Nothing: Exit Function
    Exit Function
End Function

Public Function BusinessAsk(ByVal lngModule As Long, ByVal strAskName As String, _
    ByVal strJsonAsk As String, ByVal strAskDate As String, _
    Optional strJson_out As String, Optional blnParsingJsonOut As Boolean, _
    Optional ByRef rsOutData_out As ADODB.Recordset, Optional varData_Out As Variant, _
    Optional ByVal blnShowNotErrMsg As Boolean, Optional strErrMsg_Out As String, _
    Optional ByVal blnAutoApendRecord As Boolean = False, _
    Optional objBusinessParas As clsBusinessDatas) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:业务交易
    '入参:strAskName-业务请求名称（现以过程名为准)
    '     strJsonAsk-JSON格式的入参(不需要传入Input接点，内部将自动增加)
    '     strAskDate-业务请求日期
    '     blnParsingJsonOut-是否解析输出的Json串，true-解析,False-不解析
    '     blnAutoApendRecord-自动追加数据到记录集中(多批次调用业务请求有效)
    '     blnShowNotErrMsg-不显示错误信息
    '     objBusinessParas-请求参数对象集。
    '     lngModule-模块号
    '出参:strJson_out-如果成功，返回Json格式串
    '     blnParsingJsonOut=true时，返回以下参数:
    '           rsOutData_Out-返回请求数据的记录集,blnParsingJsonOut=false时为Nothing
    '           varData_out-返回请求数据(数据组）,blnParsingJsonOut=false时为Empty
    '     strErrMsg_Out-如果失败，返回错误原因
    '返回:业务请求成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-05-23 16:20:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objHttp As WinHttp.WinHttpRequest, objBusinessAsk As clsBusinessAsk
    Dim objPara As clsBusinessData, rsTemp As ADODB.Recordset
    Dim strAccess_token As String
    Dim strInput As String, strUrl As String, strJson As String
    Dim strLogInfor As String '调试信息
    Dim lngDateDiff As Long
    Dim dtDate As Date
    Dim strFunName As String, strErr As String, lngErrNum As Long
     
    
    strFunName = "BusinessAsk"
    Call zlWritLogEx(Me, lngModule, strFunName, "开始", "入参:" & strAskName & ",Json:=" & strJsonAsk)
    
    
    If blnShowNotErrMsg Then On Error GoTo errHandle    '出错时,上级捕获错误
    
    
    strJson_out = "": strErrMsg_Out = "": strJson = strJsonAsk
    Set objBusinessAsk = GetBusinessObj(strAskName)
    If objBusinessAsk Is Nothing Then
        strErrMsg_Out = "当前请求的“" & strAskName & "”业务暂不支持，请检查!"
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回:False,错误信息:=" & strErrMsg_Out)
        GoTo ShowErrMsg:
        Exit Function
    End If
    
    If strJson <> "" Then
        'strJson = "{input:" & strJson & "}"
    ElseIf Not objBusinessParas Is Nothing Then
        If GetJsonString(objBusinessParas, strJson, strErrMsg_Out) = False Then GoTo ShowErrMsg:
        'strJson = "{input:" & strJson & "}"
    End If

    Select Case M_CONNET_TYPE
    Case 0  'webAPI:
        If OpenWebApi(objBusinessAsk, strAskDate, strJson, strJson_out, strErrMsg_Out) = False Then
             Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回:False,错误信息:=" & strErrMsg_Out)
             GoTo ShowErrMsg:
             Exit Function
        End If
    Case 1 'oracle存储过程
        If ExcuteProcedureFromJson(objBusinessAsk, strJson, strJson_out, strErrMsg_Out) = False Then
             Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回:False,错误信息:=" & strErrMsg_Out)
             GoTo ShowErrMsg:
             Exit Function
        End If
    Case Else
        Exit Function
    End Select
    
    If blnParsingJsonOut Then   '需要解析出参Json
       If objBusinessAsk.返回信息 Is Nothing Then
            Set objBusinessAsk.返回信息 = GetReturnData(objBusinessAsk.过程名)
       End If
        If BuildDataRecordFromJson(objBusinessAsk, strJson_out, rsOutData_out, varData_Out, blnShowNotErrMsg, strErrMsg_Out, blnAutoApendRecord) = False Then Exit Function
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回:True,strJson_out=" & strJson_out)
    
    BusinessAsk = True
    Exit Function
errHandle:
    lngErrNum = Err.Number: strErr = Err.Description
    If blnShowNotErrMsg Then
        Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回:False,err=" & strErr)
        Err.Raise lngErrNum, Err.Description: Exit Function
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    Call zlWritLogEx(Me, lngModule, strFunName, "结束", "返回:False,err=" & strErr)
    
    Exit Function
ShowErrMsg:
  If Not blnShowNotErrMsg Then MsgBox strErrMsg_Out, vbInformation, gstrSystemName
End Function

Private Function ExcuteProcedureFromJson(ByVal objBusinessAsk As clsBusinessAsk, ByVal strJson As String, ByRef strJson_out As String, Optional strErrMsg_Out As String) As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行存储过程
    '入参:objBusinessAsk-请求的业务对象
    '     strJson-Json入参
    '     strAskDate-请求日期
    '出参:strJson_out-获取返回的Json出参
    '     strErrMsg_out-发生错误时返回
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 09:49:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gcnOracle Is Nothing Then strErrMsg_Out = "未初始化数据库连接": ExcuteProcedureFromJson = False
    
      
    strJson_out = zlDatabase.CallProcedure(objBusinessAsk.过程名, "药品服务-" & objBusinessAsk.功能, strJson, Empty)
    ExcuteProcedureFromJson = True
 
End Function

Private Function OpenWebApi(ByVal objBusinessAsk As clsBusinessAsk, ByVal strAskDate As String, ByVal strJson As String, ByRef strJson_out As String, Optional strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开WebApi
    '入参:objBusinessAsk-请求的业务对象
    '     strJson-Json入参
    '     strAskDate-请求日期
    '出参:strJson_out-获取返回的Json出参
    '     strErrMsg_out-发生错误时返回
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 09:49:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strUrl As String, strAccess_token As String
   '业务提交
    If ValiedSecurityCertificate(strAskDate) = False Then Exit Function

   '调用业务服务
    strUrl = mstrBusinessURL & "/bizdomain/" & objBusinessAsk.发布ID

    strAccess_token = "Bearer " & mstraccess_token
    '跳过安全证书提示
    mobjHttpRequest.Option(4) = 13056
    mobjHttpRequest.Open "POST", strUrl, False      '默认是POST方式
    mobjHttpRequest.SetRequestHeader "Content-Type", "text/xml; charset=UTF-8"
    mobjHttpRequest.SetRequestHeader "AUthorization", strAccess_token

    mobjHttpRequest.Send strJson
    strJson_out = mobjHttpRequest.ResponseText
    If GetErrData(strJson_out, strErrMsg_Out) = True Then Exit Function
    OpenWebApi = True

End Function





Public Function GetJsonString(ByVal objDatas As clsBusinessDatas, ByRef strJson_out As String, ByRef strErrMsg_Out As String, _
    Optional blnReplaceNameIN As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据业务数据对象，返回Json串
    '入参:objDatas-业务数据对象
    '    blnReplaceNameIN-替换名称带_IN的
    '出参:strJson_out-Json格式串
    '     strErrMsg_Out-获取失败时，返回对应的错误信息
    '返回:获取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-06-23 16:05:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strJson As String, objdata As clsBusinessData
    Dim strName As String
    On Error GoTo errHandle
    strJson = ""
    For Each objdata In objDatas
        strName = objdata.数据名称
        If blnReplaceNameIN Then strName = Replace(UCase(strName), "_IN", "")
        
        Select Case objdata.数据类型
        Case BusinessDataType_Date
            If IsDate(objdata.数据内容) = False Or objdata.数据内容 = "" Then
                If Not objdata.是否可选 Then
                     strErrMsg_Out = strName & "必须传入一个是日期类型，格式为:yyyy-mm-dd hh:mm:ss"
                     Exit Function
                End If
                strJson = strJson & "," & """" & strName & """" & ":null"
            Else
                strJson = strJson & "," & """" & strName & """" & ":" & """" & Format(CDate(objdata.数据内容), "yyyy-mm-dd HH:MM:SS") & """"
            End If
        Case BusinessDataType_Num
               If UCase(objdata.数据名称) Like "*ID*" And Val(objdata.数据内容) = 0 Then
                    strJson = strJson & "," & """" & strName & """" & ":null"
               Else
                    strJson = strJson & "," & """" & strName & """" & ":" & IIf(objdata.是否空值, "null", Val(objdata.数据内容))
               End If
        Case Else
                strJson = strJson & "," & """" & strName & """" & ":" & IIf(Trim(objdata.数据内容) = "" Or objdata.是否空值, "null", """" & Trim(objdata.数据内容) & """" & "")
         End Select
    Next
    strJson = Mid(strJson, 2)
    strJson = "{" & """" & "input" & """" & ":" & "{" & Mid(strJson, 2) & "}}"
    strJson_out = strJson
    GetJsonString = True
    Exit Function
errHandle:
   strErrMsg_Out = Err.Description
End Function



Private Function GetErrData(ByVal strJson As String, ByRef strErrMsg_Out As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取错误信息
    '入参:
    '出参:strErrMsg_Out-返回错误信息
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2017-06-12 14:03:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant
    If (Not strJson Like "*ORA-*" And Not strJson Like "*执行领域方法出错*") Or strJson = "" Then Exit Function
    
    varData = Split(UCase(strJson), "[ZLSOFT]")
    If UBound(varData) >= 2 Then
        strErrMsg_Out = varData(1): GetErrData = True: Exit Function
    End If
    strErrMsg_Out = strJson: GetErrData = True: Exit Function
End Function
''
''Public Function BusinessAskForGet(ByVal strAskName As String, ByVal objBusinessAsks As clsBusinessDatas, ByVal strAskDate As String, _
''    Optional blnBuildRecordset As Boolean, Optional strJson_out As String, Optional ByRef rsOutData As ADODB.Recordset, _
''    Optional ByVal blnShowNotErrMsg As Boolean, Optional strErrMsg_Out As String, Optional ByVal blnAutoApendRecord As Boolean = False) As Boolean
''    '---------------------------------------------------------------------------------------------------------------------------------------------
''    '功能:业务交易(使用的是Get方式，但使用URL方式传值，会造成长度限制
''    '入参:strAskName-业务请求名称（现以过程名为准)
''    '     objBusinessAsks-传入过程的参数值列表
''    '     strAskDate-业务请求日期
''    '     blnBuildRecordset-是否生成记录集，生成返回true,否则返回False
''    '     blnShowNotErrMsg-不显示错误信息
''    '     blnAutoApendRecord-自动追回数据到记录集中(多批次调用业务请求有效)
''    '出参:strJson_out-如果成功，返回Json格式串
''    '     rsOutData-如果业务成功且blnBuildRecordset=true且有返回记录集时，返回请求数据的记录集
''    '               如果业务失败或blnBuildRecordset=False时，返回Nothing
''    '     strErrMsg_Out-如果失败，返回错误原因
''    '返回:业务请求成功返回true,否则返回False
''    '编制:刘兴洪
''    '日期:2017-05-23 16:20:43
''    '---------------------------------------------------------------------------------------------------------------------------------------------
''    Dim objHttp As WinHttp.WinHttpRequest, objBusinessAsk As clsBusinessAsk
''    Dim objPara As clsBusinessData, rsTemp As ADODB.Recordset
''    Dim strPara As String
''    Dim strInput As String, strUrl As String, strJson As String
''
''    If Not blnShowNotErrMsg Then
''        On Error GoTo errHandle
''    End If
''    strErrMsg_Out = ""
''    If objBusinessAsks Is Nothing Then Exit Function
''    Set objBusinessAsk = GetBusinessObj(strAskName)
''    If objBusinessAsk Is Nothing Then
''        strErrMsg_Out = "当前请求的“" & strAskName & "”业务暂不支持，请检查!": GoTo ShowErrMsg:
''        Exit Function
''    End If
''    For Each objPara In objBusinessAsks
''        '卫材ID_IN=8198&库房ID_IN=72&批次_IN=0&数量_IN=1"
''        Select Case objPara.数据类型
''        Case BusinessDataType_Date
''            If IsDate(objPara.数据内容) = False Or objPara.数据内容 = "" Then
''                If Not objPara.是否可选 Then
''                     strErrMsg_Out = objPara.数据名称 & "必须传入一个是日期类型，格式为:yyyy-mm-dd hh:mm:ss": GoTo ShowErrMsg:
''                     Exit Function
''                End If
''                strPara = strPara & "&" & objPara.数据名称 & "="
''            Else
''                strPara = strPara & "&" & objPara.数据名称 & "=" & Format(CDate(objPara.数据内容), "yyyy-mm-dd hh24:mi:ss")
''            End If
''        Case BusinessDataType_Num
''                strPara = strPara & "&" & objPara.数据名称 & "=" & val(objPara.数据内容)
''        Case Else
''                strPara = strPara & "&" & objPara.数据名称 & "=" & Trim(objPara.数据内容)
''        End Select
''    Next
''    strPara = Mid(strPara, 2)
''
''    '业务提交
''    If ValiedSecurityCertificate(strAskDate) = False Then Exit Function
''
''   '调用业务服务
''    strUrl = mstrBusinessURL & "/bizdomain/" & objBusinessAsk.发布ID & "/" & strPara
''
''    strInput = "Bearer " & mstraccess_token
''    '跳过安全证书提示
''    mobjHttpRequest.Option(4) = 13056
''    mobjHttpRequest.Open "GET", strUrl, False    '默认是POST方式
''    mobjHttpRequest.SetRequestHeader "Content-Type", "text/xml; charset=UTF-8"
''    mobjHttpRequest.SetRequestHeader "AUthorization", strInput
''    mobjHttpRequest.Send
''    strJson_out = mobjHttpRequest.ResponseText
''    If blnBuildRecordset Then
''        If BuildDataRecordFromJson(strAskName, strJson_out, rsOutData, blnShowNotErrMsg, strErrMsg_Out, blnAutoApendRecord) = False Then Exit Function
''    End If
''    BusinessAskForGet = True
''    Exit Function
''errHandle:
''    If ErrCenter() = 1 Then
''        Resume
''    End If
''ShowErrMsg:
''  If Not blnShowNotErrMsg Then MsgBox strErrMsg_Out, vbInformation, gstrSystemName
''End Function

Public Function BuildDataRecordFromJson(ByVal objBusinessAsk As clsBusinessAsk, ByVal strJson As String, ByRef rsOutData_out As ADODB.Recordset, Optional ByRef varData_Out As Variant, _
    Optional ByVal blnShowNotErrMsg As Boolean, Optional strErrMsg_Out As String, Optional ByVal blnAutoApendRecord As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据json串返回记录集
    '入参:objBusinessAsk-业务请求对象
    '     strJson-Json格式的字符串
    '     blnShowNotErrMsg-不显示错误信息
    '     blnAutoApendRecord-自动追回数据到记录集中(多批次调用业务请求有效)

    '出参: rsOutData_out-如果业务成功, 返回请求数据的记录集
    '               如果业务失败时，返回Nothing
    '      varData_Out-返回解析的数据
    '     strErrMsg_Out-如果失败，返回错误原因
    '返回:业务请求成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2019-08-08 12:49:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objdata As clsBusinessData
    Dim objDatasRecord As clsBusinessDatas
    Dim objDatasVar As clsBusinessDatas
    
    Dim objBusinessParsing As clsBusinessParsing
    Dim cllData As Collection
    Dim i As Long, strValue As String
    Dim blnNotValied As Boolean
    Dim j As Long
    
    Set cllData = New Collection
    
    If blnShowNotErrMsg = False Then On Error GoTo errHandle
    
    If strJson = "" Then Exit Function
    
    If rsOutData_out Is Nothing Then
        blnNotValied = True
    ElseIf rsOutData_out.State <> 1 Then
        blnNotValied = True
    End If
    
    Set objBusinessParsing = objBusinessAsk.返回信息
    If objBusinessParsing Is Nothing Then Exit Function
    
    Set objDatasRecord = objBusinessParsing.objReturnRecord
    Set objDatasVar = objBusinessParsing.objReturnData
    
    
    
    
    If (blnAutoApendRecord = False Or blnNotValied) And Not objDatasRecord Is Nothing Then
        Set rsOutData_out = Nothing
        
        If objDatasRecord.Count <> 0 Then
            Set rsOutData_out = New ADODB.Recordset
            With rsOutData_out
                For i = 1 To objDatasRecord.Count
                    Set objdata = objDatasRecord.Item(i)
                    Select Case objdata.数据类型
                    Case BusinessDataType_Num
                        .Fields.Append objdata.数据名称, adDouble, IIf(objdata.长度 = 0, 18, objdata.长度), adFldIsNullable
                    Case BusinessDataType_Date
                        .Fields.Append objdata.数据名称, adDate, 50, adFldIsNullable
                    Case Else
                        .Fields.Append objdata.数据名称, adLongVarChar, IIf(objdata.长度 = 0, 200, objdata.长度), adFldIsNullable
                    End Select
                Next
                .CursorLocation = adUseClient
                .Open , , adOpenStatic, adLockOptimistic
            End With
        End If
    End If
    With rsOutData_out
        If mobjJson.OpenJson(strJson) = False Then Exit Function
        Set cllData = New Collection
        '先取Code或message
        strValue = mobjJson.GetValue("output.code")
        cllData.Add Val(strValue), "code"
        strErrMsg_Out = mobjJson.GetValue("output.message")
        cllData.Add strErrMsg_Out, "message"
        
        If Val(strValue) = 0 Then
            If strErrMsg_Out <> "" And Not blnShowNotErrMsg Then MsgBox strErrMsg_Out, vbInformation + vbOKOnly, gstrSystemName
            Exit Function
        End If
        
        '取Json值
        If Not objDatasVar Is Nothing Then
             For i = 1 To objDatasVar.Count
                 Set objdata = objDatasVar.Item(i)
                 Select Case objdata.数据类型
                 Case BusinessDataType_Num
                    strValue = mobjJson.GetValue("output." & objdata.节点名称)
                    cllData.Add Val(strValue), objdata.节点名称
                 Case BusinessDataType_String
                    strValue = mobjJson.GetValue("output." & objdata.节点名称)
                    cllData.Add strValue, objdata.节点名称
                 Case Else
                    strValue = mobjJson.GetValue("output." & objdata.节点名称)
                    If IsDate(strValue) Then
                        cllData.Add Format(CDate(strValue), "yyyy-mm-dd HH:MM:SS"), objdata.节点名称
                    Else
                        cllData.Add "", objdata.节点名称
                    End If
                 End Select
             Next
        End If
        If cllData.Count = 0 Then
            varData_Out = Empty
         Else
           Set varData_Out = cllData
        End If
        
        If Not objDatasRecord Is Nothing Then
            If objDatasRecord.Count <> 0 Then
                For i = 0 To mobjJson.GetLength("output." & objDatasRecord.列表接点) - 1
                    .AddNew
                    For j = 1 To objDatasRecord.Count
                        Set objdata = objDatasRecord.Item(j)
                        Select Case objdata.数据类型
                        Case BusinessDataType_Num
                            strValue = mobjJson.GetValue("output." & objDatasRecord.列表接点 & "[" & i & "]." & objdata.节点名称)
                           .Fields(objdata.数据名称).Value = Val(strValue)
                        Case BusinessDataType_Date
                            strValue = mobjJson.GetValue("output." & objDatasRecord.列表接点 & "[" & i & "]." & objdata.节点名称)
                            If IsDate(strValue) Then
                                .Fields(objdata.数据名称).Value = CDate(strValue)
                            End If
                        Case Else
                            strValue = mobjJson.GetValue("output." & objDatasRecord.列表接点 & "[" & i & "]." & objdata.节点名称)
                            .Fields(objdata.数据名称).Value = strValue
                        End Select
                    Next
                    .Update
                Next
                If rsOutData_out.RecordCount <> 0 Then rsOutData_out.MoveFirst
            End If
        End If
    End With
    BuildDataRecordFromJson = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetBusinessDatas(ByVal strBusinessAskName As String) As clsBusinessParsing
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据业务请求名称，获取返回数据集对象
    '入参:BusinessAskName-业务请求名称（现以过程名为准)
    '出参:
    '返回：返回clsBusinessDatas对象
    '编制:刘兴洪
    '日期:2017-05-24 16:32:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objBusinessAsk As clsBusinessAsk
    On Error GoTo errHandle
    For Each objBusinessAsk In mobjBusinessAsks
        If UCase(objBusinessAsk.过程名) = UCase(strBusinessAskName) Then
            Set GetBusinessDatas = objBusinessAsk.返回信息: Exit Function
        End If
    Next
    Set GetBusinessDatas = New clsBusinessParsing
    Exit Function
errHandle:
    Set GetBusinessDatas = New clsBusinessParsing
End Function
Private Sub Class_Initialize()
    Call InitBusinessData
    mdtSecurity = CDate("1901-01-01")
    Set mobjJson = New clsJson
    
    If M_CONNET_TYPE = 0 Then
        Set mobjHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    End If
End Sub

Private Sub Class_Terminate()
    Err = 0: On Error Resume Next
    Set mobjJson = Nothing
    Set mobjHttpRequest = Nothing
    Set mobjBusinessAsks = Nothing
End Sub
  
