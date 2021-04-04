// zlNoticeLib.cpp: 定义 DLL 应用程序的导出函数。
//

#pragma once
#include "windows.h"
#include "ocilib.h"
#include <string>
using namespace std;

#pragma comment(lib, "ociliba.lib")
#pragma comment(lib,"comsupp.lib")

HWND formHandler;
OCI_Connection *con;
OCI_Subscription *sub;

void event_handler(OCI_Event *event);
void __stdcall error_handler(OCI_Error *err);


extern "C"  boolean  __stdcall  OCI_ConnCreate(BSTR Database, BSTR User, BSTR Pwd) {
	
	OCI_Initialize(NULL, NULL, OCI_ENV_EVENTS);
	con = OCI_ConnectionCreate((const char *)Database, (const char *)User, (const char *)Pwd, OCI_SESSION_DEFAULT);
	sub = OCI_SubscriptionRegister(con, "sub-00", OCI_CNT_ALL, event_handler, 65535, 0);
	return (con != NULL);	//Connection handle on success or NULL on failure
}

extern "C"  void   __stdcall  OCI_Register(HWND lngHandler,const char * dataTable) {
	formHandler = lngHandler;
	OCI_SetErrorHandler(error_handler);

	OCI_Statement *st;
	st = OCI_StatementCreate(con);

	OCI_Prepare(st, dataTable);
	OCI_SubscriptionAddStatement(sub, st);
}

//通知回调函数
void  event_handler(OCI_Event *event)
{
	const char *rowid;
	const char *object;
	string opration;
	string strResult;

	if (OCI_EventGetType(event) == OCI_ENT_OBJECT_CHANGED) {
		switch( OCI_EventGetOperation(event) ){
			case OCI_ONT_INSERT:
				opration = "1";
				break;
			case OCI_ONT_UPDATE:
				opration = "2";
				break;
			case OCI_ONT_DELETE:
				opration = "3";
				break;
			default:
				return;	//如果不是增改类型,就直接退出.
		}
	}
	else {
		return;	//如果不是增改类型,就直接退出.
	}

	rowid = OCI_EventGetRowid(event);
	if (strlen(rowid)<18){
		rowid="1";
	}
	object = OCI_EventGetObject(event);
	strResult = strResult + opration + "-" +  object + "-" + rowid ;

	PostMessage(formHandler, WM_USER + 1, (long)SysAllocString((BSTR)strResult.c_str()), 0);	//发送消息
}



//错误回调函数
void __stdcall error_handler(OCI_Error *err)
{
	int  err_type = OCI_ErrorGetType(err);
	const char *err_msg  = OCI_ErrorGetString(err);

	string strResult = err_type == OCI_ERR_WARNING ? "Warning:" : "Error:";
	strResult = strResult + err_msg;

	PostMessage(formHandler, WM_USER + 1, (long)SysAllocString((BSTR)strResult.c_str()), 0);
}

extern "C"  void  __stdcall OCI_UnRigister()
{
	OCI_SubscriptionUnregister(sub);
}