// Drawing.cpp : CDrawingApp �� DLL ע���ʵ�֡�

#include "stdafx.h"
#include "Drawing.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

CDrawingApp theApp;

const GUID CDECL _tlid = { 0x5310DC95, 0xD310, 0x4FAE, { 0x8A, 0x0, 0x1D, 0xC6, 0xF4, 0x45, 0x33, 0x74 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;

// CDrawingApp::InitInstance - DLL ��ʼ��

BOOL CDrawingApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO:  �ڴ�������Լ���ģ���ʼ�����롣
	}

	return bInit;
}

// CDrawingApp::ExitInstance - DLL ��ֹ

int CDrawingApp::ExitInstance()
{
	// TODO:  �ڴ�������Լ���ģ����ֹ���롣

	return COleControlModule::ExitInstance();
}

// DllRegisterServer - ������ӵ�ϵͳע���

STDAPI DllRegisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}

// DllUnregisterServer - �����ϵͳע������Ƴ�

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}