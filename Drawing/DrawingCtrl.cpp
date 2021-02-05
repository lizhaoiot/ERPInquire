// DrawingCtrl.cpp : CDrawingCtrl ActiveX �ؼ����ʵ�֡�

#include "stdafx.h"
#include "Drawing.h"
#include "DrawingCtrl.h"
#include "DrawingPropPage.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CDrawingCtrl, COleControl)

// ��Ϣӳ��

BEGIN_MESSAGE_MAP(CDrawingCtrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()

// ����ӳ��

BEGIN_DISPATCH_MAP(CDrawingCtrl, COleControl)
	DISP_FUNCTION_ID(CDrawingCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

// �¼�ӳ��

BEGIN_EVENT_MAP(CDrawingCtrl, COleControl)
END_EVENT_MAP()

// ����ҳ

// TODO: ������Ҫ��Ӹ�������ҳ�����ס���Ӽ���!
BEGIN_PROPPAGEIDS(CDrawingCtrl, 1)
	PROPPAGEID(CDrawingPropPage::guid)
END_PROPPAGEIDS(CDrawingCtrl)

// ��ʼ���๤���� guid

IMPLEMENT_OLECREATE_EX(CDrawingCtrl, "DRAWING.DrawingCtrl.1",
	0xb97fe269, 0x5b30, 0x43fe, 0xb1, 0x47, 0x41, 0xaf, 0x88, 0xb7, 0x20, 0xac)

// ����� ID �Ͱ汾

IMPLEMENT_OLETYPELIB(CDrawingCtrl, _tlid, _wVerMajor, _wVerMinor)

// �ӿ� ID

const IID IID_DDrawing = { 0x1A91DCCA, 0xB2E6, 0x4C5F, { 0xAE, 0x71, 0xC2, 0xFF, 0xB1, 0xF4, 0x33, 0x5D } };
const IID IID_DDrawingEvents = { 0xD1D6788F, 0xB824, 0x44F7, { 0x8C, 0x68, 0x9D, 0xEE, 0x32, 0x31, 0xC, 0x9A } };

// �ؼ�������Ϣ

static const DWORD _dwDrawingOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CDrawingCtrl, IDS_DRAWING, _dwDrawingOleMisc)

// CDrawingCtrl::CDrawingCtrlFactory::UpdateRegistry -
// ��ӻ��Ƴ� CDrawingCtrl ��ϵͳע�����

BOOL CDrawingCtrl::CDrawingCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO:  ��֤���Ŀؼ��Ƿ���ϵ�Ԫģ���̴߳������
	// �йظ�����Ϣ����ο� MFC ����˵�� 64��
	// ������Ŀؼ������ϵ�Ԫģ�͹�����
	// �����޸����´��룬��������������
	// afxRegApartmentThreading ��Ϊ 0��

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_DRAWING,
			IDB_DRAWING,
			afxRegApartmentThreading,
			_dwDrawingOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}


// ��Ȩ�ַ���

static const TCHAR _szLicFileName[] = _T("Drawing.lic");
static const WCHAR _szLicString[] = L"Copyright (c) 2018 ";

// CDrawingCtrl::CDrawingCtrlFactory::VerifyUserLicense -
// ����Ƿ�����û����֤

BOOL CDrawingCtrl::CDrawingCtrlFactory::VerifyUserLicense()
{
	return AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);
}

// CDrawingCtrl::CDrawingCtrlFactory::GetLicenseKey -
// ��������ʱ��Ȩ��Կ

BOOL CDrawingCtrl::CDrawingCtrlFactory::GetLicenseKey(DWORD dwReserved,
	BSTR *pbstrKey)
{
	if (pbstrKey == NULL)
		return FALSE;

	*pbstrKey = SysAllocString(_szLicString);
	return (*pbstrKey != NULL);
}


// CDrawingCtrl::CDrawingCtrl - ���캯��

CDrawingCtrl::CDrawingCtrl()
{
	InitializeIIDs(&IID_DDrawing, &IID_DDrawingEvents);
	// TODO:  �ڴ˳�ʼ���ؼ���ʵ�����ݡ�
}

// CDrawingCtrl::~CDrawingCtrl - ��������

CDrawingCtrl::~CDrawingCtrl()
{
	// TODO:  �ڴ�����ؼ���ʵ�����ݡ�
}

// CDrawingCtrl::OnDraw - ��ͼ����

void CDrawingCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& /* rcInvalid */)
{
	if (!pdc)
		return;

	// TODO:  �����Լ��Ļ�ͼ�����滻����Ĵ��롣
	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(WHITE_BRUSH)));
	pdc->Ellipse(rcBounds);
}

// CDrawingCtrl::DoPropExchange - �־���֧��

void CDrawingCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Ϊÿ���־õ��Զ������Ե��� PX_ ������
}


// CDrawingCtrl::OnResetState - ���ؼ�����ΪĬ��״̬

void CDrawingCtrl::OnResetState()
{
	COleControl::OnResetState();  // ���� DoPropExchange ���ҵ���Ĭ��ֵ

	// TODO:  �ڴ��������������ؼ�״̬��
}


// CDrawingCtrl::AboutBox - ���û���ʾ�����ڡ���

void CDrawingCtrl::AboutBox()
{
	CDialogEx dlgAbout(IDD_ABOUTBOX_DRAWING);
	dlgAbout.DoModal();
}


// CDrawingCtrl ��Ϣ�������
