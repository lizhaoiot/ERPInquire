// DrawingCtrl.cpp : CDrawingCtrl ActiveX 控件类的实现。

#include "stdafx.h"
#include "Drawing.h"
#include "DrawingCtrl.h"
#include "DrawingPropPage.h"
#include "afxdialogex.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CDrawingCtrl, COleControl)

// 消息映射

BEGIN_MESSAGE_MAP(CDrawingCtrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()

// 调度映射

BEGIN_DISPATCH_MAP(CDrawingCtrl, COleControl)
	DISP_FUNCTION_ID(CDrawingCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
END_DISPATCH_MAP()

// 事件映射

BEGIN_EVENT_MAP(CDrawingCtrl, COleControl)
END_EVENT_MAP()

// 属性页

// TODO: 根据需要添加更多属性页。请记住增加计数!
BEGIN_PROPPAGEIDS(CDrawingCtrl, 1)
	PROPPAGEID(CDrawingPropPage::guid)
END_PROPPAGEIDS(CDrawingCtrl)

// 初始化类工厂和 guid

IMPLEMENT_OLECREATE_EX(CDrawingCtrl, "DRAWING.DrawingCtrl.1",
	0xb97fe269, 0x5b30, 0x43fe, 0xb1, 0x47, 0x41, 0xaf, 0x88, 0xb7, 0x20, 0xac)

// 键入库 ID 和版本

IMPLEMENT_OLETYPELIB(CDrawingCtrl, _tlid, _wVerMajor, _wVerMinor)

// 接口 ID

const IID IID_DDrawing = { 0x1A91DCCA, 0xB2E6, 0x4C5F, { 0xAE, 0x71, 0xC2, 0xFF, 0xB1, 0xF4, 0x33, 0x5D } };
const IID IID_DDrawingEvents = { 0xD1D6788F, 0xB824, 0x44F7, { 0x8C, 0x68, 0x9D, 0xEE, 0x32, 0x31, 0xC, 0x9A } };

// 控件类型信息

static const DWORD _dwDrawingOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CDrawingCtrl, IDS_DRAWING, _dwDrawingOleMisc)

// CDrawingCtrl::CDrawingCtrlFactory::UpdateRegistry -
// 添加或移除 CDrawingCtrl 的系统注册表项

BOOL CDrawingCtrl::CDrawingCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO:  验证您的控件是否符合单元模型线程处理规则。
	// 有关更多信息，请参考 MFC 技术说明 64。
	// 如果您的控件不符合单元模型规则，则
	// 必须修改如下代码，将第六个参数从
	// afxRegApartmentThreading 改为 0。

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


// 授权字符串

static const TCHAR _szLicFileName[] = _T("Drawing.lic");
static const WCHAR _szLicString[] = L"Copyright (c) 2018 ";

// CDrawingCtrl::CDrawingCtrlFactory::VerifyUserLicense -
// 检查是否存在用户许可证

BOOL CDrawingCtrl::CDrawingCtrlFactory::VerifyUserLicense()
{
	return AfxVerifyLicFile(AfxGetInstanceHandle(), _szLicFileName,
		_szLicString);
}

// CDrawingCtrl::CDrawingCtrlFactory::GetLicenseKey -
// 返回运行时授权密钥

BOOL CDrawingCtrl::CDrawingCtrlFactory::GetLicenseKey(DWORD dwReserved,
	BSTR *pbstrKey)
{
	if (pbstrKey == NULL)
		return FALSE;

	*pbstrKey = SysAllocString(_szLicString);
	return (*pbstrKey != NULL);
}


// CDrawingCtrl::CDrawingCtrl - 构造函数

CDrawingCtrl::CDrawingCtrl()
{
	InitializeIIDs(&IID_DDrawing, &IID_DDrawingEvents);
	// TODO:  在此初始化控件的实例数据。
}

// CDrawingCtrl::~CDrawingCtrl - 析构函数

CDrawingCtrl::~CDrawingCtrl()
{
	// TODO:  在此清理控件的实例数据。
}

// CDrawingCtrl::OnDraw - 绘图函数

void CDrawingCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& /* rcInvalid */)
{
	if (!pdc)
		return;

	// TODO:  用您自己的绘图代码替换下面的代码。
	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(WHITE_BRUSH)));
	pdc->Ellipse(rcBounds);
}

// CDrawingCtrl::DoPropExchange - 持久性支持

void CDrawingCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: 为每个持久的自定义属性调用 PX_ 函数。
}


// CDrawingCtrl::OnResetState - 将控件重置为默认状态

void CDrawingCtrl::OnResetState()
{
	COleControl::OnResetState();  // 重置 DoPropExchange 中找到的默认值

	// TODO:  在此重置任意其他控件状态。
}


// CDrawingCtrl::AboutBox - 向用户显示“关于”框

void CDrawingCtrl::AboutBox()
{
	CDialogEx dlgAbout(IDD_ABOUTBOX_DRAWING);
	dlgAbout.DoModal();
}


// CDrawingCtrl 消息处理程序
