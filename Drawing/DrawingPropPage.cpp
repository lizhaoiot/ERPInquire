// DrawingPropPage.cpp : CDrawingPropPage 属性页类的实现。

#include "stdafx.h"
#include "Drawing.h"
#include "DrawingPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CDrawingPropPage, COlePropertyPage)

// 消息映射

BEGIN_MESSAGE_MAP(CDrawingPropPage, COlePropertyPage)
END_MESSAGE_MAP()

// 初始化类工厂和 guid

IMPLEMENT_OLECREATE_EX(CDrawingPropPage, "DRAWING.DrawingPropPage.1",
	0x4a3eedee, 0x7a07, 0x40fc, 0x94, 0x97, 0xb, 0x79, 0x98, 0x6f, 0xd3, 0x7b)

// CDrawingPropPage::CDrawingPropPageFactory::UpdateRegistry -
// 添加或移除 CDrawingPropPage 的系统注册表项

BOOL CDrawingPropPage::CDrawingPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_DRAWING_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}

// CDrawingPropPage::CDrawingPropPage - 构造函数

CDrawingPropPage::CDrawingPropPage() :
	COlePropertyPage(IDD, IDS_DRAWING_PPG_CAPTION)
{
}

// CDrawingPropPage::DoDataExchange - 在页和属性间移动数据

void CDrawingPropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_PostProcessing(pDX);
}

// CDrawingPropPage 消息处理程序
