// DrawingPropPage.cpp : CDrawingPropPage ����ҳ���ʵ�֡�

#include "stdafx.h"
#include "Drawing.h"
#include "DrawingPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

IMPLEMENT_DYNCREATE(CDrawingPropPage, COlePropertyPage)

// ��Ϣӳ��

BEGIN_MESSAGE_MAP(CDrawingPropPage, COlePropertyPage)
END_MESSAGE_MAP()

// ��ʼ���๤���� guid

IMPLEMENT_OLECREATE_EX(CDrawingPropPage, "DRAWING.DrawingPropPage.1",
	0x4a3eedee, 0x7a07, 0x40fc, 0x94, 0x97, 0xb, 0x79, 0x98, 0x6f, 0xd3, 0x7b)

// CDrawingPropPage::CDrawingPropPageFactory::UpdateRegistry -
// ��ӻ��Ƴ� CDrawingPropPage ��ϵͳע�����

BOOL CDrawingPropPage::CDrawingPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_DRAWING_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}

// CDrawingPropPage::CDrawingPropPage - ���캯��

CDrawingPropPage::CDrawingPropPage() :
	COlePropertyPage(IDD, IDS_DRAWING_PPG_CAPTION)
{
}

// CDrawingPropPage::DoDataExchange - ��ҳ�����Լ��ƶ�����

void CDrawingPropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_PostProcessing(pDX);
}

// CDrawingPropPage ��Ϣ�������
