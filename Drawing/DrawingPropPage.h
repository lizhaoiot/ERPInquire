#pragma once

// DrawingPropPage.h : CDrawingPropPage ����ҳ���������


// CDrawingPropPage : �й�ʵ�ֵ���Ϣ������� DrawingPropPage.cpp��

class CDrawingPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CDrawingPropPage)
	DECLARE_OLECREATE_EX(CDrawingPropPage)

// ���캯��
public:
	CDrawingPropPage();

// �Ի�������
	enum { IDD = IDD_PROPPAGE_DRAWING };

// ʵ��
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ��Ϣӳ��
protected:
	DECLARE_MESSAGE_MAP()
};

