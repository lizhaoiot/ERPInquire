#pragma once

// DrawingPropPage.h : CDrawingPropPage 属性页类的声明。


// CDrawingPropPage : 有关实现的信息，请参阅 DrawingPropPage.cpp。

class CDrawingPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CDrawingPropPage)
	DECLARE_OLECREATE_EX(CDrawingPropPage)

// 构造函数
public:
	CDrawingPropPage();

// 对话框数据
	enum { IDD = IDD_PROPPAGE_DRAWING };

// 实现
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 消息映射
protected:
	DECLARE_MESSAGE_MAP()
};

