#pragma once

// DrawingCtrl.h : CDrawingCtrl ActiveX 控件类的声明。


// CDrawingCtrl : 有关实现的信息，请参阅 DrawingCtrl.cpp。

class CDrawingCtrl : public COleControl
{
	DECLARE_DYNCREATE(CDrawingCtrl)

// 构造函数
public:
	CDrawingCtrl();

// 重写
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

// 实现
protected:
	~CDrawingCtrl();

	BEGIN_OLEFACTORY(CDrawingCtrl)        // 类工厂和 guid
		virtual BOOL VerifyUserLicense();
		virtual BOOL GetLicenseKey(DWORD, BSTR *);
	END_OLEFACTORY(CDrawingCtrl)

	DECLARE_OLETYPELIB(CDrawingCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CDrawingCtrl)     // 属性页 ID
	DECLARE_OLECTLTYPE(CDrawingCtrl)		// 类型名称和杂项状态

// 消息映射
	DECLARE_MESSAGE_MAP()

// 调度映射
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// 事件映射
	DECLARE_EVENT_MAP()

// 调度和事件 ID
public:
	enum {
	};
};

