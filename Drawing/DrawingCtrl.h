#pragma once

// DrawingCtrl.h : CDrawingCtrl ActiveX �ؼ����������


// CDrawingCtrl : �й�ʵ�ֵ���Ϣ������� DrawingCtrl.cpp��

class CDrawingCtrl : public COleControl
{
	DECLARE_DYNCREATE(CDrawingCtrl)

// ���캯��
public:
	CDrawingCtrl();

// ��д
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

// ʵ��
protected:
	~CDrawingCtrl();

	BEGIN_OLEFACTORY(CDrawingCtrl)        // �๤���� guid
		virtual BOOL VerifyUserLicense();
		virtual BOOL GetLicenseKey(DWORD, BSTR *);
	END_OLEFACTORY(CDrawingCtrl)

	DECLARE_OLETYPELIB(CDrawingCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CDrawingCtrl)     // ����ҳ ID
	DECLARE_OLECTLTYPE(CDrawingCtrl)		// �������ƺ�����״̬

// ��Ϣӳ��
	DECLARE_MESSAGE_MAP()

// ����ӳ��
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// �¼�ӳ��
	DECLARE_EVENT_MAP()

// ���Ⱥ��¼� ID
public:
	enum {
	};
};

