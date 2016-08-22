
// DlgProxy.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "MFCApplication1.h"
#include "DlgProxy.h"
#include "MFCApplication1Dlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CMFCApplication1DlgAutoProxy

IMPLEMENT_DYNCREATE(CMFCApplication1DlgAutoProxy, CCmdTarget)

CMFCApplication1DlgAutoProxy::CMFCApplication1DlgAutoProxy()
{
	EnableAutomation();
	
	// ΪʹӦ�ó������Զ��������ڻ״̬ʱһֱ���� 
	//	���У����캯������ AfxOleLockApp��
	AfxOleLockApp();

	// ͨ��Ӧ�ó����������ָ��
	//  �����ʶԻ���  ���ô�����ڲ�ָ��
	//  ָ��Ի��򣬲����öԻ���ĺ���ָ��ָ��
	//  �ô���
	ASSERT_VALID(AfxGetApp()->m_pMainWnd);
	if (AfxGetApp()->m_pMainWnd)
	{
		ASSERT_KINDOF(CMFCApplication1Dlg, AfxGetApp()->m_pMainWnd);
		if (AfxGetApp()->m_pMainWnd->IsKindOf(RUNTIME_CLASS(CMFCApplication1Dlg)))
		{
			m_pDialog = reinterpret_cast<CMFCApplication1Dlg*>(AfxGetApp()->m_pMainWnd);
			m_pDialog->m_pAutoProxy = this;
		}
	}
}

CMFCApplication1DlgAutoProxy::~CMFCApplication1DlgAutoProxy()
{
	// Ϊ������ OLE �Զ����������ж������ֹӦ�ó���
	//	������������ AfxOleUnlockApp��
	//  ���������������⣬�⻹���������Ի���
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CMFCApplication1DlgAutoProxy::OnFinalRelease()
{
	// �ͷ��˶��Զ�����������һ�����ú󣬽�����
	// OnFinalRelease��  ���ཫ�Զ�
	// ɾ���ö���  �ڵ��øû���֮ǰ�����������
	// ��������ĸ���������롣

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CMFCApplication1DlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CMFCApplication1DlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// ע��: ��������˶� IID_IMFCApplication1 ��֧��
//  ��֧������ VBA �����Ͱ�ȫ�󶨡�  �� IID ����ͬ���ӵ� .IDL �ļ��е�
//  ���Ƚӿڵ� GUID ƥ�䡣

// {F2C8AFED-F83E-422F-9A0E-C2B014B92002}
static const IID IID_IMFCApplication1 =
{ 0xF2C8AFED, 0xF83E, 0x422F, { 0x9A, 0xE, 0xC2, 0xB0, 0x14, 0xB9, 0x20, 0x2 } };

BEGIN_INTERFACE_MAP(CMFCApplication1DlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CMFCApplication1DlgAutoProxy, IID_IMFCApplication1, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 ���ڴ���Ŀ�� StdAfx.h �ж���
// {C8BF63CE-5E0A-492E-A0E7-DA91B5A31497}
IMPLEMENT_OLECREATE2(CMFCApplication1DlgAutoProxy, "MFCApplication1.Application", 0xc8bf63ce, 0x5e0a, 0x492e, 0xa0, 0xe7, 0xda, 0x91, 0xb5, 0xa3, 0x14, 0x97)


// CMFCApplication1DlgAutoProxy ��Ϣ�������
