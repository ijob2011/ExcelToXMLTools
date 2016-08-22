
// DlgProxy.cpp : 实现文件
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
	
	// 为使应用程序在自动化对象处于活动状态时一直保持 
	//	运行，构造函数调用 AfxOleLockApp。
	AfxOleLockApp();

	// 通过应用程序的主窗口指针
	//  来访问对话框。  设置代理的内部指针
	//  指向对话框，并设置对话框的后向指针指向
	//  该代理。
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
	// 为了在用 OLE 自动化创建所有对象后终止应用程序，
	//	析构函数调用 AfxOleUnlockApp。
	//  除了做其他事情外，这还将销毁主对话框
	if (m_pDialog != NULL)
		m_pDialog->m_pAutoProxy = NULL;
	AfxOleUnlockApp();
}

void CMFCApplication1DlgAutoProxy::OnFinalRelease()
{
	// 释放了对自动化对象的最后一个引用后，将调用
	// OnFinalRelease。  基类将自动
	// 删除该对象。  在调用该基类之前，请添加您的
	// 对象所需的附加清理代码。

	CCmdTarget::OnFinalRelease();
}

BEGIN_MESSAGE_MAP(CMFCApplication1DlgAutoProxy, CCmdTarget)
END_MESSAGE_MAP()

BEGIN_DISPATCH_MAP(CMFCApplication1DlgAutoProxy, CCmdTarget)
END_DISPATCH_MAP()

// 注意: 我们添加了对 IID_IMFCApplication1 的支持
//  以支持来自 VBA 的类型安全绑定。  此 IID 必须同附加到 .IDL 文件中的
//  调度接口的 GUID 匹配。

// {F2C8AFED-F83E-422F-9A0E-C2B014B92002}
static const IID IID_IMFCApplication1 =
{ 0xF2C8AFED, 0xF83E, 0x422F, { 0x9A, 0xE, 0xC2, 0xB0, 0x14, 0xB9, 0x20, 0x2 } };

BEGIN_INTERFACE_MAP(CMFCApplication1DlgAutoProxy, CCmdTarget)
	INTERFACE_PART(CMFCApplication1DlgAutoProxy, IID_IMFCApplication1, Dispatch)
END_INTERFACE_MAP()

// IMPLEMENT_OLECREATE2 宏在此项目的 StdAfx.h 中定义
// {C8BF63CE-5E0A-492E-A0E7-DA91B5A31497}
IMPLEMENT_OLECREATE2(CMFCApplication1DlgAutoProxy, "MFCApplication1.Application", 0xc8bf63ce, 0x5e0a, 0x492e, 0xa0, 0xe7, 0xda, 0x91, 0xb5, 0xa3, 0x14, 0x97)


// CMFCApplication1DlgAutoProxy 消息处理程序
