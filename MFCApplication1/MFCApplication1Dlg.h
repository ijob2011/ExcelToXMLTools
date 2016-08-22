
// MFCApplication1Dlg.h : 头文件
//

#pragma once
#include "afxeditbrowsectrl.h"
#include "IllusionExcelFile.h"
#include "afxwin.h"
#include "afxcmn.h"
#include <string>

class CMFCApplication1DlgAutoProxy;

enum OutTextFont
{
	DEFOUT = 0,//默认值 
	C_RED,//红色字体 205, 38, 38 #CD2626
	C_GREEN,//绿色 0, 100, 0 #006400
	C_BLUE, //蓝色 0, 51, 204 #0033cc
};
#define C_RED_RGB RGB(205, 38, 38)//红色字体 205, 38, 38 #CD2626
#define C_GREEN_RGB RGB(0, 100, 0)//绿色 0, 100, 0 #006400
#define C_BLUE_RGB RGB( 0, 51, 204)//蓝色 0, 51, 204 #0033cc


// CMFCApplication1Dlg 对话框
class CMFCApplication1Dlg : public CDialogEx
{
	DECLARE_DYNAMIC(CMFCApplication1Dlg);
	friend class CMFCApplication1DlgAutoProxy;
	

// 构造
public:
	CMFCApplication1Dlg(CWnd* pParent = NULL);	// 标准构造函数
	virtual ~CMFCApplication1Dlg();

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_MFCAPPLICATION1_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

// 实现
protected:
	CMFCApplication1DlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	afx_msg void OnClose();
	virtual void OnOK();
	virtual void OnCancel();
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedExec();
	afx_msg void OnEnChangeBrowseExcel();
	// 选择excel文件控件
	CMFCEditBrowseCtrl m_browse_excel;
	IllusionExcelFile* m_excel_util;
	// sheet页下拉框
	CComboBox m_boxSelect_sheet;
	afx_msg void OnCbnSelchangeComboSheet();
	// 输出消息
	CRichEditCtrl m_rich_edit_out_text;
	//往输出消息框添加文本
	void AppendLineToOutText(const CString& msg, OutTextFont fout = OutTextFont::DEFOUT);

	// 输出xml文件框
	CMFCEditBrowseCtrl m_browse_out_xml;
	afx_msg void OnBnClickedCheckSaveAll();
	// 是否保存全部sheet页
	CButton m_check_save_all;
	// 是否合并所有sheet到输出到同一个xml
	CButton m_check_merge_all;
	void AutoFillOutXmlPathName();

	void DoExport(const CString& sheetName);
	void DoExportToOneMergerXml(IllusionExcelFile& excl);
	afx_msg void OnDropFiles(HDROP hDropInfo);
	afx_msg void OnBnClickedCheckMergeAll();
};
