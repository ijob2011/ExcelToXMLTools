
// MFCApplication1Dlg.h : ͷ�ļ�
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
	DEFOUT = 0,//Ĭ��ֵ 
	C_RED,//��ɫ���� 205, 38, 38 #CD2626
	C_GREEN,//��ɫ 0, 100, 0 #006400
	C_BLUE, //��ɫ 0, 51, 204 #0033cc
};
#define C_RED_RGB RGB(205, 38, 38)//��ɫ���� 205, 38, 38 #CD2626
#define C_GREEN_RGB RGB(0, 100, 0)//��ɫ 0, 100, 0 #006400
#define C_BLUE_RGB RGB( 0, 51, 204)//��ɫ 0, 51, 204 #0033cc


// CMFCApplication1Dlg �Ի���
class CMFCApplication1Dlg : public CDialogEx
{
	DECLARE_DYNAMIC(CMFCApplication1Dlg);
	friend class CMFCApplication1DlgAutoProxy;
	

// ����
public:
	CMFCApplication1Dlg(CWnd* pParent = NULL);	// ��׼���캯��
	virtual ~CMFCApplication1Dlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_MFCAPPLICATION1_DIALOG };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��

// ʵ��
protected:
	CMFCApplication1DlgAutoProxy* m_pAutoProxy;
	HICON m_hIcon;

	BOOL CanExit();

	// ���ɵ���Ϣӳ�亯��
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
	// ѡ��excel�ļ��ؼ�
	CMFCEditBrowseCtrl m_browse_excel;
	IllusionExcelFile* m_excel_util;
	// sheetҳ������
	CComboBox m_boxSelect_sheet;
	afx_msg void OnCbnSelchangeComboSheet();
	// �����Ϣ
	CRichEditCtrl m_rich_edit_out_text;
	//�������Ϣ������ı�
	void AppendLineToOutText(const CString& msg, OutTextFont fout = OutTextFont::DEFOUT);

	// ���xml�ļ���
	CMFCEditBrowseCtrl m_browse_out_xml;
	afx_msg void OnBnClickedCheckSaveAll();
	// �Ƿ񱣴�ȫ��sheetҳ
	CButton m_check_save_all;
	// �Ƿ�ϲ�����sheet�������ͬһ��xml
	CButton m_check_merge_all;
	void AutoFillOutXmlPathName();

	void DoExport(const CString& sheetName);
	void DoExportToOneMergerXml(IllusionExcelFile& excl);
	afx_msg void OnDropFiles(HDROP hDropInfo);
	afx_msg void OnBnClickedCheckMergeAll();
};
