
// MFCApplication1Dlg.cpp : ʵ���ļ�
//

#include "stdafx.h"
#include "MFCApplication1.h"
#include "MFCApplication1Dlg.h"
#include "DlgProxy.h"
#include "afxdialogex.h"
//#include "CApplication.h"
//#include "CWorkbooks.h"
//#include "CWorkbook.h"
//#include "CWorksheets.h"
//#include "CWorksheet.h"
//#include "CRange.h"
#include <string>
#include <sstream>
#include <locale.h>
#include <fstream>
#include <iostream>
#include <ostream>

#include <list>
#include <vector>
#include <string>
#include <map>
#include <utility>

#include "stdafx.h"
#include "tinyxml2.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif

using namespace std;
using namespace tinyxml2;

// ����Ӧ�ó��򡰹��ڡ��˵���� CAboutDlg �Ի���
string UnicodeToUTF8(const wstring& str)
{
	char*     pElementText;
	int    iTextLen;
	// wide char to multi char
	iTextLen = WideCharToMultiByte(CP_UTF8,
		0,
		str.c_str(),
		-1,
		NULL,
		0,
		NULL,
		NULL);
	pElementText = new char[iTextLen + 1];
	memset((void*)pElementText, 0, sizeof(char) * (iTextLen + 1));
	::WideCharToMultiByte(CP_UTF8,
		0,
		str.c_str(),
		-1,
		pElementText,
		iTextLen,
		NULL,
		NULL);
	string strText;
	strText = pElementText;
	delete[] pElementText;
	return strText;
}

wstring UTF8ToUnicode(const string& str)
{
	int  len = 0;
	len = str.length();
	int  unicodeLen = ::MultiByteToWideChar(CP_UTF8,
		0,
		str.c_str(),
		-1,
		NULL,
		0);
	wchar_t *  pUnicode;
	pUnicode = new  wchar_t[unicodeLen + 1];
	memset(pUnicode, 0, (unicodeLen + 1) * sizeof(wchar_t));
	::MultiByteToWideChar(CP_UTF8,
		0,
		str.c_str(),
		-1,
		(LPWSTR)pUnicode,
		unicodeLen);
	wstring  rt;
	rt = (wchar_t*)pUnicode;
	delete  pUnicode;

	return  rt;
}

wstring ANSIToUnicode(const string& str)
{
	int  len = 0;
	len = str.length();
	int  unicodeLen = ::MultiByteToWideChar(CP_ACP,
		0,
		str.c_str(),
		-1,
		NULL,
		0);
	wchar_t *  pUnicode;
	pUnicode = new  wchar_t[unicodeLen + 1];
	memset(pUnicode, 0, (unicodeLen + 1) * sizeof(wchar_t));
	::MultiByteToWideChar(CP_ACP,
		0,
		str.c_str(),
		-1,
		(LPWSTR)pUnicode,
		unicodeLen);
	wstring  rt;
	rt = (wchar_t*)pUnicode;
	delete  pUnicode;

	return  rt;
}

string UnicodeToANSI(const wstring& str)
{
	char*     pElementText;
	int    iTextLen;
	// wide char to multi char
	iTextLen = WideCharToMultiByte(CP_ACP,
		0,
		str.c_str(),
		-1,
		NULL,
		0,
		NULL,
		NULL);
	pElementText = new char[iTextLen + 1];
	memset((void*)pElementText, 0, sizeof(char) * (iTextLen + 1));
	::WideCharToMultiByte(CP_ACP,
		0,
		str.c_str(),
		-1,
		pElementText,
		iTextLen,
		NULL,
		NULL);
	string strText;
	strText = pElementText;
	delete[] pElementText;
	return strText;
}

//�ַ����ָ��
std::vector<std::string> split(std::string str, std::string pattern)
{
	std::string::size_type pos;
	std::vector<std::string> result;
	str += pattern;//��չ�ַ����Է������
	int size = str.size();

	for (int i = 0; i < size; i++)
	{
		pos = str.find(pattern, i);
		if (pos < size)
		{
			std::string s = str.substr(i, pos - i);
			result.push_back(s);
			i = pos + pattern.size() - 1;
		}
	}
	return result;
}

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// �Ի�������
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV ֧��

// ʵ��
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(IDD_ABOUTBOX)
{
	EnableActiveAccessibility();
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CMFCApplication1Dlg �Ի���


IMPLEMENT_DYNAMIC(CMFCApplication1Dlg, CDialogEx);

CMFCApplication1Dlg::CMFCApplication1Dlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(IDD_MFCAPPLICATION1_DIALOG, pParent)
{
	EnableActiveAccessibility();
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
	m_pAutoProxy = NULL;
}

CMFCApplication1Dlg::~CMFCApplication1Dlg()
{
	// ����öԻ������Զ���������
	//  ���˴���ָ��öԻ���ĺ���ָ������Ϊ NULL���Ա�
	//  �˴���֪���öԻ����ѱ�ɾ����
	if (m_pAutoProxy != NULL)
		m_pAutoProxy->m_pDialog = NULL;
	m_excel_util->CloseExcelFile();
	m_excel_util->ReleaseExcel();
	delete m_excel_util;
}

void CMFCApplication1Dlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_BROWSE_EXCEL, m_browse_excel);
	DDX_Control(pDX, IDC_COMBO_SHEET, m_boxSelect_sheet);
	DDX_Control(pDX, IDC_RICHEDIT_OUT_TEXT, m_rich_edit_out_text);
	DDX_Control(pDX, IDC_BROWSE_OUT_XML, m_browse_out_xml);
	DDX_Control(pDX, IDC_CHECK_SAVE_ALL, m_check_save_all);
}

BEGIN_MESSAGE_MAP(CMFCApplication1Dlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_CLOSE()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	ON_BN_CLICKED(ID_EXEC, &CMFCApplication1Dlg::OnBnClickedExec)
	ON_EN_CHANGE(IDC_BROWSE_EXCEL, &CMFCApplication1Dlg::OnEnChangeBrowseExcel)
	ON_CBN_SELCHANGE(IDC_COMBO_SHEET, &CMFCApplication1Dlg::OnCbnSelchangeComboSheet)
	ON_BN_CLICKED(IDC_CHECK_SAVE_ALL, &CMFCApplication1Dlg::OnBnClickedCheckSaveAll)
	ON_WM_DROPFILES()
END_MESSAGE_MAP()

int testOpenExcel() {
	CApplication ExcelApp;
	CWorkbooks books;
	CWorkbook book;
	CWorksheets sheets;
	CWorksheet sheet;
	LPDISPATCH lpDisp = NULL;

	//����Excel ������(����Excel)
	if (!ExcelApp.CreateDispatch(_T("Excel.Application"), NULL))
	{
		AfxMessageBox(_T("����Excel������ʧ��!"));
		return -1;
	}

	/*�жϵ�ǰExcel�İ汾*/
	CString strExcelVersion = ExcelApp.get_Version();
	int iStart = 0;
	strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
	if (_T("11") == strExcelVersion)
	{
		AfxMessageBox(_T("��ǰExcel�İ汾��2003��"));
	}
	else if (_T("12") == strExcelVersion)
	{
		AfxMessageBox(_T("��ǰExcel�İ汾��2007��"));
	}
	else
	{
		AfxMessageBox(_T("��ǰExcel�İ汾" + strExcelVersion + "�汾��"));
	}

	ExcelApp.put_Visible(TRUE);
	ExcelApp.put_UserControl(FALSE);

	/*�õ�����������*/
	books.AttachDispatch(ExcelApp.get_Workbooks());

	/*��һ�����������粻���ڣ�������һ��������*/
	CString strBookPath = _T("C:\\12345.xlsx");
	VARIANT UpdateLinks; VARIANT ReadOnly; VARIANT Format; VARIANT Password; VARIANT WriteResPassword; VARIANT IgnoreReadOnlyRecommended; VARIANT Origin; VARIANT Delimiter; VARIANT Editable; VARIANT Notify; VARIANT Converter; VARIANT AddToMru;

	try
	{
		//lpDisp = books._Open(strBookPath, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru);
		/*��һ��������*/
		lpDisp = books.Open(strBookPath,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing);
		book.AttachDispatch(lpDisp);
	}
	catch (CMemoryException* e)
	{
		AfxMessageBox(_T("�ڴ治����"));
	}
	catch (CFileException* e)
	{
		AfxMessageBox(_T("�ļ������ڡ�"));
	}
	catch (CException* e)
	{
		AfxMessageBox(_T("�����쳣��"));
	}

	/*�õ��������е�Sheet������*/
	sheets.AttachDispatch(book.get_Sheets());

	long sheetsCount = sheets.get_Count();
	//VARIANT indexItem; indexItem.intVal = 0;
	//LPDISPATCH result= sheets.get_Item(indexItem);
	return 0;
}


struct st_excle_data_rule
{
	int nRow =0;//�����
	int nCol = 0;//�����

	int type=0; // 1, �������ݣ� 2�������ṹ

	//0,
	//int DATA_REF_ROW = 0; // �ֶ�����Ӧ����
	int VALUE_TYPE_ROW = 0; // ֵ���Ͷ�Ӧ�У���ʱδ�ã�������֤ʱ��
	//1,
	int FIELD_REF_COLUMN = 0; //�ֶ����ƶ�Ӧ����

	//0,1����
	vector<int> ignoreCol;//���Ե�����  ��������xml�ļ����У�
	vector<int> ignoreRow;//���Ե�����  ��������xml�ļ����У�
	int	DEFAULE_VALUE_ROW = 0;//Ĭ��ֵ��Ӧ����

	int FIELD_REF_ROW = 0; // �ṹ���Զ�Ӧ����
	vector<string> fieldColumnVct;// �ֶ���
	vector<string> valueTypVct;// ���ݸ�ʽ
	vector<string> defaultValueVct;//Ĭ��ֵ
};



CString AnalysisFirstRow(IllusionExcelFile& excl, const CString& sheetName, st_excle_data_rule& rule)
{
	CString errorMsg;
	if (excl.GetRowCount() <= 0) return errorMsg;

	const int nRow = excl.GetRowCount();
	const int nCol = excl.GetColumnCount();
	rule.nRow = nRow;
	rule.nCol = nCol;
	/*	������һ������ */
	for (int i = 1; i <= nRow; ++i)
	{
		CString strValue = excl.GetCellString(i, 1).Trim();
		if (strValue.Left(1) == _T('#'))
		{
			rule.ignoreRow.push_back(i);
			CString tagVal = strValue.Right(strValue.GetLength() - 1);
			if (_T("data_row") == tagVal)
			{
				rule.FIELD_REF_ROW = i;
				rule.type = 1;
			}
			else if (_T("struct_row") == tagVal)
			{
				rule.FIELD_REF_ROW = i;
				rule.type = 2;
			}
			else if (_T("value_type") == tagVal)
			{
				rule.VALUE_TYPE_ROW = i;
			}
			else if (_T("default_value") == tagVal)
			{
				rule.DEFAULE_VALUE_ROW = i;
			}
		}
	}
	if (rule.type == 0)
	{
		errorMsg.Append(_T("��A���в����ڱ�ʶ��#data_rowҲ������#struct_row"));
		return errorMsg;
	}

	
	return errorMsg;
}


void CacheCommonInfo(IllusionExcelFile& excl, st_excle_data_rule& rule)
{
	const int nRow = rule.nRow;
	const int nCol = rule.nCol;
	// ���� columnName�� valueType ��
	rule.fieldColumnVct.push_back(""); rule.fieldColumnVct.push_back("");//�������հ�ֵ
	for (int j = 2; j <= nCol; ++j)
	{
		int field = rule.FIELD_REF_ROW;
		CString strValue = excl.GetCellString(field, j).Trim();
		if (strValue.IsEmpty())
		{
			rule.ignoreCol.push_back(j);
		}
		else if (strValue.Left(1) == _T("#"))
		{
			rule.ignoreCol.push_back(j);
		}
		else if (rule.type == 2 && strValue.Left(2) == _T("$_"))
		{
			rule.FIELD_REF_COLUMN = j;
			rule.ignoreCol.push_back(j);
		}
		rule.fieldColumnVct.push_back(UnicodeToUTF8(strValue.GetString()));
	}
	if (rule.VALUE_TYPE_ROW != 0)
	{
		rule.valueTypVct.push_back(""); rule.valueTypVct.push_back("");//�������հ�ֵ
		for (int j = 2; j <= nCol; ++j)
		{
			CString strValue = excl.GetCellString(rule.VALUE_TYPE_ROW, j).Trim();
			rule.valueTypVct.push_back(UnicodeToUTF8(strValue.GetString()));
		}
	}
	if (rule.DEFAULE_VALUE_ROW != 0)
	{
		rule.defaultValueVct.push_back(""); rule.defaultValueVct.push_back("");//�������հ�ֵ
		for (int j = 2; j <= nCol; ++j)
		{
			CString strValue = excl.GetCellString(rule.DEFAULE_VALUE_ROW, j).Trim();
			rule.defaultValueVct.push_back(UnicodeToUTF8(strValue.GetString()));
		}
	}

}

CString ExportToDataXml(IllusionExcelFile& excl, const CString& sheetName, std::string filePathName, st_excle_data_rule& rule)
{
	CString errorMsg;
	const int nRow = rule.nRow;
	const int nCol = rule.nCol;

	tinyxml2::XMLDocument pDoc;
	//tinyxml2::XMLPrinter pter;
	//pDoc.Accept(&pter);
	//doc.SaveFile();

	XMLDeclaration *pDel = pDoc.NewDeclaration("xml version=\"1.0\" encoding=\"UTF-8\"");
	pDoc.LinkEndChild(pDel);

	XMLElement *rootElement = pDoc.NewElement("xml");
	pDoc.LinkEndChild(rootElement);

	const string sheetNameStr = UnicodeToUTF8(sheetName.GetString());

	vector<int>::iterator it;
	vector<int>::iterator itBegin = rule.ignoreRow.begin();
	vector<int>::iterator itEnd = rule.ignoreRow.end();
	vector<int>::iterator itCol;
	vector<int>::iterator itColBegin = rule.ignoreCol.begin();
	vector<int>::iterator itColEnd = rule.ignoreCol.end();
	for (int i = 1; i <= nRow; ++i)
	{
		it = find(itBegin, itEnd, i);
		if (it != itEnd) { //�ҵ���
			continue;
		}
		//���ж�Ϊ���ж�
		bool allEmpty = true;
		for (int j = 2; j <= nCol; ++j)
		{
			CString strValue = excl.GetCellString(i, j);
			if (!strValue.IsEmpty())
			{
				allEmpty = false;
				break;
			}
		}
		if (allEmpty) {
			continue;
		}

		XMLElement* pRowData = pDoc.NewElement(sheetNameStr.c_str());
		rootElement->LinkEndChild(pRowData);

		for (int j = 2; j <= nCol; ++j)
		{
			//����Ϊ�յ�,��#��ͷ ���������ֵ
			itCol = find(itColBegin, itColEnd, j);
			if (itCol != itColEnd) { //�ҵ���
				continue;
			}
			CString strValue = excl.GetCellString(i, j).Trim();
			string tagName = rule.fieldColumnVct[j];
			string tagValue = strValue.IsEmpty() ? (rule.DEFAULE_VALUE_ROW == 0 ? "" : rule.defaultValueVct[j]) : UnicodeToUTF8(strValue.GetString());
			XMLElement* pColumn = pDoc.NewElement(tagName.c_str());
			pRowData->LinkEndChild(pColumn);
			pColumn->SetText(tagValue.c_str());
			if (rule.VALUE_TYPE_ROW)
			{
				string valueType = rule.valueTypVct[j];
				//pColumn->SetAttribute("type", valueType.c_str()); // �������ͣ��ṹ���ж��壬�˴�ʡ��
			}
		}
	}
	pDoc.SetBOM(TRUE);
	pDoc.SaveFile(filePathName.c_str(), FALSE); // �ڶ�������ΪTRUE�Ļ���ʾȥ���ո�ѹ���洢��
	return errorMsg;
}

#define DEFAULT_ROOT_NAME  "object"
#define FIRST_NODE_NAME "properties"
CString ExportToStructXml(IllusionExcelFile& excl, const CString& sheetName, std::string filePathName, st_excle_data_rule& rule)
{
	CString errorMsg;
	const int nRow = rule.nRow;
	const int nCol = rule.nCol;

	//vector<string> fieldVct;
	//fieldVct.push_back("");//����0λ


	tinyxml2::XMLDocument pDoc;
	//tinyxml2::XMLPrinter pter;
	//pDoc.Accept(&pter);
	//doc.SaveFile();

	XMLDeclaration *pDel = pDoc.NewDeclaration("xml version=\"1.0\" encoding=\"UTF-8\"");
	pDoc.LinkEndChild(pDel);

	XMLElement *objectElement = pDoc.NewElement(DEFAULT_ROOT_NAME);
	pDoc.LinkEndChild(objectElement);

	XMLElement* pProperties = pDoc.NewElement(FIRST_NODE_NAME);
	objectElement->LinkEndChild(pProperties);

	vector<int>::iterator itEnd = rule.ignoreRow.end();
	vector<int>::iterator itColEnd = rule.ignoreCol.end();

	for (int i = 2; i <= nRow; ++i)
	{

		CString fieldName = excl.GetCellString(i, rule.FIELD_REF_COLUMN).Trim();
		vector<int>::iterator itBegin = rule.ignoreRow.begin();
		vector<int>::iterator it = find(itBegin, itEnd, i);
		if (it != itEnd) { //�ҵ���
			continue;
		}
		//���ж�Ϊ���ж�
		bool allEmpty = true;
		for (int j = 1; j <= nCol; ++j)
		{
			if (rule.FIELD_REF_ROW == j) 
			{
				continue;
			}
			CString strValue = excl.GetCellString(i, j);
			if (!strValue.IsEmpty())
			{
				allEmpty = false;
				break;
			}
		}
		if (allEmpty) {
			continue;
		}


		XMLElement* pProperty = pDoc.NewElement(UnicodeToUTF8(fieldName.GetString()).c_str());
		pProperties->LinkEndChild(pProperty);
		for (int j = 2; j <= nCol; ++j)
		{	
			vector<int>::iterator itColBegin = rule.ignoreCol.begin();
			//����Ϊ�յ�,��#��ͷ ���������ֵ
			vector<int>::iterator itCol = find(itColBegin, itColEnd, j);
			if (itCol != itColEnd) { //�ҵ���
				continue;
			}
			CString strValue = excl.GetCellString(i, j).Trim();
			pProperty->SetAttribute(rule.fieldColumnVct[j].c_str(), UnicodeToUTF8(strValue.GetString()).c_str());
		}
	}
	pDoc.SetBOM(TRUE);
	pDoc.SaveFile(filePathName.c_str(), FALSE); // �ڶ�������ΪTRUE�Ļ���ʾȥ���ո�ѹ���洢��
	return errorMsg;
}

// CMFCApplication1Dlg ��Ϣ�������

BOOL CMFCApplication1Dlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// ��������...���˵�����ӵ�ϵͳ�˵��С�

	// IDM_ABOUTBOX ������ϵͳ���Χ�ڡ�
	ASSERT((IDM_ABOUTBOX & 0xFFF0) == IDM_ABOUTBOX);
	ASSERT(IDM_ABOUTBOX < 0xF000);

	CMenu* pSysMenu = GetSystemMenu(FALSE);
	if (pSysMenu != NULL)
	{
		BOOL bNameValid;
		CString strAboutMenu;
		bNameValid = strAboutMenu.LoadString(IDS_ABOUTBOX);
		ASSERT(bNameValid);
		if (!strAboutMenu.IsEmpty())
		{
			pSysMenu->AppendMenu(MF_SEPARATOR);
			pSysMenu->AppendMenu(MF_STRING, IDM_ABOUTBOX, strAboutMenu);
		}
	}

	// ���ô˶Ի����ͼ�ꡣ  ��Ӧ�ó��������ڲ��ǶԻ���ʱ����ܽ��Զ�
	//  ִ�д˲���
	SetIcon(m_hIcon, TRUE);			// ���ô�ͼ��
	SetIcon(m_hIcon, FALSE);		// ����Сͼ��

	// TODO: �ڴ���Ӷ���ĳ�ʼ������
	m_excel_util = new IllusionExcelFile();
	if (!m_excel_util->InitExcel()) {
		AfxMessageBox(_T("����Excel������ʧ�ܣ����Ȱ�װoffice excel��"));
	}

	//IllusionExcelFile excl;
	//bool bInit = excl.InitExcel();
	////bool bRet = excl.OpenExcelFile("c:\\task.xlsx");
	//CString filePath = L"c:\\task.xlsx";
	//bool bRet = excl.OpenExcelFile(filePath);
	//CString strSheetName = excl.GetSheetName(1);
	//bool bLoad = excl.LoadSheet(strSheetName, TRUE);
	//int nRow = excl.GetRowCount();
	//int nCol = excl.GetColumnCount();

	///*for (int i = 1; i <= nRow; ++i)
	//{
	//	for (int j = 1; j <= nCol; ++j)
	//	{
	//		CString strValue = excl.GetCellString(i, j);
	//	}
	//}*/
	//

	//std::ofstream ous;// (FILENAME);
	//ous.open("c:\\112345.xml");
	//char szBOM[3] = { (char)0xEF, (char)0xBB, (char)0xBF };
	//ous.write(szBOM, 3); //����ҪҲ���Բ�Ҫ�� ����ļ��ĸ�ʽ��utf8
	//SaveToXml(excl, ous);
	//ous.close();
	//excl.CloseExcelFile();
	m_excel_util->excel_application_.put_Visible(FALSE);//Ӱ�ص�excel����
	return TRUE;  // ���ǽ��������õ��ؼ������򷵻� TRUE
}


void CMFCApplication1Dlg::OnSysCommand(UINT nID, LPARAM lParam)
{
	if ((nID & 0xFFF0) == IDM_ABOUTBOX)
	{
		CAboutDlg dlgAbout;
		dlgAbout.DoModal();
	}
	else
	{
		CDialogEx::OnSysCommand(nID, lParam);
	}
}

// �����Ի��������С����ť������Ҫ����Ĵ���
//  �����Ƹ�ͼ�ꡣ  ����ʹ���ĵ�/��ͼģ�͵� MFC Ӧ�ó���
//  �⽫�ɿ���Զ���ɡ�

void CMFCApplication1Dlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // ���ڻ��Ƶ��豸������

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// ʹͼ���ڹ����������о���
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// ����ͼ��
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//���û��϶���С������ʱϵͳ���ô˺���ȡ�ù��
//��ʾ��
HCURSOR CMFCApplication1Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// ���û��ر� UI ʱ������������Ա���������ĳ��
//  �������Զ�����������Ӧ�˳���  ��Щ
//  ��Ϣ�������ȷ����������: �����������ʹ�ã�
//  ������ UI�������ڹرնԻ���ʱ��
//  �Ի�����Ȼ�ᱣ�������

void CMFCApplication1Dlg::OnClose()
{
	if (CanExit())
		CDialogEx::OnClose();
}

void CMFCApplication1Dlg::OnOK()
{
	if (CanExit())
		CDialogEx::OnOK();
}

void CMFCApplication1Dlg::OnCancel()
{
	if (CanExit())
		CDialogEx::OnCancel();
}

BOOL CMFCApplication1Dlg::CanExit()
{
	// �����������Ա�����������Զ���
	//  �������Իᱣ�ִ�Ӧ�ó���
	//  ʹ�Ի���������������� UI ����������
	if (m_pAutoProxy != NULL)
	{
		ShowWindow(SW_HIDE);
		return FALSE;
	}

	return TRUE;
}





void CMFCApplication1Dlg::OnBnClickedExec()
{
	CString filePathName;
	GetDlgItemText(IDC_BROWSE_EXCEL, filePathName);
	if (filePathName.Right(4) != _T("xlsx") && filePathName.Right(3) != _T("xls"))
	{
		AfxMessageBox(_T("�ļ���ʽ���ԣ���ѡ��.xlsx��.xls��excel�ļ���"));
		return;
	}
	int checed = m_check_save_all.GetCheck();
	bool bRet = m_excel_util->OpenExcelFile(filePathName);

	if (checed == BST_CHECKED)
	{
		int sheetCount = m_excel_util->GetSheetCount();
		vector<CString> sheetNames;
		for (int i = 1; i <= sheetCount; i++)
		{
			CString sheetName = m_excel_util->GetSheetName(i);
			DoExport(sheetName);
		}
	}
	else
	{
		CString sheetName;
		GetDlgItemText(IDC_COMBO_SHEET, sheetName);
		DoExport(sheetName);
	}
}


void CMFCApplication1Dlg::OnEnChangeBrowseExcel()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ���������
	// ���ʹ�֪ͨ��������д CDialogEx::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
	CString filePathName;
	GetDlgItemText(IDC_BROWSE_EXCEL, filePathName);
	if (filePathName.Right(4) != _T("xlsx") && filePathName.Right(3) != _T("xls"))
	{
		AfxMessageBox(_T("�ļ���ʽ���ԣ���ѡ��.xlsx��.xls��excel�ļ���"));
	}
	bool bOpen = m_excel_util->OpenExcelFile(filePathName);
	if (!bOpen) {
		AfxMessageBox(_T("�ļ��𻵣���excle��֧�ָð汾��"));
		return;
	}
	int sheetCount = m_excel_util->GetSheetCount();
	vector<CString> sheetNames;
	for (int i= 1;i <= sheetCount;i++)
	{
		CString sheetName = m_excel_util->GetSheetName(i);
		sheetNames.push_back(sheetName);
		
	}
	m_boxSelect_sheet.ResetContent();//���ԭ������Ŀ
	for each (CString sheetName in sheetNames)
	{
		m_boxSelect_sheet.AddString(sheetName);
	}
	m_boxSelect_sheet.SetCurSel(0);//���õ�һ��sheet��Ĭ��ѡ��
	AutoFillOutXmlPathName();
}


void CMFCApplication1Dlg::OnCbnSelchangeComboSheet()
{
	AutoFillOutXmlPathName();
}

void CMFCApplication1Dlg::AppendLineToOutText(const CString& msg, OutTextFont font)
{

	m_rich_edit_out_text.SetSel(-1, -1);
	m_rich_edit_out_text.ReplaceSel(msg);
	if (OutTextFont::DEFOUT != font)
	{
		CHARFORMAT cf;
		memset(&cf, '\0', sizeof(CHARFORMAT));
		m_rich_edit_out_text.GetSelectionCharFormat(cf);
		cf.dwMask = CFM_BOLD | CFM_COLOR | CFM_FACE | CFM_ITALIC | CFM_SIZE | CFM_UNDERLINE;
		cf.dwEffects = 0;

		switch (font)
		{
		case OutTextFont::C_RED:
			cf.crTextColor = C_RED_RGB;
			break;
		case OutTextFont::C_GREEN:
			cf.crTextColor = C_GREEN_RGB;
			break;
		case OutTextFont::C_BLUE:
			cf.crTextColor = C_BLUE_RGB;
			break;
		default:
			break;
		}
		long nStart, nEnd;
		m_rich_edit_out_text.GetSel(nStart, nEnd);
		m_rich_edit_out_text.SetSel(nEnd - msg.GetLength(), nEnd); //���ô�������
		m_rich_edit_out_text.SetSelectionCharFormat(cf);
		m_rich_edit_out_text.SetFocus();

	}
	m_rich_edit_out_text.SetSel(-1, -1);
	m_rich_edit_out_text.ReplaceSel(_T("\n"));
}


void CMFCApplication1Dlg::OnBnClickedCheckSaveAll()
{
	int checed = m_check_save_all.GetCheck();
	if (checed == BST_CHECKED)
	{
		m_browse_out_xml.EnableFolderBrowseButton();
		SetDlgItemTextW(ID_EXEC, _T("����ȫ��xml"));
	}
	else
	{
		m_browse_out_xml.EnableFileBrowseButton();
		SetDlgItemTextW(ID_EXEC, _T("��������xml"));
	}
	AutoFillOutXmlPathName();
}

void CMFCApplication1Dlg::AutoFillOutXmlPathName()
{
	CString filePathName;
	GetDlgItemText(IDC_BROWSE_EXCEL, filePathName);
	if (filePathName.IsEmpty()) {
		return;
	}
	int checed = m_check_save_all.GetCheck();

	int splitIndex = filePathName.ReverseFind(_T('\\'));
	CString filePath = filePathName.Left(splitIndex + 1);
	if (checed == BST_CHECKED)
	{
		m_browse_out_xml.SetWindowTextW(filePath);
	}
	else
	{
		CString sheetName;
		// GetDlgItemText(IDC_COMBO_SHEET, sheetName); //���ȡ��ֵ�����Ӻ�bug
		int i = m_boxSelect_sheet.GetCurSel();
		m_boxSelect_sheet.GetLBText(i, sheetName);
		int splitIndex = filePathName.ReverseFind(_T('\\'));
		CString defoutOutXml = filePath + sheetName + ".xml";
		m_browse_out_xml.SetWindowTextW(defoutOutXml);
	}



}

void CMFCApplication1Dlg::DoExport(const CString& sheetName)
{
	bool bLoad = m_excel_util->LoadSheet(sheetName, TRUE);

	CString outXmlPahtName;
	GetDlgItemText(IDC_BROWSE_OUT_XML, outXmlPahtName);
	int checed = m_check_save_all.GetCheck();
	if (checed == BST_CHECKED)
	{
		outXmlPahtName.Append(_T("\\"));
		outXmlPahtName.Append(sheetName);
		outXmlPahtName.Append(_T(".xml"));
	}
	//std::ofstream ous;// (FILENAME);
	//ous.open(outXmlPahtName);
	AppendLineToOutText(_T("��ʼ����..."));
	//char szBOM[3] = { (char)0xEF, (char)0xBB, (char)0xBF };
	//ous.write(szBOM, 3); //����ҪҲ���Բ�Ҫ�� ����ļ��ĸ�ʽ��utf8
	//SaveToXmlDocument(*m_excel_util, UnicodeToANSI(outXmlPahtName.GetString()));
	//CString errorMsg = SaveToXml(*m_excel_util, UnicodeToANSI(outXmlPahtName.GetString()), sheetName);
	st_excle_data_rule rule;
	CString errorMsg  = AnalysisFirstRow(*m_excel_util, sheetName, rule);
	if (!errorMsg.IsEmpty()) goto printInfoMsg;
	CacheCommonInfo(*m_excel_util, rule);
	if (rule.type==2 && rule.FIELD_REF_COLUMN == 0)
	{
		errorMsg.Format(_T("��%d"), rule.FIELD_REF_ROW);
		errorMsg.Append(_T("�в������ֶ�����ʶ����$_��") );
		goto printInfoMsg;
	}
	if (rule.type == 1)
	{
		errorMsg = ExportToDataXml(*m_excel_util, sheetName, UnicodeToANSI(outXmlPahtName.GetString()), rule);
	}
	else if (rule.type == 2)
	{
		errorMsg = ExportToStructXml(*m_excel_util, sheetName, UnicodeToANSI(outXmlPahtName.GetString()), rule);
	}
	printInfoMsg:
	if (!errorMsg.IsEmpty())
	{
		errorMsg.Insert(0, _T("��ǩҳ")+sheetName + _T(":"));
		AppendLineToOutText(errorMsg, OutTextFont::C_RED);
	}
	else
	{
		CString successMsg;
		successMsg.Append(_T("�ɹ������ļ�"));
		successMsg.Append(sheetName + _T(".xml"));
		successMsg.Append(_T("��"));
		successMsg.Append(outXmlPahtName);
		AppendLineToOutText(successMsg, OutTextFont::C_GREEN);
	}

}


void CMFCApplication1Dlg::OnDropFiles(HDROP hDropInfo)
{
	// TODO: �ڴ������Ϣ�����������/�����Ĭ��ֵ
	wchar_t szFilePathName[_MAX_PATH + 1] = { 0 };
	CString filePathName;
	//�õ��ļ�����      
	UINT nNumOfFiles = DragQueryFile(hDropInfo, 0xFFFFFFFF, NULL, 0);

	for (UINT nIndex = 0; nIndex < nNumOfFiles; ++nIndex) {
		//���õ��ļ���   
		DragQueryFile(hDropInfo, nIndex, (LPTSTR)szFilePathName, _MAX_PATH);
		// �����ļ����Ϳ������������ˡ�:P   
		//AfxMessageBox((LPCTSTR)szFilePathName);
		CString filePathName(szFilePathName);
		if (filePathName.Right(5) == _T(".xlsx") || filePathName.Right(4) == _T(".xls")) {
			//AfxMessageBox(_T("��Ч�ļ���") + filePahtName);
			//SetDlgItemTextW(IDC_BROWSE_EXCEL, filePathName);
			m_browse_excel.SetWindowTextW(filePathName);
			break;
		}
	}

	DragFinish(hDropInfo);
	CDialogEx::OnDropFiles(hDropInfo);
}
