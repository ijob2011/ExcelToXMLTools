
// MFCApplication1Dlg.cpp : 实现文件
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

// 用于应用程序“关于”菜单项的 CAboutDlg 对话框
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

//字符串分割函数
std::vector<std::string> split(std::string str, std::string pattern)
{
	std::string::size_type pos;
	std::vector<std::string> result;
	str += pattern;//扩展字符串以方便操作
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

// 对话框数据
#ifdef AFX_DESIGN_TIME
	enum { IDD = IDD_ABOUTBOX };
#endif

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
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


// CMFCApplication1Dlg 对话框


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
	// 如果该对话框有自动化代理，则
	//  将此代理指向该对话框的后向指针设置为 NULL，以便
	//  此代理知道该对话框已被删除。
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

	//创建Excel 服务器(启动Excel)
	if (!ExcelApp.CreateDispatch(_T("Excel.Application"), NULL))
	{
		AfxMessageBox(_T("启动Excel服务器失败!"));
		return -1;
	}

	/*判断当前Excel的版本*/
	CString strExcelVersion = ExcelApp.get_Version();
	int iStart = 0;
	strExcelVersion = strExcelVersion.Tokenize(_T("."), iStart);
	if (_T("11") == strExcelVersion)
	{
		AfxMessageBox(_T("当前Excel的版本是2003。"));
	}
	else if (_T("12") == strExcelVersion)
	{
		AfxMessageBox(_T("当前Excel的版本是2007。"));
	}
	else
	{
		AfxMessageBox(_T("当前Excel的版本" + strExcelVersion + "版本。"));
	}

	ExcelApp.put_Visible(TRUE);
	ExcelApp.put_UserControl(FALSE);

	/*得到工作簿容器*/
	books.AttachDispatch(ExcelApp.get_Workbooks());

	/*打开一个工作簿，如不存在，则新增一个工作簿*/
	CString strBookPath = _T("C:\\12345.xlsx");
	VARIANT UpdateLinks; VARIANT ReadOnly; VARIANT Format; VARIANT Password; VARIANT WriteResPassword; VARIANT IgnoreReadOnlyRecommended; VARIANT Origin; VARIANT Delimiter; VARIANT Editable; VARIANT Notify; VARIANT Converter; VARIANT AddToMru;

	try
	{
		//lpDisp = books._Open(strBookPath, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru);
		/*打开一个工作簿*/
		lpDisp = books.Open(strBookPath,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing, vtMissing,
			vtMissing, vtMissing, vtMissing, vtMissing);
		book.AttachDispatch(lpDisp);
	}
	catch (CMemoryException* e)
	{
		AfxMessageBox(_T("内存不够。"));
	}
	catch (CFileException* e)
	{
		AfxMessageBox(_T("文件不存在。"));
	}
	catch (CException* e)
	{
		AfxMessageBox(_T("程序异常。"));
	}

	/*得到工作簿中的Sheet的容器*/
	sheets.AttachDispatch(book.get_Sheets());

	long sheetsCount = sheets.get_Count();
	//VARIANT indexItem; indexItem.intVal = 0;
	//LPDISPATCH result= sheets.get_Item(indexItem);
	return 0;
}


struct st_excle_data_rule
{
	int nRow =0;//最大行
	int nCol = 0;//最大列

	int type=0; // 1, 导出数据， 2，导出结构

	//0,
	//int DATA_REF_ROW = 0; // 字段名对应的行
	int VALUE_TYPE_ROW = 0; // 值类型对应行，暂时未用，数据验证时用
	//1,
	int FIELD_REF_COLUMN = 0; //字段名称对应的列

	//0,1共用
	vector<int> ignoreCol;//忽略掉得列  （不导入xml文件的列）
	vector<int> ignoreRow;//忽略掉得行  （不导入xml文件的行）
	int	DEFAULE_VALUE_ROW = 0;//默认值对应的行

	int FIELD_REF_ROW = 0; // 结构属性对应的行
	vector<string> fieldColumnVct;// 字段名
	vector<string> valueTypVct;// 数据格式
	vector<string> defaultValueVct;//默认值
};



CString AnalysisFirstRow(IllusionExcelFile& excl, const CString& sheetName, st_excle_data_rule& rule)
{
	CString errorMsg;
	if (excl.GetRowCount() <= 0) return errorMsg;

	const int nRow = excl.GetRowCount();
	const int nCol = excl.GetColumnCount();
	rule.nRow = nRow;
	rule.nCol = nCol;
	/*	分析第一列数据 */
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
		errorMsg.Append(_T("第A列中不存在标识符#data_row也不存在#struct_row"));
		return errorMsg;
	}

	
	return errorMsg;
}


void CacheCommonInfo(IllusionExcelFile& excl, st_excle_data_rule& rule)
{
	const int nRow = rule.nRow;
	const int nCol = rule.nCol;
	// 缓存 columnName， valueType 等
	rule.fieldColumnVct.push_back(""); rule.fieldColumnVct.push_back("");//放两个空白值
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
		rule.valueTypVct.push_back(""); rule.valueTypVct.push_back("");//放两个空白值
		for (int j = 2; j <= nCol; ++j)
		{
			CString strValue = excl.GetCellString(rule.VALUE_TYPE_ROW, j).Trim();
			rule.valueTypVct.push_back(UnicodeToUTF8(strValue.GetString()));
		}
	}
	if (rule.DEFAULE_VALUE_ROW != 0)
	{
		rule.defaultValueVct.push_back(""); rule.defaultValueVct.push_back("");//放两个空白值
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
		if (it != itEnd) { //找到了
			continue;
		}
		//整行都为空判断
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
			//列名为空的,已#开头 忽略这个数值
			itCol = find(itColBegin, itColEnd, j);
			if (itCol != itColEnd) { //找到了
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
				//pColumn->SetAttribute("type", valueType.c_str()); // 数据类型，结构中有定义，此处省略
			}
		}
	}
	pDoc.SetBOM(TRUE);
	pDoc.SaveFile(filePathName.c_str(), FALSE); // 第二个参数为TRUE的话表示去掉空格。压缩存储。
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
	//fieldVct.push_back("");//补充0位


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
		if (it != itEnd) { //找到了
			continue;
		}
		//整行都为空判断
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
			//列名为空的,已#开头 忽略这个数值
			vector<int>::iterator itCol = find(itColBegin, itColEnd, j);
			if (itCol != itColEnd) { //找到了
				continue;
			}
			CString strValue = excl.GetCellString(i, j).Trim();
			pProperty->SetAttribute(rule.fieldColumnVct[j].c_str(), UnicodeToUTF8(strValue.GetString()).c_str());
		}
	}
	pDoc.SetBOM(TRUE);
	pDoc.SaveFile(filePathName.c_str(), FALSE); // 第二个参数为TRUE的话表示去掉空格。压缩存储。
	return errorMsg;
}

// CMFCApplication1Dlg 消息处理程序

BOOL CMFCApplication1Dlg::OnInitDialog()
{
	CDialogEx::OnInitDialog();

	// 将“关于...”菜单项添加到系统菜单中。

	// IDM_ABOUTBOX 必须在系统命令范围内。
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

	// 设置此对话框的图标。  当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	m_excel_util = new IllusionExcelFile();
	if (!m_excel_util->InitExcel()) {
		AfxMessageBox(_T("启动Excel服务器失败，请先安装office excel！"));
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
	//ous.write(szBOM, 3); //可以要也可以不要， 标记文件的格式是utf8
	//SaveToXml(excl, ous);
	//ous.close();
	//excl.CloseExcelFile();
	m_excel_util->excel_application_.put_Visible(FALSE);//影藏掉excel窗口
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
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

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。  对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CMFCApplication1Dlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作区矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialogEx::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标
//显示。
HCURSOR CMFCApplication1Dlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

// 当用户关闭 UI 时，如果控制器仍保持着它的某个
//  对象，则自动化服务器不应退出。  这些
//  消息处理程序确保如下情形: 如果代理仍在使用，
//  则将隐藏 UI；但是在关闭对话框时，
//  对话框仍然会保留在那里。

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
	// 如果代理对象仍保留在那里，则自动化
	//  控制器仍会保持此应用程序。
	//  使对话框保留在那里，但将其 UI 隐藏起来。
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
		AfxMessageBox(_T("文件格式不对，请选择.xlsx或.xls的excel文件！"));
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
	// TODO:  如果该控件是 RICHEDIT 控件，它将不
	// 发送此通知，除非重写 CDialogEx::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
	CString filePathName;
	GetDlgItemText(IDC_BROWSE_EXCEL, filePathName);
	if (filePathName.Right(4) != _T("xlsx") && filePathName.Right(3) != _T("xls"))
	{
		AfxMessageBox(_T("文件格式不对，请选择.xlsx或.xls的excel文件！"));
	}
	bool bOpen = m_excel_util->OpenExcelFile(filePathName);
	if (!bOpen) {
		AfxMessageBox(_T("文件损坏，或excle不支持该版本！"));
		return;
	}
	int sheetCount = m_excel_util->GetSheetCount();
	vector<CString> sheetNames;
	for (int i= 1;i <= sheetCount;i++)
	{
		CString sheetName = m_excel_util->GetSheetName(i);
		sheetNames.push_back(sheetName);
		
	}
	m_boxSelect_sheet.ResetContent();//清空原来的项目
	for each (CString sheetName in sheetNames)
	{
		m_boxSelect_sheet.AddString(sheetName);
	}
	m_boxSelect_sheet.SetCurSel(0);//设置第一个sheet被默认选中
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
		m_rich_edit_out_text.SetSel(nEnd - msg.GetLength(), nEnd); //设置处理区域
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
		SetDlgItemTextW(ID_EXEC, _T("导出全部xml"));
	}
	else
	{
		m_browse_out_xml.EnableFileBrowseButton();
		SetDlgItemTextW(ID_EXEC, _T("导出单个xml"));
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
		// GetDlgItemText(IDC_COMBO_SHEET, sheetName); //这个取得值，有延后bug
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
	AppendLineToOutText(_T("开始导出..."));
	//char szBOM[3] = { (char)0xEF, (char)0xBB, (char)0xBF };
	//ous.write(szBOM, 3); //可以要也可以不要， 标记文件的格式是utf8
	//SaveToXmlDocument(*m_excel_util, UnicodeToANSI(outXmlPahtName.GetString()));
	//CString errorMsg = SaveToXml(*m_excel_util, UnicodeToANSI(outXmlPahtName.GetString()), sheetName);
	st_excle_data_rule rule;
	CString errorMsg  = AnalysisFirstRow(*m_excel_util, sheetName, rule);
	if (!errorMsg.IsEmpty()) goto printInfoMsg;
	CacheCommonInfo(*m_excel_util, rule);
	if (rule.type==2 && rule.FIELD_REF_COLUMN == 0)
	{
		errorMsg.Format(_T("第%d"), rule.FIELD_REF_ROW);
		errorMsg.Append(_T("行不存在字段名标识符“$_”") );
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
		errorMsg.Insert(0, _T("标签页")+sheetName + _T(":"));
		AppendLineToOutText(errorMsg, OutTextFont::C_RED);
	}
	else
	{
		CString successMsg;
		successMsg.Append(_T("成功导出文件"));
		successMsg.Append(sheetName + _T(".xml"));
		successMsg.Append(_T("到"));
		successMsg.Append(outXmlPahtName);
		AppendLineToOutText(successMsg, OutTextFont::C_GREEN);
	}

}


void CMFCApplication1Dlg::OnDropFiles(HDROP hDropInfo)
{
	// TODO: 在此添加消息处理程序代码和/或调用默认值
	wchar_t szFilePathName[_MAX_PATH + 1] = { 0 };
	CString filePathName;
	//得到文件个数      
	UINT nNumOfFiles = DragQueryFile(hDropInfo, 0xFFFFFFFF, NULL, 0);

	for (UINT nIndex = 0; nIndex < nNumOfFiles; ++nIndex) {
		//　得到文件名   
		DragQueryFile(hDropInfo, nIndex, (LPTSTR)szFilePathName, _MAX_PATH);
		// 有了文件名就可以想干嘛干嘛了　:P   
		//AfxMessageBox((LPCTSTR)szFilePathName);
		CString filePathName(szFilePathName);
		if (filePathName.Right(5) == _T(".xlsx") || filePathName.Right(4) == _T(".xls")) {
			//AfxMessageBox(_T("有效文件：") + filePahtName);
			//SetDlgItemTextW(IDC_BROWSE_EXCEL, filePathName);
			m_browse_excel.SetWindowTextW(filePathName);
			break;
		}
	}

	DragFinish(hDropInfo);
	CDialogEx::OnDropFiles(hDropInfo);
}
