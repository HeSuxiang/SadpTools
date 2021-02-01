
// SadpToolsDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "SadpTools.h"
#include "SadpToolsDlg.h"
#include "afxdialogex.h"

//海康威视库
#include "Sadp.h"
#pragma commment(lib,"Sadp.lib")

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// 用于应用程序“关于”菜单项的 CAboutDlg 对话框

class CAboutDlg : public CDialogEx
{
public:
	CAboutDlg();

// 对话框数据
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV 支持

// 实现
protected:
	DECLARE_MESSAGE_MAP()
};

CAboutDlg::CAboutDlg() : CDialogEx(CAboutDlg::IDD)
{
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
}

BEGIN_MESSAGE_MAP(CAboutDlg, CDialogEx)
END_MESSAGE_MAP()


// CSadpToolsDlg 对话框




CSadpToolsDlg::CSadpToolsDlg(CWnd* pParent /*=NULL*/)
	: CDialogEx(CSadpToolsDlg::IDD, pParent)
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CSadpToolsDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialogEx::DoDataExchange(pDX);
	DDX_Control(pDX, IDC_LIST1, m_programLangList);
}

BEGIN_MESSAGE_MAP(CSadpToolsDlg, CDialogEx)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
END_MESSAGE_MAP()


// CSadpToolsDlg 消息处理程序

BOOL CSadpToolsDlg::OnInitDialog()
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

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码
	 CRect rect;   
  
    // 获取编程语言列表视图控件的位置和大小   
    m_programLangList.GetClientRect(&rect);   
  
    // 为列表视图控件添加全行选中和栅格风格   
    m_programLangList.SetExtendedStyle(m_programLangList.GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES);   
  
    // 为列表视图控件添加三列   
    m_programLangList.InsertColumn(0, _T("语言"), LVCFMT_CENTER, rect.Width()/3, 0);   
    m_programLangList.InsertColumn(1, _T("2012.02排名"), LVCFMT_CENTER, rect.Width()/3, 1);   
    m_programLangList.InsertColumn(2, _T("2011.02排名"), LVCFMT_CENTER, rect.Width()/3, 2);   
  
    // 在列表视图控件中插入列表项，并设置列表子项文本   
    m_programLangList.InsertItem(0, _T("Java"));   
    m_programLangList.SetItemText(0, 1, _T("1"));   
    m_programLangList.SetItemText(0, 2, _T("1"));   
    m_programLangList.InsertItem(1, _T("C"));   
    m_programLangList.SetItemText(1, 1, _T("2"));   
    m_programLangList.SetItemText(1, 2, _T("2"));   
    m_programLangList.InsertItem(2, _T("C#"));   
    m_programLangList.SetItemText(2, 1, _T("3"));   
    m_programLangList.SetItemText(2, 2, _T("6"));   
    m_programLangList.InsertItem(3, _T("C++"));   
    m_programLangList.SetItemText(3, 1, _T("4"));   
    m_programLangList.SetItemText(3, 2, _T("3"));   

	    // 在列表视图控件中插入列表项，并设置列表子项文本   
    m_programLangList.InsertItem(4, _T("Java"));   
    m_programLangList.SetItemText(4, 1, _T("1"));   
    m_programLangList.SetItemText(4, 2, _T("1"));   
    m_programLangList.InsertItem(5, _T("C"));   
    m_programLangList.SetItemText(5, 1, _T("2"));   
    m_programLangList.SetItemText(5, 2, _T("2"));   
    m_programLangList.InsertItem(6, _T("C#"));   
    m_programLangList.SetItemText(6, 1, _T("3"));   
    m_programLangList.SetItemText(6, 2, _T("6"));   
    m_programLangList.InsertItem(7, _T("C++"));   
    m_programLangList.SetItemText(7, 1, _T("4"));   
    m_programLangList.SetItemText(7, 2, _T("3"));   

	    // 在列表视图控件中插入列表项，并设置列表子项文本   
   
    m_programLangList.InsertItem(8, _T("C"));   
    m_programLangList.SetItemText(8, 1, _T("2"));   
    m_programLangList.SetItemText(8, 2, _T("2"));   
    m_programLangList.InsertItem(9, _T("C#"));   
    m_programLangList.SetItemText(9, 1, _T("3"));   
    m_programLangList.SetItemText(9, 2, _T("6"));   
    m_programLangList.InsertItem(10, _T("C++"));   
    m_programLangList.SetItemText(10, 1, _T("4"));   
    m_programLangList.SetItemText(10, 2, _T("3"));   

	   m_programLangList.InsertItem(11, _T("Java"));   
    m_programLangList.SetItemText(11, 1, _T("1"));   
    m_programLangList.SetItemText(11, 2, _T("1")); 

	    // 在列表视图控件中插入列表项，并设置列表子项文本   
    m_programLangList.InsertItem(12, _T("Java"));   
    m_programLangList.SetItemText(12, 1, _T("1"));   
    m_programLangList.SetItemText(12, 2, _T("1"));   
    m_programLangList.InsertItem(13, _T("C"));   
    m_programLangList.SetItemText(13, 1, _T("2"));   
    m_programLangList.SetItemText(13, 2, _T("2"));   
    m_programLangList.InsertItem(14, _T("C#"));   
    m_programLangList.SetItemText(14, 1, _T("3"));   
    m_programLangList.SetItemText(14, 2, _T("6"));   
    m_programLangList.InsertItem(15, _T("C++"));   
    m_programLangList.SetItemText(15, 1, _T("4"));   
    m_programLangList.SetItemText(15, 2, _T("3"));   


	    // 在列表视图控件中插入列表项，并设置列表子项文本   
    m_programLangList.InsertItem(16, _T("Java"));   
    m_programLangList.SetItemText(16, 1, _T("1"));   
    m_programLangList.SetItemText(16, 2, _T("1"));   
    m_programLangList.InsertItem(17, _T("C"));   
    m_programLangList.SetItemText(17, 1, _T("2"));   
    m_programLangList.SetItemText(17, 2, _T("2"));   
    m_programLangList.InsertItem(18, _T("C#"));   
    m_programLangList.SetItemText(18, 1, _T("3"));   
    m_programLangList.SetItemText(18, 2, _T("6"));   
    m_programLangList.InsertItem(19, _T("C++"));   
    m_programLangList.SetItemText(19, 1, _T("4"));   
    m_programLangList.SetItemText(19, 2, _T("3"));   

	    // 在列表视图控件中插入列表项，并设置列表子项文本   
   m_programLangList.InsertItem(20, _T("C#"));   
    m_programLangList.SetItemText(20, 1, _T("3"));   
    m_programLangList.SetItemText(20, 2, _T("6"));   
	m_programLangList.InsertItem(21, _T("C#"));   
    m_programLangList.SetItemText(21, 1, _T("3"));   
    m_programLangList.SetItemText(21, 2, _T("6"));   
	m_programLangList.InsertItem(22, _T("C#"));   
    m_programLangList.SetItemText(22, 1, _T("3"));   
    m_programLangList.SetItemText(22, 2, _T("6"));   
	m_programLangList.InsertItem(23, _T("C#"));   
    m_programLangList.SetItemText(23, 1, _T("3"));   
    m_programLangList.SetItemText(23, 2, _T("6"));

	unsigned version =  SADP_GetSadpVersion();
	CString str;
	str.Format(_T("%x"), version);
	//SetDlgItemText(IDC_STATIC,str);

	SetDlgItemText(IDC_EDIT1,str);
  
	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

void CSadpToolsDlg::OnSysCommand(UINT nID, LPARAM lParam)
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
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CSadpToolsDlg::OnPaint()
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
HCURSOR CSadpToolsDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

