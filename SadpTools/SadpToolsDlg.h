
// SadpToolsDlg.h : 头文件
//

#pragma once
#include "afxcmn.h"

//海康威视库
#include "Sadp.h"

//Excel库
#include "ExcelOp.h"

//void CALLBACK DeviceInfoCallback(const SADP_DEVICE_INFO_V40 *lpDeviceInfo, void *pUserData);

// CSadpToolsDlg 对话框
class CSadpToolsDlg : public CDialogEx
{
// 构造
public:
	CSadpToolsDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_SADPTOOLS_DIALOG };

	//静态对象指针
	static CSadpToolsDlg* pThis;

	//静态回调函数
	static void CALLBACK DeviceInfoCallback(const SADP_DEVICE_INFO_V40 *lpDeviceInfo, void *pUserData);
	
	//统计设备数量
	int DeviceCount;

	//更新数据
	BOOL UpdateSadpData(CString Ipv4Address,CString MacAddress,  CString DeviceType, CString InfoTpye );


	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持

// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()

public:
	CListCtrl m_programLangList;
	afx_msg void OnBnClickedButton1();

	CExcelOp * m_Excel;
	afx_msg void OnBnClickedButton2();
};
