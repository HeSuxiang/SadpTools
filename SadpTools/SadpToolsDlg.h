
// SadpToolsDlg.h : ͷ�ļ�
//

#pragma once
#include "afxcmn.h"

//�������ӿ�
#include "Sadp.h"

//Excel��
#include "ExcelOp.h"

//void CALLBACK DeviceInfoCallback(const SADP_DEVICE_INFO_V40 *lpDeviceInfo, void *pUserData);

// CSadpToolsDlg �Ի���
class CSadpToolsDlg : public CDialogEx
{
// ����
public:
	CSadpToolsDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_SADPTOOLS_DIALOG };

	//��̬����ָ��
	static CSadpToolsDlg* pThis;

	//��̬�ص�����
	static void CALLBACK DeviceInfoCallback(const SADP_DEVICE_INFO_V40 *lpDeviceInfo, void *pUserData);
	
	//ͳ���豸����
	int DeviceCount;

	//��������
	BOOL UpdateSadpData(CString Ipv4Address,CString MacAddress,  CString DeviceType, CString InfoTpye );


	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��

// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
