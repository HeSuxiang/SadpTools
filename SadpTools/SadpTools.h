
// SadpTools.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CSadpToolsApp:
// �йش����ʵ�֣������ SadpTools.cpp
//

class CSadpToolsApp : public CWinApp
{
public:
	CSadpToolsApp();

// ��д
public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CSadpToolsApp theApp;