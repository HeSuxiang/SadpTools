#pragma once

//http://www.blogjava.net/jinheking/archive/2005/07/19/5150.html �ϲ���Ԫ��
//http://blog.csdn.net/qinghezhen/article/details/9906023
//http://blog.csdn.net/gyssoft/article/details/1592104
//http://read.pudn.com/downloads100/doc/fileformat/411493/C++/ComTypeLibfor7/comexcel/excel10/comexcel.cpp__.htm
//http://www.cnblogs.com/fullsail/archive/2012/12/28/2837952.html ʹ��OLE���ٶ�дEXCEL��Դ��
//http://blog.csdn.net/handsing/article/details/5461070
//http://www.cnblogs.com/xianyunhe/archive/2011/09/13/2174703.html
//http://www.update8.com/Program/C++/27229.html
//http://bbs.csdn.net/topics/390091827 
#include "stdafx.h" 
//#include "Debug/excel.tlh"

//OLE��ͷ�ļ�
#include "CRange.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CApplication.h"

class CExcelOp
{
public:
	CExcelOp(void);
	~CExcelOp(void);


public:
	void ShowInExcel(BOOL bShow);

	///���һ��CELL�Ƿ����ַ���
	//BOOL    IsCellString(long iRow, long iColumn);
	///���һ��CELL�Ƿ�����ֵ
	//BOOL    IsCellInt(long iRow, long iColumn);

	void SetCell(long irow, long icolumn,CString new_string);
	///�õ�һ��CELL��String
	CString GetCellString(long iRow, long iColumn);
	CString GetCellStringByName(CString rowName,CString colName);

	///�õ�����
	//int     GetCellInt(long iRow, long iColumn);
	///�õ�double������
	//double  GetCellDouble(long iRow, long iColumn);

	///ȡ���е�����
	int GetRowCount();
	///ȡ���е�����
	int GetColumnCount();

	///ʹ��ĳ��shet��shit��shit
	BOOL LoadSheet(long table_index,BOOL pre_load = FALSE);
	///ͨ������ʹ��ĳ��sheet��
	BOOL LoadSheet(const TCHAR* sheet,BOOL pre_load = FALSE);
	///ͨ�����ȡ��ĳ��Sheet������
	CString GetSheetName(long table_index);

	///�õ�Sheet������
	int GetSheetCount();

	///���ļ�
	BOOL OpenExcelFile(const TCHAR * file_name);
	///�رմ򿪵�Excel �ļ�����ʱ���EXCEL�ļ���Ҫ
	void CloseExcelFile(BOOL if_save = FALSE);
	//���Ϊһ��EXCEL�ļ�
	void SaveasXSLFile(const CString &xls_file);
	///ȡ�ô��ļ�������
	CString GetOpenFileName();
	///ȡ�ô�sheet������
	CString GetLoadSheetName();

	///д��һ��CELLһ��int
	//void SetCellInt(long irow, long icolumn,int new_int);
	///д��һ��CELLһ��string
	//void SetCellString(long irow, long icolumn,CString new_string);

public:
	///��ʼ��EXCEL OLE
	static BOOL InitExcel();
	///�ͷ�EXCEL�� OLE
	static void ReleaseExcel();
	///ȡ���е����ƣ�����27->AA
	static char *GetColumnName(long iColumn);

protected:

	//Ԥ�ȼ���
	void PreLoadSheet(); 

public:
	///�򿪵�EXCEL�ļ�����
	CString       open_excel_file_;

	CWorkbooks    m_Books;
	CWorkbook     m_Book;
	CWorksheets   m_sheets;
	CWorksheet    m_sheet;
	CRange        m_Rge;

 	static CApplication m_app;

	///�Ƿ��Ѿ�Ԥ������ĳ��sheet������
	BOOL          already_preload_;

	///Create the SAFEARRAY from the VARIANT ret.
	COleSafeArray ole_safe_array_;
 

	LPDISPATCH m_lpDisp;  
 
};
