#pragma once

//http://www.blogjava.net/jinheking/archive/2005/07/19/5150.html 合并单元格
//http://blog.csdn.net/qinghezhen/article/details/9906023
//http://blog.csdn.net/gyssoft/article/details/1592104
//http://read.pudn.com/downloads100/doc/fileformat/411493/C++/ComTypeLibfor7/comexcel/excel10/comexcel.cpp__.htm
//http://www.cnblogs.com/fullsail/archive/2012/12/28/2837952.html 使用OLE高速读写EXCEL的源码
//http://blog.csdn.net/handsing/article/details/5461070
//http://www.cnblogs.com/xianyunhe/archive/2011/09/13/2174703.html
//http://www.update8.com/Program/C++/27229.html
//http://bbs.csdn.net/topics/390091827 
#include "stdafx.h" 
//#include "Debug/excel.tlh"

//OLE的头文件
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

	///检查一个CELL是否是字符串
	//BOOL    IsCellString(long iRow, long iColumn);
	///检查一个CELL是否是数值
	//BOOL    IsCellInt(long iRow, long iColumn);

	void SetCell(long irow, long icolumn,CString new_string);
	///得到一个CELL的String
	CString GetCellString(long iRow, long iColumn);
	CString GetCellStringByName(CString rowName,CString colName);

	///得到整数
	//int     GetCellInt(long iRow, long iColumn);
	///得到double的数据
	//double  GetCellDouble(long iRow, long iColumn);

	///取得行的总数
	int GetRowCount();
	///取得列的总数
	int GetColumnCount();

	///使用某个shet，shit，shit
	BOOL LoadSheet(long table_index,BOOL pre_load = FALSE);
	///通过名称使用某个sheet，
	BOOL LoadSheet(const TCHAR* sheet,BOOL pre_load = FALSE);
	///通过序号取得某个Sheet的名称
	CString GetSheetName(long table_index);

	///得到Sheet的总数
	int GetSheetCount();

	///打开文件
	BOOL OpenExcelFile(const TCHAR * file_name);
	///关闭打开的Excel 文件，有时候打开EXCEL文件就要
	void CloseExcelFile(BOOL if_save = FALSE);
	//另存为一个EXCEL文件
	void SaveasXSLFile(const CString &xls_file);
	///取得打开文件的名称
	CString GetOpenFileName();
	///取得打开sheet的名称
	CString GetLoadSheetName();

	///写入一个CELL一个int
	//void SetCellInt(long irow, long icolumn,int new_int);
	///写入一个CELL一个string
	//void SetCellString(long irow, long icolumn,CString new_string);

public:
	///初始化EXCEL OLE
	static BOOL InitExcel();
	///释放EXCEL的 OLE
	static void ReleaseExcel();
	///取得列的名称，比如27->AA
	static char *GetColumnName(long iColumn);

protected:

	//预先加载
	void PreLoadSheet(); 

public:
	///打开的EXCEL文件名称
	CString       open_excel_file_;

	CWorkbooks    m_Books;
	CWorkbook     m_Book;
	CWorksheets   m_sheets;
	CWorksheet    m_sheet;
	CRange        m_Rge;

 	static CApplication m_app;

	///是否已经预加载了某个sheet的数据
	BOOL          already_preload_;

	///Create the SAFEARRAY from the VARIANT ret.
	COleSafeArray ole_safe_array_;
 

	LPDISPATCH m_lpDisp;  
 
};
