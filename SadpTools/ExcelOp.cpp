#include "StdAfx.h"
#include "ExcelOp.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif



CExcelOp::CExcelOp(void):already_preload_(FALSE)
{

}

CExcelOp::~CExcelOp(void)
{
 
}

COleVariant
covTrue((short)TRUE),
covFalse((short)FALSE),
covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
 
CApplication CExcelOp::m_app;

//初始化EXCEL文件，
BOOL CExcelOp::InitExcel()
{ 
	BOOL bRtn =FALSE ;
	TCHAR* ptchsExcel[]={
		_T("Excel.Application") //创建Excel 2000服务器(启动Excel) 
		,_T("Excel.Application.8")//Excel 97
		,_T("Excel.Application.9")//Excel 2000
		,_T("Excel.Application.10")//Excel xp
		,_T("Excel.Application.11")//Excel 2003
		,_T("Excel.Application.12")//Excel 2007
		,_T("Excel.Application.14")//Excel 2010

	};

	CLSID clsid;
	HRESULT hr = S_FALSE;
	for (int i= sizeof(ptchsExcel)/sizeof(*ptchsExcel)-1;i>=0;i--)
	{
#if 0
		
		if (m_app.CreateDispatch(ptchsExcel[i],NULL)) 
		{
			break;
		}
		等同于下面的操作
#else
		hr = CLSIDFromProgID(ptchsExcel[i], &clsid);
		if(SUCCEEDED(hr))
		{ 
			hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&m_app);
			if(SUCCEEDED(hr)) 
			{
				bRtn = TRUE;
			}
			break;
		} 
#endif
	}

	if (!bRtn)
	{
		AfxMessageBox(_T("创建Excel服务失败,你可能没有安装EXCEL，请检查!")); 
		return bRtn;
	}

	m_app.put_DisplayAlerts(FALSE); 
	return bRtn;
}

//
void CExcelOp::ReleaseExcel()
{ 
	m_app.Quit();
	m_app.ReleaseDispatch();
	m_app=NULL;
}


//关闭打开的Excel 文件,默认情况不保存文件
void CExcelOp::CloseExcelFile(BOOL if_save)
{
	//如果已经打开，关闭文件
	if (open_excel_file_.IsEmpty() == FALSE)
	{
		//如果保存,交给用户控制,让用户自己存，如果自己SAVE，会出现莫名的等待
		if (if_save)
		{
			ShowInExcel(TRUE);
		}
		else
		{
			m_Book.Close(COleVariant(short(FALSE)),COleVariant(open_excel_file_),covOptional);
			m_Books.Close();
		}

		//打开文件的名称清空
		open_excel_file_.Empty();
	}

	m_sheets.ReleaseDispatch();
	m_sheet.ReleaseDispatch();
	m_Rge.ReleaseDispatch();
	m_Book.ReleaseDispatch();
	m_Books.ReleaseDispatch();
}

//打开excel文件
BOOL CExcelOp::OpenExcelFile(const TCHAR *file_name)
{
	//先关闭
	CloseExcelFile();

	//m_Books.AttachDispatch(m_app.get_Workbooks(),1);  
	//COleVariant varPath(file_name);  
	//m_Book.AttachDispatch(m_Books.Add(varPath));  

	//利用模板文件建立新文档 
	m_Books.AttachDispatch(m_app.get_Workbooks(),true); 

	//LPDISPATCH lpDis = NULL;
	m_lpDisp = m_Books.Add(COleVariant(file_name)); 
	if (m_lpDisp)
	{
		m_Book.AttachDispatch(m_lpDisp); 
		//得到Worksheets 
		m_sheets.AttachDispatch(m_Book.get_Worksheets(),true); 

		//记录打开的文件名称
		open_excel_file_ = file_name;

		return TRUE;
	}

	return TRUE;
}



//
void CExcelOp::ShowInExcel(BOOL bShow)
{
	m_app.put_Visible(bShow);
	m_app.put_UserControl(bShow);
}

void CExcelOp::SaveasXSLFile(const CString &xls_file)
{
	m_Book.SaveAs(COleVariant(xls_file),
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		0,
		covOptional,
		covOptional,
		covOptional,
		covOptional,
		covOptional);
	return;
}


int CExcelOp::GetSheetCount()
{
	return m_sheets.get_Count();
}


CString CExcelOp::GetSheetName(long table_index)
{
	CWorksheet sheet;
	sheet.AttachDispatch(m_sheets.get_Item(COleVariant((long)table_index)),true);
	CString name = sheet.get_Name();
	sheet.ReleaseDispatch();
	return name;
}

//按照序号加载Sheet表格,可以提前加载所有的表格内部数据
BOOL CExcelOp::LoadSheet(long table_index,BOOL pre_load)
{
	LPDISPATCH lpDis = NULL;
	m_Rge.ReleaseDispatch();
	m_sheet.ReleaseDispatch();
	lpDis = m_sheets.get_Item(COleVariant((long)table_index));
	if (lpDis)
	{
		m_sheet.AttachDispatch(lpDis,true);
		m_Rge.AttachDispatch(m_sheet.get_Cells(), true);
	}
	else
	{
		return FALSE;
	}

	already_preload_ = FALSE;
	//如果进行预先加载
	if (pre_load)
	{
		PreLoadSheet();
		already_preload_ = TRUE;
	}

	return TRUE;
}


//按照名称加载Sheet表格,可以提前加载所有的表格内部数据
BOOL CExcelOp::LoadSheet(const TCHAR* sheet,BOOL pre_load)
{
	LPDISPATCH lpDis = NULL;
	m_Rge.ReleaseDispatch();
	m_sheet.ReleaseDispatch();
	lpDis = m_sheets.get_Item(COleVariant(sheet));
	if (lpDis)
	{
		m_sheet.AttachDispatch(lpDis,true);
		m_Rge.AttachDispatch(m_sheet.get_Cells(), true);

	}
	else
	{
		return FALSE;
	}
	//
	already_preload_ = FALSE;
	//如果进行预先加载
	if (pre_load)
	{
		already_preload_ = TRUE;
		PreLoadSheet();
	}

	return TRUE;
}



//得到列的总数
int CExcelOp::GetColumnCount()
{
	CRange range;
	CRange usedRange;
	usedRange.AttachDispatch(m_sheet.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Columns(), true);
	int count = range.get_Count();
	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();
	return count;
}

//得到行的总数
int CExcelOp::GetRowCount()
{
	CRange range;
	CRange usedRange;
	usedRange.AttachDispatch(m_sheet.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Rows(), true);
	int count = range.get_Count();
	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();
	return count;
}

//返回打开的EXCEL文件名称
CString CExcelOp::GetOpenFileName()
{
	return open_excel_file_;
}

//取得打开sheet的名称
CString CExcelOp::GetLoadSheetName()
{
	return m_sheet.get_Name();
}


//取得列的名称，比如27->AA
char *CExcelOp::GetColumnName(long icolumn)
{   
	static char column_name[64];
	size_t str_len = 0;

	while(icolumn > 0)
	{
		int num_data = icolumn % 26;
		icolumn /= 26;
		if (num_data == 0)
		{
			num_data = 26;
			icolumn--;
		}
		column_name[str_len] = (char)((num_data-1) + 'A' );
		str_len ++;
	}
	column_name[str_len] = '\0';
	//反转
	_strrev(column_name);

	return column_name;
}

//预先加载
void CExcelOp::PreLoadSheet()
{
	CRange used_range;
	used_range = m_sheet.get_UsedRange();    

	VARIANT ret_ary = used_range.get_Value2();
	if (!(ret_ary.vt & VT_ARRAY))
	{
		return;
	}
	ole_safe_array_.Clear();
	ole_safe_array_.Attach(ret_ary); 
}

void CExcelOp::SetCell(long irow, long icolumn,CString new_string)
{
	COleVariant new_value(new_string);
	CRange start_range = m_sheet.get_Range(COleVariant(_T("A1")),covOptional);
	CRange write_range = start_range.get_Offset(COleVariant((long)irow -1),COleVariant((long)icolumn -1) );
	write_range.put_Value2(new_value);
	start_range.ReleaseDispatch();
	write_range.ReleaseDispatch();
}

CString CExcelOp::GetCellString(long iRow, long iColumn)  
{  
	//_variant_t varRow(iRow);  
	//_variant_t varCol(iColumn); 

	//COleVariant value;  
	//range.AttachDispatch(m_sheet.get_Cells(),TRUE);  
	//value=range.get_Item(varRow,varCol);                    //返回的类型是VT_DISPATCH 这是一个指针  
	//range.AttachDispatch(value.pdispVal,TRUE);  
	//VARIANT value2=range.get_Text();  
	//CString strValue=value2.bstrVal;  
	//return strValue;  

	COleVariant vResult ; 
	//字符串
	if (already_preload_ == FALSE)
	{
		m_Rge.AttachDispatch(m_Rge.get_Item (COleVariant((long)iRow),COleVariant((long)iColumn)).pdispVal, true);
		vResult =m_Rge.get_Value2();
	}
	//如果数据依据预先加载了
	else
	{
		long read_address[2];
		VARIANT val;
		read_address[0] = iRow;
		read_address[1] = iColumn;
		ole_safe_array_.GetElement(read_address, &val);
		vResult = val;
	}

	CString str;
	if(vResult.vt == VT_BSTR)       //字符串
	{
		str=vResult.bstrVal;
	}
	else if (vResult.vt==VT_INT)
	{
		str.Format(_T("%d"),vResult.pintVal);
	}
	else if (vResult.vt==VT_R8)     //8字节的数字
	{
		str.Format(_T("%0.0f"),vResult.dblVal);
		//str.Format("%.0f",vResult.dblVal);
		//str.Format("%1f",vResult.fltVal);
	}
	else if(vResult.vt==VT_DATE)    //时间格式
	{
		SYSTEMTIME st;
		VariantTimeToSystemTime(vResult.date, &st);
		CTime tm(st);
		str=tm.Format(_T("%Y-%m-%d"));

	}
	else if(vResult.vt==VT_EMPTY)   //单元格空的
	{
		str=_T("");
	} 

	m_Rge.ReleaseDispatch();

	return str;
} 

CString CExcelOp::GetCellStringByName(CString rowName,CString colName)  
{  
	COleVariant value;  
	CString strValue;  
	long row=0,col=0;  
	long re_row=0,re_col=0;

	m_Rge.AttachDispatch(m_sheet.get_Cells(),TRUE);  
	for (row=1,col=1;col<m_Rge.get_Column();col++)  
	{  
		value=m_Rge.get_Item(_variant_t(row),_variant_t(col));                  //返回的类型是VT_DISPATCH 这是一个指针  
		m_Rge.AttachDispatch(value.pdispVal,TRUE);  
		VARIANT value2=m_Rge.get_Text();  
		CString strValue=value2.bstrVal;  
		if (strValue==colName)  
			break;  
	}  
	re_col=col;  
	for (row=1,row=1;row<m_Rge.get_Row();row++)  
	{  
		value=m_Rge.get_Item(_variant_t(row),_variant_t(col));                  //返回的类型是VT_DISPATCH 这是一个指针  
		m_Rge.AttachDispatch(value.pdispVal,TRUE);  
		VARIANT value2=m_Rge.get_Text();  
		CString strValue=value2.bstrVal;  
		if (strValue==rowName)        
			break;  
	}  
	re_row=row;  
	return GetCellString(re_row,re_col);  
}  
