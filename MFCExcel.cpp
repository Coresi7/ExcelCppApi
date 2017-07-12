/*
 * MFCExcel.cpp
 * Author: hearstzhang (hearstzhang@tencent.com)
 * Please let me know if you have any question about this.
 * Some version of Microsoft Excel may failed to generate the proper
 * ole header files for this. At least capable with Microsoft Excel
 * version 2010, 2013, 2016.
 * Your right of copy and redistribute is granted
 * as long as you keep this copyright annoucement.
 */

 // ----------------------MFCExcel.cpp----------------

#include "stdafx.h"
#include "MFCExcel.h"

//MFCExcelSingletonAdmin class implementation area
unsigned int MFCExcelSingletonAdmin::m_nReference = 0;
MFCExcelFile * MFCExcelSingletonAdmin::m_pInstance = nullptr;

MFCExcelSingletonAdmin::MFCExcelSingletonAdmin()
{
	if (!m_pInstance)
	{
		m_pInstance = new MFCExcelFile();
	}
	if (!m_pInstance->InitExcel())
	{
		//This computer didn't install Microsoft Excel or WPS
		delete m_pInstance;
		m_pInstance = nullptr;
	}
	else
	{
		m_nReference++;
	}
}

MFCExcelSingletonAdmin::~MFCExcelSingletonAdmin()
{
	if (m_pInstance != nullptr)
	{
		m_nReference--;
		if (!m_nReference)
		{
			if (!m_pInstance->GetOpenedFileName().IsEmpty())
			{
				//close file first
				//because this will release resources.
				m_pInstance->CloseExcelFile();
			}
			//close Excel applicaton
			m_pInstance->ReleaseExcel();
			delete m_pInstance;
			m_pInstance = nullptr;
		}
	}
	else
	{
		m_nReference = 0;
		return;
	}
}

// return instance
// will return nullptr if failed to create instance.
MFCExcelFile * MFCExcelSingletonAdmin::GetInstance()
{
	return m_pInstance;
}

//MFCExcelFile class implementation area
enum XLDeleteMethod {
	xlShiftToLeft = -4159L,
	xlShiftUp = -4162L
};

COleVariant
covTrue((short)TRUE),
covFalse((short)FALSE),
covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

// the Application of the excel
CApplication MFCExcelFile::excel_application_;

bool MFCExcelFile::isInited = false;
bool MFCExcelFile::isReleased = true;

MFCExcelFile::MFCExcelFile()
	:already_preload_(FALSE), isUpdate(false)
{
	return;
}

MFCExcelFile::~MFCExcelFile()
{
	return;
}

// initial excel server
// and create excel application here.
BOOL MFCExcelFile::InitExcel()
{
	try
	{
		if (MFCExcelFile::isInited)
		{
			return TRUE;
		}
		// Create Excel server.
#ifdef UNICODE
		if (!excel_application_.CreateDispatch(L"Excel.Application", NULL))
#else
		if (!excel_application_.CreateDispatch("Excel.Application", NULL))
#endif
		{
			// You may not install Microsoft Excel or Kingsoft WPS.
			return FALSE;
		}

		excel_application_.put_DisplayAlerts(FALSE);

		//initialized successfully.
		MFCExcelFile::isInited = true;
		MFCExcelFile::isReleased = false;
		return TRUE;
	}
	catch (...)
	{
		return FALSE;
	}
}

// Release excel and close excel server.
void MFCExcelFile::ReleaseExcel()
{
	if (MFCExcelFile::isReleased)
	{
		return;
	}
	MFCExcelFile::isReleased = true;
	MFCExcelFile::isInited = false;
	excel_application_.Quit();
	excel_application_.ReleaseDispatch();
	excel_application_ = NULL;
}

// open excel file
// wile create one in memory if failed to find it
// save to take effect.
BOOL MFCExcelFile::OpenExcelFile(const TCHAR *file_name)
{
	try
	{
		// get relative path.
		CString tempFileName = file_name;
		GetRelativePathIfNot(tempFileName);

		// if can not find such file
		// create one.
		if (!PathFileExists(tempFileName))
		{
			CloseExcelFile();
			excel_books_.AttachDispatch(excel_application_.get_Workbooks(), TRUE);
			excel_work_book_.AttachDispatch(excel_books_.Add(_variant_t(CString())));
			excel_sheets_.AttachDispatch(excel_work_book_.get_Worksheets(), TRUE);
			// insert a sheet if the sheet count is zero.
			if (excel_sheets_.get_Count() == 0)
			{
				excel_sheets_.Add(covOptional, covOptional, _variant_t(1), covOptional);
			}
			open_excel_file_ = tempFileName;
			return open_excel_file_.IsEmpty() ? FALSE : TRUE;
		}
		else
		{
			// close first
			CloseExcelFile();

			// create new file with templete
			excel_books_.AttachDispatch(excel_application_.get_Workbooks(), TRUE);

			LPDISPATCH lpDis = NULL;
			lpDis = excel_books_.Add(COleVariant(tempFileName));
			if (lpDis)
			{
				excel_work_book_.AttachDispatch(lpDis);
				// get Worksheets 
				excel_sheets_.AttachDispatch(excel_work_book_.get_Worksheets(), TRUE);
				// use this to judge if you have already opened an excel.
				// record opened file name
				open_excel_file_ = tempFileName;

				return TRUE;
			}
			return FALSE;
		}
	}
	catch (...)
	{
		return FALSE;
	}
}

// close current excel file
// wont save it by default.
void MFCExcelFile::CloseExcelFile(BOOL if_save)
{
	// close first
	if (open_excel_file_.IsEmpty() == FALSE)
	{
		// save file.
		if (if_save)
		{
			// ShowInExcel(TRUE);
			// excel_work_book_.Save();
			// excel_work_book_.Close(COleVariant(short(TRUE)), COleVariant(open_excel_file_), covOptional);
			SaveasXLSFile(open_excel_file_);
			excel_books_.Close();
		}
		else
		{
			// 
			excel_work_book_.Close(COleVariant(short(FALSE)), COleVariant(open_excel_file_), covOptional);
			excel_books_.Close();
		}

		// clear file name
		open_excel_file_.Empty();
	}

	//release resources
	excel_sheets_.ReleaseDispatch();
	excel_work_sheet_.ReleaseDispatch();
	excel_current_range_.ReleaseDispatch();
	excel_work_book_.ReleaseDispatch();
	excel_books_.ReleaseDispatch();
	if (already_preload_)
	{
		already_preload_ = FALSE;
		isUpdate = false;
		ole_safe_array_.DestroyData();
	}
}

// save excel file.
void MFCExcelFile::SaveasXLSFile(const CString &xls_file)
{
	// make it a relative path if not.
	CString tempFileName = xls_file;
	GetRelativePathIfNot(tempFileName);

	excel_work_book_.SaveAs(COleVariant(tempFileName),
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

// save and reload current file and sheet.
void MFCExcelFile::Save()
{
	if (open_excel_file_.IsEmpty())
	{
		return;
	}
	else
	{
		CString tempFileName = open_excel_file_;
		CString tempSheetName = GetLoadSheetName();
		BOOL isPre = already_preload_;
		CloseExcelFile(TRUE);
		OpenExcelFile(tempFileName);
		LoadSheet(tempSheetName, isPre);
		return;
	}
}

// get sheet numbers
long MFCExcelFile::GetSheetCount()
{
	return excel_sheets_.get_Count();
}

// get the name of a sheet by its index
CString MFCExcelFile::GetSheetName(long table_index)
{
	CWorksheet sheet;
	sheet.AttachDispatch(excel_sheets_.get_Item(COleVariant((long)table_index)), true);
	CString name = sheet.get_Name();
	sheet.ReleaseDispatch();
	return name;
}

// load sheet content by its index
// index start from 1
// will create sheet until the index provided with name "SheetN"
// if the index is bigger than the biggest index of the current file.
// preload mode by default
BOOL MFCExcelFile::LoadSheet(long table_index, BOOL pre_load)
{
	try
	{
		if (table_index <= 0)
		{
			return FALSE;
		}
		LPDISPATCH lpDis = NULL;
		excel_current_range_.ReleaseDispatch();
		excel_work_sheet_.ReleaseDispatch();
		auto sheetcount = excel_sheets_.get_Count();
		if (sheetcount < table_index)
		{
			for (auto i = sheetcount; i < table_index; i++)
			{
				excel_work_sheet_ = excel_sheets_.get_Item(_variant_t(i));
				excel_sheets_.Add(covOptional, _variant_t(excel_work_sheet_), _variant_t(1), covOptional);
			}
		}
		lpDis = excel_sheets_.get_Item(COleVariant((long)table_index));
		if (lpDis)
		{
			excel_work_sheet_.AttachDispatch(lpDis, true);
			excel_current_range_.AttachDispatch(excel_work_sheet_.get_Cells(), true);
		}
		else
		{
			return FALSE;
		}

		already_preload_ = FALSE;
		// preload
		if (pre_load)
		{
			PreLoadSheet();
			already_preload_ = TRUE;
		}

		return TRUE;
	}
	catch (...)
	{
		return FALSE;
	}
}

// load sheet content by its name
// will create one with the name provided if failed to find such sheet.
// preload mode by default
BOOL MFCExcelFile::LoadSheet(const TCHAR* sheet, BOOL pre_load)
{
	try
	{
		LPDISPATCH lpDis = NULL;
		excel_current_range_.ReleaseDispatch();
		excel_work_sheet_.ReleaseDispatch();
		try
		{
			lpDis = excel_sheets_.get_Item(COleVariant(sheet));

			excel_work_sheet_.AttachDispatch(lpDis, true);
			auto tmp = excel_work_sheet_.get_Cells();
			excel_current_range_.AttachDispatch(tmp, true);

		}
		catch (...)
		{
			auto sheetcount = excel_sheets_.get_Count();
			excel_work_sheet_ = excel_sheets_.get_Item(_variant_t(sheetcount));
			lpDis = excel_sheets_.Add(covOptional, _variant_t(excel_work_sheet_), _variant_t(1), covOptional);
			excel_work_sheet_.AttachDispatch(lpDis);
			excel_work_sheet_.put_Name(sheet);
		}

		// 
		already_preload_ = FALSE;
		// preload
		if (pre_load)
		{
			PreLoadSheet();
			already_preload_ = TRUE;
		}

		return TRUE;
	}
	catch (...)
	{
		return FALSE;
	}
}

// get total count of cols
long MFCExcelFile::GetColumnCount()
{
	CRange range;
	CRange usedRange;
	usedRange.AttachDispatch(excel_work_sheet_.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Columns(), true);
	auto count = range.get_Count();
	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();
	return count;
}

// get total count of Rows
long MFCExcelFile::GetRowCount()
{
	CRange range;
	CRange usedRange;
	usedRange.AttachDispatch(excel_work_sheet_.get_UsedRange(), true);
	range.AttachDispatch(usedRange.get_Rows(), true);
	auto count = range.get_Count();
	usedRange.ReleaseDispatch();
	range.ReleaseDispatch();
	return count;
}

// Detect if a cell contains string.
BOOL MFCExcelFile::IsCellString(long irow, long icolumn)
{
	CRange range;
	range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
	COleVariant vResult = range.get_Value2();
	range.ReleaseDispatch();
	if (vResult.vt == VT_BSTR)
	{
		return TRUE;
	}
	return FALSE;
}

// Detect if a cell contains number.
BOOL MFCExcelFile::IsCellInt(long irow, long icolumn)
{
	CRange range;
	range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
	COleVariant vResult = range.get_Value2();
	range.ReleaseDispatch();
	// VT_R8 8 bytes of real number
	if (vResult.vt == VT_INT || vResult.vt == VT_R8)
	{
		return TRUE;
	}
	return FALSE;
}

// Get a cell content as CString
CString MFCExcelFile::GetCellString(long irow, long icolumn)
{
	COleVariant vResult;
	CString str;
	if (already_preload_ == FALSE)
	{
		CRange range;
		range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
		vResult = range.get_Value2();
		range.ReleaseDispatch();
	}
	// if preload mode
	else
	{
		// Update first to get latest changes
		if (isUpdate)
		{
			PreLoadSheet();
		}
		if (!ole_safe_array_.GetDim())
		{
#ifdef UNICODE
			return L"";
#else
			return "";
#endif
		}
		long read_address[2];
		VARIANT val;
		read_address[0] = irow;
		read_address[1] = icolumn;
		long uBoundRow, uBoundCol;
		ole_safe_array_.GetUBound(1, &uBoundRow);
		ole_safe_array_.GetUBound(2, &uBoundCol);
		if (irow > uBoundRow || icolumn > uBoundCol
			|| irow <= 0 || icolumn <= 0)
		{
#ifdef UNICODE
			return L"";
#else
			return "";
#endif
		}
		ole_safe_array_.GetElement(read_address, &val);
		vResult = val;
	}

	if (vResult.vt == VT_BSTR)
	{
		str = vResult.bstrVal;
	}
	// interger
	else if (vResult.vt == VT_INT)
	{
#ifdef UNICODE
		str.Format(L"%d", vResult.pintVal);
#else
		str.Format("%d", vResult.pintVal);
#endif
	}
	// 8 byte real number
	else if (vResult.vt == VT_R8)
	{
#ifdef UNICODE
		str.Format(L"%0.8f", vResult.dblVal);
#else
		str.Format("%0.8f", vResult.dblVal);
#endif
	}
	// time format
	else if (vResult.vt == VT_DATE)
	{
		SYSTEMTIME st;
		VariantTimeToSystemTime(vResult.date, &st);
		CTime tm(st);
#ifdef UNICODE
		str = tm.Format(L"%Y-%m-%d");
#else
		str = tm.Format("%Y-%m-%d");
#endif

	}
	// empty cell
	else if (vResult.vt == VT_EMPTY)
	{
#ifdef UNICODE
		str = L"";
#else
		str = "";
#endif
	}

	return str;
}

// Get cell content as a double number
double MFCExcelFile::GetCellDouble(long irow, long icolumn)
{
	double rtn_value = 0.0;
	COleVariant vresult;
	// string
	if (already_preload_ == FALSE)
	{
		CRange range;
		range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
		vresult = range.get_Value2();
		range.ReleaseDispatch();
	}
	// if preload mode
	else
	{
		// Update first to get latest changes
		if (isUpdate)
		{
			PreLoadSheet();
		}
		if (!ole_safe_array_.GetDim())
		{
			return 0.0;
		}
		long read_address[2];
		VARIANT val;
		read_address[0] = irow;
		read_address[1] = icolumn;
		long uBoundRow, uBoundCol;
		ole_safe_array_.GetUBound(1, &uBoundRow);
		ole_safe_array_.GetUBound(2, &uBoundCol);
		if (irow > uBoundRow || icolumn > uBoundCol
			|| irow <= 0 || icolumn <= 0)
		{
			return 0.0;
		}

		ole_safe_array_.GetElement(read_address, &val);
		vresult = val;
	}

	if (vresult.vt == VT_R8)
	{
		rtn_value = vresult.dblVal;
	}

	return rtn_value;
}

// VT_R8
int MFCExcelFile::GetCellInt(long irow, long icolumn)
{
	int num;
	COleVariant vresult;

	if (already_preload_ == FALSE)
	{
		CRange range;
		range.AttachDispatch(excel_current_range_.get_Item(COleVariant((long)irow), COleVariant((long)icolumn)).pdispVal, true);
		vresult = range.get_Value2();
		range.ReleaseDispatch();
	}
	// Preload mode
	else
	{
		// Update first to get latest changes
		if (isUpdate)
		{
			PreLoadSheet();
		}
		if (!ole_safe_array_.GetDim())
		{
			return 0;
		}
		long read_address[2];
		VARIANT val;
		read_address[0] = irow;
		read_address[1] = icolumn;
		long uBoundRow, uBoundCol;
		ole_safe_array_.GetUBound(1, &uBoundRow);
		ole_safe_array_.GetUBound(2, &uBoundCol);
		if (irow > uBoundRow || icolumn > uBoundCol
			|| irow <= 0 || icolumn <= 0)
		{
			return 0;
		}
		ole_safe_array_.GetElement(read_address, &val);
		vresult = val;
	}
	// Get the content
	num = static_cast<int>(vresult.dblVal);

	return num;
}

// Using a string to set cell content
void MFCExcelFile::SetCellString(long irow, long icolumn, CString new_string)
{
	if (irow <= 0 || icolumn <= 0)
	{
		return;
	}

	COleVariant new_value(new_string);
	CRange start_range = excel_work_sheet_.get_Range(COleVariant(GetCellName(irow, icolumn)), covOptional);
	start_range.put_Value2(new_value);
	start_range.ReleaseDispatch();

	// Ask for refresh memory data at next reading function.
	if (!isUpdate && already_preload_)
	{
		isUpdate = true;
	}
}

// Using an interget to set cell content
void MFCExcelFile::SetCellInt(long irow, long icolumn, int new_int)
{
	if (irow <= 0 || icolumn <= 0)
	{
		return;
	}

	COleVariant new_value((long)new_int);
	CRange start_range = excel_work_sheet_.get_Range(COleVariant(GetCellName(irow, icolumn)), covOptional);
	start_range.put_Value2(new_value);
	start_range.ReleaseDispatch();

	// Ask for refresh memory data at next reading function.
	if (!isUpdate && already_preload_)
	{
		isUpdate = true;
	}
}

// set a range of cell content
// with a high performance
bool MFCExcelFile::SetRangeCellString(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
	const std::vector<CString>& new_string)
{
	if (!iRowStart || !iRowEnd)
	{
		return false;
	}
	if (iRowStart > iRowEnd || iColStart > iColEnd)
	{
		return false;
	}

	// Calculate size of vector and excel modify content.
	DWORD numCols = iColEnd - iColStart + 1;
	DWORD numRows = iRowEnd - iRowStart + 1;

	CString temp1, temp2;
	//temp1 will be reused as a blank string.
	temp1 = MFCExcelFile::GetCellName(iRowStart, iColStart);
	temp2 = MFCExcelFile::GetCellName(iRowEnd, iColEnd);

	CRange range = excel_work_sheet_.get_Range(_variant_t(temp1), _variant_t(temp2));

	COleSafeArray saRet;
	DWORD numElements[2];
	numElements[0] = numRows;
	numElements[1] = numCols;
	saRet.Create(VT_BSTR, 2, numElements);

	//reuse temp1
#ifdef UNICODE
	temp1 = L"";
#else
	temp1 = "";
#endif

	long index[2];
	size_t vectorsize = new_string.size();
	long tempRecur;
	// Rows
	for (index[0] = 0; index[0] < static_cast<long>(numRows); index[0]++)
	{
		// Cols
		for (index[1] = 0; index[1] < static_cast<long>(numCols); index[1]++)
		{
			BSTR bstr;
			tempRecur = index[0] * (numCols)+index[1];
			if (tempRecur < static_cast<long>(vectorsize))
			{
				bstr = new_string[tempRecur].AllocSysString();
			}
			else
			{
				// reuse temp1 here
				bstr = temp1.AllocSysString();
			}
			saRet.PutElement(index, bstr);
			SysFreeString(bstr);
		}
	}

	range.put_Value(covOptional, (COleVariant)saRet);
	range.ReleaseDispatch();
	saRet.DestroyData();

	// Ask for refresh memory data at next reading function.
	if (!isUpdate && already_preload_)
	{
		isUpdate = true;
	}
	return true;
}

// set a range of cell string
// and clear rest rows
bool MFCExcelFile::SetRangeCellStringAndClearRestRows(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
	const std::vector<CString>& new_string)
{
	if (!SetRangeCellString(iRowStart, iRowEnd, iColStart, iColEnd, new_string))
	{
		return false;
	}
	auto temp = GetRowCount();
	if (temp >= (long)iRowEnd + 1)
	{
		DeleteRangedRow(iRowEnd + 1, temp);
	}
	return true;
}

// set a range of cell string
// and clear rest cols
bool MFCExcelFile::SetRangeCellStringAndClearRestCols(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
	const std::vector<CString>& new_string)
{
	if (!SetRangeCellString(iRowStart, iRowEnd, iColStart, iColEnd, new_string))
	{
		return false;
	}
	auto temp = GetColumnCount();
	if (temp >= (long)iColEnd + 1)
	{
		DeleteRangedCol(iColEnd + 1, temp);
	}
	return true;
}

// Delete a range of columns
bool MFCExcelFile::DeleteRangedCol(long iColStart, long iColEnd)
{
	if (iColEnd < iColStart)
	{
		return false;
	}

	long iRangeRow = GetRowCount();

	CString temp1, temp2;
	temp1 = GetCellName(1, iColStart);
	temp2 = GetCellName(iRangeRow, iColEnd);
	CRange range = excel_work_sheet_.get_Range(_variant_t(temp1), _variant_t(temp2));
	range.Delete(COleVariant((long)xlShiftToLeft));
	range.ReleaseDispatch();
	// Ask for refresh memory copy at next reading function.
	if (!isUpdate && already_preload_)
	{
		PreLoadSheet();
	}
	return true;
}

// Delete a range of rows
bool MFCExcelFile::DeleteRangedRow(long iRowStart, long iRowEnd)
{
	if (iRowEnd < iRowStart)
	{
		return false;
	}

	long iRangeCol = GetColumnCount();

	CString temp1, temp2;
	temp1 = GetCellName(iRowStart, 1);
	temp2 = GetCellName(iRowEnd, iRangeCol);
	CRange range = excel_work_sheet_.get_Range(_variant_t(temp1), _variant_t(temp2));
	range.Delete(COleVariant((long)xlShiftUp));
	range.ReleaseDispatch();
	// Ask for refresh memory copy at next reading function.
	if (!isUpdate && already_preload_)
	{
		PreLoadSheet();
	}
	return true;
}

// 
void MFCExcelFile::ShowInExcel(BOOL bShow)
{
	excel_application_.put_Visible(bShow);
	excel_application_.put_UserControl(bShow);
}

// return the name of excel
CString MFCExcelFile::GetOpenedFileName()
{
	return open_excel_file_;
}

// return the name of current sheet
CString MFCExcelFile::GetLoadSheetName()
{
	return excel_work_sheet_.get_Name();
}

// return the name of column, 27->AA, 3->C etc
TCHAR *MFCExcelFile::GetColumnName(long icolumn)
{
	static TCHAR column_name[64];
	size_t str_len = 0;

	while (icolumn > 0)
	{
		int num_data = icolumn % 26;
		icolumn /= 26;
		if (num_data == 0)
		{
			num_data = 26;
			icolumn--;
		}
#ifdef UNICODE
		column_name[str_len] = (TCHAR)((num_data - 1) + L'A');
#else
		column_name[str_len] = (TCHAR)((num_data - 1) + 'A');
#endif
		str_len++;
	}
#ifdef UNICODE
	column_name[str_len] = L'\0';
#else
	column_name[str_len] = '\0';
#endif

	// reverse current string
#ifdef UNICODE
	_wcsrev(column_name);
#else
	_strrev(column_name);
#endif

	return column_name;
}

// Get current name of the cell
// (1, 27)->AA1, (2, 3)->C2 etc
CString MFCExcelFile::GetCellName(long iRow, long iCol)
{
#ifdef UNICODE
	if (iRow <= 0 || iCol <= 0)
	{
		return L"";
	}
	CString ret = L"";
	ret.Format(L"%ld", iRow);
#else
	if (iRow <= 0 || iCol <= 0)
	{
		return "";
	}
	CString ret = "";
	ret.Format("%ld", iRow);
#endif
	ret = MFCExcelFile::GetColumnName(iCol) + ret;
	return ret;
}

// preload sheet
void MFCExcelFile::PreLoadSheet()
{
	// clear update status and refresh memory content
	isUpdate = false;

	CRange used_range;

	used_range = this->excel_work_sheet_.get_UsedRange();

	VARIANT ret_ary = used_range.get_Value2();
	if (!(ret_ary.vt & VT_ARRAY))
	{
		return;
	}
	// clear the preloaded ole save array this first
	ole_safe_array_.DestroyData();
	ole_safe_array_.Attach(ret_ary);
}

// remain path untouched if the path is absolute
void MFCExcelFile::GetRelativePathIfNot(CString &path)
{
	// replace / to \\ to compatible with default way of path
	// clear unusable splashes
	// to avoid such problem "\\1.xlsx" "\\\\\\\\1.xlsx" .etc
#ifdef UNICODE
	path.Replace(L'/', L'\\');
	while (path.Find(L'\\') == 0)
#else
	path.Replace('/', '\\');
	while (path.Find('\\') == 0)
#endif
	{
		path = path.Right(path.GetLength() - 1);
	}
	// test if the path is absolute
	if (PathIsRelative(path))
	{
		// the path is not absolute, add current program path.
		CString exePathCString;
		TCHAR exePath[MAX_PATH];
		memset(exePath, 0, MAX_PATH);
		GetModuleFileName(NULL, exePath, MAX_PATH);
		exePathCString = exePath;
#ifdef UNICODE
		exePathCString = exePathCString.Left(exePathCString.ReverseFind(L'\\') + 1);
#else
		exePathCString = exePathCString.Left(exePathCString.ReverseFind('\\') + 1);
#endif
		path = exePathCString + path;
	}
}


#if 0
void demoExcelOperationFunction()
{
	//MFCExcelSingletonAdmin通过引用计数自动管理MFCExcelFile对象。
	//通过它，你不用担心它会造成内存驻留残余的Excel App。
	MFCExcelSingletonAdmin test;
	auto excel = test.GetInstance();
	if (excel == nullptr)
	{
		//你的计算机没有安装Excel或WPS
		return;
	}
	excel->OpenExcelFile(L"1.xlsx");//打开程序所在目录的1.xlsx，如果没有则创建一个新的在内存中，保存生效。
	excel->LoadSheet(L"Sheet1", TRUE);//预加载模式加载Sheet1，如没有这个Sheet会自动创建一个空的。
	excel->SetCellString(1, 1, L"Hello from Visual C++ Excel OLE！");//向A1单元格填写这串字符串
	std::vector<CString> inputElements;
	for (int i = 0; i < 4; i++)
		inputElements.push_back(L"Hello World!");
	excel->SetRangeCellString(3, 4, 1, 2, inputElements);//向A3:B4写入"Hello World"字符串
	excel->CloseExcelFile(TRUE);//相当于Save(); CloseExcelFile(FALSE);
}//作用域结束的时候，如果发现MFCExcelSingletonAdmin引用计数为0（也就是所有MFCExcelSingletonAdmin对象全都析构了），将会自动释放Excel App和关闭Excel文件。
#endif