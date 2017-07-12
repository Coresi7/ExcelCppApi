/*
 * MFCExcel.h
 * Author: hearstzhang (hearstzhang@tencent.com)
 * A safely MFCExcelFile which added many useful functions
 * and mechanisms to make it much faster compared to the gyssoft's
 * original version of MFCExcelFile.
 * It will finish read and write for millions of cells in seconds, if you use it well.
 * Preload excel sheet is recommended, unless memory problem.
 * Your permission for copy and redistribute is granted as long as you keep
 * these copyright announcements.
 * Requires Microsoft Excel or Kingsoft WPS installed. Compatible with
 * MBCS and Unicode character set.
 * Please let me know if you have some question about this.
 */

/*
 * v1.0.0 initial version, compatible to both MBCS and UNICODE
 * v1.0.1 add two batch set cell string function
 * v1.0.2 fixed bugs for pretential crash at preload mode
 * v1.0.3 fixed bugs for set cell string function
 * v1.0.4 fixed bugs for pretential crash at init and release excel application
 * v1.0.5 add two batch set cell string function, which could clear rest rows and cols
 * v1.0.6 add exception handlers for multiple function which could make it stable
 * v1.0.7 fixed bugs for batch set cell string function.
 * v1.0.8 increase performance of batch set cell string function.
 * v1.0.9 fixed bugs for pretential crash.
 * v1.0.10 add delete row and col function
 * v1.0.11 delete row and col function does not require preload
 * v1.0.12 batch set cell and delete rest row and col function does not require preload
 * v1.0.13 fixed bugs for delete row and col function
 * v1.0.14 add get cell name function
 * v1.0.15 increase performance of batch set cell string
 * v1.0.16 batch set cell string does not do nothing if vector is not big enough.fill rest cells as blank instead.
 * v1.0.17 get cell name function does not requries preload.
 * v1.0.18 increase performance for reading for preload mode
 * v1.0.19 increase performance for set cell value one by one.
 * v1.0.20 loadsheet function could automatically create a blank sheet if target sheet does not exist.
 * v1.0.21 openexcelfile function could automatically create a excel file in memory for operation if target excel file does not exist.But this file will save to disk only if you save it.
 * v1.0.22 fixed a bug which delete ranged row may not take effect in some circumstances.
 * v1.0.23 add save function, could save the excel to disk and reload it immediately to refresh status.
 * v1.0.24 move delete xls method enum to source file to avoid global variable define.
 * v1.0.25 add singleton adminstration class llusionExcelSingletonAdmin for MFCExcelFile.But compatible with old codes.
 * v1.0.26 make modification to MFCExcelFile class constructor and destructor to compatible with llusionExcelSingletonAdmin class.
 * v1.0.27 fix a bug which may cause argument error if preload an empty sheet.
 * v1.0.28 PreloadSheet is a friend function now, which means, no allowed to use outside the class.
 * v1.0.29 Solve pretential memory leak problem.
 * v1.0.30 Preload mode does not require save function to refresh memory data.This option will be done by any read functions automatically.
 * v1.0.31 Update copyright information.
 * v1.0.32 Support absolate and not absolate path.
 * v1.0.33 Add demo excel operation function. See MFCExcel.cpp for details.
 */

// -----------------------MFCExcel.h------------------------
#pragma once

#include "CApplication.h"
#include "CWorkbook.h"
#include "CWorkbooks.h"
#include "CWorksheet.h"
#include "CWorksheets.h"
#include "CRange.h"
#include <vector>

// 前置声明
class MFCExcelFile;

// 用于管理MFCExcel的单例类
// 建议你使用这个东西来管理MFCExcelFile类。但你需要注意，这个类本身不是单例模式的类，
// 当引用计数为0的时候，自动析构掉MFCExcelFile对象
// 保证你的工具在退出的时候，可以正确释放掉Excel App
class MFCExcelSingletonAdmin
{
protected:
	// MFCExcelFile单例对象。
	static MFCExcelFile * m_pInstance;
	// 引用计数器
	static unsigned int m_nReference;

public:
	// 构造和析构函数
	// 设为public是因为这个类本身不是单例模式的类
	MFCExcelSingletonAdmin();
	~MFCExcelSingletonAdmin();
	// 如果目标计算机没有安装Excel或WPS则返回nullptr
	// 否则返回一个正常的MFCExcelFile对象。
	MFCExcelFile * GetInstance();
};

// 建议使用它的单例模式管理类MFCExcelSingletonAdmin管理
// 因为一个程序只能调用一个excel进程
// 也就是说，如果你想读取另一个excel文件，你就要关闭当前文件，但不必ReleaseExcel。
// index从1而不是从0开始。例如col = 1, row = 1对应单元格A1，sheetindex = 1对应第一个sheet等。
class MFCExcelFile
{
	// 单例模式管理类友元
	friend class MFCExcelSingletonAdmin;
protected:
	// 预加载函数
	// 在预加载模式下，任何对Excel的写入都会在下一次读取内容的操作的时候
	// 自动调用它刷新内存副本。
	void PreLoadSheet();
	// 自动把相对路径转换成绝对路径
	// 自动把\转换成//以和MFC兼容。
	// 如果本身就是绝对路径，则不会对path做任何改变，除了转换\
	// 否则把输入的path转换成绝对路径，以当前程序所在路径为起始
	void GetRelativePathIfNot(CString &path);

public:
	// 构造函数和析构函数
	MFCExcelFile();
	virtual ~MFCExcelFile();

protected:
	// 打开的EXCEL文件名称
	CString open_excel_file_;
	// EXCEL BOOK集合，（多个文件时）
	CWorkbooks excel_books_;
	// 当前使用的BOOK，当前处理的文件
	CWorkbook excel_work_book_;
	// EXCEL的sheets集合
	CWorksheets excel_sheets_;
	// 当前使用sheet
	CWorksheet excel_work_sheet_;
	// 当前的操作区域
	CRange excel_current_range_;
	// 是否已经预加载了某个sheet的数据
	BOOL already_preload_;
	// 预加载模式下使用OLE数组存储Excel内容
	COleSafeArray ole_safe_array_;
	// 维护Excel服务是否已经被初始化或者释放。
	static bool isInited;
	static bool isReleased;
	// EXCEL的进程实例
	static CApplication excel_application_;
	// 预加载模式下，用它标记是否需要刷新预加载内容，为true则需要刷新
	bool isUpdate;

public:
	// 调用Excel打开当前文档
	void ShowInExcel(BOOL bShow = TRUE);
	// 检查一个CELL是否是字符串
	BOOL IsCellString(long iRow, long iColumn);
	// 检查一个CELL是否是数值
	BOOL IsCellInt(long iRow, long iColumn);
	// 得到一个CELL的String，如果在预加载的情况下，尝试访问越界数据，则会返回空白值。
	CString GetCellString(long iRow, long iColumn);
	// 得到整数，如果在预加载的情况下，尝试访问越界数据，则会返回0
	int GetCellInt(long iRow, long iColumn);
	// 得到double的数据，如果在预加载的情况下，尝试访问越界数据，则会返回0.0
	double GetCellDouble(long iRow, long iColumn);
	// 取得行的总数
	long GetRowCount();
	// 取得列的总数
	long GetColumnCount();
	// 加载Sheet以供使用，如果没有，则自动创建一个。如果你录入的Sheet序号高于
	// 当前已有的Sheet数量，则会自动在最后创建Sheet直到到达你输入的index。
	// 注意index从1开始，如果输入0则直接返回FALSE.
	// Preload为TRUE是预加载模式，建议如果读取内容，尽可能使用该模式，除非Sheet非常大使得预加载超级消耗计算机资源
	BOOL LoadSheet(long table_index, BOOL pre_load = TRUE);
	// 通过名称使用某个sheet，
	// 如果没有这个名字命名的Sheet，则根据你提供的名称自动创建一个放在最后面。
	// Preload是预加载模式，建议如果读取内容，尽可能使用该模式，除非Sheet非常大使得预加载超级消耗计算机资源。
	BOOL LoadSheet(const TCHAR* sheet, BOOL pre_load = TRUE);
	// 通过序号取得某个Sheet的名称
	CString GetSheetName(long table_index);
	// 得到Sheet的总数
	long GetSheetCount();
	// 打开文件，如果对应位置没有这个excel文件，则自动在内存创建一个空白的excel文件，保存生效，返回FALSE则打开失败
	// 支持绝对和相对路径。如果是相对路径，则从这个程序所在的当前路径为起始。
	// 支持正斜杠和反斜杠的路径（例如"..\\1.xlsx"、"../1.xlsx"）
	BOOL OpenExcelFile(const TCHAR * file_name);
	// 关闭打开的Excel 文件。如果if_save为TRUE则保存文件
	void CloseExcelFile(BOOL if_save = FALSE);
	// 另存为一个EXCEL文件
	// 支持绝对和相对路径。如果是相对路径，则从这个程序所在的当前路径为起始。
	// 支持正斜杠和反斜杠的路径（例如"..\\1.xlsx"、"../1.xlsx"）
	void SaveasXLSFile(const CString &xls_file);
	// 立刻保存到磁盘，同时会关闭再打开当前sheet以刷新状态。
	void Save();
	// 取得打开文件的名称
	CString GetOpenedFileName();
	// 取得打开sheet的名称
	CString GetLoadSheetName();
	// 写入一个CELL一个int，尽可能不要用于批量写入因为该函数代价较高
	void SetCellInt(long irow, long icolumn, int new_int);
	// 写入一个CELL一个string，尽可能不要用于批量写入因为该函数代价较高
	void SetCellString(long irow, long icolumn, CString new_string);
	// 写入一定范围的string，例如从A2写到D10，则传入2, 10, 1, 4。其中传入的vector负责存储数据，并且遍历方式是new_string[irow * colsize + icol]
	// 如果iRowStart <= 0或iColStart <= 0，则这个函数什么都不做直接返回false。
	// vector不够大没关系，会将超出范围的数据写成空白，不过要小心如果你范围设置的不正确则会用空白错误的覆盖掉对应内容。
	// 这个是相当高效的函数，比如说，写入千万量级的数据，在i7-4790，16GDDR3 32位Excel 2016下只需要数秒。不要求一定预读取。
	// 但如果真的是千万级数据。。。你还是分批写入吧，容易爆内存
	bool SetRangeCellString(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
		const std::vector<CString>& new_string);
	// 这个函数会把超出写入范围的行清空。但不会清空其写入范围之前的内容。用法参见SetRangeCellString，这里略去
	// 例如，从A2写到D10，它会清空第10行以后的数据。你要是写工具并且涉及全表写入，你会知道这个函数会在哪里派上用场的。
	// 如果第10行及以后为空，则什么操作都不做。
	bool SetRangeCellStringAndClearRestRows(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
		const std::vector<CString>& new_string);
	// 这个函数会把超出写入范围的列清空。但不会清空其写入范围之前的内容。用法参见SetRangeCellString，这里略去
	// 例如，从A2写到D10，它会清空第4列以后的数据。你要是写工具并且涉及全表写入，你会知道这个函数会在哪里派上用场的。
	// 如果第4列及以后为空，则什么操作都不做。
	bool SetRangeCellStringAndClearRestCols(DWORD iRowStart, DWORD iRowEnd, DWORD iColStart, DWORD iColEnd,
		const std::vector<CString>& new_string);
	// 删除整列，并且后面的列前移。但是注意只有保存才会生效。
	bool DeleteRangedCol(long iColStart, long iColEnd);
	// 删除整行，并且后面的行上移。但是注意只有保存才会生效。
	bool DeleteRangedRow(long iRowStart, long iRowEnd);

public:
	// 初始化EXCEL OLE，并且会在内存当中启动一个Excel服务器。注意一个程序只能允许启动一个Excel App
	// 这也是为什么我为它写了个单例模式管理类。
	static BOOL InitExcel();
	// 释放EXCEL的 OLE，关闭Excel服务器。
	// 释放了之后下次任何MFCExcelFile对象想要使用Excel必须要Init。
	static void ReleaseExcel();
	// 取得单元格名称，例如 (1,1)对应A1
	// 如果输入不合法的数值（例如0，0）则直接返回空字符串。
	static CString GetCellName(long iRow, long iCol);
	// 取得列的名称，比如27->AA
	static TCHAR *GetColumnName(long iColumn);
};
