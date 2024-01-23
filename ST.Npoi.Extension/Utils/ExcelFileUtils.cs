///************************************************************************************
///   Author:Tony Stark
///   CreateTime:2023/3/30 11:44:05
///   Mail:2609639898@qq.com
///   GitHub:https://github.com/getup700
///
///   Description:
///
///************************************************************************************

using System;
using System.IO;
using System.Data;
using System.Linq;
using NPOI.SS.UserModel;
using System.Diagnostics;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Windows.Forms;
using ST.Net.Extension.Utils;
using System.Collections.Generic;
using ST.Net.Extension.Extensions;
using ST.NPOI.Extension.Extensions;
using NPOI.POIFS.FileSystem;
using NPOI.POIFS.Crypt;
using NPOI.HSSF.Record.Crypto;

namespace ST.NPOI.Extension.Utils;

public static class ExcelFileUtils
{
    public static IWorkbook OpenExcel(string fullFileName, string password = null)
    {
        if (!File.Exists(fullFileName))
        {
            throw new Exception($"当前文件不存在，请检查文件路径后重试。\n{fullFileName}");
        }
        FileStream fileStream = null;
        try
        {
            fileStream = new FileStream(fullFileName, FileMode.Open, FileAccess.Read);
        }
        catch (Exception)
        {
            throw new Exception($"当前文档被其他程序占用，请关闭文档后重试。\n{fullFileName}");
        }
        IWorkbook workbook = null;
        string extension = Path.GetExtension(fullFileName);
        if (extension == ".xls")
        {
            workbook = new HSSFWorkbook(fileStream);
        }
        else if (extension == ".xlsx")
        {
            workbook = new XSSFWorkbook(fullFileName);
        }
        fileStream.Close();
        fileStream.Dispose();
        return workbook;
    }

    public static IWorkbook OpenCreateWorkbookDialog(string fileName, string fileExtension, out string fullFilePath, Action<FolderBrowserDialog> action = null)
    {
        IWorkbook workbook = null;
        var folderBrowserDialogSelectedPath = FileUtils.OpenFolderBrowserDialog(action);
        if (folderBrowserDialogSelectedPath == null)
        {
            fullFilePath = null;
            return workbook;
        }
        fullFilePath = Path.Combine(folderBrowserDialogSelectedPath, $"{fileName}{fileExtension}");

        if (File.Exists(fullFilePath))
        {
            var result = MessageBox.Show($"已存在[{fileName}{fileExtension}]，是否继续删除并继续?\n\n", "Export Tips", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                File.Delete(fullFilePath);
                workbook = CreateWorkbook(fullFilePath);
            }
            else if (result == DialogResult.No)
            {
                return workbook;
            }
        }
        else
        {
            workbook = CreateWorkbook(fullFilePath);
        }
        return workbook;
    }

    /// <summary>
    /// Open sheet with the specified name.If there is no name for sheet,open the first one.
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="sheetName"></param>
    /// <param name="action"></param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    /// <exception cref="ArgumentNullException"></exception>
    public static IWorkbook OpenWorkbookDialog(out ISheet sheet, string sheetName = null, Action<OpenFileDialog> action = null)
    {
        string fullFileName = string.Empty;
        OpenFileDialog openFileDialog = new OpenFileDialog()
        {
            Filter = "Excel文件|*.xls;*.xlsx",
            RestoreDirectory = true,
            FilterIndex = 1
        };
        action?.Invoke(openFileDialog);
        if (openFileDialog.ShowDialog() == DialogResult.OK)
        {
            fullFileName = openFileDialog.FileName;
        }
        if (fullFileName == "")
        {
            sheet = null;
            return null;
        }

        var workbook = OpenExcel(fullFileName);

        if (sheetName == null)
        {
            sheet = workbook.GetSheetAt(0);
        }
        else
        {
            sheet = workbook.GetSheet(sheetName);
        }

        if (sheet == null)
        {
            throw new Exception($"未找到指定工作表，检查导入的工作表后重试");
        }

        return workbook;
    }

    public static IWorkbook CreateWorkbook(string excelPath)
    {
        IWorkbook Workbook = null;
        var extension = Path.GetExtension(excelPath);
        if (extension.Equals(".xls"))
        {
            Workbook = new HSSFWorkbook();
        }
        else if (extension.Equals(".xlsx"))
        {
            Workbook = new XSSFWorkbook();
        }
        else
        {
            throw new ArgumentException(nameof(extension), "input extension is invalid");
        }
        return Workbook;
    }

    public static IWorkbook ReadWorkbook(string excelPath)
    {
        IWorkbook Workbook = null;
        var extension = Path.GetExtension(excelPath);
        FileStream fileStream = null;
        try
        {
            fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.ReadWrite);
        }
        catch (Exception)
        {
            throw new Exception($"当前文档被其他程序占用，请关闭文档后重试。\n{excelPath}");
        }
        if (extension.Equals(".xls"))
        {
            Workbook = new HSSFWorkbook(fileStream);
        }
        else if (extension.Equals(".xlsx"))
        {
            Workbook = new XSSFWorkbook(fileStream);
        }
        else
        {
            throw new ArgumentException(nameof(extension), "input extension is invalid");
        }
        fileStream.Close();
        return Workbook;
    }

    public static void OpenFileExplorerWindow(string filePath, bool openFile = false)
    {
        if (!openFile)
        {
            filePath = Path.GetDirectoryName(filePath);
        }
        var fileExplorer = new ProcessStartInfo(filePath);
        Process.Start(fileExplorer);
    }

    public static void DataTableToSheet(DataTable dataTable, ISheet sheet, bool isSetColumnNames = true)
    {
        if (dataTable == null)
        {
            return;
        }
        var count = dataTable.Columns.Count;
        if (count == 0)
        {
            return;
        }
        var startIndex = 0;
        var endCount = dataTable.Rows.Count;
        if (isSetColumnNames)
        {
            var names = dataTable.GetColumnNames().ToList();
            var titleRow = sheet.CreateRow(0);
            for (var i = 0; i < names.Count; i++)
            {
                var name = names[i];
                titleRow.CreateCell(i).SetCellValue(name);
            }

            for (int i = 1; i < dataTable.Rows.Count + 1; i++)
            {
                IRow dataRow = sheet.CreateRow(i);
                for (int j = 0; j < count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dataTable.Rows[i - 1][j].ToString());
                }
            }
        }
        else
        {
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i);
                for (int j = 0; j < count; j++)
                {
                    dataRow.CreateCell(j).SetCellValue(dataTable.Rows[i][j].ToString());
                }
            }
        }

    }

    public static void DataTableToSheet2D(DataTable dataTable, ISheet sheet)
    {
        if (dataTable == null || dataTable.Columns.Count == 0)
        {
            return;
        }
        var firstRow = sheet.CreateRow(0);
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
            firstRow.CreateCell(i).SetCellValue(dataTable.Columns[i].ColumnName);
        }

        for (int i = 0; i < dataTable.Rows.Count; i++)
        {
            var row = sheet.CreateRow(i + 1);
            for (int j = 0; j < dataTable.Columns.Count; j++)
            {
                var value = dataTable.Rows[i][j].ToString();
                row.CreateCell(j).SetCellValue(value);
            }

        }
    }

    public static DataTable SheetToDataTable(ISheet sheet, DataTable dataTable = null)
    {
        if (sheet == null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }
        dataTable ??= new DataTable();
        var columnsCount = dataTable.Columns.Count;
        var headerRow = sheet.GetRows().FirstOrDefault();
        var headerCells = headerRow?.GetCells();
        var contentRows = sheet.GetRows();
        contentRows.RemoveAt(0);
        if (columnsCount == 0)
        {
            if (headerCells == null)
            {
                return null;
            }
            foreach (var item in headerCells)
            {
                var title = item.GetCellValue();
                dataTable.Columns.Add(title);
            }
        }
        //var titles = new List<string>();
        //for (int i = 0; i < dataTable.Columns.Count; i++)
        //{
        //    var title = dataTable.Columns[i].ColumnName;
        //    titles.Add(title);
        //}

        foreach (var currentRow in contentRows)
        {
            var cells = currentRow.GetCells();
            var items = cells.Count - headerCells.Count;
            if (items > 0)
            {
                cells.RemoveRange(cells.Count - items, items);
            }
            var values = new List<string>();
            foreach (var currentCell in cells)
            {
                var cellValue = currentCell.GetCellValue();
                values.Add(cellValue);
            }
            var array = values.ToArray();
            dataTable.Rows.Add(array);
        }
        return dataTable;
    }

    public static void ExportToExcel(this DataTable dataTable, IWorkbook workbook = null, string bookName = "Export")
    {
        if (dataTable == null)
        {
            throw new ArgumentNullException(nameof(dataTable));
        }
        workbook = ExcelFileUtils.OpenCreateWorkbookDialog(bookName, ".xls", out string filePath);
        if (workbook == null)
        {
            return;
        }
        var headerStyle = workbook.CommonHeaderStyle();
        var contentStyle = workbook.CommonTextStyle();

        var sheet = workbook.GetSheet(dataTable.TableName);
        sheet ??= workbook?.CreateSheet(dataTable.TableName);
        var headerRow = sheet.GetRow(0);
        if (headerRow == null)
        {
            headerRow = sheet.CreateRow(0);
            var excelColumnName = new List<string>();
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                excelColumnName.Add(dataTable.Columns[i].ColumnName);
            }
            headerRow.AssignValuesInOrder(excelColumnName);
            headerRow.RowStyle = headerStyle;
        }
        headerRow.SetRowStyle(workbook.CommonHeaderStyle());

        for (int i = 0; i < dataTable.Rows.Count; i++)
        {
            var currentRow = sheet.CreateRow(i + 1);
            for (int j = 0; j < dataTable.Columns.Count; j++)
            {
                var cell = currentRow.CreateCell(j);
                cell.SetCellValue(dataTable.Rows[i][j].ToString());
                cell.CellStyle = contentStyle;
            }
        }

        using (var fileStream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
        {
            workbook.Write(fileStream);
            fileStream.Close();
        }

        var dialogResult = MessageBox.Show("导出成功，是否打开文件所在目录", "", MessageBoxButtons.YesNo);
        if (dialogResult == DialogResult.Yes)
        {
            ExcelFileUtils.OpenFileExplorerWindow(filePath);
        }

    }
}
