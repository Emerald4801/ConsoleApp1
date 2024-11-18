using Microsoft.Office.Interop.Excel;
using System.Data.Common;
using System.Runtime.InteropServices;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Ex = Microsoft.Office.Interop.Excel;
namespace ExTest
{
    internal class ExcelHelper : IDisposable
    {
        private Ex.Application _app;

        private Ex.Workbook _workbook;

        private Ex.Workbooks _workbooks;

        private Ex.Worksheet _worksheet;

        private string _path;


        public ExcelHelper()
        {
            _app = new Ex.Application();
        }


        public bool Open(string path, int numOfSheet = 1)
        {
            try
            {
                _path = path;
                _workbooks = _app.Workbooks;
                if (!File.Exists(path))
                {
                    _workbook = _workbooks.Add();
                }
                else
                {
                    _workbook = _workbooks.Open(path);
                }

                OpenWorksheet(numOfSheet);

                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("ашибка: " + e);
            }
            return false;
        }


        private void OpenWorksheet(int numOfSheet)
        {
            try
            {
                while (_workbook.Worksheets.Count < numOfSheet)
                    _workbook.Worksheets.Add();
                _worksheet = _workbook.Worksheets[numOfSheet];
            }
            catch (Exception e)
            {
                Console.WriteLine("ашибка: " + e);
            }

        }


        #region Setters

        internal bool Set(int row, int column, object data)
        {
            try
            {
                _worksheet.Cells[row, column].Value = data;
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("ашибка " + e);
            }
            return false;
        }

        //internal bool SetRow(int rowIdx, Row row, int StartCol = 1)
        //{
        //    try
        //    {
        //        for (int i = StartCol; i < StartCol + row._Length; i++)
        //        {
        //            Set(rowIdx + 1, i, row.GetValue(i).Value);
        //        }
        //        return true;
        //    }
        //    catch (Exception e)
        //    {
        //        Console.WriteLine(e);
        //    }
        //    return false;
        //}

        #endregion


        #region Getters

        public string Get(int row, int column)
        {
            return _worksheet.Cells[row, column].Value;
        }

        //internal Row GetRow(int index, int StartCol, int EndCol)
        //{
        //    var row = new Row(EndCol - StartCol);

        //    for (int i = StartCol; i < EndCol; i++)
        //    {
        //        row.SetValue(i - StartCol, new CellValue(Get(index, i), CellType.String));
        //    }
        //    return row;
        //}

        //public List<string> GetRow(int row)
        //{
        //    var a = _worksheet.Columns;
        //    foreach (var item in a)
        //    {
                
        //    }
        //    int last_row = _worksheet.Cells.Find("*", _worksheet.Cells[1, 1], Ex.XlFindLookIn.xlFormulas, Ex.XlLookAt.xlPart,
        //        Ex.XlSearchOrder.xlByRows, Ex.XlSearchDirection.xlPrevious);

        //    List<string> values = new();
        //    int column = 1;
        //    do
        //    {
        //        values.Add(_worksheet.Cells[row, column].Value);
        //        column++;
        //    } while (_worksheet.Cells[row, column] != last_row);

        //    return values;
        //}

        #endregion


        #region Design

        internal bool Merge(string firstColumn, int firstRow, string secondColumn, int secondRow)
        {
            try
            {
                string Cell1 = $"{firstColumn}{firstRow}";
                string Cell2 = $"{secondColumn}{secondRow}";
                ((Ex.Worksheet)_app.ActiveSheet).Range[Cell1, Cell2].Merge(Type.Missing);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("ашибка " + e);
            }
            return false;
        }

        internal bool Merge(string Cell1, string Cell2)
        {
            try
            {
                ((Ex.Worksheet)_app.ActiveSheet).Range[Cell1, Cell2].Merge(Type.Missing);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("ашибка " + e);
            }
            return false;
        }

        internal bool AutoFit(string StartColumn, string EndColumn)
        {
            try
            {
                //((Excel.Worksheet)_excel.ActiveSheet).Range[$"{StartColumn}1:{EndColumn}1"].AutoFit;
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("ашибка " + e);
            }
            return false;
        }

        #endregion

        
        public void CreateDropDownList(int row, int column, List<string> variants)
        {
            var flatList = string.Join(",", variants.ToArray());

            var cell = (Ex.Range)_worksheet.Cells[row, column];
            cell.Validation.Delete();
            cell.Validation.Add(
               XlDVType.xlValidateList,
               XlDVAlertStyle.xlValidAlertInformation,
               XlFormatConditionOperator.xlBetween,
               flatList,
               Type.Missing);

            cell.Validation.IgnoreBlank = true;
            cell.Validation.InCellDropdown = true;
        }

        public void Save()
        {
            try
            {
                if (!File.Exists(_path))
                    _workbook.SaveAs(_path);
                else
                    _workbook.Save();
            }
            catch (Exception e)
            {
                Console.WriteLine("ашибка " + e);
            }
        }


        public void Dispose()
        {
            //Освобождение _worksheet
            if (_worksheet != null)
            {
                while (Marshal.ReleaseComObject(_worksheet) != 0) { }
                _worksheet = null;
            }

            //Освобождение _workbook
            if (_workbook != null)
            {
                _workbook.Close();
                while (Marshal.ReleaseComObject(_workbook) != 0) { }
                _worksheet = null;
            }

            //Освобождение _workbooks
            if (_workbooks != null)
            {
                _workbooks.Close();
                while (Marshal.ReleaseComObject(_workbooks) != 0) { }
                _workbooks = null;
            }

            //Освобождение _app
            if (_app != null)
            {
                _app.Quit();
                while (Marshal.ReleaseComObject(_app) != 0) { }
                _app = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}