using System.Runtime.InteropServices;
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


        public string Get(int row, int col)
        {
            return _worksheet.Cells[row, col].Value;
        }


        public void Set(int row, int col, string value)
        {
            _worksheet.Cells[row, col].Value = value;
        }


        //public List<string> GetRow(int row)
        //{
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

        //public void SetRow(int row, int col, string value)
        //{
        //    _worksheet.Cells[row, col].Value = value;
        //}


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