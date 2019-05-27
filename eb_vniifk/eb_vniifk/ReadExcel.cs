using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Linq;
using Microsoft.Office.Interop;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Xml;
using System.IO;
using System.Data.OleDb;
using System.Security;
using System.Threading.Tasks;



namespace eb_vniifk
{
    public class ReadExcel
    {
        private string path = "";
        private string test = "";
        private DataTable internalTable;
        private _Excel.Application xlApp; //приложение ексель
        private _Excel.Workbook xlWorkBook; //книга
        private _Excel.Worksheet xlWorkSheet; //лист
        private _Excel.Range xlRange; // ячейка
        private object misValue = System.Reflection.Missing.Value;
        private DataRow row;
        private Files Files;
        public ReadExcel(string _path = "")
        {
            path = _path;
            Files = new eb_vniifk.Files();
            OpenXlApp();
            xlApp.Visible = true;
            //TakeIntervalFromExcel(path);
            TakeIntervalOleDb(path);
            test = "1";

        }
        private void OpenXlApp()
        {
            xlApp = new _Excel.Application();
        }
        private void TakeIntervalFromExcel(string path)
        {
            string sheet = "";
            //xlWorkBook = xlApp.Workbooks;
            if (xlWorkBook == null || path.Length > 0)
            {
                try
                {
                    //xlWorkBook = xlApp.Workbooks.Open(path, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                    xlWorkBook = xlApp.Workbooks.Open(path, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    

                    //xlWorkBook = xlApp.Workbooks.Open(path);
                }
                catch (System.Exception ex)
                {
                    test = "2";
                }
            }
        }
        private void OpenExcelSheet(string sheet = "", int l = -1, string path = "")
        {
            //try
            //{

            int error = 0;
            if (xlWorkBook == null) TakeIntervalFromExcel(path);

            try
            {
                for (int i = 0; i < xlWorkBook.Worksheets.Count; i++)
                {
                    int index = i + 1;
                    _Excel.Worksheet testSheet = (_Excel.Worksheet)xlWorkBook.Worksheets[index];
                    if (sheet == testSheet.Name) l = index;
                }
                if (l > 0)
                {
                    xlWorkSheet = (_Excel.Worksheet)xlWorkBook.Worksheets[l];
                }
            }
            catch (System.Exception err)
            {
                error = 1;
            }
            if (error <= 0)
            {
                //xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets[@strExcelTable];
                try
                {
                    xlRange = xlWorkSheet.UsedRange;
                }
                catch (System.Exception ex)
                {
                    error = 1;
                    return;
                }
                if (xlRange != null || xlRange.Rows.Count > 1)
                {
                    internalTable = new DataTable();
                    object[,] valueArray = (object[,])xlRange.get_Value(_Excel.XlRangeValueDataType.xlRangeValueDefault);
                    for (int j = 0; j < xlRange.Columns.Count; j++)
                    {
                        internalTable.Columns.Add((j + 1).ToString(), typeof(string));
                    }
                    for (int i = 0; i < xlRange.Rows.Count; i++)
                    {
                        row = internalTable.NewRow();
                        for (int j = 0; j < xlRange.Columns.Count; j++)
                        {
                            if (valueArray != null && valueArray.GetValue(i + 1, j + 1) != null)
                            {
                                row[j] = valueArray.GetValue(i + 1, j + 1).ToString();
                            }
                        }
                        internalTable.Rows.Add(row);
                    }
                }
            }
        }
        private void TakeIntervalOleDb(string path, string sheet = "")
        {
            Files.ExcelOleDbConn = new OleDbConnection(Files.connectionString(path));
            Files.ExcelOleDbConn.Open();
            Files.ExcelAdapter = new OleDbDataAdapter("Select * from [" + sheet + "];", Files.ExcelOleDbConn);
            Files.ExcelOleDbConn.Close();
            internalTable = null;
            internalTable = new DataTable();
            try
            {
                Files.ExcelAdapter.Fill(internalTable);
            }
            catch (System.Exception err)
            {
            }
        }
        private void CloseExcelBook()
        {
            try
            {
                xlWorkBook.Close(true, misValue, misValue);
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
            }
            catch (System.Exception err)
            {
            }
        }
        private void CloseExcelAppl()
        {
            /*
             * Метод для завершения работы с файлом и его закрытие
             * */
            try
            {
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (System.Exception err)
            {
            }
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                string err = "Unable to release the Object " + ex.ToString();
            }
            finally
            {
                GC.Collect();
            }
        }
    }
    public class Files
    {
        public Files()
        {
        }
        /*
         * Этот класс объектов отвечает за присоединение к файлу
         * эксель с помощью адаптера оле дб. Этот необходимо для более быстрого
         * считывания файлов и последующая их запись в базу.
         * */
        //провайдер
        private string strConnectionProv8 = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
        private string strConnectionProv12 = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
        //тип соединения
        private string strConnectionType8 = ";Extended Properties='Excel 8.0'";
        private string strConnectionType12 = ";OLE DB Services=-1;Extended Properties='Excel 12.0 xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
        private string ll = "";
        private string ll1 = ";HDR=NO;IMEX=0";
        private string ll2 = ";HDR=Yes;IMEX=1";
        private string ll3 = ";HDR=Yes;IMEX=0";
        private string ll4 = ";IMEX=0";
        private string ll5 = ";HDR=NO";
        private string ll6 = ";HDR=NO;IMEX=1";
        private string ll7 = ";HDR=YES";
        //Можно использовать, если количество строк менее 65536
        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes'
        //Если строк больше 65536
        //Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties="Excel 12.0 Xml;HDR=YES";
        //Переменные для базы и табличных значений
        private OleDbConnection excelOleDbConn; //подключение через оледб
        private OleDbDataAdapter excelAdapter; //адаптер для подключения
        private DataTable dtExcel = new DataTable(); //таблица под эксель лист
        private DataSet dsExcel = new DataSet(); //коллекция таблиц
        private DataTable dtExcelMeta = new DataTable(); //Структура документа Ексель
        //Глобальные переменные
        public string StrConnectionProv8 { get { return strConnectionProv8; } }
        public string StrConnectionProv12 { get { return strConnectionProv12; } }
        public string StrConnectionType8 { get { return strConnectionType8; } }
        public string StrConnectionType12 { get { return strConnectionType12; } }
        public OleDbConnection ExcelOleDbConn { get { return excelOleDbConn; } set { excelOleDbConn = value; } }
        public OleDbDataAdapter ExcelAdapter { get { return excelAdapter; } set { excelAdapter = value; } }
        public DataTable DtExcel { get { return dtExcel; } set { dtExcel = value; } }
        public DataSet DsExcel { get { return dsExcel; } set { dsExcel = value; } }
        public DataTable DtExcelMeta { get { return dtExcelMeta; } set { dtExcelMeta = value; } }
        //
        //
        //
        public string connectionString(string path)
        {
            /*
             * Эта фунция возвращает строку подключения к екселю.
             * */
            string connect = @strConnectionProv8 + path + strConnectionType8;
            return connect;
        }

    }
}
