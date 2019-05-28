using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Linq;
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
        private object misValue = System.Reflection.Missing.Value;
        private DataRow row;
        private Files Files;
        public ReadExcel(string _path = "")
        {
            path = _path;
            Files = new eb_vniifk.Files();
            //TakeIntervalFromExcel(path);
            TakeIntervalOleDb(path);
            //readExcel(path);
            test = "1";

        }
        private void readExcel(string path)
        {
            //C:\Users\Ilya\Google Диск\Projects\ВНИИФК\Загрузчик в ЭБ\Тест2.xls
            string connectionString = @"provider = Microsoft.ACE.OLEDB.12.0; 
                            data source = C:\Users\Ilya\Google Диск\Projects\ВНИИФК\Загрузчик в ЭБ\Тест2.xls; 
                            Extended Properties = 'Excel 12.0'";
           // connectionString = Files.connectionString(path);
            OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
            try
            {
                oleDbConnection.Open();
                //MessageBox.Show("Connection Successful");
                string err3 = "good";
            }
            catch (System.Exception err)
            {
                string err2 = err.ToString();
                //MessageBox.Show("Connection failed");

            }
        }
      private void TakeIntervalOleDb(string path, string sheet = "")
        {
            //path = path.Replace("\\", @"\");
            string connectionString = Files.connectionString(path);
            OleDbConnection oleDbConnection = new OleDbConnection(connectionString);
            oleDbConnection.Open();
            string err3 = "good";
            //DataTable dt = new DataTable();
            //Files.ExcelOleDbConn = new OleDbConnection(Files.connectionString(path));
            //Files.ExcelOleDbConn.Open();
            //Files.ExcelAdapter = new OleDbDataAdapter("Select * from [" + sheet + "];", Files.ExcelOleDbConn);
            //OleDbCommand comm = new OleDbCommand();
            //comm.CommandText = "Select * from [Лист1$]";
            //comm.Connection = Files.ExcelOleDbConn;
            //Files.ExcelAdapter = new OleDbDataAdapter("Select * from [Лист1]", Files.ExcelOleDbConn);
            //Files.ExcelAdapter.SelectCommand = comm;
            //Files.ExcelAdapter.Fill(dt);


            //Files.ExcelOleDbConn.Close();
            //internalTable = null;
            //internalTable = new DataTable();
            //try
            //{
            //    Files.ExcelAdapter.Fill(internalTable);
            //}
            //catch (System.Exception err)
            //{
            //}
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
        private string strConnectionProv12 = @"provider = Microsoft.ACE.OLEDB.12.0; data source = ";
        //тип соединения
        private string strConnectionType8 = ";Extended Properties='Excel 8.0'";
        //private string strConnectionType12 = ";Extended Properties='Excel 12.0 xml;HDR=YES;IMEX=1;MAXSCANROWS=0'";
        private string strConnectionType12 = "; Extended Properties = 'Excel 12.0'";
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
            string connect = "";
            string extension = Path.GetExtension(path);
            if(extension.IndexOf("xls") > 0)
            {
                connect = strConnectionProv8 + path + strConnectionType8;
            }
            if (extension.IndexOf("xlsx") > 0)
            {
                connect = strConnectionProv12 + path + strConnectionType12;
            }
            connect = strConnectionProv12 + path + strConnectionType12;
            connect = @"provider = Microsoft.ACE.OLEDB.12.0; 
                            data source = C:\Users\Ilya\Google Диск\Projects\ВНИИФК\Загрузчик в ЭБ\Тест2.xls; 
                            Extended Properties = 'Excel 12.0'";
            return connect;
        }

    }
}
