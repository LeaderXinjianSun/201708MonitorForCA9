using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BingLibrary.hjb;
using BingLibrary.hjb.Intercepts;
using System.ComponentModel.Composition;
using System.Collections.ObjectModel;
using System.Windows.Threading;
using System.IO;
using System.Windows;
using OfficeOpenXml;
using 臻鼎科技OraDB;
using System.Data;
using Microsoft.Win32;
using System.Drawing;

namespace MonitorUIforCA9
{
    [BingAutoNotify]
    public class Class1 : DataSource
    {
        #region 属性
        public virtual ObservableCollection<MachineStatus> MachineStatusCollection { set; get; } = new ObservableCollection<MachineStatus>();
        public virtual int ItemsCount { set; get; }
        public virtual string SQL_ora_server { set; get; } = "qwer";
        public virtual string SQL_ora_user { set; get; } = "sfcabar";
        public virtual string SQL_ora_pwd { set; get; } = "sfcabar*168";
        #endregion
        #region 变量
        private List<MachineStatus> machineStatusList = new List<MachineStatus>();
        private static DispatcherTimer dispatcherTimer = new DispatcherTimer();
        private string mashineIDExcelFile = System.Environment.CurrentDirectory + "\\MashineName.xlsx";
        #endregion
        #region 构造函数
        public Class1()
        {
            //machineStatusList.Add(new MachineStatus { MachineID = "M1",  YieldCount = 10, AlarmCount = 1, AlmPer = 91.2, UpdateTime = "201708050607", RemoteIP = "192.168.0.2" });
            //ItemsCount = 0;
            dispatcherTimer.Tick += new EventHandler(DispatcherTimerTickUpdateUi);
            dispatcherTimer.Interval = new TimeSpan(0, 0, 1);
            dispatcherTimer.Start();
            bool r = ImportMashineIDExcelFile();
            if (r)
            {
                RunLoop();
            }
        }
        #endregion
        #region 命令函数

        public void ExpoerCommand()
        {
            #region choose export file
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Filter = "Microsoft Excel 2013|*.xlsx";
            dialog.DefaultExt = "xlsx";
            dialog.AddExtension = true;
            dialog.Title = "Save Excel";
            dialog.InitialDirectory = "D:\\";
            dialog.FileName = DateTime.Now.ToString("yyyyMMdd") + DateTime.Now.ToString("HHmmss");
            bool? result = dialog.ShowDialog();
            if (result == null || result.Value == false)
            {
                return;
            }
            #endregion
            #region check file if it is openning
            FileStream stream = null;
            try
            {
                stream = new FileStream(dialog.FileName, FileMode.Create);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            #endregion
            #region write excel
            using (stream)
            {
                ExcelPackage package = new ExcelPackage(stream);

                package.Workbook.Worksheets.Add("CA9");
                ExcelWorksheet sheet = package.Workbook.Worksheets[1];

                #region write header
                sheet.Cells[1, 1].Value = "MachineID";
                sheet.Cells[1, 2].Value = "UserID";
                sheet.Cells[1, 3].Value = "YieldCount";
                sheet.Cells[1, 4].Value = "AlarmCount";
                sheet.Cells[1, 5].Value = "AlmPer";
                sheet.Cells[1, 6].Value = "UpdateTime";
                sheet.Cells[1, 7].Value = "RemoteIP";

                using (ExcelRange range = sheet.Cells[1, 1, 1, 7])
                {
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(Color.Gray);
                    range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                    range.Style.Border.Bottom.Color.SetColor(Color.Black);
                    range.AutoFitColumns(4);
                }
                #endregion

                #region write content
                int pos = 2;
                foreach (MachineStatus item in machineStatusList)
                {
                    sheet.Cells[pos, 1].Value = item.MachineID;
                    sheet.Cells[pos, 2].Value = item.UserID;
                    sheet.Cells[pos, 3].Value = item.YieldCount.ToString();
                    sheet.Cells[pos, 4].Value = item.AlarmCount.ToString();
                    sheet.Cells[pos, 5].Value = item.AlmPer.ToString();
                    sheet.Cells[pos, 6].Value = item.UpdateTime;
                    sheet.Cells[pos, 7].Value = item.RemoteIP;

                    if (item.AlmPer > 90)
                    {
                        using (ExcelRange range = sheet.Cells[pos, 1, pos, 5])
                        {
                            range.Style.Font.Color.SetColor(Color.Red);
                        }
                    }

                    using (ExcelRange range = sheet.Cells[pos, 1, pos, 7])
                    {
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        range.Style.Border.Bottom.Color.SetColor(Color.Black);
                        range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                    }

                    pos++;
                }
                #endregion

                package.Save();
            }
            #endregion

            MessageBox.Show("Export Successfully!");
        }

        #endregion
        #region 功能函数

        private bool ImportMashineIDExcelFile()
        {
            #region check file if exists
            FileStream stream = null;
            try
            {
                //stream = new FileStream(dialog.FileName, FileMode.Open);
                stream = File.OpenRead(mashineIDExcelFile);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            #endregion
            #region read excel
            using (stream)
            {
                ExcelPackage package = new ExcelPackage(stream);

                ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                #region check excel format
                if (sheet == null)
                {
                    MessageBox.Show("Excel format error!");
                    return false;
                }
                if (!sheet.Cells[1, 1].Value.Equals("MachineID"))
                {
                    MessageBox.Show("Excel format error!");
                    return false;
                }
                #endregion

                #region get last row index
                int lastRow = sheet.Dimension.End.Row;
                while (sheet.Cells[lastRow, 1].Value == null)
                {
                    lastRow--;
                }
                #endregion

                #region read datas
                for (int i = 2; i <= lastRow; i++)
                {
                    machineStatusList.Add(new MachineStatus
                    {
                        MachineID = sheet.Cells[i, 1].Value.ToString()
                    });
                }
                return true;
                #endregion

            }
            #endregion
        }
        private string SelectDt(string MashineID)
        {
            string[] arrField = new string[1];
            string[] arrValue = new string[1];
            string rtstr = "error";
            DataTable dt;
            try
            {
                string tablename = "sfcdata.barautbind";
                OraDB oraDB = new OraDB(SQL_ora_server, SQL_ora_user, SQL_ora_pwd);
                if (oraDB.isConnect())
                {
                    arrField[0] = "BB01";
                    arrValue[0] = MashineID;
                    DataSet s = oraDB.selectSQLwithOrder(tablename.ToUpper(), arrField, arrValue);
                    dt = s.Tables[0];
                    if (dt.Rows.Count > 0)
                    {
                        rtstr = (string)dt.Rows[0]["BB01"] + ";" + (string)dt.Rows[0]["BB02"] + ";" + (string)dt.Rows[0]["BB03"] + ";" + (string)dt.Rows[0]["BB04"] + ";" + (string)dt.Rows[0]["BLUID"] + ";" + (string)dt.Rows[0]["BB06"] + ";" + (string)dt.Rows[0]["BB07"];
                    }
                }
                oraDB.disconnect();
            }
            catch (Exception ex)
            {
                Log.Default.Error("Class1.SelectDt", ex.Message);
            }
            return rtstr;
        }
        private async void RunLoop()
        {
            while (true)
            {
                await Task.Delay(100);
                try
                {
                    foreach (MachineStatus item in machineStatusList)
                    {
                        string s = SelectDt(item.MachineID);
                        string[] ss = s.Split(new string[] { ";"}, StringSplitOptions.None);
                        if (ss.Length == 7)
                        {
                            item.UserID = ss[1];
                            item.YieldCount = int.Parse(ss[2]);
                            item.AlarmCount = int.Parse(ss[3]);
                            item.AlmPer = double.Parse(ss[4]);
                            item.UpdateTime = ss[5];
                            item.RemoteIP = ss[6];
                        }
                        await Task.Delay(200);
                    }
                }
                catch
                {

                }
            }
        }
        #endregion
        #region 事件函数

        private void DispatcherTimerTickUpdateUi(Object sender, EventArgs e)
        {
            if (machineStatusList.Count > 0)
            {
                MachineStatusCollection.Clear();
            }
            foreach (MachineStatus item in machineStatusList)
            {
                MachineStatusCollection.Add(item);
            }
            ItemsCount = machineStatusList.Count;
        }

        #endregion
    }
    public class MachineStatus
    {
        public string MachineID { set; get; }
        public string UserID { set; get; }
        public int YieldCount { set; get; }
        public int AlarmCount { set; get; }
        public double AlmPer { set; get; }
        public string UpdateTime { set; get; }
        public string RemoteIP { set; get; }
    }
    class VMManager
    {
        [Export(MEF.Contracts.Data)]
        [ExportMetadata(MEF.Key, "md")]
        Class1 md = Class1.New<Class1>();
    }
}