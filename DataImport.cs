using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Xml;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.SqlClient;
using System.Threading;

//当前为main分支
namespace DataClient
{
    public partial class DataImport : Form
    {
        public Microsoft.Office.Interop.Excel.ApplicationClass m_ExcelApplication = null;//启动Excel
        Workbook m_xBook = null;//打开Excel文件
        Worksheet workSheet = null;//打开Sheet文件
        MdbHelper m_mdbHelper = null;
        string m_pointNoPrefix = "";

        public DataImport()
        {
            InitializeComponent();
            OpenExcel();
            m_pointNoPrefix = GenRandomCode();
        }


        private void SetExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            ofd.Title = "请选择Excel管线数据文件！";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                ExcelPath.Text = ofd.FileName;
                string currentDirectory = Environment.CurrentDirectory;
                m_ExcelApplication.Visible = false;

                this.m_xBook = m_ExcelApplication.Workbooks.Open(ofd.FileName, Missing.Value, Missing.Value, 
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                fillLayer();
            }
        }

        public void OutPut(string strVal)
        {
            lbOuput.Items.Add(strVal);
            lbOuput.SetSelected(lbOuput.Items.Count - 1, true);
        }

        void fillLayer()
        {
            chkLayerList.Items.Clear();

            int nCount = m_xBook.Sheets.Count;
            for (int i = 1; i <= nCount; i++)
            {
                Worksheet worksheet = (Worksheet)m_xBook.Sheets[i];
                string strSheetName = worksheet.Name;
                chkLayerList.Items.Add(strSheetName);
            }
        }


        private void SetOutMDB_Click(object sender, EventArgs e)
        {
            if(savePersonMDB.ShowDialog()==DialogResult.OK)
            {
                PersonMDB.Text = savePersonMDB.FileName;
            }
        }

        System.Collections.Generic.List<string> DXArray = new List<string>();


        private void DataImport_Load(object sender, EventArgs e)
        {
            
        }
        private void AddLog(string value)
        {

        }


        ///处理两点一线数据
        ///
        private void AllButton_Click(object sender, EventArgs e)
        {
            if (ExcelPath.Text == "")
            {
                MessageBox.Show("未选择Excel数据文件！");
                return;
            }

            if (PersonMDB.Text == "")
            {
                MessageBox.Show("未选择输出的mdb文件！");
                return;
            }

            if (chkLayerList.CheckedItems.Count <= 0)
            {
                MessageBox.Show("请选择要处理的图层！");
                return;
            }

            if (!File.Exists(PersonMDB.Text))
            {
                bool bSucc = MdbHelper.CreateAccessDatabase(PersonMDB.Text);
                if (bSucc)
                {
                    OutPut("创建mdb文件成功！");
                }
                else
                {
                    OutPut("创建mdb文件失败！");
                    MessageBox.Show("mdb路径不合法！");
                    return;
                }
            }
            m_mdbHelper = new MdbHelper(PersonMDB.Text);

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            string strProjectName = txtProjectName.Text;
            int nCount = chkLayerList.CheckedIndices.Count;
            for (int i = 0; i < nCount; i++)
            {
                int idx = chkLayerList.CheckedIndices[i];
                
                Worksheet ws = (Worksheet)m_xBook.Sheets[idx + 1];
                string strSheetName = ws.Name;
                Console.WriteLine((idx + 1) + strSheetName);

                readSheet rs = new readSheet();
                rs.PointNoPrefix = m_pointNoPrefix;
                rs.ProjectName = strProjectName;
                rs.MdbHelper = m_mdbHelper;
                rs.Output = lbOuput;
                rs.DealSheet(ws);
                rs.ExecuteSql();
            }
            OutPut("处理完毕,用时" + stopWatch.Elapsed + "秒！");
        }

        //创建打开Excel环境
        private bool OpenExcel()
        {
            m_ExcelApplication = new Microsoft.Office.Interop.Excel.ApplicationClass();
            try
            {
                if (m_ExcelApplication == null)
                {
                    MessageBox.Show("无法启动Excel,可能您的电脑未安装Excel");
                    return false;
                }
                m_ExcelApplication.Visible = false;
                m_ExcelApplication.UserControl = true;
                m_ExcelApplication.DisplayAlerts = false;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            finally
            {

            }
            return true;
        }

        //关闭打开Excel环境
        private bool CloseExcel()
        {
            try
            {
                if (m_ExcelApplication == null)
                {
                    return true;
                }
                m_ExcelApplication.Quit();
                KeyMyExcelProcess.Kill(m_ExcelApplication);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(m_ExcelApplication);
                m_ExcelApplication = null;
                GC.Collect();
                //MarshalObject(ExcelApplication);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        //打开ExcelFile文件
        private bool OpenExcelFile(string strPath)
        {
            try
            {
                if (m_ExcelApplication == null)
                {
                    MessageBox.Show("请创建打开环境");
                    return false;
                }
                m_xBook = m_ExcelApplication.Workbooks.Open(strPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        //关闭ExcelFile文件
        private bool CloseExcelFile()
        {
            try
            {
                if (m_xBook == null)
                {
                    return true;
                }
                m_xBook.Close();
                //MarshalObject(workBook);
                if (m_xBook != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(m_xBook);
                }
                
                m_xBook = null;
                GC.Collect();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        //打开ExcelSheet文件
        private bool OpenSheet(int iSelectIndex)
        {
            if (m_xBook == null)
            {
                MessageBox.Show("请打开Excel");
                return false;
            }
            try
            {
                workSheet = m_xBook.Sheets[iSelectIndex] as Worksheet;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        //打开ExcelSheet文件
        private bool CloseSheet()
        {
            if (workSheet == null)
            {
                return true;
            }
            try
            {
                //MarshalObject(workSheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workSheet);
                workSheet = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

        private void DataImport_FormClosing(object sender, FormClosingEventArgs e)
        {
            CloseExcelFile();
            CloseExcel();
        }

        private void MarshalObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);
                }
                obj = null;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chkLayerList.Items.Count; i++)
            {
                chkLayerList.SetItemChecked(i, true);
            }
        }

        private void btnSelectNone_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chkLayerList.Items.Count; i++)
            {
                chkLayerList.SetItemChecked(i, false);
            }
        }

        private void btnSelectInvert_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < chkLayerList.Items.Count; i++)
            {
                if (chkLayerList.GetItemChecked(i))
                    chkLayerList.SetItemChecked(i, false);
                else
                {
                    chkLayerList.SetItemChecked(i, true);
                }
            }
        }


        public string GenRandomCode()
        {
            Thread.Sleep(50);
            Random rom = new Random((int)DateTime.Now.Ticks);
            char[] allcheckRandom ={'0','1','2','3','4','5','6','7','8','9',
                                    'A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W',
                                    'X','Y','Z','a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q',
                                    'r','s','t','u','v','w','x','y','z'};
            string Randomcode = "";
            for (int i = 0; i < 6; i++)
            {
                Randomcode += allcheckRandom[rom.Next(allcheckRandom.Length)];
            }

            return Randomcode;
        }
    }

    //kill Excel进程
    public class KeyMyExcelProcess
    {
        [System.Runtime.InteropServices.DllImport("User32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
        {
            try
            {
                IntPtr t = new IntPtr(excel.Hwnd);   //得到这个句柄，具体作用是得到这块内存入口
                int k = 0;
                GetWindowThreadProcessId(t, out k);   //得到本进程唯一标志k
                System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);   //得到对进程k的引用
                p.Kill();     //关闭进程k
            }
            catch (System.Exception ex)
            {
                throw ex;
            }
        }
    }
}


