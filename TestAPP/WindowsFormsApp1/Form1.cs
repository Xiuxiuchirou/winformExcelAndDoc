using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        } 
        public static class Excel
        {

            /// <summary>
            /// 将excel内容读到DataTable中
            /// </summary>
            /// <param name="dt">存放数据DataTable</param>
            /// <param name="fileName">excel文件</param>
            /// <param name="iSheet">第几sheet页（0开头）</param>
            /// <param name="isFirstRowColumn">excel第一行为标题</param>
            /// <returns>数据行数</returns>
            public static int ReadExcel(ref DataTable dt, string fileName, int iSheet, bool isFirstRowColumn)
            {
                IWorkbook workbook = null;
                FileStream fs = null;
                ISheet sheet = null;
                int startRow = 0;
                try
                {
                    fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
                    if (fileName.ToLower().IndexOf(".xlsx") > 0) // 2007版本
                        workbook = new XSSFWorkbook(fs);
                    else if (fileName.ToLower().IndexOf(".xls") > 0) // 2003版本
                        workbook = new HSSFWorkbook(fs);

                    //if (sheetName != null)
                    //{
                    //    sheet = workbook.GetSheet(sheetName);
                    //    if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                    //    {
                    //        sheet = workbook.GetSheetAt(0);
                    //    }
                    //}

                    sheet = workbook.GetSheetAt(iSheet);

                    if (sheet != null)
                    {
                        IRow firstRow = sheet.GetRow(0);
                        int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                        if (isFirstRowColumn)
                        {
                            for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                            {
                                ICell cell = firstRow.GetCell(i);
                                if (cell != null)
                                {
                                    string cellValue = cell.StringCellValue;
                                    if (cellValue != null)
                                    {
                                        DataColumn column = new DataColumn(cellValue);
                                        dt.Columns.Add(column);
                                    }
                                }
                            }
                            startRow = sheet.FirstRowNum + 1;
                        }
                        else
                        {
                            startRow = sheet.FirstRowNum;
                        }

                        //最后一列的标号
                        int rowCount = sheet.LastRowNum;
                        for (int i = startRow; i <= rowCount; ++i)
                        {
                            IRow row = sheet.GetRow(i);
                            //当调整行高但没有输入具体信息时，row!=null and row.FirstCellNum==-1，因此也要增加对row.FirstCellNum值是否为-1的判断
                            if (row == null || row.FirstCellNum == -1)
                                continue; //没有数据的行默认是null　　　　　　　

                            DataRow dataRow = dt.NewRow();
                            for (int j = row.FirstCellNum; j < cellCount; ++j)
                            {
                                if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                    dataRow[j] = row.GetCell(j).ToString();
                            }
                            dt.Rows.Add(dataRow);
                        }
                    }
                    return dt.Rows.Count;
                }
                catch (Exception ex)
                {
                    throw new Exception("读取excel失败:" + ex.Message);
                }
            }
            /// <summary>
            /// 将excel内容读到DataTable中,读第一Sheet页
            /// </summary>
            /// <param name="dt">存放数据DataTable</param>
            /// <param name="sPathFile">excel文件</param>
            /// <param name="bTitle">excel第一行为标题</param>
            /// <returns>数据行数</returns>
            public static int ReadExcel(ref DataTable dt, string sPathFile, bool bTitle)
            {
                return ReadExcel(ref dt, sPathFile, 0, bTitle);
            }

            /// <summary>
            /// 将excel内容读到DataTable中,读第一Sheet页，第一行作为标题
            /// </summary>
            /// <param name="dt">存放数据DataTable</param>
            /// <param name="sPathFile">excel文件</param>
            /// <returns>数据行数</returns>
            public static int ReadExcel(ref DataTable dt, string sPathFile)
            {
                return ReadExcel(ref dt, sPathFile, 0, true);
            }

            /// <summary>
            /// 将excel内容读到DataTable中
            /// </summary>
            /// <param name="sPathFile">excel文件</param>
            /// <param name="iSheet">第几sheet页（0开头）</param>
            /// <param name="bTitle">excel第一行为标题</param>
            /// <returns></returns>
            public static DataTable ReadExcel(string sPathFile, int iSheet, bool bTitle)
            {
                DataTable dt = new DataTable();
                ReadExcel(ref dt, sPathFile, iSheet, bTitle);
                return dt;
            }


            /// <summary>
            /// 将excel内容读到DataTable中,第一行作字段名
            /// </summary>
            /// <param name="sPathFile">excel文件</param>
            /// <param name="iSheet">第几sheet页（0开头）</param>
            /// <returns></returns>
            public static DataTable ReadExcel(string sPathFile, int iSheet)
            {
                return ReadExcel(sPathFile, iSheet, true);
            }

            /// <summary>
            /// 将excel第一sheet页内容读到DataTable中,第一行作字段名
            /// </summary>
            /// <param name="sPathFile">excel文件</param>
            /// <returns></returns>
            public static DataTable ReadExcel(string sPathFile)
            {
                return ReadExcel(sPathFile, 0, true);
            }

           

          
        }

        /// <summary>
        /// 从excel中取数据，并整理字段
        /// </summary>
        /// <param name="sExcelFile"></param>



        public DataTable ExceltoDataSet(string path)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1';";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            System.Data.DataTable schemaTable = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
            string tableName = schemaTable.Rows[0][2].ToString().Trim();

            string strExcel = "";
            OleDbDataAdapter myCommand = null;
            DataSet ds = null;
            strExcel = "Select   *   From   [" + tableName + "]";
            myCommand = new OleDbDataAdapter(strExcel, strConn);

            ds = new DataSet();

            myCommand.Fill(ds, tableName);
            System.Data.DataTable dt = ds.Tables[0];
             
            return dt;

        }
        private void button2_Click(object sender, EventArgs e)
        {
           
            using (OpenFileDialog openFileDialog = new OpenFileDialog() { Filter = "Excel 97-2003 Workbook|*.xls|Excel Workbook|*.xlsx" })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                { 
                    string importExcelPath =  openFileDialog.FileName; 
                    DataTable dt = new DataTable();
                    Excel.ReadExcel(ref dt, importExcelPath);
                    DataColumn dcdelete;
                    if (dt.Columns.Contains("答案"))
                    {
                        foreach (DataColumn dc in dt.Columns)
                        {
                            if (dc.ColumnName == "答案")
                            {
                                dcdelete = dc; 
                                dt.Columns.Remove(dcdelete);
                                break;
                            } 
                        }
                        
                    }
                    dataGridView1.DataSource = dt;
                    
                     
                }
            }
        }
        private static void KillSpecialExcel()
        {
            foreach (System.Diagnostics.Process theProc in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                if (!theProc.HasExited)
                {
                    bool b = theProc.CloseMainWindow();
                    if (b == false)
                    {
                        theProc.Kill();
                    }
                    theProc.Close();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            try
            {
                dt = (DataTable)dataGridView1.DataSource;
            }
            catch (Exception ec)
            {
                MessageBox.Show(ec.Message);
                return;
            }
            if (dt == null || dt.Rows.Count <10)
                return;
            int number = (int)numericUpDown1.Value;
            if (number == 0)
            {
                number = 1;
            }
            for (int i = 0; i < number; i++)
            {
                //生成试卷
                int tmp = 10;
                List<int> list = new List<int>();
                int ix = 10000;
                while (list.Count < tmp && ix > 0)
                {
                    Random rd = new Random();  //无参即为使用系统时钟为种子
                    int index = rd.Next(0, dt.Rows.Count);
                    if (!list.Contains(index))
                    {
                        list.Add(index);
                    }
                }
                if (list.Count == tmp)
                {
                    List<string> testList = new List<string>();
                    foreach (var item in list)
                    {
                        if (dt.Rows.Count > item)
                            testList.Add(dt.Rows[item]["题目"].ToString());
                    }
                    createDocFiles(testList,i.ToString());
                } 

            }
            MessageBox.Show("试卷生成成功,请查看桌面文件夹");

        }


        public void createDocFiles(List<string> testList,string fileName)
        {
            XWPFDocument doc = new XWPFDocument();      //创建新的word文档

            XWPFParagraph p1 = doc.CreateParagraph();   //向新文档中添加段落
            p1.SetAlignment(ParagraphAlignment.CENTER); //段落对其方式为居中
            XWPFRun r1 = p1.CreateRun();                //向该段落中添加文字
            r1.SetText("配电信息化应用试题");
            r1.SetBold(true);//设置粗体 
            r1.SetFontSize(16);//设置字体大小


            XWPFParagraph space = doc.CreateParagraph();   //向新文档中添加段落
            space.SetAlignment(ParagraphAlignment.CENTER); //段落对其方式为居中
            XWPFRun sp1 = space.CreateRun();                //向该段落中添加文字
            sp1.SetText(@" 
");
         



            XWPFParagraph pX = doc.CreateParagraph();   //向新文档中添加段落
            pX.SetAlignment(ParagraphAlignment.LEFT); //段落对其方式为居中
            XWPFRun rX = pX.CreateRun();                //向该段落中添加文字
            rX.SetText("单位:                                                        姓名:                         ");
            rX.SetBold(true);//设置粗体 
            rX.SetFontSize(12);//设置字体大小

            XWPFParagraph space1 = doc.CreateParagraph();   //向新文档中添加段落
            space1.SetAlignment(ParagraphAlignment.CENTER); //段落对其方式为居中
            XWPFRun sp11 = space.CreateRun();                //向该段落中添加文字
            sp11.SetText(@"




");

            int i = 0;
            foreach (var item in testList)
            {
                XWPFParagraph p2 = doc.CreateParagraph();
                p2.SetAlignment(ParagraphAlignment.LEFT);

                XWPFRun r2 = p2.CreateRun();
                r2.SetText((++i).ToString()+"."+item.ToString().Trim()+ @"                                                        
                                                       
                                                       
                                                       
                                                       ");
                r2.SetFontSize(12);//设置字体大小
                //r2.SetBold(true);//设置粗体 
                p2.SetAlignment(ParagraphAlignment.LEFT);
            }
            FileStream sw = File.Create(Environment.GetFolderPath(Environment.SpecialFolder.Desktop)+"/测试题"+DateTime.Now.ToString("HHMMssmmm")+ fileName + ".docx");  
            doc.Write(sw);                              //...
            sw.Close();
        }
 
 



    }
}
