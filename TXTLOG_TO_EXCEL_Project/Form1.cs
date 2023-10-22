using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Threading;
using System.Collections;
using TXTLOG_TO_EXCEL_Project.Module;
using HorizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment;
using BorderStyle = NPOI.SS.UserModel.BorderStyle;

namespace TXTLOG_TO_EXCEL_Project
{
    public partial class Form1 : Form
    {

        /// <summary>
        /// 取得Sync_Dynamic_80611_20200716_Multi_S各項數據
        /// </summary>
        ArrayList arrayList_Dynamic = new ArrayList();

        /// <summary>
        /// 取得Input Output Eff Noise Multi or Single各項數據
        /// </summary>
        ArrayList arrayList_Ripple = new ArrayList();


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            #region 1.  Checked DKDMS.exe Execution
            //取得此process的名稱
            String name = Process.GetCurrentProcess().ProcessName;
            //取得所有與目前process名稱相同的process
            Process[] ps = Process.GetProcessesByName(name);
            //ps.Length > 1 表示此proces已重複執行
            if (ps.Length > 1)
            {
                System.Environment.Exit(System.Environment.ExitCode);
            }
            #endregion
        }
   

        private void btn_Report_Click(object sender, EventArgs e)
        {

            if (txt_TXTPATH.Text == "")
            {
                MessageBox.Show("請先選擇TXT LOG檔!");
                return;
            }
            this.Cursor = Cursors.WaitCursor;
            string strReadline = ""; //取得內容
            StreamReader reader = new StreamReader(txt_TXTPATH.Text, System.Text.Encoding.Default); //作業系統目前 ANSI 字碼頁的編碼方式               
            if ((strReadline = reader.ReadToEnd()) != null)
            {
                string[] ReadlineArray = Regex.Split(strReadline, "================================================================================");
                if (DataToExcel(ReadlineArray, dlg_TXT.FileName.Replace("txt", "xlsx")))
                {
                    MessageBox.Show("匯出成功!");
                }
            }
            reader.Close();
            this.Cursor = Cursors.Default;
        }

        private void lbl_TXT_Click(object sender, EventArgs e)
        {
            dlg_TXT.Title = "開啟LOG(txt)文件";
            dlg_TXT.Filter = "txt files (*.txt)|*.txt";
            dlg_TXT.FilterIndex = 1;
            dlg_TXT.RestoreDirectory = true;
            dlg_TXT.Multiselect = false;
            if (dlg_TXT.ShowDialog() == DialogResult.OK)
            {               
                txt_TXTPATH.Text = dlg_TXT.FileName;               
            }
        }

        #region 匯出成Excel_Sheet(Item ALL)
        /// <summary>
        /// Datable匯出成Excel
        /// </summary>
        /// <param name="dt">內容列</param>
        /// <param name="file">檔名</param>
        private bool DataToExcel(string[] arraystr, string file)
        {
            if(arrayList_Dynamic.Count > 0)
            {
                arrayList_Dynamic.Clear();               
            }

            if (arrayList_Ripple.Count > 0)
            {
                arrayList_Ripple.Clear();
            }

            int iCurrentRow = 0;  //記憶已被使用的列
            IWorkbook workbook;

            try
            {
                workbook = new XSSFWorkbook();
            }
            catch
            {
                workbook = new HSSFWorkbook();
                file = file.Replace("xlsx", "xls");
            }

            //超連結字體
            XSSFFont hyperlink_font = (XSSFFont)workbook.CreateFont();
            hyperlink_font.FontName = "Calibri";    //字型
            hyperlink_font.FontHeightInPoints = 12;  //字體大小
            hyperlink_font.Color = NPOI.HSSF.Util.HSSFColor.Blue.Index;
            hyperlink_font.Underline = NPOI.SS.UserModel.FontUnderlineType.Single;  //底線

            //正常字體
            XSSFFont normal_font = (XSSFFont)workbook.CreateFont();
            normal_font.FontName = "Calibri";    //字型
            normal_font.FontHeightInPoints = 12;  //字體大小

            //正常藍色字體
            XSSFFont normal_Blue_font = (XSSFFont)workbook.CreateFont();
            normal_Blue_font.FontName = "Calibri";    //字型
            normal_Blue_font.FontHeightInPoints = 12;  //字體大小    
            normal_Blue_font.Color = NPOI.HSSF.Util.HSSFColor.Blue.Index;


            for (int i = 1; i < arraystr.Length; i++)
            {
                string sDetail = arraystr[i].Trim();
                string[] arrayDetail = Regex.Split(sDetail, "\r\n|\n");

                #region 取得每項Sync_Dynamic_80611_20200716_Multi_S測項數據區
                if (arrayDetail[0].ToString().Contains("Sync_Dynamic_80611_20200716_Multi_S"))
                {
                    try
                    {
                        for (int j = 25; j < 34; j++)
                        {
                            string input = arrayDetail[j].ToString().Trim();
                            //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                            string pattern = @"\s+";
                            string replacement = " ";
                            string result = Regex.Replace(input, pattern, replacement);
                            string[] arrayresult = result.Split(' ');

                            switch (arrayresult[0])
                            {
                                case "48V":
                                    Dynamic_Column.col_48V_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_48V_Vpk_Low = arrayresult[6].ToString();
                                    break;
                                case "12V":
                                    Dynamic_Column.col_12V_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_12V_Vpk_Low = arrayresult[6].ToString();
                                    break;
                                case "24V":
                                    Dynamic_Column.col_24V_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_24V_Vpk_Low = arrayresult[6].ToString();
                                    break;
                                case "5V_CMB":
                                    Dynamic_Column.col_5V_CMB_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_5V_CMB_Vpk_Low = arrayresult[6].ToString();
                                    break ;
                                case "5CMB":
                                    Dynamic_Column.col_5V_CMB_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_5V_CMB_Vpk_Low = arrayresult[6].ToString();
                                    break;
                                case "5V":
                                    Dynamic_Column.col_5V_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_5V_Vpk_Low = arrayresult[6].ToString();
                                    break;
                                case "3.3V":
                                    Dynamic_Column.col_3V_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_3V_Vpk_Low = arrayresult[6].ToString();
                                    break;
                                case "PWOK_D2D_1":
                                    Dynamic_Column.col_PWOK_D2D_1_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_PWOK_D2D_1_Vpk_Low = arrayresult[6].ToString();
                                    break;
                                case "PWOK_D2D_2":
                                    Dynamic_Column.col_PWOK_D2D_2_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_PWOK_D2D_2_Vpk_Low = arrayresult[6].ToString();
                                    break;
                                case "SMBAlert":
                                    Dynamic_Column.col_SMBAlert_Vpk_High = arrayresult[3].ToString();
                                    Dynamic_Column.col_SMBAlert_Vpk_Low = arrayresult[6].ToString();
                                    break;
                            }
                        }
                        //取得個欄位數據集和欄位名稱
                        string sAll_Data = "";
                        string sAll_Col = "";
                        //取得模組裡各個變數與值
                        Type dynamicColumnType = typeof(Dynamic_Column);
                        PropertyInfo[] properties = dynamicColumnType.GetProperties(BindingFlags.Public | BindingFlags.Static);
                        foreach (PropertyInfo property in properties)
                        {
                            string propertyName = property.Name;
                            if(propertyName.Split('_')[1] == "3V")
                            {
                                propertyName = propertyName.Replace("3V", "3.3V");
                            }
                            sAll_Col += propertyName.Replace("col_", "") + "|";
                            string propertyValue = (string)property.GetValue(null); // 使用null表示靜態屬性值
                            sAll_Data += propertyValue + "|";
                        }

                        if (sAll_Col != "")
                        {
                            sAll_Col = sAll_Col.Trim('|');
                        }

                        if (sAll_Data != "")
                        {
                            sAll_Data = sAll_Data.Trim('|');
                        }

                        if (arrayList_Dynamic.Count == 0)
                        {
                            arrayList_Dynamic.Add(sAll_Col);
                            arrayList_Dynamic.Add(sAll_Data);
                        }
                        else
                        {
                            arrayList_Dynamic.Add(sAll_Data);
                        }
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show("抓取Dynamic數據失敗:" + ex.Message);
                    }                    

                }
                #endregion

                #region 取得每項Ripple測項數據區
                if (arrayDetail[0].ToString().Contains("Input Output Eff Noise Multi or Single"))
                {
                    try
                    {
                        for (int j = 75; j < 81; j++)
                        {
                            string input = arrayDetail[j].ToString().Trim();
                            //\s 是一個特殊的元字符，用於匹配各種空白字符，包括空格、制表符、换行符等。而 + 是量詞，表示匹配前面的元素一次或多次。
                            string pattern = @"\s+";
                            string replacement = " ";
                            string result = Regex.Replace(input, pattern, replacement);
                            string[] arrayresult = result.Split(' ');

                            switch (arrayresult[0])
                            {
                                case "48V":
                                    Eff_Noise_Column.col_48V_Vpp_Noise = arrayresult[3].ToString();
                                    break;
                                case "12V":
                                    Eff_Noise_Column.col_12V_Vpp_Noise = arrayresult[3].ToString();
                                    break;
                                case "24V":
                                    Eff_Noise_Column.col_24V_Vpp_Noise = arrayresult[3].ToString();
                                    break;
                                case "5CMB":
                                    Eff_Noise_Column.col_5V_CMB_Vpp_Noise = arrayresult[3].ToString();
                                    break;
                                case "5V":
                                    Eff_Noise_Column.col_5V_Vpp_Noise = arrayresult[3].ToString();
                                    break;
                                case "3.3V":
                                    Eff_Noise_Column.col_3V_Vpp_Noise = arrayresult[3].ToString();
                                    break;

                            }

                        }
                        //取得個欄位數據集和欄位名稱
                        string sAll_Data = "";
                        string sAll_Col = "";
                        //取得模組裡各個變數與值
                        Type EffColumnType = typeof(Eff_Noise_Column);
                        PropertyInfo[] properties = EffColumnType.GetProperties(BindingFlags.Public | BindingFlags.Static);
                        foreach (PropertyInfo property in properties)
                        {
                            string propertyName = property.Name;
                            sAll_Col += propertyName.Replace("col_", "") + "(mV)" + "|";
                            string propertyValue = (string)property.GetValue(null); // 使用null表示靜態屬性值
                            sAll_Data += propertyValue + "|";
                        }

                        if (sAll_Col != "")
                        {
                            sAll_Col = sAll_Col.Trim('|');
                        }

                        if (sAll_Data != "")
                        {
                            sAll_Data = sAll_Data.Trim('|');
                        }

                        if (arrayList_Ripple.Count == 0)
                        {
                            arrayList_Ripple.Add(sAll_Col);
                            arrayList_Ripple.Add(sAll_Data);
                        }
                        else
                        {
                            arrayList_Ripple.Add(sAll_Data);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("抓取Ripple數據失敗:" + ex.Message);
                    }                  
                }
                #endregion
            }

            //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 黑字體
            XSSFCellStyle Dynamic_Item_style = (XSSFCellStyle)workbook.CreateCellStyle();
            Dynamic_Item_style.VerticalAlignment = VerticalAlignment.Center;
            Dynamic_Item_style.Alignment = HorizontalAlignment.Center;
            Dynamic_Item_style.BorderTop = BorderStyle.Medium;
            Dynamic_Item_style.BorderBottom = BorderStyle.Medium;
            Dynamic_Item_style.BorderLeft = BorderStyle.Medium;
            Dynamic_Item_style.BorderRight = BorderStyle.Medium;
            Dynamic_Item_style.WrapText = true;
            Dynamic_Item_style.SetFont(normal_font);

            //垂直水平置中 儲存格的邊框樣式 文字控制自動換列 藍字體
            XSSFCellStyle Dynamic_Data_style = (XSSFCellStyle)workbook.CreateCellStyle();
            Dynamic_Data_style.VerticalAlignment = VerticalAlignment.Center;
            Dynamic_Data_style.Alignment = HorizontalAlignment.Center;
            Dynamic_Data_style.BorderTop = BorderStyle.Medium;
            Dynamic_Data_style.BorderBottom = BorderStyle.Medium;
            Dynamic_Data_style.BorderLeft = BorderStyle.Medium;
            Dynamic_Data_style.BorderRight = BorderStyle.Medium;
            Dynamic_Data_style.WrapText = true;
            Dynamic_Data_style.SetFont(normal_Blue_font);

            #region 寫入Sheet(Dynamic)內容

            ISheet sheet_Dynamic = workbook.CreateSheet("Dynamic");
            iCurrentRow = 0;

            for (int i = 0; i < arrayList_Dynamic.Count; i++)
            {
                string sDynamic_Data = arrayList_Dynamic[i].ToString();
                string[] arrayDynamic_Data = sDynamic_Data.Split('|');
                IRow Dynamic_row1 = sheet_Dynamic.CreateRow(iCurrentRow);
                Dynamic_row1.HeightInPoints = 30;  //設定每個儲存格列高

                try
                {
                    for (int j = 0; j < arrayDynamic_Data.Length; j++)
                    {
                        ICell Dynamic_cell = Dynamic_row1.CreateCell(j);

                        #region SetColumnWidth  需* 256用途
                        //在 NPOI 中，設置列寬時，使用的單位是 1 / 256 字符寬度。這是因為 Excel 的列寬度是以字符寬度為基準的，
                        //為什麼要使用 1 / 256 字符寬度作為單位呢？這是因為 Excel 的列寬是以列寬度單元格的總數為基準的。
                        //在代碼中，15 * 256 的計算結果表示將列寬設置為 15 個字符的寬度。通過乘以 256，我們將該值轉換為 Excel 中使用的單位。
                        //一個單元格的列寬度為 1，並且可以使用更小的單位來調整列寬度。使用 1 / 256 字符寬度作為基本單位，可以更精確地調整列寬，以符合特定的需求。
                        #endregion

                        sheet_Dynamic.SetColumnWidth(j, 14 * 256);  //設定每個儲存格欄寬 
                        if (i == 0)
                        {
                            Dynamic_cell.CellStyle = Dynamic_Item_style;
                            if (arrayDynamic_Data[j].ToString().Contains("*"))
                            {
                                Dynamic_cell.SetCellValue(arrayDynamic_Data[j].ToString().Replace(arrayDynamic_Data[j].ToString(), "*"));
                            }
                            else
                            {
                                Dynamic_cell.SetCellValue(arrayDynamic_Data[j].ToString());
                            }
                        }
                        else
                        {
                            Dynamic_cell.CellStyle = Dynamic_Data_style;
                            if (arrayDynamic_Data[j].ToString().Contains("*"))
                            {
                                Dynamic_cell.SetCellValue(arrayDynamic_Data[j].ToString().Replace(arrayDynamic_Data[j].ToString(), "*"));
                            }
                            else
                            {
                                Dynamic_cell.SetCellValue(Convert.ToDouble(arrayDynamic_Data[j].ToString()));
                            }
                        }
                    }

                    iCurrentRow += 1;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Dynamic資料處理失敗:" + ex.Message);
                }

            }

            #endregion

            #region 寫入Sheet(Ripple)內容

            ISheet sheet_Ripple = workbook.CreateSheet("Ripple");
            iCurrentRow = 0;

            for (int i = 0; i < arrayList_Ripple.Count; i++)
            {
                string sRipple_Data = arrayList_Ripple[i].ToString();
                string[] arrayRipple_Data = sRipple_Data.Split('|');
                IRow Ripple_row1 = sheet_Ripple.CreateRow(iCurrentRow);
                Ripple_row1.HeightInPoints = 30;  //設定每個儲存格列高

                try
                {
                    for (int j = 0; j < arrayRipple_Data.Length; j++)
                    {
                        ICell Ripple_cell = Ripple_row1.CreateCell(j);

                        #region SetColumnWidth  需* 256用途
                        //在 NPOI 中，設置列寬時，使用的單位是 1 / 256 字符寬度。這是因為 Excel 的列寬度是以字符寬度為基準的，
                        //為什麼要使用 1 / 256 字符寬度作為單位呢？這是因為 Excel 的列寬是以列寬度單元格的總數為基準的。
                        //在代碼中，15 * 256 的計算結果表示將列寬設置為 15 個字符的寬度。通過乘以 256，我們將該值轉換為 Excel 中使用的單位。
                        //一個單元格的列寬度為 1，並且可以使用更小的單位來調整列寬度。使用 1 / 256 字符寬度作為基本單位，可以更精確地調整列寬，以符合特定的需求。
                        #endregion

                        sheet_Ripple.SetColumnWidth(j, 14 * 256);  //設定每個儲存格欄寬 
                        if (i == 0)
                        {
                            Ripple_cell.CellStyle = Dynamic_Item_style;
                            if (arrayRipple_Data[j].ToString().Contains("*"))
                            {
                                Ripple_cell.SetCellValue(arrayRipple_Data[j].ToString().Replace(arrayRipple_Data[j].ToString(), "*"));
                            }
                            else
                            {
                                Ripple_cell.SetCellValue(arrayRipple_Data[j].ToString());
                            }
                        }
                        else
                        {
                            Ripple_cell.CellStyle = Dynamic_Data_style;
                            if (arrayRipple_Data[j].ToString().Contains("*"))
                            {
                                Ripple_cell.SetCellValue(arrayRipple_Data[j].ToString().Replace(arrayRipple_Data[j].ToString(), "*"));
                            }
                            else
                            {
                                Ripple_cell.SetCellValue(Convert.ToDouble(arrayRipple_Data[j].ToString()));
                            }
                        }

                    }

                    iCurrentRow += 1;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ripple資料處理失敗:" + ex.Message);
                }

            }

            #endregion

            try
            {
                MemoryStream stream = new MemoryStream();
                workbook.Write(stream);
                byte[] buf = stream.ToArray();
                stream.Flush();

                //儲存為Excel檔案  
                using (FileStream fs = new FileStream(file, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(buf, 0, buf.Length);
                    fs.Flush();
                }                
            }
            catch (Exception ex)
            {
                MessageBox.Show("寫入EXCEL檔失敗:" + ex.Message);
                return false;
            }

            Thread.Sleep(200);

            //畫折線圖
            if (!Table_Chart.Draw_Table(file))
            {
                return false;
            }

            return true;
        }
        #endregion

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            GC.Collect();
        }        
    }
}
