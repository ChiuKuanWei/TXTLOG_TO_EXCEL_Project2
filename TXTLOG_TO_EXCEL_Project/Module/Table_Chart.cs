using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;

namespace TXTLOG_TO_EXCEL_Project.Module
{
    public static class Table_Chart
    {
        public static string[] EngArray = new string[]
            {
                "","A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M","N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"
            };
        /// <summary>
        /// 繪製圖表
        /// </summary>
        /// <param name="sFileName">指定執行檔案及儲存位置</param>
        public static bool Draw_Table(string sFileName)
        {
            
            Excel.Application myExcel = null;
            Excel.Workbook myBook = null;
            Excel.Worksheet mySheet = null;
            string strRange = ""; //設定數據範圍

            try
            {
                myExcel = new Excel.Application();
                myBook = myExcel.Workbooks.Open(sFileName);
                myExcel.Visible = true; //顯示Excel程式

                // 遍歷工作簿中的所有Sheet
                foreach (Excel.Worksheet worksheet in myBook.Sheets)
                {
                    //引用指定工作表
                    mySheet = (Excel.Worksheet)myBook.Worksheets[worksheet.Name];

                    if (worksheet.Name == "Dynamic")
                    {
                        //取得工作表包含數據的區域
                        Excel.Range usedRange = worksheet.UsedRange;

                        //取得包含數據的行數和列數
                        int rowCount = usedRange.Rows.Count;
                        int columnCount = usedRange.Columns.Count;

                        //取得模組裡各個變數
                        Type dynamicColumnType = typeof(Dynamic_Column);
                        PropertyInfo[] properties = dynamicColumnType.GetProperties(BindingFlags.Public | BindingFlags.Static);

                        //設定每張圖表的間距
                        int iShapes_Item_Top = 29;

                        //創建一個圖表索引變數
                        int chartIndex = 1;

                        //遍歷每一列 ->
                        for (int column = 1; column <= columnCount; column+=2)
                        {                           
                            string propertyName = properties[column].Name;
                            // col_48V_Vpk_High
                            propertyName = propertyName.Replace("col_", "");
                            string[] arraypropertyName = propertyName.Split('_');
                            propertyName = "";
                            int arraypropertyName_Length = arraypropertyName.Length-2;
                            for (int i = 0; i < arraypropertyName_Length; i++)
                            {
                                if (arraypropertyName[0] == "3V")
                                {
                                    arraypropertyName[0] = arraypropertyName[0].Replace("3V", "3.3V");
                                }
                                propertyName += arraypropertyName[i] + "_";
                            }

                            string sChartIndex = "Chart " + chartIndex;                           
                            strRange = EngArray[column] + "1:" + EngArray[column + 1] + rowCount.ToString();
                            Setworkbook(strRange, mySheet, myBook, propertyName.TrimEnd('_'), iShapes_Item_Top, sChartIndex);
                            Thread.Sleep(100);
                            iShapes_Item_Top += 310;
                            // 增加圖表索引
                            chartIndex++;
                        }
                    }
                    else if(worksheet.Name == "Ripple")
                    {

                    }
                }

                //儲存工作簿
                myBook.SaveAs(sFileName);

                //關閉工作簿和 Excel 應用程式
                myBook.Close();
                myExcel.Quit();

                //釋放COM，防止内存泄漏
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(mySheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myBook);

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("繪製圖表失敗:" + ex.Message);
                myExcel.Visible = true;
                return false;
            }
        }

        /// <summary>
        /// 設定圖表
        /// </summary>
        /// <param name="strRange">畫折線圖範圍</param>
        /// <param name="mySheet">指定畫圖表的sheet</param>
        /// <param name="myBook">取得EXCEL檔</param>
        /// <param name="sTitle">圖表標頭</param>
        /// <param name="iShapes_Item_Top">設定每張圖表(上邊距)的間距</param>
        /// <param name="sChartIndex">給每張圖表命名不同的Chart名稱</param>
        private static void Setworkbook(string strRange, Excel.Worksheet mySheet, Excel.Workbook myBook, string sTitle, int iShapes_Item_Top, string sChartIndex)
        {
            //在工作簿 新增一張 統計圖表，單獨放在一個分頁裡面
            myBook.Charts.Add(Type.Missing, Type.Missing, 1, Type.Missing);          

            //在循環內部，在添加新圖表後，為其設定唯一的名稱
            myBook.ActiveChart.Name = sChartIndex;

            //設定折線圖樣式
            myBook.ActiveChart.ChartType = Excel.XlChartType.xlLine;

            //設定 統計圖表 的 數據範圍內容
            myBook.ActiveChart.SetSourceData(mySheet.get_Range(strRange), Excel.XlRowCol.xlColumns);
            //將新增的統計圖表 插入到 指定位置
            myBook.ActiveChart.Location(Excel.XlChartLocation.xlLocationAsObject, mySheet.Name);

            mySheet.Shapes.Item(sChartIndex).Width = 500;   //調整圖表寬度
            mySheet.Shapes.Item(sChartIndex).Height = 310;  //調整圖表高度
            mySheet.Shapes.Item(sChartIndex).Top = iShapes_Item_Top; //調整圖表在分頁中的高度(上邊距) 位置
            mySheet.Shapes.Item(sChartIndex).Left = 1380;    //調整圖表在分頁中的左右(左邊距) 位置
           
           
            //設定 繪圖區 的 背景顏色
            myBook.ActiveChart.PlotArea.Interior.Color = ColorTranslator.ToOle(Color.White);
            //設定 繪圖區 的 邊框線條樣式
            myBook.ActiveChart.PlotArea.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            //設定 繪圖區 的 寬度
            myBook.ActiveChart.PlotArea.Width = 440;
            //設定 繪圖區 的 高度
            myBook.ActiveChart.PlotArea.Height = 250;
            //設定 繪圖區 在 圖表中的 高低位置(上邊距)
            myBook.ActiveChart.PlotArea.Top = 30;
            //設定 繪圖區 在 圖表中的 左右位置(左邊距)
            myBook.ActiveChart.PlotArea.Left = 15;
            //設定 繪圖區 的 x軸名稱下方 顯示y軸的 數據資料
            myBook.ActiveChart.HasDataTable = false;

            //設定 圖表的 背景顏色 使用color
            myBook.ActiveChart.ChartArea.Interior.Color = ColorTranslator.ToOle(Color.White);
            //設定 圖表的 邊框顏色 使用color
            myBook.ActiveChart.ChartArea.Border.Color = ColorTranslator.ToOle(Color.Black);
            //設定 圖表的 邊框樣式 (實線)
            myBook.ActiveChart.ChartArea.Border.LineStyle = Excel.XlLineStyle.xlContinuous;

            //設定 圖例項目 的 背景色彩
            myBook.ActiveChart.Legend.Interior.Color = ColorTranslator.ToOle(Color.White);
            myBook.ActiveChart.Legend.Width = 55;        //設定 圖例 的 寬度
            myBook.ActiveChart.Legend.Height = 20;       //設定 圖例 的 高度
            myBook.ActiveChart.Legend.Font.Size = 11;    //設定 圖例 的 字體大小 
            myBook.ActiveChart.Legend.Font.Bold = true;  //設定 圖例 的 字體樣式=粗體
            myBook.ActiveChart.Legend.Font.Name = "細明體";//設定 圖例 的 字體字型=細明體
            myBook.ActiveChart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;//設訂 圖例 的 位置靠下
            myBook.ActiveChart.Legend.Border.LineStyle = Excel.XlLineStyle.xlDash;//設定 圖例 的 邊框線條

            //設定 圖表 x 軸 內容
            Excel.Axis xAxis = (Excel.Axis)myBook.ActiveChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            //設定 圖表 x軸 橫向線條 線條樣式
            xAxis.MajorGridlines.Border.LineStyle = Excel.XlLineStyle.xlContinuous;
            //設定 圖表 x軸 橫向線條顏色
            xAxis.MajorGridlines.Border.Color = ColorTranslator.ToOle(Color.Gray);
            xAxis.HasTitle = false;  //設定 x軸 座標軸標題 = false(不顯示)，不打就是不顯示
            xAxis.TickLabels.Font.Name = "標楷體"; //設定 x軸 字體字型=標楷體
            xAxis.TickLabels.Font.Size = 12;       //設定 x軸 字體大小
            xAxis.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone; //座標軸邊框設定無線條

            //設定 圖表 y軸 內容
            Excel.Axis yAxis = (Excel.Axis)myBook.ActiveChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
            yAxis.TickLabels.Font.Name = "標楷體"; //設定 y軸 字體字型=標楷體 
            yAxis.TickLabels.Font.Size = 12;       //設定 y軸 字體大小
            yAxis.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone; //座標軸邊框設定無線條


            //設定 圖表 標題 顯示 = false(關閉)
            myBook.ActiveChart.HasTitle = true;
            //設定 圖表 標題 = 匯率
            myBook.ActiveChart.ChartTitle.Text = sTitle;
            //設定 圖表 標題 陰影 = false(關閉)
            myBook.ActiveChart.ChartTitle.Shadow = false;
            //設定 圖表 標題 邊框樣式
            myBook.ActiveChart.ChartTitle.Border.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
        }
    }
}
