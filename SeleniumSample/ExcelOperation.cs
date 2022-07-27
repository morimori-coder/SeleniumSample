using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace SeleniumSample
{
    internal class ExcelOperation
    {
        // コード引用元 ：https://qiita.com/hakuaneko/items/332cf7dd9fcd70ccc052
        public void Excel_OutPutEx(List<string> texts)
        {
            //Excelオブジェクトの初期化
            Excel.Application ExcelApp = null;
            Excel.Workbooks wbs = null;
            Excel.Workbook wb = null;
            Excel.Sheets shs = null;
            Excel.Worksheet ws = null;

            try
            {
                //Excelシートのインスタンスを作る
                ExcelApp = new Excel.Application();
                wbs = ExcelApp.Workbooks;
                wb = wbs.Add();

                shs = wb.Sheets;
                ws = shs[1];
                ws.Select(Type.Missing);

                ExcelApp.Visible = false;

                // エクセルファイルにデータをセットする
                int counter = 1;
                foreach(var text in texts)
                {
                    // Excelのcell指定
                    Excel.Range w_rgn = ws.Cells;
                    Excel.Range rgn = w_rgn[counter, 1];

                    try
                    {
                        // Excelにデータをセット
                        rgn.Value2 = text;
                    }
                    finally
                    {
                        // Excelのオブジェクトはループごとに開放する
                        Marshal.ReleaseComObject(w_rgn);
                        Marshal.ReleaseComObject(rgn);
                        w_rgn = null;
                        rgn = null;
                    }
                    counter++;
                }

                //excelファイルの保存
                wb.SaveAs(@"ファイルパス");
                wb.Close(false);
                ExcelApp.Quit();
            }
            finally
            {
                //Excelのオブジェクトを開放し忘れているとプロセスが落ちないため注意
                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(shs);
                Marshal.ReleaseComObject(wb);
                Marshal.ReleaseComObject(wbs);
                Marshal.ReleaseComObject(ExcelApp);
                ws = null;
                shs = null;
                wb = null;
                wbs = null;
                ExcelApp = null;

                GC.Collect();
            }
        }


    }
}
