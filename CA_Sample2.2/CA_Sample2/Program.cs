using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace CA_Sample2
{
    class Program
    {
        // const values
        public static class Define
        {
            public const string csStrExcelExtensions = ".xlsx";
            public const string csStrColon = ":";
            public const string csStrUnderbar = "_";
            public const string csStrDateTimeFmt = "yyyyMMddHHmmss";
            public const string csStrMsgForBasicFile = "コピー元ファイル：";
            public const string csStrMsgForCopyFile = "コピー先ファイル：";
            public const string csStrMsgInputEnterKey = "Enterを入力してください";
            public const string csStrErrMsgForNotFindFile = "Excelファイルが見つかりません。";
        }

        // ファイル取得処理
        static string GetFileName()
        {
            //配列の先頭には実行ファイルのパスが入っているので、2番目以降がドロップされたファイルのパスになる
            string[] files = System.Environment.GetCommandLineArgs();
            if (files.Length <= 1 || System.IO.Path.GetExtension(@files[1]) != Define.csStrExcelExtensions)
            {
                return "";
            }
            return files[1];
        }

        // メイン処理
        static void Main(string[] args)
        {
            int iBasicStartTime = 600;
            int iBasicEndTime = 1140;
            string strBeforeFileNm = null;
            string strBeforeDir = null;
            string strBeforePath = null;
            string strAfterFileNm = null;
            string strAfterDir = null;
            string strAfterPath = null;

            Random cRandom = new System.Random();

            strBeforePath = GetFileName();
            //strBeforePath = "C:\\Users\\keitashinohara\\Desktop\\2017年2月勤務管理表（専門職）.xlsx";
            if (strBeforePath == "")
            {
                Console.WriteLine(Define.csStrErrMsgForNotFindFile);
                Console.WriteLine(Define.csStrMsgInputEnterKey);
                Console.ReadLine();
                return;
            }

            strBeforeFileNm = System.IO.Path.GetFileName(strBeforePath);
            strBeforeDir = System.IO.Path.GetDirectoryName(strBeforePath);
            strAfterFileNm = DateTime.Now.ToString(Define.csStrDateTimeFmt) + Define.csStrUnderbar + Environment.UserName + Define.csStrUnderbar + System.IO.Path.GetFileName(strBeforePath);
            strAfterDir = System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            strAfterPath = System.IO.Path.Combine(strAfterDir, strAfterFileNm);

            // Excel操作用オブジェクト
            Microsoft.Office.Interop.Excel.Application xlApp = null;
            Microsoft.Office.Interop.Excel.Workbooks xlBooks = null;
            Microsoft.Office.Interop.Excel.Workbook xlBook = null;
            Microsoft.Office.Interop.Excel.Sheets xlSheets = null;
            Microsoft.Office.Interop.Excel.Worksheet xlSheet = null;
            Microsoft.Office.Interop.Excel.Range xlRange1 = null;
            Microsoft.Office.Interop.Excel.Range xlRange2 = null;
            Microsoft.Office.Interop.Excel.Range xlRange3 = null;
            Microsoft.Office.Interop.Excel.Range xlCells = null;

            // Excelアプリケーション生成
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            // Excelファイルの警告を無視する
            xlApp.DisplayAlerts = false;
            // 対象のExcelブックを開く
            xlBooks = xlApp.Workbooks;
            xlBook = xlBooks.Open(strBeforePath);
            // シートを選択する
            xlSheets = xlBook.Worksheets;
            xlSheet = xlSheets[2] as Microsoft.Office.Interop.Excel.Worksheet;
            // 表示
            xlApp.Visible = false;

            try
            {
                xlCells = xlSheet.Cells;

                // Excelファイル編集処理
                // ループ処理は、C9～C40で決め打ち                   
                for (int iRowCnt = 9; iRowCnt <= 40; iRowCnt++)
                {
                    xlRange1 = xlCells[iRowCnt, 1] as Microsoft.Office.Interop.Excel.Range;
                    xlRange2 = xlCells[iRowCnt, 3] as Microsoft.Office.Interop.Excel.Range;
                    xlRange3 = xlCells[iRowCnt, 4] as Microsoft.Office.Interop.Excel.Range;

                    // Aカラムに日にちが入力されていない場合、ループを抜ける
                    if (xlRange1.Value == null || xlRange1.Value.ToString() == "") { break; }

                    // 条件１：行の背景色が設定されていない場合、下記の時間計算処理を行う。
                    if (xlRange1.Interior.ColorIndex > 0) { continue; }

                    // 出社時間・帰宅時間を計算
                    //Random cRandom = new System.Random();
                    int iWorkTimeForStart = iBasicStartTime + cRandom.Next(-60, 60);
                    int iWorkTimeForEnd = iBasicEndTime + cRandom.Next(-60, 60);
                    string strStartTime = (iWorkTimeForStart / 60).ToString() + Define.csStrColon + (iWorkTimeForStart % 60).ToString();
                    string strEndTime = (iWorkTimeForEnd / 60).ToString() + Define.csStrColon + (iWorkTimeForEnd % 60).ToString();

                    // 出社時間・帰宅時間を入力
                    // 条件２：対象のセルに値が入っていない場合、上記で計算した時間を入力
                    xlRange2.Value = (xlRange2.Value == null || xlRange2.Value.ToString() == "") ? strStartTime : xlRange2.Value;
                    xlRange3.Value = (xlRange3.Value == null || xlRange3.Value.ToString() == "") ? strEndTime : xlRange3.Value;
                }
            }
            finally
            {
                // Excelファイルの警告を戻す
                //xlApp.DisplayAlerts = true;

                // 名前を付けてデスクトップに保存
                xlBook.SaveAs(@strAfterPath);

                // Excelプロセス解放
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange1);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange2);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange3);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCells);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                xlBook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                // コンソールに、元ファイルのパスとコピーファイルのパスを表示
                Console.WriteLine(Define.csStrMsgForBasicFile + strBeforePath);
                Console.WriteLine(Define.csStrMsgForCopyFile + strAfterPath);
                Console.WriteLine(Define.csStrMsgInputEnterKey);
                Console.ReadLine();
            }
        }
    }
}
