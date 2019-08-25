using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFontChanger
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length < 2)
            {
                return 0;
            }

            // EXCELファイルパス
            string filePath = args[0];
            // EXCELファイルフルパス
            string fullPath = System.IO.Path.GetFullPath(filePath);
            // ファイルが無い場合終了
            if (!System.IO.File.Exists(fullPath))
            {
                return 0;
            }

            // シート名
            string strWorksheetName = "";
            // フォント名
            string fontName = args[1];
            //末端の改行文字を取り除く
            fontName.Trim();

            Microsoft.Office.Interop.Excel.Application xlApp = null;
            Microsoft.Office.Interop.Excel.Workbooks xlBooks = null;
            Microsoft.Office.Interop.Excel.Workbook xlBook = null;
            Microsoft.Office.Interop.Excel.Sheets xlSheets = null;
            Microsoft.Office.Interop.Excel.Worksheet xlSheet = null;
            Microsoft.Office.Interop.Excel.Range xlRange = null;
            Microsoft.Office.Interop.Excel.Font xlFont = null;
            Microsoft.Office.Interop.Excel.Shapes xlShapes = null;
            Microsoft.Office.Interop.Excel.Shape xlShape = null;
            Microsoft.Office.Interop.Excel.TextFrame2 xlTextFrame2 = null;
            Microsoft.Office.Core.TextRange2 xlTextRange2 = null;
            Microsoft.Office.Core.Font2 xlFont2 = null;

            try
            {
                // EXCEL起動
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                // EXCELは非表示にする
                xlApp.Visible = false;
                // 保存確認なしにする
                xlApp.DisplayAlerts = false;
                // Workbook開く
                xlBooks = xlApp.Workbooks;
                xlBook = xlBooks.Open(fullPath);
                // Worksheet開く
                xlSheets = xlBook.Sheets;
                for (int ii = 1; ii <= xlSheets.Count; ++ii)
                {
                    xlSheet = xlSheets.Item[ii];
                    strWorksheetName = xlSheet.Name;

                    // セルのフォントを変更する
                    xlRange = xlSheet.UsedRange;
                    xlFont = xlRange.Font;
                    xlFont.Name = fontName;
                    if (xlFont != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlFont);
                        xlFont = null;
                    }
                    if (xlRange != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                        xlRange = null;
                    }

                    // 図形のフォントを変更する
                    xlShapes = xlSheet.Shapes;
                    for (int jj = 1; jj <= xlShapes.Count; ++jj)
                    {
                        xlShape = xlShapes.Item(jj);
                        xlTextFrame2 = xlShape.TextFrame2;
                        if (xlTextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            xlTextRange2 = xlTextFrame2.TextRange;
                            xlFont2 = xlTextRange2.Font;

                            // 右から左へ記述するフォントの設定(例:アラビア語)
                            xlFont2.NameComplexScript = fontName;
                            // 全角フォントの設定(例:日本語)
                            xlFont2.NameFarEast = fontName;
                            // 半角フォントの設定(例:英語)
                            xlFont2.Name = fontName;

                            xlRange = xlShape.TopLeftCell;
                            System.Console.WriteLine(xlRange.Address + " [" + xlShape.Name + "]" + "のフォントを変更しました。");
                            if (xlRange != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                                xlRange = null;
                            }

                            if (xlFont2 != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlFont2);
                                xlFont2 = null;
                            }
                            if (xlTextRange2 != null)
                            {
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlTextRange2);
                                xlTextRange2 = null;
                            }
                        }
                        if (xlTextFrame2 != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlTextFrame2);
                            xlTextFrame2 = null;
                        }
                        if (xlShape != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlShape);
                            xlShape = null;
                        }
                    }
                    if (xlShapes != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlShapes);
                        xlShapes = null;
                    }
                    if (xlSheet != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                        xlSheet = null;
                    }
                    System.Console.WriteLine("処理が終了しました。シート名:" + strWorksheetName);
                }
                if (xlSheets != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                    xlSheets = null;
                }

                // Workbook保存
                xlBook.Save();
                System.Console.WriteLine("処理が終了しました。ファイル名:" + xlBook.Name);
            }
            catch
            {
                System.Console.WriteLine("ランタイムエラーが発生しました。シート名:" + strWorksheetName);
            }
            finally
            {
                if (xlFont2 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlFont2);
                    xlFont2 = null;
                }

                if (xlTextRange2 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlTextRange2);
                    xlTextRange2 = null;
                }

                if (xlTextFrame2 != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlTextFrame2);
                    xlTextFrame2 = null;
                }

                if (xlShape != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlShape);
                    xlShape = null;
                }

                if (xlShapes != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlShapes);
                    xlShapes = null;
                }

                if (xlFont != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlFont);
                    xlFont = null;
                }

                if (xlRange != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                    xlRange = null;
                }

                if (xlSheet != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                    xlSheet = null;
                }

                if (xlSheets != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);
                    xlSheets = null;
                }

                if (xlBook != null)
                {
                    xlBook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                    xlBook = null;
                }

                if (xlBooks != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);
                    xlBooks = null;
                }

                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                    xlApp = null;
                }
            }
            return 0;
        }
    }
}
