using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic
{

    public partial class ThisAddIn
    {
        public static class Global
        {
            //是否执行重命名、删除、移动文件等功能标识,0为关闭，1为打开
            public static int readFile;

            //聚光灯功能开关标识，0为关闭，1为打开
            public static int spotlight;
            public static int spotlightColorIndex;

            //聚光灯功能打开前保存的彩色单元格位置和填充颜色
            public static Dictionary<string, int> cellColor = new Dictionary<string, int>();

            public static bool created_qr_sheet = false;

            public static Dictionary<string, int> GetColorDictionary(Excel.Range usedRange)
            {
                Dictionary<string, int> cellColorDict = new Dictionary<string, int>();
                foreach (Excel.Range cell in usedRange)
                {
                    if (cell.Interior.ColorIndex > 0)
                    {
                        string cellAddress = cell.Address;
                        int cellColorIndex = cell.Interior.ColorIndex;
                        cellColorDict.Add(cellAddress, cellColorIndex);
                    }
                }
                return cellColorDict;
            }
        }

        public static Excel.Application app;         //声明一个Excel的Application变量
        public static string hotKey;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 兜底程序集解析：安装环境下加载项的 .dll.config 绑定重定向可能没有生效，
            // 这里按本地目录中的实际版本解析（例如 Npgsql 依赖的 System.Text.Json 8.x -> 9.x）
            AppDomain.CurrentDomain.AssemblyResolve += CurrentDomain_AssemblyResolve;

            // 最后一道防线：VSTO 宿主下未处理异常会直接终止 Excel 进程。
            // 这里把 WinForms UI 线程异常和 AppDomain 异常都拦下并写日志，
            // 避免个别事件处理器的问题导致整个 Excel 退出。
            try
            {
                System.Windows.Forms.Application.SetUnhandledExceptionMode(
                    System.Windows.Forms.UnhandledExceptionMode.CatchException);
            }
            catch { }
            System.Windows.Forms.Application.ThreadException += Application_ThreadException;
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            app = Globals.ThisAddIn.Application;
            Global.spotlight = 0;

            //监听选定单元格变化所触发事件
            app.SheetSelectionChange += app_SheetSelectionChange;

            app.SheetBeforeDelete += new Excel.AppEvents_SheetBeforeDeleteEventHandler(Application_SheetBeforeDelete);

        }

        /// <summary>
        /// WinForms UI 线程未处理异常的兜底处理：只写日志，不再让异常终结 Excel。
        /// </summary>
        private void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            try
            {
                DbHelper.LogError("WinForms 线程未处理异常", e.Exception);
            }
            catch { }
        }

        /// <summary>
        /// AppDomain 级未处理异常的兜底处理（例如后台线程异常）。
        /// 注意：此类异常无法阻止进程终止，但至少能留下一份可排查的日志。
        /// </summary>
        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                var ex = e.ExceptionObject as Exception;
                DbHelper.LogError("AppDomain 未处理异常（IsTerminating=" + e.IsTerminating + "）",
                    ex ?? new Exception(e.ExceptionObject == null ? "未知异常对象" : e.ExceptionObject.ToString()));
            }
            catch { }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        /// <summary>
        /// 当运行时找不到指定版本的依赖程序集时，用加载项目录中的同名程序集兜底。
        /// 这样即使加载项的 .dll.config 没有随安装包部署、或 VSTO 没有加载绑定重定向，
        /// Npgsql 等第三方库也能正常加载（而不是在 try 之外抛出被宿主吞掉的异常）。
        /// </summary>
        private static Assembly CurrentDomain_AssemblyResolve(object sender, ResolveEventArgs args)
        {
            try
            {
                var requested = new AssemblyName(args.Name);
                if (string.IsNullOrEmpty(requested.Name))
                {
                    return null;
                }

                // 同名的程序集已经加载过，直接复用
                foreach (var loaded in AppDomain.CurrentDomain.GetAssemblies())
                {
                    if (string.Equals(loaded.GetName().Name, requested.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        return loaded;
                    }
                }

                string baseDir = Path.GetDirectoryName(typeof(ThisAddIn).Assembly.Location);
                if (string.IsNullOrEmpty(baseDir))
                {
                    return null;
                }

                string assemblyPath = Path.Combine(baseDir, requested.Name + ".dll");
                if (File.Exists(assemblyPath))
                {
                    return Assembly.LoadFrom(assemblyPath);
                }
            }
            catch
            {
                // 解析失败时交回运行时处理
            }

            return null;
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        //选定单元格变化事件
        private void app_SheetSelectionChange(object sh, Excel.Range Target)
        {
            //当聚光灯功能打开时，变更选定单元格即触发
            if (Global.spotlight == 1)
            {
                Excel.Range selectedRange = Target;
                Excel.Worksheet activesheet = ThisAddIn.app.ActiveSheet;

                //如果选择多个单元格时退出事件，选择1个单元格时触发事件
                if (selectedRange.Count != 1)
                {
                    return;
                }
                else
                {
                    activesheet.Cells.Interior.ColorIndex = 0;
                    ThisAddIn.app.ScreenUpdating = false;
                    selectedRange.EntireRow.Interior.ColorIndex = Global.spotlightColorIndex;
                    selectedRange.EntireColumn.Interior.ColorIndex = Global.spotlightColorIndex;
                    //selectedRange.EntireRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204)); ;
                    //selectedRange.EntireColumn.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(204, 255, 204)); ;
                    ThisAddIn.app.ScreenUpdating = true;
                }
            }
        }

        //监听删除表事件
        private void Application_SheetBeforeDelete(object Sh)
        {
            Excel.Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet ws = Globals.ThisAddIn.Application.ActiveSheet;
            if (ws.Name == "_rename")
            {
                Global.readFile = 0;
            }
            if (ws.Name == "_QR_Scan")
            {
                Global.created_qr_sheet = false;
            }
        }
    }
}
