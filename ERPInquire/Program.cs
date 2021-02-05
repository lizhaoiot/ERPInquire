using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Threading;

namespace ERPInquire
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            //加载空间中文字库
            Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("zh-cn");

            //隐藏文件夹
            DirectoryInfo dirInfo1 = new DirectoryInfo("data");
            dirInfo1.Attributes = FileAttributes.Hidden;
            DirectoryInfo dirInfo2 = new DirectoryInfo("Template");
            dirInfo2.Attributes = FileAttributes.Hidden;
            DirectoryInfo dirInfo3 = new DirectoryInfo("Log");
            dirInfo3.Attributes = FileAttributes.Hidden;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainFrm());
        }
    }
}
