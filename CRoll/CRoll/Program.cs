using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CRoll
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            //总图片空间10M，临时数据空间3M
            DataClass.msFile = new System.IO.MemoryStream(20000000);
            DataClass.msTmpFile = new System.IO.MemoryStream(5000000);
            DataClass.frmMain = new FrmMain();
            Application.Run(DataClass.frmMain);
            //Application.Run(new FrmMain());
           
        }
    }
}
