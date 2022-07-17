using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//新添加的头文件
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

namespace Demo1
{
    public class Class1
    {
        //加一个命令
        [CommandMethod("demo1")]
        //写函数
        public void demo1()
        {
            //声明命令行对象
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("我是无敌小可爱");
        }
    }
}
