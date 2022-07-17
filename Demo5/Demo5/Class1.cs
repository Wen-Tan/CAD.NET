using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace Texts
{
    public class Texts
    {
        [CommandMethod("AddText")]
        public void AddText()
        {
            Database db = HostApplicationServices.WorkingDatabase;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                // 以只读方式打开块表
                BlockTable BlkTb1;
                BlkTb1 = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                // 以写方式打开模型空间块表记录
                BlockTableRecord BlkTb1Rec;
                BlkTb1Rec = trans.GetObject(BlkTb1[BlockTableRecord.ModelSpace],OpenMode.ForWrite) as BlockTableRecord;

                //创建第一个单行文字
                DBText firsttext = new DBText();
                //文字位置为原点
                firsttext.Position = new Point3d();
                //文字高度
                firsttext.Height = 15;

                //文字内容
                firsttext.TextString = "This is firsttext!";
                //水平垂直居中
                firsttext.HorizontalMode = TextHorizontalMode.TextCenter;
                firsttext.VerticalMode = TextVerticalMode.TextVerticalMid;
                //对齐点
                firsttext.AlignmentPoint = firsttext.Position;
                BlkTb1Rec.AppendEntity(firsttext);
                trans.AddNewlyCreatedDBObject(firsttext, true);

                trans.Commit();
            }
        }
    }
}
