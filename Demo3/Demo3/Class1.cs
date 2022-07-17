using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace Lines
{
    public class Lines
    {
        [CommandMethod("FirstLine")]
        public static void FirstLine()
        {
            //获取当前活动图形数据库
            Database db = HostApplicationServices.WorkingDatabase;
            //直线起点
            Point3d startPoint = new Point3d(0, 100, 0);
            //直线终点
            Point3d endPoint = new Point3d(100, 100, 0);
            //新建一直线对象
            Line line = new Line(startPoint, endPoint);
            //定义一个指向当前数据库的事务处理，以添加直线
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //以读方式打开块表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                //以写方式打开模型空间块表记录
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                //将图形对象的信息添加到块表记录中,并返回ObjectId对象
                btr.AppendEntity(line);
                //把对象添加到事务处理中
                trans.AddNewlyCreatedDBObject(line, true);
                //事务提交
                trans.Commit();
            }
        }
    }
}
