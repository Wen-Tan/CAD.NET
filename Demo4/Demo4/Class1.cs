using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace AddCircle
{
    public class AddCircle
    {
        [CommandMethod("AddCircle")]
        public static void addCircle()
        {
            // 获得当前文档和数据库 
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // 启动一个事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                //以只读方式打开块表
                // 定义一个块表
                BlockTable acBlkTb1;
                acBlkTb1 = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                // 以写方式打开模型空间块表记录 
                // 定义一个存放圆的块表记录
                BlockTableRecord acBlkTb1Reccircle;
                //定义一个存放圆的块表记录
                BlockTableRecord acBlkTb1Recarc;
                acBlkTb1Reccircle = acTrans.GetObject(acBlkTb1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                
                acBlkTb1Recarc = acTrans.GetObject(acBlkTb1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                // 创建一个中心点在 (2,3,0) ，半径为4.25 的圆
                //实例一个圆
                Circle acCirc = new Circle();
                acCirc.SetDatabaseDefaults();
                //圆心
                acCirc.Center = new Point3d(2, 3, 0);
                //半径
                acCirc.Radius = 4.25;

                // 创建一个中心点在 (6.25,9.125,0)，半径为6，起始角度为1.117（64度），终点角度为3.5605（204度）的圆弧。
                Arc acArc = new Arc(new Point3d(6.25, 9.125, 0), 6, 1.117, 3.5605);
                //添加新对象到块表记录和事务中
                acBlkTb1Reccircle.AppendEntity(acCirc);
                acBlkTb1Recarc.AppendEntity(acArc);
                acTrans.AddNewlyCreatedDBObject(acCirc, true);
                acTrans.AddNewlyCreatedDBObject(acArc, true);

                acTrans.Commit();
            }
        }
    }
}
