using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;



namespace Demo2
{
    public class Class1
    {
        [CommandMethod("NewLine")]
        public static void NewLine()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                //获取模型空间
                BlockTable blockTb1 = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                BlockTableRecord modelSpace = tr.GetObject(blockTb1[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                //创建直线
                Line line = new Line();
                line.StartPoint = new Point3d(0, 0, 0);
                line.EndPoint = new Point3d(1000, 1000, 0);

                //图层：test_layer
                string layerName = "test_layer";
                LayerTable layerTbl = tr.GetObject(db.LayerTableId, OpenMode.ForRead) as LayerTable;
                if (layerTbl.Has(layerName))
                {
                    line.Layer = layerName;
                }

                //-------------------------------
                // 颜色
                //-------------------------------
                // 颜色索引
                line.Color = Color.FromColorIndex(ColorMethod.ByAci, 4);
                // 随块
                line.Color = Color.FromColorIndex(ColorMethod.ByBlock, 0);
                // 随层
                line.Color = Color.FromColorIndex(ColorMethod.ByLayer, 256);
                // RGB
                line.Color = Color.FromRgb(255, 120, blue: 237);

                //-------------------------------
                // 线型: Center
                //-------------------------------
                string linetypeName = "Center";
                LinetypeTable linetypeTbl = tr.GetObject(db.LinetypeTableId, OpenMode.ForRead) as LinetypeTable;
                if (!linetypeTbl.Has(linetypeName))
                {
                    db.LoadLineTypeFile(linetypeName, "acad.lin");
                }
                if (linetypeTbl.Has(linetypeName))
                {
                    line.Linetype = linetypeName;
                }

                //-------------------------------
                // 线宽
                //-------------------------------
                line.LineWeight = LineWeight.ByLayer;
                line.LineWeight = LineWeight.ByBlock;
                line.LineWeight = LineWeight.ByLineWeightDefault;
                line.LineWeight = LineWeight.LineWeight211;

                //-------------------------------
                // 添加到模型空间并提交到数据库
                //-------------------------------
                modelSpace.AppendEntity(line);
                tr.AddNewlyCreatedDBObject(line, true);
                tr.Commit();
            }
        }

    }
}
