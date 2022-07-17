using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;



namespace 求一个钻孔和等高线的交点
{
    public class Class1
    {
        [CommandMethod("PLVCG")]
        //需要在命令行输入这个命令PLVCG


        public void PLineVertexCoordsGet()
        {
            //需要访问Database的操作 需首先将该文档进行锁定，操作完成后，在最后进行释放
            DocumentLock docLock = Application.DocumentManager.MdiActiveDocument.LockDocument();
            // 对话框窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 数据库对象
            Database db = HostApplicationServices.WorkingDatabase;

            try
            {
                // 开启事务处理
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;//根据cad里面的结构定义的对象，主要还是entity还有entity下面的各类线段
                    BlockTableRecord btr = trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    Entity entity1 = null;
                    Point3dCollection points = new Point3dCollection();//其实也可以创建二维的，如果为了运算方便的话可以改为二维的。
                    Point3dCollection points1 = new Point3dCollection();
                    Polyline3d denggaoxian = new Polyline3d();
                    List<double> zcha = new List<double>();
                    List<double> zhenliex = new List<double>();
                    List<double> zhenliey = new List<double>();
                    List<double> zhenliez = new List<double>();
                    List<Point3d> jiaodian = new List<Point3d>();
                    double zcha1 = 0, zhenliex1 = 0, zhenliey1 = 0, zhenliez1 = 0;
                    double angle_zhibeizhen = 0, angle_zuankong = 0, angle = 0;
                    //计算指北针角度
                    angle_zhibeizhen = getCompassAngle(btr, trans);
                    //找钻孔
                    entity1 = SearchDrill(btr, trans);
                    //钻孔角度
                    angle_zuankong = getDrillAngle(entity1);
                    //找到钻孔直线之后，找等高线，其实可以写在一起作为一个双重循环的。分开写的话便于理解。
                    isoheight(btr, trans, zcha, zhenliex, zhenliey, zhenliez, entity1, points, points1);
                    zhenliex.Sort();
                    zhenliey.Sort();
                    //输出所有交点或进行其他处理。目前能想到的方法是先做两次找交点（一次是都不延申的points1，一次是只延申原实体的points2），然后对points2进行排序或查找操作。
                    Process_jiaodian(points, points1, jiaodian, zhenliez);
                    //计算巷道间隔
                    double hangdao_jiange = getHangdaojiange(points1);
                    //生成等高线
                    zhenliex1 = System.Math.Abs(zhenliex[zhenliex.Count - 1] - zhenliex[0]);
                    zcha.Sort();
                    List<double> newzcha = zcha.Distinct().ToList();
                    double s = System.Math.Abs(newzcha[newzcha.Count - 1] - zcha[0]);
                    GenerateIsoheight(btr, trans, newzcha, s, zhenliex1);

                    List<Point3d> jiaodian1 = new List<Point3d>();
                    Point3dCollection point3d2 = new Point3dCollection();
                    Point3dCollection point3d22 = new Point3dCollection();
                    Point3dCollection point3d = new Point3dCollection();
                    //计算图例中煤层与等高线的交点
                    Calculate_jiaodian(btr, trans, jiaodian, jiaodian1, point3d);
                    //Polyline3d poly = new Polyline3d(Poly3dType.SimplePoly, point3d, true) ;


                    jiaodian1.Sort((x, y) => { return x.X.CompareTo(y.X); });
                    //上下煤层拟合点
                    CoalSeam(jiaodian1, point3d2, point3d22);
                    //Spline spline = new Spline(point3d2,vecTan,vecTan ,4, 0.0);
                    Spline spline = new Spline(point3d2, 4, 0.0);
                    Spline spline1 = new Spline(point3d22, 4, 0.0);
                    //在CAD中生成煤层
                    GenerateCoalSeam(btr, trans, spline, spline1);
                    //poly.Layer = "0";
                    //spline.Layer = "0";
                    //spline1.Layer = "0";
                    /*btr.AppendEntity(poly);
                    trans.AddNewlyCreatedDBObject(poly, true);*/

                    /*foreach (Point3d po in point3d2)
                    {
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n多段线点{0}", po));
                    }*/

                    Line zuankong = entity1 as Line;
                    //计算巷道坐标
                    double zhangdao = CalculHangdaoz(jiaodian1, zuankong);
                    double yhangdao = 800 - s + zhangdao - newzcha[0];
                    double xhangdao = zuankong.StartPoint.X;

                    //生成巷道
                    GenerateTunnel(btr, trans, zhangdao, xhangdao, yhangdao, spline, hangdao_jiange);

                    //计算钻孔坐标
                    double zuankong_z = CalculDrillz(jiaodian1, zuankong);
                    double zuankong_x = zuankong.EndPoint.X;
                    double zuankong_y = 800 - s + zuankong_z - newzcha[0] + 1.3;

                    //生成钻孔
                    Point3d zuankongEnd = new Point3d(zuankong_x, zuankong_y, 0);
                    Point3d zuankongStart = new Point3d(xhangdao, yhangdao + 1.3, 0);
                    Line zuankong_line = new Line(zuankongStart, zuankongEnd);
                    GenerateDrill(btr, trans, zuankongEnd, zuankongStart, zuankong_line);

                    //计算方位角
                    angle = CalculAzimuth(angle_zhibeizhen, angle_zuankong);

                    //计算煤层倾角
                    double tanAngleValue = CalculDip(zuankongEnd, zuankongStart);

                    //插入文字到指定位置
                    InsertText(btr, trans, angle, tanAngleValue, zuankong_line, zhenliex1);

                    /*Line hangdaoline = new Line(phangdao1, phangdao2);
                    hangdaoline.Layer = "0";
                    btr.AppendEntity(hangdaoline);
                    trans.AddNewlyCreatedDBObject(hangdaoline, true);*/

                    //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n钻孔起点{0}", zuankong.StartPoint));
                    //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n钻孔终点{0}", zuankong.EndPoint));
                    /*for (int i = 0; i < zcha.Count; i++)
                    {
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n等高线z轴{0}", zcha[i]));
                        if (i > 0)
                        {
                            if (zcha1 < System.Math.Abs((zcha[i] - zcha[i - 1])))
                            {
                                zcha1 = System.Math.Abs((zcha[i] - zcha[i - 1]));
                            }
                        }
                    }
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\nz轴差值最大{0}", zcha1));*/
                    trans.Commit();
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                Application.ShowAlertDialog(e.Message);
            }
            // 解锁文档
            docLock.Dispose();
        }

        public void Test()
        {
            //需要访问Database的操作 需首先将该文档进行锁定，操作完成后，在最后进行释放
            DocumentLock docLock = Application.DocumentManager.MdiActiveDocument.LockDocument();
            // 对话框窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 数据库对象
            Database db = HostApplicationServices.WorkingDatabase;
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\ntest111"));

            docLock.Dispose();
        }
        //计算指北针角度
        public double getCompassAngle(BlockTableRecord btr, Transaction trans)
        {
            double angle_zhibeizhen = 0;
            foreach (ObjectId id in btr)
            {

                Entity entity = (Entity)trans.GetObject(id, OpenMode.ForRead);//实体化对象，entity对象非常关键
                //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", entity.GetType()));
                /*if(entity.BlockName == "指北针2" && entity.GetType() == typeof(Line))
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", entity.BlockName));
                    Line zhibeizhen = entity as Line;
                    angle_zhibeizhen = zhibeizhen.Angle;
                }*/
                if (entity.GetType() == typeof(BlockReference))
                {
                    BlockReference blockReference = entity as BlockReference;
                    if (blockReference.Name == "指北针4")
                    {
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", blockReference.Name));
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", blockReference.Rotation));
                        angle_zhibeizhen = blockReference.Rotation / Math.PI * 180;
                    }
                    //BlockReference br = (BlockReference)trans.GetObject(id, OpenMode.ForRead);

                }
                /*if (entity.ColorIndex == 3 && entity.GetType() == typeof(Line))
                {
                    Line zhibeizhen = entity as Line;
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", zhibeizhen.StartPoint));
                    angle_zhibeizhen = zhibeizhen.Angle;
                }*/
            }
            return angle_zhibeizhen;
        }
        //返回钻孔实体
        public Entity SearchDrill(BlockTableRecord btr, Transaction trans)
        {
            foreach (ObjectId id in btr)
            {
                //Access to the entity
                //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("{0}\n", entity.GetType()));
                Entity entity = (Entity)trans.GetObject(id, OpenMode.ForRead);//实体化对象，entity对象非常关键
                String layername = entity.Layer;
                //找钻孔
                if (layername == "设计校检钻孔")//找钻孔的图层
                {
                    //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", entity.Layer));//在cad命令行上面打印出图层信息，除了打印交点部分，其他打印的部分都可以注释掉
                    //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("{0}", entity.ColorIndex));//打印颜色索引，这个部分要注意，有的索引及其不规范。。。。。。
                    if ((entity.GetType() == typeof(Line)) && (entity.ColorIndex != 5))//找类型是直线并且颜色是白色的的实体，有的颜色索引是0，255也是白色，但是其实是不一样的，一定要注意
                    {
                        //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("!!!"));

                        return entity;
                    }
                }
            }
            return null;
        }
        //钻孔角度
        public double getDrillAngle(Entity entity)
        {
            double angle_zuankong = 0;
            Line line1 = entity as Line;//定义一个直线对象，这个比较关键，虽然看着没啥
            angle_zuankong = line1.Angle;
            //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("{0}",line1.StartPoint));//输出这条直线的起点信息。
            return angle_zuankong;
        }
        //等高线
        public void isoheight(BlockTableRecord btr, Transaction trans, List<double> zcha, List<double> zhenliex, List<double> zhenliey, List<double> zhenliez, Entity entity1, Point3dCollection points, Point3dCollection points1)
        {
            foreach (ObjectId id in btr)
            {
                Entity entity = (Entity)trans.GetObject(id, OpenMode.ForRead);
                //Access to the entity
                //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("{0}\n", entity.GetType()));
                String layername = entity.Layer;
                if (layername == "223等高线")
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", entity.Layer));
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("{0}", entity.ColorIndex));
                    Polyline3d polyline = entity as Polyline3d;
                    zcha.Add(polyline.StartPoint.Z);
                    zhenliex.Add(polyline.StartPoint.X);
                    zhenliey.Add(polyline.StartPoint.Y);
                    zhenliez.Add(polyline.StartPoint.Z);
                    if (entity1 != null)
                    {
                        //entity1.IntersectWith(entity, Intersect.OnBothOperands, new Plane(), points, IntPtr.Zero, IntPtr.Zero);
                        entity1.IntersectWith(entity, Intersect.ExtendThis, new Plane(), points, IntPtr.Zero, IntPtr.Zero);
                        //这个部分也非常非常的关键，里面的参数也很多，一定要注意不同场景的用法。
                        /*
                         OnBothOperands :两个实体（这里是线段）都不延申
                        ExtendBoth :两个实体都延申
                        ExtendArgument :只延申作为参数的实体（第一个参数entity）
                        ExtendThis :只延申原实体（entity1）
                         */
                        //points是得到的交点信息。后面要跟两个IntPtr.Zero，如果输入0的话，是已经弃用的方法，不推荐。
                    }
                }
                else if (layername == "133煤巷")
                {
                    //entity1.IntersectWith(entity, Intersect.OnBothOperands, new Plane(), points1, IntPtr.Zero, IntPtr.Zero);
                    entity1.IntersectWith(entity, Intersect.ExtendThis, new Plane(), points1, IntPtr.Zero, IntPtr.Zero);
                }
            }
            /*foreach (Point3d po in points1)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n巷道线点{0}", po));
            }*/
        }
        //处理交点
        public void Process_jiaodian(Point3dCollection points, Point3dCollection points1, List<Point3d> jiaodian, List<double> zhenliez)
        {
            int num = 0;
            foreach (Point3d point3D in points)
            {
                //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", point3D));
                Point3d p = new Point3d(point3D.X, point3D.Y, zhenliez[num]);
                jiaodian.Add(p);
                num = num + 1;
            }
            foreach (Point3d point3D in points)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n巷道交点{0}", point3D));
            }
        }
        //计算巷道间隔
        public double getHangdaojiange(Point3dCollection points1)
        {
            double hangdao_jiange = 3.5;

            if (points1.Count == 2)
            {
                hangdao_jiange = System.Math.Abs(points1[1].X - points1[0].X);
                foreach (Point3d point3D in points1)
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n巷道交点{0}", point3D));
                }
            }
            return hangdao_jiange;
        }
        //生成等高线
        public void GenerateIsoheight(BlockTableRecord btr, Transaction trans, List<double> newzcha, double s, double zhenliex1)
        {

            /*for (int i = 0; i < newzcha.Count; i++)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", newzcha[i]));
            }*/

            for (int i = 0; i < newzcha.Count; i++)
            {
                Point3d startpoint = new Point3d(2240, 800 - s + i * 20, newzcha[i]);
                Point3d endpoint = new Point3d(2240 + zhenliex1, 800 - s + i * 20, newzcha[i]);
                Point3d textpoint = new Point3d(2220, 800 - s + i * 20, newzcha[i]);
                MText inserDimension = new MText();
                inserDimension.Location = textpoint;
                String dimension = newzcha[i].ToString();
                inserDimension.Contents = dimension;
                inserDimension.TextHeight = 4;
                inserDimension.Attachment = AttachmentPoint.MiddleCenter;
                inserDimension.Layer = "0";

                Point3d startpoint1 = new Point3d(2240, 800 - s + i * 20 + 10, newzcha[i] + 10);
                Point3d endpoint1 = new Point3d(2240 + zhenliex1, 800 - s + i * 20 + 10, newzcha[i] + 10);
                Point3d textpoint1 = new Point3d(2220, 800 - s + i * 20 + 10, newzcha[i] + 10);
                MText inserDimension1 = new MText();
                inserDimension1.Location = textpoint1;
                String dimension1 = (newzcha[i] + 10).ToString();
                inserDimension1.Contents = dimension1;
                inserDimension1.TextHeight = 4;
                inserDimension1.Attachment = AttachmentPoint.MiddleCenter;
                inserDimension1.Layer = "0";

                Line line = new Line(startpoint, endpoint);
                Line line1 = new Line(startpoint1, endpoint1);
                line.Layer = "0";
                line1.Layer = "0";
                btr.AppendEntity(line);//以图形对象的信息添加到块表记录中
                btr.AppendEntity(line1);
                btr.AppendEntity(inserDimension);
                btr.AppendEntity(inserDimension1);
                trans.AddNewlyCreatedDBObject(line, true);//把对象添加到事务处理中
                trans.AddNewlyCreatedDBObject(line1, true);
                trans.AddNewlyCreatedDBObject(inserDimension, true);
                trans.AddNewlyCreatedDBObject(inserDimension1, true);
            }
        }
        //计算交点
        public void Calculate_jiaodian(BlockTableRecord btr, Transaction trans, List<Point3d> jiaodian, List<Point3d> jiaodian1, Point3dCollection point3d)
        {
            for (int i = 0; i < jiaodian.Count; i++)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n{0}", jiaodian[i]));
                foreach (ObjectId id in btr)
                {
                    Entity entity = (Entity)trans.GetObject(id, OpenMode.ForRead);
                    if (entity.Layer == "0" && entity.GetType() == typeof(Line))
                    {
                        Line denggaoline = entity as Line;
                        if (denggaoline.StartPoint.Z == jiaodian[i].Z)
                        {
                            Point3d k = new Point3d(jiaodian[i].X, denggaoline.StartPoint.Y, jiaodian[i].Z);
                            jiaodian1.Add(k);
                            point3d.Add(k);
                        }
                    }
                }
            }
        }
        //计算煤层拟合点
        public void CoalSeam(List<Point3d> jiaodian1, Point3dCollection point3d2, Point3dCollection point3d22)
        {
            foreach (Point3d po in jiaodian1)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\njiaodian线点{0}", po));
                Point3d po1 = new Point3d(po.X, po.Y + 2.6, po.Z);
                point3d22.Add(po1);
                point3d2.Add(po);
            }
        }
        //生成煤层线
        public void GenerateCoalSeam(BlockTableRecord btr, Transaction trans, Spline spline, Spline spline1)
        {
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n点{0}", spline.StartPoint));
            spline.Layer = "0";
            spline.SetDatabaseDefaults();
            btr.AppendEntity(spline);
            trans.AddNewlyCreatedDBObject(spline, true);
            spline1.Layer = "0";
            spline1.SetDatabaseDefaults();
            btr.AppendEntity(spline1);
            trans.AddNewlyCreatedDBObject(spline1, true);
        }
        //计算巷道坐标z
        public double CalculHangdaoz(List<Point3d> jiaodian1, Line zuankong)
        {
            int biaoji = 0;
            for (int i = 0; i < jiaodian1.Count; i++)
            {
                if (jiaodian1[i].X > zuankong.StartPoint.X)
                {
                    biaoji = i;
                    break;
                }
            }
            double zhangdao = jiaodian1[biaoji - 1].Z + 20 * (zuankong.StartPoint.X - jiaodian1[biaoji - 1].X) / (jiaodian1[biaoji].X - jiaodian1[biaoji - 1].X);
            return zhangdao;
        }
        //生成巷道
        public void GenerateTunnel(BlockTableRecord btr, Transaction trans, double zhangdao, double xhangdao, double yhangdao, Spline spline, double hangdao_jiange)
        {
            Point3d phangdao_zhi1 = new Point3d(xhangdao, yhangdao, zhangdao);
            Point3d phangdao_zhi2 = new Point3d(xhangdao, yhangdao + 1, zhangdao);
            Line hangdao_zhixian = new Line(phangdao_zhi1, phangdao_zhi2);
            Point3dCollection hangdao_nihe = new Point3dCollection();
            hangdao_zhixian.IntersectWith(spline, Intersect.ExtendThis, new Plane(), hangdao_nihe, IntPtr.Zero, IntPtr.Zero);
            Point3d phangdao1 = new Point3d();
            if (hangdao_nihe.Count == 1)
            {
                phangdao1 = hangdao_nihe[0];
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n巷道直线映射交点{0}", phangdao1));
            }

            //Point3d phangdao1 = new Point3d(xhangdao,yhangdao,zhangdao);
            /*Point3d phangdao2 = new Point3d(xhangdao, yhangdao+3.6, zhangdao);
            Point3d phangdao3 = new Point3d(xhangdao- hangdao_jiange, yhangdao+2.6, zhangdao);
            Point3d phangdao4 = new Point3d(xhangdao-hangdao_jiange, yhangdao, zhangdao);*/
            Point3d phangdao2 = new Point3d(phangdao1.X, phangdao1.Y + 3.6, phangdao1.Z);
            Point3d phangdao3 = new Point3d(phangdao1.X - hangdao_jiange, phangdao1.Y + 2.6, phangdao1.Z);
            Point3d phangdao4 = new Point3d(phangdao1.X - hangdao_jiange, phangdao1.Y, phangdao1.Z);
            Point3dCollection hangdao_tixing = new Point3dCollection();
            hangdao_tixing.Add(phangdao1);
            hangdao_tixing.Add(phangdao2);
            hangdao_tixing.Add(phangdao3);
            hangdao_tixing.Add(phangdao4);
            Polyline3d hangdao_polyline = new Polyline3d(Poly3dType.SimplePoly, hangdao_tixing, true);
            hangdao_polyline.Layer = "0";
            hangdao_polyline.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
            btr.AppendEntity(hangdao_polyline);
            trans.AddNewlyCreatedDBObject(hangdao_polyline, true);
        }
        //计算钻孔z坐标
        public double CalculDrillz(List<Point3d> jiaodian1, Line zuankong)
        {
            int biaoji1 = 0;
            for (int i = 0; i < jiaodian1.Count; i++)
            {
                if (jiaodian1[i].X > zuankong.EndPoint.X)
                {
                    biaoji1 = i;
                    break;
                }
            }
            double zuankong_z = jiaodian1[biaoji1 - 1].Z + 20 * (zuankong.EndPoint.X - jiaodian1[biaoji1 - 1].X) / (jiaodian1[biaoji1].X - jiaodian1[biaoji1 - 1].X);
            return zuankong_z;
        }
        //生成钻孔
        public void GenerateDrill(BlockTableRecord btr, Transaction trans, Point3d zuankongEnd, Point3d zuankongStart, Line zuankong_line)
        {
            /*Point3d zuankongEnd = new Point3d(zuankong_x,zuankong_y,zuankong_z);
            Point3d zuankongStart = new Point3d(xhangdao, yhangdao + 1.3, zhangdao);*/

            zuankong_line.Layer = "0";
            zuankong_line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;


            btr.AppendEntity(zuankong_line);
            trans.AddNewlyCreatedDBObject(zuankong_line, true);
        }
        //计算方位角
        public double CalculAzimuth(double angle_zhibeizhen, double angle_zuankong)
        {
            double angle = 0;
            if (angle_zhibeizhen >= angle_zuankong)
            {
                angle = angle_zhibeizhen - angle_zuankong;
            }
            else
            {
                angle = 360 + angle_zhibeizhen - angle_zuankong;
            }
            return angle;
        }
        //计算倾角
        public double CalculDip(Point3d zuankongEnd, Point3d zuankongStart)
        {
            double tanvalue = System.Math.Abs((zuankongEnd.Y - zuankongStart.Y) / (zuankongEnd.X - zuankongStart.X));
            double tanRadianValue2 = Math.Atan(tanvalue);//求弧度值
            double tanAngleValue = tanRadianValue2 / Math.PI * 180;//求角度
            return tanAngleValue;
        }
        //插入文字到指定位置
        public void InsertText(BlockTableRecord btr, Transaction trans, double angle, double tanAngleValue, Line zuankong_line, double zhenliex1)
        {
            //保留两位小数
            double z_angle = Math.Round(angle, 2);
            double tan_angle = Math.Round(tanAngleValue, 2);
            double lengh_angle = Math.Round(zuankong_line.Length, 2);
            //Point3d textpoint_angle = new Point3d(2240 + zhenliex1 / 2, 800 - 2 * 20, 0);
            Point3d textpoint_angle = new Point3d(2240 + zhenliex1 / 4, 800, 0);
            MText inserDimension_angle = new MText();
            inserDimension_angle.Location = textpoint_angle;
            String dimension_angle = null;
            //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n指北针角度{0}", angle_zhibeizhen));
            //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(string.Format("\n钻孔角度{0}", angle_zuankong));


            if (zuankong_line.Angle > 180)
            {
                //tanAngleValue = -tanAngleValue;
                dimension_angle = "X4  " + z_angle + "°  ∠ - " + tan_angle + "° " + lengh_angle + "m";
            }
            else
            {
                dimension_angle = "X4  " + z_angle + "°  ∠ + " + tan_angle + "° " + lengh_angle + "m";
            }
            //String dimension_angle = "X4  " + angle + "∠" + ;
            inserDimension_angle.Contents = dimension_angle;
            inserDimension_angle.TextHeight = 4;
            inserDimension_angle.Attachment = AttachmentPoint.MiddleCenter;
            inserDimension_angle.Layer = "0";
            btr.AppendEntity(inserDimension_angle);
            trans.AddNewlyCreatedDBObject(inserDimension_angle, true);
        }
    }
}