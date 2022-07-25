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

namespace CAD.DrawProfile4
{
    public class ReadDrawCADBatch
    {
        //自动绘图之前的校验步骤。设置默认没有图层为0；正确为1；类型错误为2；
        [CommandMethod("TESTDWG")]
        public void TESTDWG()
        {
            //需要访问Database的操作 需首先将该文档进行锁定，操作完成后，在最后进行释放
            DocumentLock docLock = Application.DocumentManager.MdiActiveDocument.LockDocument();
            // 对话框窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 数据库对象
            Database db = HostApplicationServices.WorkingDatabase;
            Dictionary<String, int> test = new Dictionary<string, int> {
                {"工作面巷道注记",0 },{"导线点",0 },{"断层上盘",0 },{"断层下盘",0 },{"标注",0 },{"223等高线",0 },{"岩巷/煤巷",0 },
                {"229断层注记",0 },{"取样点",0 },{"导线点高程",0 },{"导线点名称",0 },{"224等高线注记",0 },
                {"工作面信息",0 }, {"煤层厚度",0},{"抽采半径",0},
                {"指北针",0 },
                {"设计校验钻孔",0 },
            };
            //储存数据库数据
            //List<List<String>> sql = new List<List<String>>();
            //InsertSql(sql);
            //CAD处理
            try
            {
                using (TransactionHost host = new TransactionHost(db))
                {
                    List<String> names = new List<String>();
                    foreach (ObjectId id in host.ModelSpace)
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                        //工作面巷道注记
                        if (entity.Layer == "工作面巷道注记")
                        {
                            if (entity.GetType() == typeof(DBText) && test["工作面巷道注记"] != 2)
                            {
                                test["工作面巷道注记"] = 1;
                            }
                            else test["工作面巷道注记"] = 2;
                            continue;
                        }
                        //一般命名为:工作面巷道注记 + 导线点(2901上顺槽导线点)
                        if (entity.Layer.EndsWith("导线点"))
                        {
                            if (entity.GetType() == typeof(Circle) && test["导线点"] != 2)
                            {
                                test["导线点"] = 1;
                            }
                            else test["导线点"] = 2;
                            continue;
                        }
                        //断层上下盘
                        if (entity.Layer == "断层上盘")
                        {
                            if (entity.GetType() == typeof(Spline) && test["断层上盘"] != 2)
                            {
                                test["断层上盘"] = 1;
                            }
                            else test["断层上盘"] = 2;
                            continue;
                        }
                        if (entity.Layer == "断层下盘")
                        {
                            if (entity.GetType() == typeof(Spline) && test["断层下盘"] != 2)
                            {
                                test["断层下盘"] = 1;
                            }
                            else test["断层下盘"] = 2;
                            continue;
                        }
                        //标注部分
                        if (entity.Layer == "标注")
                        {
                            if ((entity.GetType() == typeof(MText) || entity.GetType() == typeof(DBText)) && test["标注"] != 2)
                            {
                                test["标注"] = 1;
                            }
                            else test["标注"] = 2;
                            continue;
                        }
                        //等高线，有两种线形。
                        if (entity.Layer == "223等高线")
                        {
                            if ((entity.GetType() == typeof(Polyline3d) || entity.GetType() == typeof(Polyline)) && test["223等高线"] != 2)
                            {
                                test["223等高线"] = 1;
                            }
                            else test["223等高线"] = 2;
                            continue;
                        }
                        //文本有两种类型，DBText单行，MText多行
                        if (entity.Layer == "229断层注记")
                        {
                            if ((entity.GetType() == typeof(MText) || entity.GetType() == typeof(DBText)) && test["229断层注记"] != 2)
                            {
                                test["229断层注记"] = 1;
                            }
                            else test["229断层注记"] = 2;
                            continue;
                        }
                        if (entity.Layer == "导线点高程")
                        {
                            if ((entity.GetType() == typeof(MText) || entity.GetType() == typeof(DBText)) && test["导线点高程"] != 2)
                            {
                                test["导线点高程"] = 1;
                            }
                            else test["导线点高程"] = 2;
                            continue;
                        }
                        if (entity.Layer == "导线点名称")
                        {
                            if ((entity.GetType() == typeof(MText) || entity.GetType() == typeof(DBText)) && test["导线点名称"] != 2)
                            {
                                test["导线点名称"] = 1;
                            }
                            else test["导线点名称"] = 2;
                            continue;
                        }
                        if (entity.Layer == "224等高线注记")
                        {
                            if ((entity.GetType() == typeof(MText) || entity.GetType() == typeof(DBText)) && test["224等高线注记"] != 2)
                            {
                                test["224等高线注记"] = 1;
                            }
                            else test["224等高线注记"] = 2;
                            continue;
                        }
                        //取样点校验
                        if (entity.Layer == "取样点")
                        {
                            if (entity.GetType() == typeof(Circle) && test["取样点"] != 2)
                            {
                                test["取样点"] = 1;
                            }
                            else test["取样点"] = 2;
                            continue;
                        }
                        //煤巷或岩巷
                        if (entity.Layer == "133煤巷" || entity.Layer == "134岩巷" || entity.Layer == "设计煤巷" || entity.Layer == "设计岩巷" || entity.Layer == "煤巷" || entity.Layer == "岩巷")
                        {
                            test["岩巷/煤巷"] = 1;
                            continue;
                        }
                        //工作面信息
                        if (entity.Layer == "工作面信息")
                        {
                            if (entity.GetType() == typeof(MText) && test["工作面信息"] != 2)
                            {
                                test["工作面信息"] = 1;
                            }
                            else test["工作面信息"] = 2;
                            continue;
                        }
                        //指北针
                        if (entity.GetType() == typeof(BlockReference))
                        {
                            BlockReference blockReference = entity as BlockReference;
                            if (blockReference.Name == "指北针")
                            {
                                test["指北针"] = 1;
                            }
                            continue;
                        }
                        //设计校验钻孔(之后要修改，不一定是校验孔)
                        if (entity.Layer.Contains("设计校验钻孔"))
                        {
                            if ((entity.GetType() == typeof(Line) || entity.GetType() == typeof(Polyline)) && test["设计校验钻孔"] != 2)
                            {
                                test["设计校验钻孔"] = 1;
                            }
                            else test["设计校验钻孔"] = 2;
                            continue;
                        }
                    }
                }
                //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("# tunnel " + tunnel_name[i] + "\n");
                Double test_sum = 0;
                foreach (String name_test in test.Keys)
                {
                    if (test[name_test] == 0)
                    {
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("* " + "缺少必要图层: " + name_test + "\n");
                    }
                    else if (test[name_test] == 1)
                    {
                        test_sum = test_sum + 1;
                    }
                    else if (test[name_test] == 2)
                    {
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("* " + name_test + "图层中的线型或者文本类型错误" + "\n");
                    }
                }
                if (test_sum == 17)
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("# 校验完成");
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                Application.ShowAlertDialog(e.Message);
            }
            // 解锁文档
            docLock.Dispose();
        }

        //生成一个顺层钻孔平面图的所有钻孔剖面图
        [CommandMethod("DPDWG")]
        public void DrawProfileDWG()
        {
            //需要访问Database的操作 需首先将该文档进行锁定，操作完成后，在最后进行释放
            DocumentLock docLock = Application.DocumentManager.MdiActiveDocument.LockDocument();
            // 对话框窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 数据库对象
            Database db = HostApplicationServices.WorkingDatabase;
            //储存数据库数据
            List<List<String>> sql = new List<List<String>>();
            InsertSql(sql);
            //CAD处理
            try
            {
                Dictionary<String, List<Circle>> Daoxiandian = new Dictionary<String, List<Circle>>();
                SearchDaoxiandian(db, Daoxiandian);
                sql[3].Add(GetCompassAngle(db).ToString());
                List<Line> drills = new List<Line>();//钻孔
                Point3d p0 = new Point3d(0, 0, 0);//剖面图起始位置
                GenerateLayer(db);
                drills = SearchDrill(db);
                List<String> names = new List<String>();
                GetName(db, names, drills);
                SortDrill(drills, names);
                GenerateBackground(db, drills.Count);
                for (int i = 0; i < drills.Count; i++)
                {
                    //ed.WriteMessage("#Drill "+names[i]+" StartPoint:"+drills[i].StartPoint+" EndPoint:"+ drills[i].EndPoint+"\n");
                    //ed.WriteMessage("# " + names[i] + " " + drills[i].StartPoint.X + " " + drills[i].StartPoint.Y + " " + drills[i].StartPoint.Z + " " + drills[i].EndPoint.X + " " + drills[i].EndPoint.Y + " " + drills[i].EndPoint.Z + "\n");
                    p0 = new Point3d((300 * i) % 1500, -200 * (i / 5), 0);
                    sql[2].Add(names[i]);
                    DrawSingleDrillProfile(db, p0, drills[i], names[i], sql, Daoxiandian);
                }
                PrintSql(sql);
                dataTable(db, sql, new Point3d(1600, 200, 0), 10);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                Application.ShowAlertDialog(e.Message);
            }
            // 解锁文档
            docLock.Dispose();
        }

        //搜索巷道导向点
        public void SearchDaoxiandian(Database db, Dictionary<String, List<Circle>> mark)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                List<String> names = new List<String>();
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    if (entity.Layer == "工作面巷道注记" && entity.GetType() == typeof(DBText))
                    {
                        DBText dBText = (DBText)entity;
                        if (!names.Contains(dBText.TextString))
                        {
                            names.Add(dBText.TextString);
                        }
                    }
                }
                foreach (String name in names)
                {
                    mark.Add(name, new List<Circle>());
                }
                for (int i = 0; i < names.Count; i++)
                {
                    foreach (ObjectId id in host.ModelSpace)
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                        if (entity.Layer == names[i] + "导线点" && entity.GetType() == typeof(Circle))
                        {
                            Circle c = (Circle)entity;
                            mark[names[i]].Add(c);
                        }
                    }
                }
                for (int i = 0; i < names.Count; i++)
                {
                    mark[names[i]].Sort((a, b) => a.Center.X.CompareTo(b.Center.X));
                }
            }
        }

        //数据库添加实体
        public void InsertSql(List<List<String>> sql)
        {
            List<String> sqlTunnel = new List<String>();//巷道名称
            List<String> sqlSite = new List<String>();//钻场编号，钻场编码
            List<String> sqlDrill = new List<String>();//钻孔编号，钻孔编码，钻孔方位角，倾角，孔深，开孔高度，标点，相对标点位置，距标点距离
            List<String> sqlCompass = new List<String>();//旋转角度
            List<String> sqlContour = new List<String>();//等高线高度
            sql.Add(sqlTunnel);
            sql.Add(sqlSite);
            sql.Add(sqlDrill);
            sql.Add(sqlCompass);
            sql.Add(sqlContour);
        }

        //按照钻孔名称给钻孔排序
        public void SortDrill(List<Line> drills, List<String> names)
        {
            List<String> name = new List<String>();
            for (int i = 0; i < names.Count; i++)
            {
                if (!name.Contains(names[i][0].ToString()))
                {
                    name.Add(names[i][0].ToString());
                }
            }
            int n1 = 0;
            for (int i = 0; i < name.Count; i++)
            {
                int n2 = 0;
                for (int j = 0; j < names.Count; j++)
                {
                    if (name[i] == names[j][0].ToString())
                    {
                        n2++;
                    }
                }
                int n3 = 1;
                while (n3 != 0)
                {
                    n3 = 0;
                    for (int j = 0; j < names.Count; j++)
                    {
                        if (name[i] == names[j][0].ToString())
                        {
                            int loc = Convert.ToInt32(names[j].Substring(1)) + n1 - 1;
                            if (j != loc)
                            {
                                n3++;
                                var temp1 = names[j];
                                names[j] = names[loc];
                                names[loc] = temp1;
                                var temp2 = drills[j];
                                drills[j] = drills[loc];
                                drills[loc] = temp2;
                            }
                        }
                    }
                }
                n1 += n2;
            }
        }

        //获取煤层厚度
        public double GetCoalThick(Database db)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = host.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                    if (entity != null && entity.GetType() == typeof(MText))
                    {
                        MText text = (MText)entity;
                        String coalthick = text.Contents;
                        if (text.Layer == "工作面信息" && coalthick.Substring(0, 4) == "煤层厚度")
                        {
                            int Indexofm = coalthick.IndexOf('m');
                            return Convert.ToDouble(coalthick.Substring(5, Indexofm - 5));
                        }
                    }
                }
            }
            return 3;
        }

        //获取抽采半径
        public double GetDrillR(Database db)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = host.GetObject(id, OpenMode.ForRead) as Autodesk.AutoCAD.DatabaseServices.Entity;
                    if (entity != null && entity.GetType() == typeof(MText))
                    {
                        MText text = (MText)entity;
                        String coalthick = text.Contents;
                        if (text.Layer == "工作面信息" && coalthick.Substring(0, 4) == "抽采半径")
                        {
                            int Indexofm = coalthick.IndexOf('m');
                            return Convert.ToDouble(coalthick.Substring(5, Indexofm - 5));
                        }
                    }
                }
            }
            return 8;
        }

        //添加图层
        public void GenerateLayer(Database db)
        {
            AddIn(db, "图例");
            AddIn(db, "等高线注记");
            AddIn(db, "煤层");
            AddIn(db, "刨面钻孔");
            AddIn(db, "煤巷");
            AddIn(db, "煤巷剖面");
            AddIn(db, "设计钻孔注记");
            AddIn(db, "断层剖面");
            AddIn(db, "底抽巷钻孔");
        }
        //画出一个钻孔的剖面图(顺层)
        public void DrawSingleDrillProfile(Database db, Point3d p0, Line drill, String name, List<List<String>> sql, Dictionary<String, List<Circle>> Daoxiandian)
        {
            List<double> height = new List<double>();
            Point3dCollection points = new Point3dCollection();//points 是钻孔延长线和等高线的交点
            //钻孔和煤巷有两个交点
            Point3d point_road1 = new Point3d();
            Point3d point_road2 = drill.StartPoint;//钻孔的起点也是钻孔与煤巷的交点
            height = Isoheight(db, drill, points);
            String tunnel_name = drill.Layer.Substring(0, drill.Layer.Length - 6);
            sql[0].Add(tunnel_name);
            point_road1 = RoadwayIntersection(db, drill);
            double tunnelwidth = point_road1.GetDistanceBetweenToPoint(point_road2);
            InsertText(db, new Point3d(p0.X + 80, p0.Y + 136, 0), tunnel_name + name + "钻孔剖面图   （1：1000）", "图例", 3);
            //计算剖面图中煤层和等高线的交点
            List<Point3d> edges = FindEdge(points);
            Point3d point_start = FindStartPoint(edges, point_road2);
            int constant = 20;
            //煤层剖面下底板拟合点
            Point3dCollection coal = Coal1(point_start, point_road2, p0, points, constant, height);
            Point3dCollection coal1 = new Point3dCollection();
            Collection(coal, coal1);
            SortCoal(coal1);
            double thick = GetCoalThick(db);//煤层厚度
            String DaoxiandianName = "";//标点名称
            String DaoxiandianLocation = "里";//相对标点位置
            double DaoxiandianDis = 0; //距离标点距离
            DaoxiandianLocation = DrillDaoxiandian2(db, drill, Daoxiandian);
            DaoxiandianName = DrillDaoxiandian1(db, drill, Daoxiandian);
            DaoxiandianDis = DrillDaoxiandian(db, drill, Daoxiandian);
            double drillheight = DrillingHeight(db, drill, Daoxiandian);
            //Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("#drillheight " + drillheight +"\n");
            //判断是否通过断层
            if (HasFault(db, drill, coal1))//通过断层
            {

                Spline faultup = new Spline();
                Spline faultdown = new Spline();
                Dictionary<String, Point3d> jiaodian = new Dictionary<string, Point3d>();
                jiaodian.Add("断层上盘", new Point3d());
                jiaodian.Add("断层下盘", new Point3d());
                faultup = GetFaultUp(db, drill, jiaodian);
                faultdown = GetFaultDown(db, drill, jiaodian, faultup);
                String info;
                info = GetInfo(db, faultup);
                if (info != null)
                {
                    info = info.Trim();
                }
                if (info == null)
                {
                    info = GetInfo(db, faultdown).ToString().Trim();
                }
                int index = info.IndexOf('=');
                String luo = info.Substring(index + 1, info.Count() - 2 - index);

                double luocha = Convert.ToDouble(luo);
                if (luocha > 10)
                {
                    luocha = Math.Sqrt(luocha);//调整落差，避免落差过大
                }
                bool direction = GetFaultDir(drill, jiaodian);
                Coal10Fault(coal1, jiaodian, drill, p0);
                CutSpline(coal1, p0);
                double min = InsertHeightText(db, p0, height, coal, coal1, drillheight);
                CoalStart(coal1, drillheight, min, p0);
                DrawFault(db, jiaodian, coal1, luocha, p0, direction, thick, drill, name, sql, info, min, drillheight, tunnelwidth);

            }
            else//未通过断层
            {
                Coal10(coal1);
                CutSpline(coal1, p0);
                double min = InsertHeightText(db, p0, height, coal, coal1, drillheight);
                CoalStart(coal1, drillheight, min, p0);
                Spline spl1, spl2;
                spl1 = GenerateCoalSeam(db, coal1);
                //煤层剖面上地板拟合点
                Point3dCollection coal2 = Coal2(coal1, thick);
                spl2 = GenerateCoalSeam(db, coal2);
                Point3d tunnelpoint = new Point3d(coal1[0].X, p0.Y + 30 + drillheight - min, 0);
                Line drill1 = GenerateDrill(db, drill, thick, coal1, tunnelpoint);

                GenerateHangdao(db, tunnelpoint, tunnelwidth, UpOrDown(name));
                GenerateText(db, p0, drill, drill1, name, sql);
                DrillColor(db, drill1, spl1, spl2);
            }
            sql[2][sql[2].Count - 1] += " " + DaoxiandianName + " " + DaoxiandianLocation + " " + Math.Round(DaoxiandianDis, 1);
        }

        //判断是上顺槽还是下顺槽
        public bool UpOrDown(String name)
        {
            if (name.Contains('S'))
                return true;
            return false;
        }

        //判断钻孔从上盘进入下盘还是从下盘进入上盘,true表示由下盘进入上盘
        public bool GetFaultDir(Line drill, Dictionary<String, Point3d> jiaodian)
        {
            if (drill.StartPoint.Y < drill.EndPoint.Y)
            {
                if (jiaodian["断层下盘"].Y < jiaodian["断层上盘"].Y)
                    return true;
                else
                    return false;
            }
            else
            {
                if (jiaodian["断层下盘"].Y < jiaodian["断层上盘"].Y)
                    return false;
                else
                    return true;
            }
        }
        //画出断层
        public void DrawFault(Database db, Dictionary<String, Point3d> jiaodian, Point3dCollection coal, double luocha, Point3d p0, bool direction, double thick, Line drill, String name, List<List<String>> sql, String info, double min, double drillheight, double tunnelwidth)
        {
            double x1, x2;
            x1 = BaseTool.GetDistanceBetweenToPoint(drill.StartPoint, jiaodian["断层上盘"]);
            x2 = BaseTool.GetDistanceBetweenToPoint(drill.StartPoint, jiaodian["断层下盘"]);
            Point3dCollection points1 = GetPointByX(p0.X + x1 + 20, coal);
            Point3dCollection points2 = GetPointByX(p0.X + x2 + 20, coal);
            Point3d p1 = points1[0];
            Point3d p2 = points2[0];
            p2 = new Point3d(p2.X, p1.Y + luocha, p2.Z);//根据落差获得下盘在剖面图的交点坐标
            Point3dCollection coal1 = new Point3dCollection();
            Point3dCollection coal2 = new Point3dCollection();
            if (p1.X < p2.X)
            {
                for (int i = 0; i < coal.Count; i++)
                {
                    if (coal[i].X < p1.X)
                        coal1.Add(coal[i]);
                    if (coal[i].X > p2.X)
                        coal2.Add(coal[i]);
                }
                coal1.Add(p1);
                coal2.Insert(0, p2);
                //Coal10(coal1);
                //Coal10(coal2);
                Spline spl11 = GenerateCoalSeam(db, coal1);
                Spline spl21 = GenerateCoalSeam(db, coal2);

                Point3dCollection coal12 = Coal2(coal1, thick);
                Point3dCollection coal22 = Coal2(coal2, thick);
                StartEndPointFault(p1, p2, coal12, coal22);

                Spline spl12 = GenerateCoalSeam(db, coal12);
                Spline spl22 = GenerateCoalSeam(db, coal22);
                MarkFault(db, p1, p2, false);//false表示断层左边低右边高

                Point3d tunnelpoint = new Point3d(coal1[0].X, p0.Y + 30 + drillheight - min, 0);
                Line drill1 = GenerateDrill1(db, drill, thick, coal1, coal2, tunnelpoint);
                GenerateHangdao(db, tunnelpoint, tunnelwidth, UpOrDown(name));
                GenerateText(db, p0, drill, drill1, name, sql);
                DrillColorFault(db, drill1, spl11, spl12, spl21, spl22);
            }
            if (p1.X > p2.X)
            {
                for (int i = 0; i < coal.Count; i++)
                {
                    if (coal[i].X < p2.X)
                        coal1.Add(coal[i]);
                    if (coal[i].X > p1.X)
                        coal2.Add(coal[i]);
                }
                if (direction)//钻孔由下盘进入上盘
                {
                    coal1.Add(p2);
                    coal2.Insert(0, p1);
                }
                //Coal10(coal1);
                //Coal10(coal2);
                Spline spl11 = GenerateCoalSeam(db, coal1);
                Spline spl21 = GenerateCoalSeam(db, coal2);


                Point3dCollection coal12 = Coal2(coal1, thick);
                Point3dCollection coal22 = Coal2(coal2, thick);
                StartEndPointFault(p1, p2, coal12, coal22);


                Spline spl12 = GenerateCoalSeam(db, coal12);
                Spline spl22 = GenerateCoalSeam(db, coal22);

                MarkFault(db, p1, p2, true);//true表示断层左边高右边低

                Point3d tunnelpoint = new Point3d(coal1[0].X, p0.Y + 30 + drillheight - min, 0);
                Line drill1 = GenerateDrill1(db, drill, thick, coal1, coal2, tunnelpoint);
                GenerateHangdao(db, tunnelpoint, tunnelwidth, UpOrDown(name));
                GenerateText(db, p0, drill, drill1, name, sql);
                DrillColorFault(db, drill1, spl11, spl12, spl21, spl22);
            }
            InsertText(db, new Point3d(p1.X, p0.Y + 30, p0.Z), info, "图例", 3);
        }

        //处理煤层线靠近断层的起点和终点(顺层)
        public void StartEndPointFault(Point3d p1, Point3d p2, Point3dCollection coal12, Point3dCollection coal22)
        {
            Line line = new Line(p1, p2);
            Point3dCollection point3DCollection = new Point3dCollection();
            Spline spline = new Spline(coal12, 4, 0.0);
            line.IntersectWith(spline, Intersect.ExtendThis, new Plane(), point3DCollection, IntPtr.Zero, IntPtr.Zero);
            if (point3DCollection.Count == 1)
            {

                if (coal12[coal12.Count - 2].X < point3DCollection[0].X)
                {
                    coal22.Insert(0, new Point3d(coal22[0].X + (point3DCollection[0].X - coal12[coal12.Count - 1].X), coal22[0].Y + (point3DCollection[0].Y - coal12[coal12.Count - 1].Y), 0));
                    if (coal22[1].X - coal22[0].X < 10)
                        coal22.RemoveAt(1);
                    coal12.RemoveAt(coal12.Count - 1);
                }

                else
                {
                    coal22.Insert(0, new Point3d(coal22[0].X + (point3DCollection[0].X - coal12[coal12.Count - 1].X), coal22[0].Y + (point3DCollection[0].Y - coal12[coal12.Count - 1].Y), 0));
                    if (coal22[1].X - coal22[0].X < 10)
                        coal22.RemoveAt(1);
                    coal12.RemoveAt(coal12.Count - 2);
                    coal12.RemoveAt(coal12.Count - 1);
                }
                coal12.Add(point3DCollection[0]);

            }
            else
            {
                spline = new Spline(coal22, 4, 0.0);
                line.IntersectWith(spline, Intersect.ExtendThis, new Plane(), point3DCollection, IntPtr.Zero, IntPtr.Zero);
                if (point3DCollection.Count == 1)
                {
                    if (coal22[1].X > point3DCollection[0].X)
                    {
                        coal12.Add(new Point3d(coal12[coal12.Count - 1].X + (point3DCollection[0].X - coal22[0].X), coal12[coal12.Count - 1].Y + (point3DCollection[0].Y - coal22[0].Y), 0));
                        if (coal12[coal12.Count - 1].X - coal12[coal12.Count - 2].X < 10)
                            coal12.RemoveAt(coal12.Count - 2);
                        coal22.RemoveAt(0);
                    }
                    else
                    {
                        coal12.Add(new Point3d(coal12[coal12.Count - 1].X + (point3DCollection[0].X - coal22[0].X), coal12[coal12.Count - 1].Y + (point3DCollection[0].Y - coal22[0].Y), 0));
                        if (coal12[coal12.Count - 1].X - coal12[coal12.Count - 2].X < 10)
                            coal12.RemoveAt(coal12.Count - 2);
                        coal22.RemoveAt(1);
                        coal22.RemoveAt(0);
                    }
                    coal22.Insert(0, point3DCollection[0]);
                }
            }
        }

        //钻孔颜色改变(顺层有断层)
        public void DrillColorFault(Database db, Line drill, Spline spl1, Spline spl2, Spline spl3, Spline spl4)
        {
            Point3dCollection point3DCollection1 = new Point3dCollection();
            Point3dCollection point3DCollection2 = new Point3dCollection();
            Point3dCollection point3DCollection3 = new Point3dCollection();
            Point3dCollection point3DCollection4 = new Point3dCollection();
            using (TransactionHost host = new TransactionHost(db))
            {
                if (drill.EndPoint.X <= spl1.EndPoint.X)
                {
                    drill.IntersectWith(spl1, Intersect.OnBothOperands, new Plane(), point3DCollection1, IntPtr.Zero, IntPtr.Zero);
                    drill.IntersectWith(spl2, Intersect.OnBothOperands, new Plane(), point3DCollection2, IntPtr.Zero, IntPtr.Zero);
                    if (point3DCollection1.Count == 2)
                    {
                        Line line = new Line(point3DCollection1[0], point3DCollection1[1]);
                        line.ColorIndex = 30;
                        line.Layer = "刨面钻孔";
                        line.Linetype = "ByLayer";
                        line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                        host.AddEntity(line);
                    }
                    if (point3DCollection2.Count == 2)
                    {
                        Line line = new Line(point3DCollection2[0], point3DCollection2[1]);
                        line.ColorIndex = 30;
                        line.Layer = "刨面钻孔";
                        line.Linetype = "ByLayer";
                        line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                        host.AddEntity(line);
                    }
                    host.Commit();
                }
                else
                {
                    drill.IntersectWith(spl1, Intersect.OnBothOperands, new Plane(), point3DCollection1, IntPtr.Zero, IntPtr.Zero);
                    drill.IntersectWith(spl2, Intersect.OnBothOperands, new Plane(), point3DCollection2, IntPtr.Zero, IntPtr.Zero);
                    drill.IntersectWith(spl3, Intersect.OnBothOperands, new Plane(), point3DCollection3, IntPtr.Zero, IntPtr.Zero);
                    drill.IntersectWith(spl4, Intersect.OnBothOperands, new Plane(), point3DCollection4, IntPtr.Zero, IntPtr.Zero);
                    if (point3DCollection1.Count == 1 && point3DCollection4.Count == 1)
                    {
                        Line line = new Line(point3DCollection1[0], point3DCollection4[0]);
                        line.ColorIndex = 30;
                        line.Layer = "刨面钻孔";
                        line.Linetype = "ByLayer";
                        line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                        host.AddEntity(line);
                    }
                    if (point3DCollection2.Count == 1 && point3DCollection3.Count == 1)
                    {
                        Line line = new Line(point3DCollection2[0], point3DCollection3[0]);
                        line.ColorIndex = 30;
                        line.Layer = "刨面钻孔";
                        line.Linetype = "ByLayer";
                        line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                        host.AddEntity(line);
                    }
                    host.Commit();
                }

            }
        }

        //获取和钻孔相交的断层上盘实体(顺层)
        public Spline GetFaultUp(Database db, Line drill, Dictionary<String, Point3d> jiaodian)
        {
            Spline Up = new Spline();
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Point3dCollection points = new Point3dCollection();
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    if (entity.Layer == "断层上盘")
                    {
                        //求交点，钻孔和断层线必须共面
                        Line line = new Line(new Point3d(drill.StartPoint.X, drill.StartPoint.Y, 0), new Point3d(drill.EndPoint.X, drill.EndPoint.Y, 0));
                        line.IntersectWith(entity, Intersect.ExtendThis, points, IntPtr.Zero, IntPtr.Zero);
                        if (points.Count > 0)
                        {
                            if (drill.StartPoint.Y > drill.EndPoint.Y)
                            {
                                if (drill.StartPoint.GetDistanceBetweenToPoint(points[0]) < 160 && (int)points[0].Y < (int)drill.StartPoint.Y)
                                {
                                    Up = entity as Spline;
                                    jiaodian["断层上盘"] = points[0];
                                    break;
                                }
                            }
                            if (drill.StartPoint.Y < drill.EndPoint.Y)
                            {
                                if (drill.StartPoint.GetDistanceBetweenToPoint(points[0]) < 160 && (int)points[0].Y > (int)drill.StartPoint.Y)
                                {
                                    Up = entity as Spline;
                                    jiaodian["断层上盘"] = points[0];
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return Up;
        }

        //获取和钻孔相交的断层下盘实体(顺层)
        public Spline GetFaultDown(Database db, Line drill, Dictionary<String, Point3d> jiaodian, Spline Up)
        {
            Spline Down = new Spline();
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Point3dCollection points = new Point3dCollection();
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    if (entity.Layer == "断层下盘")
                    {
                        //求交点，钻孔和断层线必须共面
                        Line line = new Line(new Point3d(drill.StartPoint.X, drill.StartPoint.Y, 0), new Point3d(drill.EndPoint.X, drill.EndPoint.Y, 0));
                        line.IntersectWith(entity, Intersect.ExtendThis, points, IntPtr.Zero, IntPtr.Zero);
                        if (points.Count > 0 && isSameFault(entity as Spline, Up))
                        {
                            if (drill.StartPoint.Y > drill.EndPoint.Y)
                            {
                                if ((int)points[0].Y < (int)drill.StartPoint.Y)
                                {
                                    Down = entity as Spline;
                                    jiaodian["断层下盘"] = points[0];
                                    break;
                                }
                            }
                            if (drill.StartPoint.Y < drill.EndPoint.Y)
                            {
                                if ((int)points[0].Y > (int)drill.StartPoint.Y)
                                {
                                    Down = entity as Spline;
                                    jiaodian["断层下盘"] = points[0];
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return Down;
        }

        //判断上盘和下盘是否是一个断层
        public bool isSameFault(Spline up, Spline down)
        {
            if ((int)up.StartPoint.X == (int)down.StartPoint.X && (int)up.StartPoint.Y == (int)down.StartPoint.Y)
                return true;
            if ((int)up.StartPoint.X == (int)down.EndPoint.X && (int)up.StartPoint.Y == (int)down.EndPoint.Y)
                return true;
            if ((int)up.EndPoint.X == (int)down.StartPoint.X && (int)up.EndPoint.Y == (int)down.StartPoint.Y)
                return true;
            if ((int)up.EndPoint.X == (int)down.EndPoint.X && (int)up.EndPoint.Y == (int)down.EndPoint.Y)
                return true;
            return false;
        }

        //获取开孔点巷道高度
        public double DrillingHeight(Database db, Line drill, Dictionary<String, List<Circle>> Daoxiandian)
        {
            double drillheight = 0;
            String TunnelName = drill.Layer.Substring(0, drill.Layer.Count() - 6);
            List<Circle> Circles = Daoxiandian[TunnelName];
            List<DBText> gaocheng = new List<DBText>();
            GetHeightMark(db, gaocheng);
            for (int i = Circles.Count - 1; i >= 0; i--)
            {
                if (i != Circles.Count - 1 && Circles[i].Center.X < drill.StartPoint.X)
                {
                    String s1 = NearText(db, Circles[i], gaocheng).TextString;
                    String s2 = NearText(db, Circles[i + 1], gaocheng).TextString;
                    double height1 = Convert.ToDouble(s1);
                    double height2 = Convert.ToDouble(s2);
                    if (s1.Count() - s1.IndexOf('.') - 1 == 3)
                    {
                        height1 -= 2.6;
                    }
                    if (s2.Count() - s2.IndexOf('.') - 1 == 3)
                    {
                        height2 -= 2.6;
                    }
                    drillheight = height1 + (height2 - height1) * (drill.StartPoint.X - Circles[i].Center.X) / (Circles[i + 1].Center.X - Circles[i].Center.X);
                    break;
                }
            }
            if (drillheight == 0)
            {
                String s1 = NearText(db, Circles[0], gaocheng).TextString;
                drillheight = Convert.ToDouble(s1);
                if (s1.Count() - s1.IndexOf('.') - 1 == 3)
                {
                    drillheight -= 2.6;
                }
            }
            return drillheight;
        }

        //获取钻孔的据标点距离(顺层)
        public double DrillDaoxiandian(Database db, Line drill, Dictionary<String, List<Circle>> Daoxiandian)
        {

            double DaoxiandianDis = 0;
            String TunnelName = drill.Layer.Substring(0, drill.Layer.Count() - 6);
            List<Circle> Circles = Daoxiandian[TunnelName];
            //List<DBText> gaocheng = new List<DBText>();
            List<DBText> mark = new List<DBText>();
            //GetHeightMark(db,gaocheng);
            GetDaoxianMark(db, mark);
            for (int i = Circles.Count - 1; i >= 0; i--)
            {
                if (Circles[i].Center.X < drill.StartPoint.X)
                {
                    DaoxiandianDis = drill.StartPoint.X - Circles[i].Center.X;

                    break;
                }
            }
            if (DaoxiandianDis == 0)
            {
                DaoxiandianDis = Circles[0].Center.X - drill.StartPoint.X;
            }
            return DaoxiandianDis;
        }

        //获取钻孔的导线点编号(顺层)
        public string DrillDaoxiandian1(Database db, Line drill, Dictionary<String, List<Circle>> Daoxiandian)
        {
            String DaoxiandianName = "";
            String TunnelName = drill.Layer.Substring(0, drill.Layer.Count() - 6);
            List<Circle> Circles = Daoxiandian[TunnelName];
            //List<DBText> gaocheng = new List<DBText>();
            List<DBText> mark = new List<DBText>();
            //GetHeightMark(db, gaocheng);
            GetDaoxianMark(db, mark);
            for (int i = Circles.Count - 1; i >= 0; i--)
            {
                if (Circles[i].Center.X < drill.StartPoint.X)
                {
                    DaoxiandianName = NearText(db, Circles[i], mark).TextString;
                    break;
                }
            }
            if (DaoxiandianName == "")
            {
                DaoxiandianName = NearText(db, Circles[0], mark).TextString;
            }
            return DaoxiandianName;
        }

        //获取钻孔的导线点编号(顺层)
        public string DrillDaoxiandian2(Database db, Line drill, Dictionary<String, List<Circle>> Daoxiandian)
        {

            String TunnelName = drill.Layer.Substring(0, drill.Layer.Count() - 6);
            List<Circle> Circles = Daoxiandian[TunnelName];
            if (Circles[0].Center.X > drill.StartPoint.X)
                return "外";
            return "里";
        }

        //获取断层的信息(顺层)
        public String GetInfo(Database db, Spline fault)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    if (entity != null)
                    {
                        Point3dCollection point3DCollection = new Point3dCollection();
                        if (entity.Layer == "229断层注记")
                        {
                            DBText text = entity as DBText;

                            if (((int)text.Position.X == (int)fault.StartPoint.X && (int)text.Position.Y == (int)fault.StartPoint.Y) || ((int)text.Position.X == (int)fault.EndPoint.X && (int)text.Position.Y == (int)fault.EndPoint.Y))
                            {
                                return text.TextString;
                            }
                        }
                    }
                }
            }
            return null;
        }

        //向剖面图添加倾角方位角孔长信息
        public void GenerateText(Database db, Point3d p0, Line drill, Line drill1, String name, List<List<String>> sql)
        {
            Point3d ptext = new Point3d(p0.X + 100, p0.Y + 70, p0.Z);
            double drill_angle = Math.Round(GetDrillAngle(drill1), 1);
            double azimuth_angle = Math.Round(CalculAzimuth(GetCompassAngle(db), drill), 1);
            double length = Math.Round(drill1.Length, 1);
            String dimension_angle = name + "  " + azimuth_angle + "°  ∠ " + drill_angle + "° " + length + "m";
            sql[2][sql[2].Count - 1] += " " + azimuth_angle + " " + drill_angle + " " + length + " 1.0";
            InsertText(db, ptext, dimension_angle, "设计钻孔注记", 3);
        }


        //更改钻孔颜色(顺层无断层)
        public void DrillColor(Database db, Line drill, Spline spl1, Spline spl2)
        {
            Point3dCollection point3DCollection1 = new Point3dCollection();
            Point3dCollection point3DCollection2 = new Point3dCollection();
            Line line1 = new Line(new Point3d(drill.StartPoint.X, drill.StartPoint.Y, 0), new Point3d(drill.EndPoint.X, drill.EndPoint.Y, 0));
            using (TransactionHost host = new TransactionHost(db))
            {
                spl1.IntersectWith(line1, Intersect.OnBothOperands, new Plane(), point3DCollection1, IntPtr.Zero, IntPtr.Zero);
                spl2.IntersectWith(line1, Intersect.OnBothOperands, new Plane(), point3DCollection2, IntPtr.Zero, IntPtr.Zero);
                if (point3DCollection1.Count == 2)
                {
                    Line line = new Line(point3DCollection1[0], point3DCollection1[1]);
                    line.ColorIndex = 30;
                    line.Layer = "刨面钻孔";
                    line.Linetype = "ByLayer";
                    line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                    host.AddEntity(line);
                }
                if (point3DCollection2.Count == 2)
                {
                    Line line = new Line(point3DCollection2[0], point3DCollection2[1]);
                    line.ColorIndex = 30;
                    line.Layer = "刨面钻孔";
                    line.Linetype = "ByLayer";
                    line.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                    host.AddEntity(line);
                }
                host.Commit();
            }
        }

        //煤层拟合点Y间隔(等高线间距设置无端层)
        public void Coal10(Point3dCollection coal)
        {
            int num = coal.Count - 1;
            for (int i = 0; i < num; i++)
            {
                Point3d point = BaseTool.GetCenterPointBetweenTwoPoint(coal[2 * i], coal[2 * i + 1]);
                if (Math.Abs((int)coal[2 * i].Y - (int)coal[2 * i + 1].Y) == 20)
                    coal.Insert(2 * i + 1, point);
                else if (Math.Abs((int)coal[2 * i].Y - (int)coal[2 * i + 1].Y) != 0)
                {
                    coal.Insert(2 * i + 1, coal[2 * i]);
                }
                else
                {
                    if (10 * (coal[2 * i + 1].X - coal[2 * i].X) / 200 > 10)
                        coal.Insert(2 * i + 1, new Point3d(point.X, point.Y - 10, point.Z));
                    else
                        coal.Insert(2 * i + 1, new Point3d(point.X, point.Y - 10 * (coal[2 * i + 1].X - coal[2 * i].X) / 200, point.Z));
                }

            }
            for (int j = 0; j < 5; j++)
            {
                for (int i = 0; i < coal.Count - 1; i++)
                {
                    if (coal[i] == coal[i + 1])
                    {
                        coal.RemoveAt(i);
                        break;
                    }
                }
            }
        }

        //煤层拟合点Y间隔(等高线间距设置有端层)
        public void Coal10Fault(Point3dCollection coal, Dictionary<String, Point3d> jiaodian, Line drill, Point3d p0)
        {
            double x1, x2;
            x1 = BaseTool.GetDistanceBetweenToPoint(drill.StartPoint, jiaodian["断层上盘"]);
            x2 = BaseTool.GetDistanceBetweenToPoint(drill.StartPoint, jiaodian["断层下盘"]);

            int num = coal.Count - 1;
            for (int i = 0; i < num; i++)
            {
                Point3d point = BaseTool.GetCenterPointBetweenTwoPoint(coal[2 * i], coal[2 * i + 1]);
                if (Math.Abs((int)coal[2 * i].Y - (int)coal[2 * i + 1].Y) == 20)
                {
                    if (p0.X + x1 + 20 < coal[2 * i].X || p0.X + x1 + 20 > coal[2 * i + 1].X)
                    {
                        if (p0.X + x2 + 20 < coal[2 * i].X || p0.X + x2 + 20 > coal[2 * i + 1].X)
                            coal.Insert(2 * i + 1, point);
                        else
                        {
                            coal.Insert(2 * i + 1, coal[2 * i]);
                        }
                    }
                    else
                    {
                        coal.Insert(2 * i + 1, coal[2 * i]);
                    }
                }
                else if (Math.Abs((int)coal[2 * i].Y - (int)coal[2 * i + 1].Y) != 0)
                {
                    coal.Insert(2 * i + 1, coal[2 * i]);
                }
                else
                {
                    if (10 * (coal[2 * i + 1].X - coal[2 * i].X) / 100 > 10)
                        coal.Insert(2 * i + 1, new Point3d(point.X, point.Y - 10, 0));
                    else
                        coal.Insert(2 * i + 1, new Point3d(point.X, point.Y - 10 * (coal[2 * i + 1].X - coal[2 * i].X) / 100, 0));
                }

            }
            for (int j = 0; j < 5; j++)
            {
                for (int i = 0; i < coal.Count - 1; i++)
                {
                    if (coal[i] == coal[i + 1])
                    {
                        coal.RemoveAt(i);
                        break;
                    }
                }
            }
        }

        //计算方位角
        public double CalculAzimuth(double angle_zhibeizhen, Line drill)
        {
            double drill_angle = drill.Angle / Math.PI * 180;
            double angle = 0;
            if (angle_zhibeizhen >= drill_angle)
            {
                angle = angle_zhibeizhen - drill_angle;
            }
            else
            {
                angle = 360 + angle_zhibeizhen - drill_angle;
            }
            return angle;
        }

        //生成填充图案
        public void GenerateHatch(Database db, Point3d p0, int num, int lines)
        {
            ObjectId id;
            ObjectId hatchId;
            Point3d p = new Point3d(p0.X, p0.Y, p0.Z);
            for (int i = 0; i < lines; i++)
            {
                using (TransactionHost host = new TransactionHost(db))
                {
                    Point3dCollection points = new Point3dCollection();
                    points.Add(p);
                    points.Add(new Point3d(p.X + 1.5, p.Y, p.Z));
                    points.Add(new Point3d(p.X + 1.5, p.Y + 10, p.Z));
                    points.Add(new Point3d(p.X, p.Y + 10, p.Z));
                    Polyline3d polyline = new Polyline3d(Poly3dType.SimplePoly, points, true);
                    polyline.Layer = "图例";
                    polyline.Linetype = "ByLayer";
                    id = host.AddEntity(polyline);
                    host.Commit();
                    // 声明一个图案填充对象
                    Hatch hatch = new Hatch();
                    // 设置填充比例
                    hatch.PatternScale = 1;
                    // 设置填充类型和图案名称
                    hatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    // 加入图形数据库
                    hatch.Layer = "图例";
                    hatch.Linetype = "ByLayer";
                    hatchId = host.AddEntity(hatch);
                    // 设置关联
                    hatch.Associative = true;
                    // 设置边界图形和填充方式
                    ObjectIdCollection obIds = new ObjectIdCollection();
                    obIds.Add(id);
                    hatch.AppendLoop(HatchLoopTypes.Outermost, obIds);
                    // 计算填充并显示
                    hatch.EvaluateHatch(true);
                    // 提交事务
                    host.Commit();
                }
                p = new Point3d(p.X + 1.5 * Math.Pow(-1, i + num), p.Y + 10, p.Z);
            }
        }

        //生成背景图
        public void GenerateBackground(Database db, int num)
        {
            Point3d p1;
            for (int i = 0; i < num; i++)
            {
                p1 = new Point3d((300 * i) % 1500, -200 * (i / 5), 0);
                GenerateHatch(db, new Point3d(p1.X - 3, p1.Y, p1.Z), 0, 12);
                GenerateHatch(db, new Point3d(p1.X + 241.5, p1.Y, p1.Z), 1, 12);
                GenerateLine(db, p1, 12, 240);
            }

        }

        //生成背景图
        public void GenerateBackground1(Database db, int num)
        {
            Point3d p1;
            for (int i = 0; i < num; i++)
            {
                p1 = new Point3d((230 * i) % 1150, -100 * (i / 5), 0);
                GenerateHatch(db, new Point3d(p1.X - 3, p1.Y, p1.Z), 0, 4);
                GenerateHatch(db, new Point3d(p1.X + 171.5, p1.Y, p1.Z), 1, 4);
                GenerateLine(db, p1, 4, 170);
            }

        }

        //生成高度刻度网格横线
        public void GenerateLine(Database db, Point3d p0, int lines, int length)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                Line line1 = new Line(new Point3d(p0.X + length / 3, p0.Y + 7 + lines * 10, p0.Z), new Point3d(p0.X + 2 * length / 3, p0.Y + 7 + 10 * lines, p0.Z));
                line1.Layer = "图例";
                host.AddEntity(line1);
                for (int i = 0; i < lines + 1; i++)
                {
                    Line line = new Line(new Point3d(p0.X - 3, p0.Y + i * 10, p0.Z), new Point3d(p0.X + 3 + length, p0.Y + i * 10, p0.Z));
                    line.Linetype = "ByLayer";
                    line.Layer = "图例";
                    host.AddEntity(line);
                }
                Line line2 = new Line(new Point3d(p0.X, p0.Y, p0.Z), new Point3d(p0.X, p0.Y + 10 * lines, p0.Z));
                Line line3 = new Line(new Point3d(p0.X + length, p0.Y, p0.Z), new Point3d(p0.X + length, p0.Y + 10 * lines, p0.Z));
                Line line4 = new Line(new Point3d(p0.X + 3 + length, p0.Y, p0.Z), new Point3d(p0.X + 3 + length, p0.Y + 10 * lines, p0.Z));
                Line line5 = new Line(new Point3d(p0.X - 3, p0.Y, p0.Z), new Point3d(p0.X - 3, p0.Y + 10 * lines, p0.Z));
                line2.Layer = "图例";
                line3.Layer = "图例";
                line4.Layer = "图例";
                line5.Layer = "图例";
                line1.Linetype = "ByLayer";
                line2.Linetype = "ByLayer";
                line3.Linetype = "ByLayer";
                line4.Linetype = "ByLayer";
                line5.Linetype = "ByLayer";
                host.AddEntity(line2);
                host.AddEntity(line3);
                host.AddEntity(line4);
                host.AddEntity(line5);
                host.Commit();
            }
        }

        //计算指北针角度
        public double GetCompassAngle(Database db)
        {
            double angle_compass = 0;
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);//实体化对象，entity对象非常关键
                    if (entity.GetType() == typeof(BlockReference))
                    {
                        BlockReference blockReference = entity as BlockReference;
                        if (blockReference.Name == "指北针")
                        {
                            angle_compass = BaseTool.AngleToDegree(blockReference.Rotation);

                        }
                    }
                }
            }
            return angle_compass;
        }

        //返回所有钻孔实体集合
        public List<Line> SearchDrill(Database db)
        {
            List<Line> lines = new List<Line>();
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);//实体化对象，entity对象非常关键
                    //找钻孔
                    if (entity.GetType() == typeof(Line) && entity.Layer.Contains("设计校验钻孔"))//找钻孔的图层
                    {
                        lines.Add((Line)entity);
                    }
                    if (entity.GetType() == typeof(Polyline) && entity.Layer.Contains("设计校验钻孔"))//找钻孔的图层
                    {
                        Polyline polyline = (Polyline)entity;
                        Point3d start = new Point3d(polyline.StartPoint.X, polyline.StartPoint.Y, 0);
                        Point3d end = new Point3d(polyline.EndPoint.X, polyline.EndPoint.Y, 0);
                        Line line = new Line(start, end);
                        line.Layer = polyline.Layer;
                        lines.Add(line);
                    }
                }
            }
            return lines;
        }

        //取样点
        public bool HasCircle(Database db, Line drill)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);//实体化对象，entity对象非常关键
                    if (entity.GetType() == typeof(Circle) && entity.Layer == "取样点")
                    {
                        Circle circle = (Circle)entity;
                        if ((int)drill.EndPoint.X == (int)circle.Center.X && (int)drill.EndPoint.Y == (int)circle.Center.Y)
                            return true;
                    }
                }
            }
            return false;
        }

        //复制Point3dCollection
        public void Collection(Point3dCollection c1, Point3dCollection c2)
        {
            foreach (Point3d point in c1)
            {
                c2.Add(point);
            }
        }

        //插入高度刻度(顺层)
        public double InsertHeightText(Database db, Point3d p0, List<double> height, Point3dCollection coal, Point3dCollection coal_cut, double drillheight)
        {

            List<double> height_cut = new List<double>();
            for (int i = 0; i < height.Count; i++)
            {
                foreach (Point3d point in coal_cut)
                {
                    if (coal[i] == point)
                        height_cut.Add(height[i]);
                }
            }
            double min;
            if (height_cut.Count > 0)
            {
                min = height_cut.Min() - 10;
            }
            else
            {
                min = height.Min();
            }
            if (drillheight < min)
                min = (int)drillheight / 10 * 10;
            for (int i = 0; i < coal_cut.Count; i++)
            {
                coal_cut[i] = new Point3d(coal_cut[i].X, coal_cut[i].Y - min, coal_cut[i].Z);
            }
            using (TransactionHost host = new TransactionHost(db))
            {
                for (int i = 0; i < 13; i++)
                {
                    MText insertDimension = new MText();
                    insertDimension.Location = new Point3d(p0.X - 10, p0.Y + 2 + i * 10, 0);
                    insertDimension.Contents = (min + (i - 3) * 10).ToString();
                    insertDimension.TextHeight = 2.1;
                    insertDimension.Attachment = AttachmentPoint.TopLeft;
                    insertDimension.Layer = "等高线注记";
                    host.AddEntity(insertDimension);

                    MText insertDimension1 = new MText();
                    insertDimension1.Location = new Point3d(p0.X + 244, p0.Y + 2 + i * 10, 0);
                    insertDimension1.Contents = (min + (i - 3) * 10).ToString();
                    insertDimension1.TextHeight = 2.1;
                    insertDimension1.Attachment = AttachmentPoint.TopLeft;
                    insertDimension1.Layer = "等高线注记";
                    host.AddEntity(insertDimension1);
                }
                host.Commit();
            }
            return min;
        }



        //生成巷道剖面
        public void GenerateHangdao(Database db, Point3d tunnelpoint, double tunnelwidth, bool UpOrDown)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                Point3dCollection points = new Point3dCollection();
                if (UpOrDown)
                {
                    points.Add(tunnelpoint);
                    points.Add(new Point3d(tunnelpoint.X - tunnelwidth, tunnelpoint.Y, 0));
                    points.Add(new Point3d(tunnelpoint.X - tunnelwidth, tunnelpoint.Y + 3.8, 0));
                    points.Add(new Point3d(tunnelpoint.X, tunnelpoint.Y + 2.6, 0));
                }
                else
                {
                    points.Add(tunnelpoint);
                    points.Add(new Point3d(tunnelpoint.X - tunnelwidth, tunnelpoint.Y, 0));
                    points.Add(new Point3d(tunnelpoint.X - tunnelwidth, tunnelpoint.Y + 2.6, 0));
                    points.Add(new Point3d(tunnelpoint.X, tunnelpoint.Y + 3.8, 0));
                }

                Polyline3d polyline = new Polyline3d(Poly3dType.SimplePoly, points, true);
                polyline.Layer = "煤巷";
                polyline.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                host.AddEntity(polyline);
                host.Commit();
            }
        }


        //Line和Circle的交点
        public Point3d GetInter(Line line, Point3d p0)
        {
            Arc arc = new Arc(p0, 2.4, 0, Math.PI);
            Line line1 = new Line(new Point3d(p0.X - 2.4, p0.Y, p0.Z), new Point3d(p0.X - 2.4, p0.Y - 1.5, p0.Z));
            Line line2 = new Line(new Point3d(p0.X + 2.4, p0.Y, p0.Z), new Point3d(p0.X + 2.4, p0.Y - 1.5, p0.Z));
            Line line3 = new Line(new Point3d(p0.X - 2.4, p0.Y - 1.5, p0.Z), new Point3d(p0.X + 2.4, p0.Y - 1.5, p0.Z));
            Point3dCollection point3DCollection = new Point3dCollection();
            line.IntersectWith(arc, Intersect.OnBothOperands, new Plane(), point3DCollection, IntPtr.Zero, IntPtr.Zero);
            if (point3DCollection.Count == 1)
            {
                return point3DCollection[0];
            }
            line.IntersectWith(line1, Intersect.OnBothOperands, new Plane(), point3DCollection, IntPtr.Zero, IntPtr.Zero);
            if (point3DCollection.Count == 1)
            {
                return point3DCollection[0];
            }
            line.IntersectWith(line2, Intersect.OnBothOperands, new Plane(), point3DCollection, IntPtr.Zero, IntPtr.Zero);
            if (point3DCollection.Count == 1)
            {
                return point3DCollection[0];
            }
            line.IntersectWith(line3, Intersect.OnBothOperands, new Plane(), point3DCollection, IntPtr.Zero, IntPtr.Zero);
            if (point3DCollection.Count == 1)
            {
                return point3DCollection[0];
            }
            return p0;
        }

        //生成单个钻孔(无断层)
        public Line GenerateDrill(Database db, Line drill, double thick, Point3dCollection coal, Point3d tunnelpoint)
        {
            double drill_x = coal[0].X + drill.Length;
            Point3dCollection points_drill = GetPointByX(drill_x, coal);
            Point3d drill_start = new Point3d(tunnelpoint.X, tunnelpoint.Y + 1.3, 0);
            Point3d drill_end = new Point3d(drill_x, points_drill[0].Y + thick / 5.0, 0);

            Line drill1 = new Line(drill_start, drill_end);
            using (TransactionHost host = new TransactionHost(db))
            {
                drill1.Layer = "刨面钻孔";
                drill1.Linetype = "ByLayer";
                drill1.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                host.AddEntity(drill1);
                host.Commit();
            }
            return drill1;
        }

        //生成单个钻孔(有断层)
        public Line GenerateDrill1(Database db, Line drill, double thick, Point3dCollection coal1, Point3dCollection coal2, Point3d tunnelpoint)
        {
            double drill_x = coal1[0].X + drill.Length;
            Point3dCollection points_drill = GetPointByX(drill_x, coal1);
            if (points_drill.Count == 0)
            {
                points_drill = GetPointByX(drill_x, coal2);
            }
            Point3d drill_start = new Point3d(tunnelpoint.X, tunnelpoint.Y + 1.3, 0);
            Point3d drill_end = new Point3d(drill_x, points_drill[0].Y + thick / 5.0, 0);
            Line drill1 = new Line(drill_start, drill_end);
            using (TransactionHost host = new TransactionHost(db))
            {
                drill1.Layer = "刨面钻孔";
                drill1.Linetype = "ByLayer";
                drill1.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
                host.AddEntity(drill1);
                host.Commit();
            }
            return drill1;
        }

        //煤层剖面线起点位置
        public void CoalStart(Point3dCollection coal, double drillheight, double min, Point3d p0)
        {
            if (coal[0].Y == coal[1].Y)
            {
                coal.RemoveAt(0);
                coal.Insert(0, new Point3d(p0.X + 20, p0.Y + drillheight - min + 30, 0));
            }
        }

        //去除多余的煤层曲线部分(顺层)
        public void CutSpline(Point3dCollection coal, Point3d p0)
        {
            double x1 = p0.X + 20;
            double x2 = p0.X + 240;
            Point3dCollection points1 = GetPointByX(x1, coal);
            Point3dCollection points2 = GetPointByX(x2, coal);
            Point3d p1 = new Point3d();
            Point3d p2 = new Point3d();
            if (points1.Count > 0)
                p1 = points1[0];
            else
                p1 = new Point3d(x1, coal[0].Y, 0);
            if (points2.Count > 0)
                p2 = points2[0];
            else
                p2 = new Point3d(x2, coal[coal.Count - 1].Y, 0);
            Point3dCollection delete = new Point3dCollection();
            foreach (Point3d point in coal)
            {
                if (point.X >= x2 || point.X <= x1)
                {
                    delete.Add(point);
                }
            }
            foreach (Point3d point in delete)
            {
                coal.Remove(point);
            }
            coal.Insert(0, p1);
            coal.Add(p2);
        }

        //根据X坐标获取完整坐标
        public Point3dCollection GetPointByX(double x, Point3dCollection points)
        {
            Point3d p1 = new Point3d(x, 0, 0);
            Point3d p2 = new Point3d(x, 1, 0);
            Line line = new Line(p1, p2);
            Point3dCollection points1 = new Point3dCollection();
            Spline spline = new Spline(points, 4, 0.0);
            line.IntersectWith(spline, Intersect.ExtendBoth, new Plane(), points1, IntPtr.Zero, IntPtr.Zero);
            return points1;
        }

        //根据x的大小对拟合点排序        
        public void SortCoal(Point3dCollection coal)
        {
            for (int i = 0; i < coal.Count; i++)
            {
                for (int j = 0; j < coal.Count - i - 1; j++)
                {
                    if (coal[j].X > coal[j + 1].X)
                    {
                        Point3d p = coal[j];
                        coal.RemoveAt(j);
                        coal.Insert(j + 1, p);

                    }
                }
            }
        }



        //根据y的大小对拟合点排序        
        public void SortCoalY(Point3dCollection coal)
        {
            for (int i = 0; i < coal.Count; i++)
            {
                for (int j = 0; j < coal.Count - i - 1; j++)
                {
                    if (coal[j].Y > coal[j + 1].Y)
                    {
                        Point3d p = coal[j];
                        coal.RemoveAt(j);
                        coal.Insert(j + 1, p);

                    }
                }
            }
        }

        //生成煤层剖面线(顺层)
        public Spline GenerateCoalSeam(Database db, Point3dCollection coal)
        {
            Spline spline = new Spline(coal, 4, 0.0);
            spline.Layer = "煤层";
            spline.Linetype = "ByLayer";
            spline.ColorIndex = 8;
            spline.LineWeight = Autodesk.AutoCAD.DatabaseServices.LineWeight.LineWeight030;
            spline.SetDatabaseDefaults();
            using (TransactionHost host = new TransactionHost(db))
            {
                host.AddEntity(spline);
                host.Commit();
            }
            return spline;
        }



        //煤层剖面上底板拟合点
        public Point3dCollection Coal2(Point3dCollection coal, double thick)
        {
            Point3dCollection coal2 = new Point3dCollection();
            foreach (Point3d point in coal)
            {
                coal2.Add(new Point3d(point.X, point.Y + thick, 0));
            }
            return coal2;
        }

        //煤层剖面下底板拟合点(顺层)
        public Point3dCollection Coal1(Point3d startpoint, Point3d point_road, Point3d p0, Point3dCollection points, int constant, List<double> height)
        {
            double dis1 = BaseTool.GetDistanceBetweenToPoint(point_road, startpoint);
            Point3dCollection coal_points = new Point3dCollection();
            for (int i = 0; i < points.Count; i++)
            {
                double dis2 = BaseTool.GetDistanceBetweenToPoint(points[i], startpoint);
                coal_points.Add(new Point3d(p0.X + constant + dis2 - dis1, p0.Y + 30 + height[i], 0));
            }
            return coal_points;
        }


        //获取开孔点巷道高度
        public double GroupHeight(Database db, Line group, List<Circle> Circles)
        {
            double drillheight = 0;

            List<DBText> gaocheng = new List<DBText>();
            GetHeightMark(db, gaocheng);
            for (int i = Circles.Count - 1; i >= 0; i--)
            {
                if (i != Circles.Count - 1 && Circles[i].Center.X < group.StartPoint.X)
                {
                    String s1 = NearText(db, Circles[i], gaocheng).TextString;
                    String s2 = NearText(db, Circles[i + 1], gaocheng).TextString;
                    double height1 = Convert.ToDouble(s1);
                    double height2 = Convert.ToDouble(s2);
                    if (s1.Count() - s1.IndexOf('.') - 1 == 3)
                    {
                        height1 -= 2.6;
                    }
                    if (s2.Count() - s2.IndexOf('.') - 1 == 3)
                    {
                        height2 -= 2.6;
                    }
                    drillheight = height1 + (height2 - height1) * (group.StartPoint.X - Circles[i].Center.X) / (Circles[i + 1].Center.X - Circles[i].Center.X);
                    break;
                }
            }
            if (drillheight == 0)
            {
                String s1 = NearText(db, Circles[0], gaocheng).TextString;
                drillheight = Convert.ToDouble(s1);
                if (s1.Count() - s1.IndexOf('.') - 1 == 3)
                {
                    drillheight -= 2.6;
                }
            }
            return drillheight;
        }

        //确定煤层两侧边界坐标(穿层)
        public void Edge(Point3dCollection coal_points, Point3d p0)
        {
            if (coal_points[0].X > p0.X)
            {
                if (coal_points.Count == 1)
                    coal_points.Insert(0, new Point3d(p0.X, coal_points[0].Y, 0));
                else
                {
                    double deltY = (coal_points[1].Y - coal_points[0].Y) * (coal_points[0].X - p0.X) / (coal_points[1].X - coal_points[0].X);
                    if (deltY > 10)
                        coal_points.Insert(0, new Point3d(p0.X, coal_points[0].Y - 10, 0));
                    else
                        coal_points.Insert(0, new Point3d(p0.X, coal_points[0].Y - deltY, 0));
                }
            }
            else
            {
                Point3dCollection points_drill = GetPointByX(p0.X, coal_points);
                coal_points.RemoveAt(0);
                coal_points.Insert(0, new Point3d(p0.X, points_drill[0].Y, p0.Z));
            }
            if (coal_points[coal_points.Count - 1].X < p0.X + 170)
            {
                if (coal_points.Count == 1)
                    coal_points.Add(new Point3d(p0.X + 170, coal_points[coal_points.Count - 1].Y, 0));
                else
                {
                    int i = coal_points.Count - 1;
                    double deltY = (coal_points[i].Y - coal_points[i - 1].Y) * (p0.X + 170 - coal_points[i].X) / (coal_points[i].X - coal_points[i - 1].X);
                    if (deltY > 10)
                    {
                        coal_points.Add(new Point3d(p0.X + 170, coal_points[i].Y + 10, 0));
                    }
                    else
                    {
                        coal_points.Add(new Point3d(p0.X + 170, coal_points[i].Y + deltY, 0));
                    }

                }
            }
            else
            {
                Point3dCollection points_drill = GetPointByX(p0.X + 170, coal_points);
                coal_points.RemoveAt(coal_points.Count - 1);
                coal_points.Add(new Point3d(p0.X + 170, points_drill[0].Y, p0.Z));
            }

        }

        //找到起始点
        public Point3d FindStartPoint(List<Point3d> edges, Point3d point_road)
        {
            if (BaseTool.GetDistanceBetweenToPoint(edges[0], point_road) > BaseTool.GetDistanceBetweenToPoint(edges[0], edges[1]))
                return point_road;
            if (BaseTool.GetDistanceBetweenToPoint(edges[1], point_road) > BaseTool.GetDistanceBetweenToPoint(edges[0], edges[1]))
                return point_road;
            if (BaseTool.GetDistanceBetweenToPoint(edges[0], point_road) < BaseTool.GetDistanceBetweenToPoint(edges[1], point_road))
                return edges[0];
            else
                return edges[1];
        }

        //找到剖面边界
        public List<Point3d> FindEdge(Point3dCollection points)
        {
            double dis = 0;
            List<Point3d> edges = new List<Point3d>();
            for (int i = 0; i < points.Count; i++)
            {
                for (int j = i + 1; j < points.Count; j++)
                {
                    if (dis < BaseTool.GetDistanceBetweenToPoint(points[i], points[j]))
                    {
                        edges.Clear();
                        dis = BaseTool.GetDistanceBetweenToPoint(points[i], points[j]);
                        edges.Add(points[i]);
                        edges.Add(points[j]);
                    }
                }
            }
            return edges;
        }

        //判断该条钻孔剖面是否存在断层(顺层)
        public bool HasFault(Database db, Line drill, Point3dCollection coal)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Point3dCollection points = new Point3dCollection();
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    if (entity.Layer == "断层上盘")
                    {
                        //求交点，钻孔和断层线必须共面
                        Line line = new Line(new Point3d(drill.StartPoint.X, drill.StartPoint.Y, 0), new Point3d(drill.EndPoint.X, drill.EndPoint.Y, 0));
                        line.IntersectWith(entity, Intersect.ExtendThis, points, IntPtr.Zero, IntPtr.Zero);
                        if (points.Count > 0)
                        {
                            if (drill.StartPoint.Y > drill.EndPoint.Y)
                            {
                                if (drill.StartPoint.GetDistanceBetweenToPoint(points[0]) < 160 && (int)points[0].Y < (int)drill.StartPoint.Y)
                                    return true;
                            }
                            if (drill.StartPoint.Y < drill.EndPoint.Y)
                            {
                                if (drill.StartPoint.GetDistanceBetweenToPoint(points[0]) < 160 && (int)points[0].Y > (int)drill.StartPoint.Y)
                                    return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        //获取每个钻孔对应的编号名称
        public void GetName(Database db, List<String> names, List<Line> drills)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                for (int i = 0; i < drills.Count; i++)
                {
                    Point3d point3D = drills[i].EndPoint;
                    foreach (ObjectId id in host.ModelSpace)
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                        if (entity != null)
                        {
                            if (entity.Layer == "标注" && entity.GetType() == typeof(MText))
                            {
                                MText text = entity as MText;
                                if ((int)text.Location.X == (int)point3D.X && (int)text.Location.Y == (int)point3D.Y)
                                {
                                    names.Add(text.Contents.Trim());
                                }
                            }
                            if (entity.Layer == "标注" && entity.GetType() == typeof(DBText))
                            {
                                DBText text = entity as DBText;
                                if ((int)text.Position.X == (int)point3D.X && (int)text.Position.Y == (int)point3D.Y)
                                {
                                    names.Add(text.TextString.ToString().Trim());
                                }
                            }
                        }
                    }
                }
            }
        }

        //钻孔倾角
        public double GetDrillAngle(Line drill)
        {
            double angle_drill = drill.Angle.AngleToDegree();
            if (angle_drill > 180)
            {
                return -1 * (360.0 - angle_drill);
            }
            return angle_drill;
        }

        //记录钻孔与等高线的交点
        public List<double> Isoheight(Database db, Line line, Point3dCollection points)
        {
            List<double> height = new List<double>();
            int i = 0;
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    String layername = entity.Layer;
                    if (layername == "223等高线")
                    {
                        if (entity.GetType() == typeof(Polyline3d))
                        {
                            Polyline3d polyline = entity as Polyline3d;
                            if (line != null)
                            {
                                //entity1.IntersectWith(entity, Intersect.OnBothOperands, new Plane(), points, IntPtr.Zero, IntPtr.Zero);
                                /*
                                 OnBothOperands :两个实体（这里是线段）都不延申
                                ExtendBoth :两个实体都延申
                                ExtendArgument :只延申作为参数的实体（第一个参数entity）
                                ExtendThis :只延申原实体（entity1）
                                 */
                                //points是得到的交点信息。后面要跟两个IntPtr.Zero，如果输入0的话，是已经弃用的方法，不推荐。
                                line.IntersectWith(entity, Intersect.ExtendThis, new Plane(), points, IntPtr.Zero, IntPtr.Zero);
                                while (points.Count > i)
                                {
                                    height.Add(polyline.StartPoint.Z);
                                    i++;
                                }
                            }
                        }
                        else if (entity.GetType() == typeof(Polyline))
                        {
                            Polyline polyline = entity as Polyline;
                            if (line != null)
                            {
                                line.IntersectWith(entity, Intersect.ExtendThis, new Plane(), points, IntPtr.Zero, IntPtr.Zero);
                                while (points.Count > i)
                                {
                                    height.Add(polyline.Elevation);
                                    i++;
                                }
                            }
                        }
                    }
                }
            }
            return height;
        }

        //记录钻孔与煤巷的交点
        public Point3d RoadwayIntersection(Database db, Line line)
        {
            Point3dCollection points = new Point3dCollection();//points1 是钻孔延长线和煤巷的交点
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    String layername = entity.Layer;
                    if (layername == "133煤巷" || layername == "134岩巷" || layername == "设计煤巷" || layername == "设计岩巷" || layername == "煤巷" || layername == "岩巷")
                    {
                        line.IntersectWith(entity, Intersect.ExtendThis, new Plane(), points, IntPtr.Zero, IntPtr.Zero);
                    }
                }
            }
            int dis1 = 100;
            Point3d p = new Point3d();
            foreach (Point3d point in points)
            {
                int dis2 = (int)BaseTool.GetDistanceBetweenToPoint(point, line.StartPoint);

                if (dis2 < dis1 && dis2 != 0)
                {
                    dis1 = dis2;
                    p = point;
                }
            }
            return p;
        }

        //插入text到指定位置
        public void InsertText(Database db, Point3d p0, String contents, String layername, double TextHeight)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                MText inserDimension1 = new MText();
                inserDimension1.Location = p0;
                inserDimension1.Contents = contents;
                inserDimension1.TextHeight = TextHeight;
                inserDimension1.Attachment = AttachmentPoint.TopLeft;
                inserDimension1.Layer = layername;
                host.AddEntity(inserDimension1);
                host.Commit();
            }
        }

        //导入外部dwg文件中自定义的块
        public static void ImportBlocksFromDWG(string dwgFileName, Database currentDB)
        {
            Database tempDB = new Database(false, true);//新建一个临时图形数据库
            try
            {
                tempDB.ReadDwgFile(dwgFileName, System.IO.FileShare.Read, true, null);//读取外部DWG到临时数据库
                ObjectIdCollection ids = new ObjectIdCollection();//存放块OjbectId的集合
                using (Transaction trans = tempDB.TransactionManager.StartTransaction())
                {
                    BlockTable bt = trans.GetObject(tempDB.BlockTableId, OpenMode.ForRead) as BlockTable;
                    foreach (ObjectId item in bt)
                    {
                        BlockTableRecord btr = trans.GetObject(item, OpenMode.ForRead) as BlockTableRecord;
                        if (!btr.IsAnonymous && !btr.IsLayout)//筛选自定义的命名块,剔除匿名块和布局块
                        {
                            ids.Add(item);
                        }
                    }
                }
                //复制临时数据库中的块到当前数据库的块表中
                tempDB.WblockCloneObjects(ids, currentDB.BlockTableId, new IdMapping(), DuplicateRecordCloning.Replace, false);
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                Application.ShowAlertDialog("错误：" + ex.Message);
            }
            tempDB.Dispose();
        }

        //新建一个给定名字的图层
        public ObjectId AddIn(Database db, String layername)
        {
            ObjectId layerId = ObjectId.Null;
            using (TransactionHost host = new TransactionHost(db))
            {
                LayerTable lt = (LayerTable)host.GetObject(db.LayerTableId, OpenMode.ForWrite);
                if (!lt.Has(layername))
                {
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layername;
                    layerId = lt.Add(ltr);
                    host.AddNewlyCreatedDBObject(ltr, true);

                }
                host.Commit();
            }
            return layerId;
        }

        //断层标记
        public void MarkFault(Database db, Point3d p1, Point3d p2, bool isLeft)
        {
            Line line = new Line(p1, p2);
            Line line1 = new Line(new Point3d(0, 0, 0), new Point3d(90, 0, 0));
            Line line2 = new Line(new Point3d(65, -3, 0), new Point3d(90, -3, 0));
            Line line3 = new Line(new Point3d(65, 3, 0), new Point3d(90, 3, 0));
            Line line4;
            Line line5;

            if (isLeft)
            {
                line4 = new Line(new Point3d(65, -3, 0), new Point3d(75, -4.5, 0));
                line5 = new Line(new Point3d(90, 3, 0), new Point3d(80, 4.5, 0));
            }
            else
            {
                line4 = new Line(new Point3d(90, -3, 0), new Point3d(80, -4.5, 0));
                line5 = new Line(new Point3d(65, 3, 0), new Point3d(75, 4.5, 0));
            }

            using (TransactionHost host = new TransactionHost(db))
            {
                Matrix3d mt = Matrix3d.Rotation(line.Angle, Vector3d.ZAxis, new Point3d(0, 0, 0));
                line1.TransformBy(mt);
                line2.TransformBy(mt);
                line3.TransformBy(mt);
                line4.TransformBy(mt);
                line5.TransformBy(mt);
                Point3d midpoint = BaseTool.GetCenterPointBetweenTwoPoint(line1.StartPoint, line1.EndPoint);
                Vector3d vector3D = p1 - midpoint;
                line1.StartPoint = line1.StartPoint.Add(vector3D);
                line1.EndPoint = line1.EndPoint.Add(vector3D);
                line2.StartPoint = line2.StartPoint.Add(vector3D);
                line2.EndPoint = line2.EndPoint.Add(vector3D);
                line3.StartPoint = line3.StartPoint.Add(vector3D);
                line3.EndPoint = line3.EndPoint.Add(vector3D);
                line4.StartPoint = line4.StartPoint.Add(vector3D);
                line4.EndPoint = line4.EndPoint.Add(vector3D);
                line5.StartPoint = line5.StartPoint.Add(vector3D);
                line5.EndPoint = line5.EndPoint.Add(vector3D);
                line1.ColorIndex = 1;
                line2.ColorIndex = 1;
                line3.ColorIndex = 1;
                line4.ColorIndex = 1;
                line5.ColorIndex = 1;
                line1.Layer = "断层剖面";
                line2.Layer = "断层剖面";
                line3.Layer = "断层剖面";
                line4.Layer = "断层剖面";
                line5.Layer = "断层剖面";
                line1.Linetype = "ByLayer";
                line2.Linetype = "ByLayer";
                line3.Linetype = "ByLayer";
                line4.Linetype = "ByLayer";
                line5.Linetype = "ByLayer";
                host.AddEntity(line1);
                host.AddEntity(line2);
                host.AddEntity(line3);
                host.AddEntity(line4);
                host.AddEntity(line5);
                host.Commit();
            }
        }

        //获取直线和Spline的交点
        public Point3d GetJiaodian(Line line, Point3dCollection coal)
        {
            Point3dCollection point3DCollection = new Point3dCollection();
            Spline spline = new Spline(coal, 4, 0.0);
            line.IntersectWith(spline, Intersect.ExtendThis, point3DCollection, IntPtr.Zero, IntPtr.Zero);
            if (point3DCollection.Count > 0)
                return point3DCollection[0];
            else
                return new Point3d(0, 0, 0);
        }

        //获取Spline和圆的交点
        public Point3d GetJiaodianCircle(Point3dCollection coal, Circle circle)
        {
            Point3dCollection point3DCollection = new Point3dCollection();
            Spline spline = new Spline(coal, 4, 0.0);
            circle.IntersectWith(spline, Intersect.OnBothOperands, point3DCollection, IntPtr.Zero, IntPtr.Zero);
            if (point3DCollection.Count == 1 && point3DCollection[0].X > circle.Center.X)
                return point3DCollection[0];
            else if (point3DCollection.Count == 2)
            {
                if (point3DCollection[1].X > point3DCollection[0].X)
                    return point3DCollection[1];
                else
                    return point3DCollection[0];
            }
            return new Point3d(1000000, 1000000, 0);
        }

        //添加直线到CAD数据库
        public void AddLine(Database db, Line drill1)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                host.AddEntity(drill1);
                host.Commit();
            }
        }

        //获得平面图中所有高程(穿层/顺层)
        public void GetHeightMark(Database db, List<DBText> texts)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity1 = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    if (entity1.Layer == "导线点高程" && entity1.GetType() == typeof(DBText))
                    {
                        texts.Add(entity1 as DBText);
                    }
                }
            }
        }

        //获得平面图中所有导线点编码(顺层/穿层)
        public void GetDaoxianMark(Database db, List<DBText> texts)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity1 = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    if (entity1.Layer == "导线点名称" && entity1.GetType() == typeof(DBText))
                    {
                        texts.Add(entity1 as DBText);
                    }
                }
            }
        }


        //获得钻孔的位置
        public void GetDrillPosition(Database db, Line drill, Point3dCollection points)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                foreach (ObjectId id in host.ModelSpace)
                {
                    Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                    if (entity != null)
                    {
                        //if(entity.Layer== "六采区灾害治理措施巷钻孔" && entity.GetType() == typeof(Circle))
                        if (entity.Layer.EndsWith("钻孔") && entity.GetType() == typeof(Circle))
                        {
                            Point3dCollection point3DCollection = new Point3dCollection();
                            drill.IntersectWith(entity, Intersect.OnBothOperands, new Plane(), point3DCollection, IntPtr.Zero, IntPtr.Zero);
                            if (point3DCollection.Count > 0)
                            {
                                Circle circle = (Circle)entity;
                                points.Add(circle.Center);
                            }
                        }
                    }
                }
            }
            SortCoalY(points);
        }


        //生成巷道剖面图
        public void GenerateTunnel(Database db, Point3d p0)
        {
            using (TransactionHost host = new TransactionHost(db))
            {
                Arc arc = new Arc(p0, 2.4, 0, Math.PI);
                Line line1 = new Line(new Point3d(p0.X - 2.4, p0.Y, p0.Z), new Point3d(p0.X - 2.4, p0.Y - 1.5, p0.Z));
                Line line2 = new Line(new Point3d(p0.X + 2.4, p0.Y, p0.Z), new Point3d(p0.X + 2.4, p0.Y - 1.5, p0.Z));
                Line line3 = new Line(new Point3d(p0.X - 2.4, p0.Y - 1.5, p0.Z), new Point3d(p0.X + 2.4, p0.Y - 1.5, p0.Z));
                arc.ColorIndex = 1;
                line1.ColorIndex = 1;
                line2.ColorIndex = 1;
                line3.ColorIndex = 1;
                host.AddEntity(arc);
                host.AddEntity(line1);
                host.AddEntity(line2);
                host.AddEntity(line3);
                host.Commit();
            }
        }

        //延长直线
        public void ProLongLine(Line line, double num)
        {
            line.EndPoint = new Point3d(line.EndPoint.X + num * Math.Cos(line.Angle), line.EndPoint.Y + num * Math.Sin(line.Angle), 0);
        }

        //输出数据库(顺层)
        public void PrintSql(List<List<String>> sql)
        {
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("# compass " + sql[3][0] + "\n");
            List<String> tunnel_name = new List<String>();
            foreach (String name in sql[0])
            {
                if (!tunnel_name.Contains(name))
                    tunnel_name.Add(name);
            }
            for (int i = 0; i < tunnel_name.Count; i++)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("# tunnel " + tunnel_name[i] + "\n");
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("# site " + "1" + " " + tunnel_name[i] + "钻场\n");
                int n = 1;
                for (int j = 0; j < sql[0].Count; j++)
                {
                    if (sql[0][j] == tunnel_name[i])
                    {
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("# drilling " + n.ToString() + " " + sql[2][j] + "\n");
                        n += 1;
                    }
                }
            }
        }

        //插入EXCEL表(顺层)
        public void dataTable(Database db, List<List<String>> sql, Point3d p0, int textHeight)
        {

            Table table = new Table();
            table.SetSize(sql[0].Count() + 1, 11);
            table.SetRowHeight(20); // 设置行高
            table.SetColumnWidth(60); // 设置列宽
            table.Columns[1].Width = 100;

            table.Position = p0; // 设置插入点
            table.Cells[0, 0].TextString = "序号";
            table.Cells[0, 0].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 1].TextString = "打钻地点";
            table.Cells[0, 1].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 2].TextString = "钻孔类型";
            table.Cells[0, 2].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 3].TextString = "孔号";
            table.Cells[0, 3].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 4].TextString = "方位角";
            table.Cells[0, 4].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 5].TextString = "倾角";
            table.Cells[0, 5].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 6].TextString = "孔深";
            table.Cells[0, 6].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 7].TextString = "钻孔距离底板高度";
            table.Cells[0, 7].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 8].TextString = "标点";
            table.Cells[0, 8].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 9].TextString = "相对标点方向";
            table.Cells[0, 9].TextHeight = textHeight; //设置文字高度
            table.Cells[0, 10].TextString = "据标点距离";
            table.Cells[0, 10].TextHeight = textHeight; //设置文字高度

            List<String> tunnel_name = new List<String>();
            foreach (String name in sql[0])
            {
                if (!tunnel_name.Contains(name))
                    tunnel_name.Add(name);
            }

            int r = 1;//记录所在行数
            for (int i = 0; i < tunnel_name.Count; i++)
            {
                for (int j = 0; j < sql[0].Count; j++)
                {
                    if (sql[0][j] == tunnel_name[i])
                    {
                        string[] sArray = sql[2][j].Split(' ');
                        table.Cells[r, 0].TextString = r.ToString();
                        table.Cells[r, 0].TextHeight = textHeight;
                        table.Cells[r, 0].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 1].TextString = tunnel_name[i];
                        table.Cells[r, 1].TextHeight = textHeight;
                        table.Cells[r, 1].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 2].TextString = "顺层钻孔";
                        table.Cells[r, 2].TextHeight = textHeight;
                        table.Cells[r, 2].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 3].TextString = sArray[0];
                        table.Cells[r, 3].TextHeight = textHeight;
                        table.Cells[r, 3].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 4].TextString = sArray[1];
                        table.Cells[r, 4].TextHeight = textHeight;
                        table.Cells[r, 4].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 5].TextString = sArray[2];
                        table.Cells[r, 5].TextHeight = textHeight;
                        table.Cells[r, 5].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 6].TextString = sArray[3];
                        table.Cells[r, 6].TextHeight = textHeight;
                        table.Cells[r, 6].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 7].TextString = sArray[4];
                        table.Cells[r, 7].TextHeight = textHeight;
                        table.Cells[r, 7].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 8].TextString = sArray[5];
                        table.Cells[r, 8].TextHeight = textHeight;
                        table.Cells[r, 8].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 9].TextString = sArray[6];
                        table.Cells[r, 9].TextHeight = textHeight;
                        table.Cells[r, 9].Alignment = CellAlignment.MiddleCenter;
                        table.Cells[r, 10].TextString = sArray[7];
                        table.Cells[r, 10].TextHeight = textHeight;
                        table.Cells[r, 10].Alignment = CellAlignment.MiddleCenter;
                        r += 1;
                    }
                }
            }

            //Color color = Color.FromColorIndex(ColorMethod.ByAci, 3); // 声明颜色
            //table.Cells[0, 0].BackgroundColor = color; // 设置背景颜色
            //color = Color.FromColorIndex(ColorMethod.ByAci, 1);
            //table.Cells[0, 0].ContentColor = color; //内容颜色
            using (TransactionHost host = new TransactionHost(db))
            {
                host.AddEntity(table);
                host.Commit();
            }
        }

        //寻找实体最近的text(穿层)
        public DBText NearText(Database db, Autodesk.AutoCAD.DatabaseServices.Entity entity, List<DBText> texts)
        {
            DBText dBText = new DBText();
            double dis = 10000;
            if (entity.GetType() == typeof(Circle))
            {
                Circle circle = (Circle)entity;
                foreach (DBText text in texts)
                {
                    if (text.Position.GetDistanceBetweenToPoint(circle.Center) < dis)
                    {
                        dis = text.Position.GetDistanceBetweenToPoint(circle.Center);
                        dBText = text;
                    }
                }
            }

            return dBText;
        }

        [CommandMethod("PRINTSQL")]
        //读取Excel表输出数据(顺层)
        public void readExcelPrintSql()
        {
            DocumentLock docLock = Application.DocumentManager.MdiActiveDocument.LockDocument();
            // 对话框窗口
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            // 数据库对象
            Database db = HostApplicationServices.WorkingDatabase;
            try
            {
                using (TransactionHost host = new TransactionHost(db))
                {
                    foreach (ObjectId id in host.ModelSpace)
                    {
                        Autodesk.AutoCAD.DatabaseServices.Entity entity = (Autodesk.AutoCAD.DatabaseServices.Entity)host.GetObject(id, OpenMode.ForRead);
                        if (entity != null)
                        {
                            if (entity.GetType() == typeof(Table))
                            {
                                Table table = (Table)entity;
                                int row = table.Rows.Count();
                                List<String> tunnel_name = new List<String>();

                                for (int i = 1; i < row; i++)
                                {
                                    if (!tunnel_name.Contains(table.Cells[i, 1].GetTextString(FormatOption.IgnoreMtextFormat)))
                                        tunnel_name.Add(table.Cells[i, 1].GetTextString(FormatOption.IgnoreMtextFormat));
                                }
                                for (int i = 0; i < tunnel_name.Count; i++)
                                {
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("# tunnel " + tunnel_name[i] + "\n");
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("# site " + "1" + " " + tunnel_name[i] + "钻场\n");
                                    int n = 1;
                                    for (int j = 1; j < row; j++)
                                    {
                                        if (table.Cells[j, 1].GetTextString(FormatOption.IgnoreMtextFormat) == tunnel_name[i])
                                        {
                                            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("# drilling " + n.ToString() + " " + table.Cells[j, 3].GetTextString(FormatOption.IgnoreMtextFormat) + " " + table.Cells[j, 4].GetTextString(FormatOption.IgnoreMtextFormat)
                                                + " " + table.Cells[j, 5].GetTextString(FormatOption.IgnoreMtextFormat) + " " + table.Cells[j, 6].GetTextString(FormatOption.IgnoreMtextFormat) + " " + table.Cells[j, 7].GetTextString(FormatOption.IgnoreMtextFormat) + " " + table.Cells[j, 8].GetTextString(FormatOption.IgnoreMtextFormat)
                                                + " " + table.Cells[j, 9].GetTextString(FormatOption.IgnoreMtextFormat) + " " + table.Cells[j, 10].GetTextString(FormatOption.IgnoreMtextFormat) + "\n");
                                            n += 1;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception e)
            {
                Application.ShowAlertDialog(e.Message);
            }


            docLock.Dispose();
        }


    }
}
