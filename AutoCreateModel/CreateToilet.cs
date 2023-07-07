using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using static AutoCreateModel.FindLevel;
using Document = Autodesk.Revit.DB.Document;

namespace AutoCreateModel
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreateToilet : IExternalEventHandler
    {
        public double prjNS = 0.0; // 專案基準點：N/S
        public double prjWE = 0.0; // 專案基準點：W/E
        public double angle = 0.0; // 旋轉角度
        public void Execute(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string folderPath = ReadJsonForm.folderPath; // Json檔案資料夾路徑
            JsonData jsonData = ReadJson.ReadJsonFile(folderPath);

            FindLevel findLevel = new FindLevel();
            Tuple<List<LevelElevation>, LevelElevation, double> multiValue = findLevel.FindDocViewLevel(doc);
            List<LevelElevation> levelElevList = multiValue.Item1; // 全部樓層

            ProjectBasePoint(doc); // 專案基準點
            List<FamilyInstance> manDataModels = CreateToiletFamilyInstance(doc); // 抓取建置廁所模型所需的族群
            CreateMenRestroom(doc, jsonData, levelElevList, manDataModels); // 建立男廁模型
        }
        // 查詢專案基準點
        private void ProjectBasePoint(Document doc)
        {

            // 專案基準點, 暫定距離原點最大偏移的為BasePoint
            List<BasePoint> allPrjLocations = new FilteredElementCollector(doc).OfClass(typeof(BasePoint)).WhereElementIsNotElementType().Cast<BasePoint>().ToList();
            List<BasePoint> prjLocations = allPrjLocations.Where(x => x.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM) != null).ToList();
            BasePoint prjLocation = prjLocations.Where(x => x.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble() ==
                                    prjLocations.Max(y => y.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble())).FirstOrDefault();
            prjNS = Convert.ToDouble(prjLocation.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsValueString()); // 南北
            prjWE = Convert.ToDouble(prjLocation.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsValueString()); // 東西
            try
            {
                string angleton = prjLocation.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsValueString();
                if (angleton != null)
                {
                    angle = -Convert.ToDouble(angleton.Remove(angleton.Length - 1)); // 至正北的角度
                }
            }
            catch (Exception)
            {

            }
        }
        // 抓取建置廁所模型所需的族群
        private List<FamilyInstance> CreateToiletFamilyInstance(Document doc)
        {
            IList<ElementFilter> filters = new List<ElementFilter>();
            ElementCategoryFilter genericModelFilter = new ElementCategoryFilter(BuiltInCategory.OST_GenericModel); // 一般模型
            ElementCategoryFilter plumbingFixturesFilter = new ElementCategoryFilter(BuiltInCategory.OST_PlumbingFixtures); // 衛工裝置
            filters.Add(genericModelFilter);
            filters.Add(plumbingFixturesFilter);
            LogicalOrFilter chooseFilters = new LogicalOrFilter(filters);
            List<Element> elems = new List<Element>();
            List<FamilyInstance> toiletModels = new FilteredElementCollector(doc).WherePasses(chooseFilters).WhereElementIsNotElementType().Cast<FamilyInstance>().ToList();
            // 男廁模型
            List<FamilyInstance> manDataModels = toiletModels.Where(x => x.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)") ||
                                                                         x.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)") ||
                                                                         x.Symbol.FamilyName.Contains("MRT_小便斗群組(SinoBIM-第1版)")).Distinct(new RemoveDuplicatesComparer()).ToList();
            return manDataModels;
        }
        // 建立男廁模型
        private void CreateMenRestroom(Document doc, JsonData jsonData, List<LevelElevation> levelElevList, List<FamilyInstance> manDataModels)
        {
            TransactionGroup transGroup = new TransactionGroup(doc, "建立男廁模型");
            transGroup.Start();

            List<FamilyInstance> menRestroom = PutMenRestroomElems(doc, jsonData, levelElevList, manDataModels); // 放置男廁元件
            EditMenRestroomElems(doc, menRestroom, jsonData); // 排列元件組合

            transGroup.Assimilate();
        }
        // 放置男廁元件
        private List<FamilyInstance> PutMenRestroomElems(Document doc, JsonData jsonData, List<LevelElevation> levelElevList, List<FamilyInstance> manDataModels)
        {
            List<FamilyInstance> menRestroom = new List<FamilyInstance>();
            using (Transaction trans = new Transaction(doc, "放置元件"))
            {
                trans.Start();
                foreach (ManData manData in jsonData.ManDataList)
                {
                    double levelElevation = levelElevList[5].Height;
                    foreach (FamilyInstance manDataModel in manDataModels)
                    {
                        FamilySymbol symbol = manDataModel.Symbol;
                        XYZ xyz = new XYZ(manData.RestroomMan_x, manData.RestroomMan_y, levelElevation);
                        FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, symbol, StructuralType.NonStructural);
                        if (instance.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)"))
                        {
                            Parameter counts = instance.LookupParameter("一般坐式數量");
                            counts.Set(manData.Toilet_Count);
                        }
                        else if (instance.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)"))
                        {
                            Parameter counts = instance.LookupParameter("兒童洗面盆數量");
                            if (manData.Washbasin_Count > 1)
                            {
                                counts.Set(1);
                                counts = instance.LookupParameter("一般洗面盆數量");
                                counts.Set(manData.Washbasin_Count - 1);
                            }
                            else
                            {
                                counts.Set(0);
                                counts = instance.LookupParameter("一般洗面盆數量");
                                counts.Set(manData.Washbasin_Count);
                            }
                        }
                        else if (instance.Symbol.FamilyName.Contains("MRT_小便斗群組(SinoBIM-第1版)"))
                        {
                            Parameter counts = instance.LookupParameter("兒童(無障礙)小便斗數量");
                            if (manData.Washbasin_Count > 1)
                            {
                                counts.Set(1);
                                counts = instance.LookupParameter("一般小便斗數量");
                                counts.Set(manData.Washbasin_Count - 1);
                            }
                            else
                            {
                                counts.Set(0);
                                counts = instance.LookupParameter("一般小便斗數量");
                                counts.Set(manData.Washbasin_Count);
                            }
                        }
                        menRestroom.Add(instance);
                    }
                }
                trans.Commit();
            }

            return menRestroom;
        }
        // 排列元件組合
        private void EditMenRestroomElems(Document doc, List<FamilyInstance> menRestroom, JsonData jsonData)
        {
            using (Transaction trans = new Transaction(doc, "排列組合"))
            {
                trans.Start();
                FamilyInstance toilet = menRestroom.Where(x => x.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)")).FirstOrDefault();
                FamilyInstance washbasin = menRestroom.Where(x => x.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)")).FirstOrDefault();
                FamilyInstance urinal = menRestroom.Where(x => x.Symbol.FamilyName.Contains("MRT_小便斗群組(SinoBIM-第1版)")).FirstOrDefault();

                // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                double toiletTotalWidth = toilet.LookupParameter("一般坐式總寬度").AsDouble();
                double toiletTotalDepth = toilet.LookupParameter("總深度").AsDouble();
                double partitionThickness = toilet.LookupParameter("隔板厚度").AsDouble() / 2;

                // MRT_洗手台群組(SinoBIM-第3版)
                double washbasinTotalWidth = washbasin.LookupParameter("檯面標準寬度").AsDouble();
                double washbasinTotalDepth = washbasin.LookupParameter("總深度").AsDouble();

                // MRT_小便斗群組(SinoBIM-第1版)
                double urinalTotalWidth = urinal.LookupParameter("總寬度").AsDouble();
                double urinalTotalDepth = urinal.LookupParameter("總深度").AsDouble();

                foreach (ManData manData in jsonData.ManDataList)
                {
                    double menRestroomLength = UnitUtils.ConvertToInternalUnits(manData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                    double menRestroomWidth = UnitUtils.ConvertToInternalUnits(manData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度

                    if (manData.Rotate_id.Equals(1) || manData.Rotate_id.Equals(2))
                    {
                        // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        XYZ offset = new XYZ(0, -toiletTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, toilet.Id, offset);

                        // MRT_洗手台群組(SinoBIM-第3版)
                        offset = new XYZ(menRestroomLength - washbasinTotalWidth - partitionThickness, -washbasinTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);

                        // MRT_小便斗群組(SinoBIM-第1版)
                        LocationPoint lp = urinal.Location as LocationPoint;
                        Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                        double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                        ElementTransformUtils.RotateElement(doc, urinal.Id, line, rotation);
                        offset = new XYZ(urinalTotalWidth, -menRestroomWidth + urinalTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, urinal.Id, offset);
                    }
                    else if(manData.Rotate_id.Equals(3) || manData.Rotate_id.Equals(4))
                    {
                        // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        XYZ offset = new XYZ(menRestroomLength - toiletTotalWidth - partitionThickness, -toiletTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, toilet.Id, offset);

                        // MRT_洗手台群組(SinoBIM-第3版)
                        offset = new XYZ(0, -washbasinTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);

                        // MRT_小便斗群組(SinoBIM-第1版)
                        LocationPoint lp = urinal.Location as LocationPoint;
                        Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                        double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                        ElementTransformUtils.RotateElement(doc, urinal.Id, line, rotation);
                        offset = new XYZ(menRestroomLength, -menRestroomWidth + urinalTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, urinal.Id, offset);
                    }
                }

                trans.Commit();
            }
        }
        // 比對Symbol.FamilyName名稱移除重複
        private class RemoveDuplicatesComparer : IEqualityComparer<FamilyInstance>
        {
            public bool Equals(FamilyInstance x, FamilyInstance y)
            {
                return x.Symbol.FamilyName == y.Symbol.FamilyName;
            }
            public int GetHashCode(FamilyInstance obj)
            {
                return obj.Symbol.FamilyName.GetHashCode();
            }
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
