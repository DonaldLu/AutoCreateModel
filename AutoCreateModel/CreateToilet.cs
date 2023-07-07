using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
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
            uiapp.Application.FailuresProcessing += FaliureProcessor; // 關閉警示視窗

            string folderPath = ReadJsonForm.folderPath; // Json檔案資料夾路徑
            JsonData jsonData = ReadJson.ReadJsonFile(folderPath);

            FindLevel findLevel = new FindLevel();
            Tuple<List<LevelElevation>, LevelElevation, double> multiValue = findLevel.FindDocViewLevel(doc);
            List<LevelElevation> levelElevList = multiValue.Item1; // 全部樓層

            ProjectBasePoint(doc); // 專案基準點
            List<FamilyInstance> dataElems = DataElems(doc); // 抓取建置廁所模型所需的族群

            TransactionGroup transGroup = new TransactionGroup(doc, "建立廁所模型");
            transGroup.Start();
            Tuple<List<FamilyInstance>, List<FamilyInstance>> restroomElems = CreateRestroom(doc, jsonData, levelElevList, dataElems); // 建立廁所模型
            EditParameter(doc, restroomElems, jsonData); // 排列元件組合
            transGroup.Assimilate();
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
        private List<FamilyInstance> DataElems(Document doc)
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
            List<FamilyInstance> dataElems = toiletModels.Where(x => x.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)") ||
                                                                         x.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)") ||
                                                                         x.Symbol.FamilyName.Contains("MRT_小便斗群組(SinoBIM-第1版)")).Distinct(new RemoveDuplicatesComparer()).ToList();
            return dataElems;
        }
        // 建立廁所模型
        private Tuple<List<FamilyInstance>, List<FamilyInstance>> CreateRestroom(Document doc, JsonData jsonData, List<LevelElevation> levelElevList, List<FamilyInstance> dataElems)
        {
            List<FamilyInstance> manRestrooms = new List<FamilyInstance>();
            List<FamilyInstance> womanRestrooms = new List<FamilyInstance>();
            using (Transaction trans = new Transaction(doc, "放置元件"))
            {
                trans.Start();
                manRestrooms = CreateManRestroom(doc, jsonData, levelElevList, dataElems); // 建立男廁元件
                womanRestrooms = CreateWomanRestroom(doc, jsonData, levelElevList, dataElems); // 建立女廁元件
                trans.Commit();
            }

            return Tuple.Create(manRestrooms, womanRestrooms);
        }
        // 建立男廁元件
        private List<FamilyInstance> CreateManRestroom(Document doc, JsonData jsonData, List<LevelElevation> levelElevList, List<FamilyInstance> dataElems)
        {
            List<FamilyInstance> manRestrooms = new List<FamilyInstance>();

            foreach (ManData manData in jsonData.ManDataList)
            {
                double levelElevation = levelElevList[5].Height; // 穿堂層
                XYZ xyz = new XYZ(manData.RestroomMan_x, manData.RestroomMan_y, levelElevation);

                FamilyInstance toilet = dataElems.Where(x => x.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)")).FirstOrDefault();
                FamilyInstance washbasin = dataElems.Where(x => x.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)")).FirstOrDefault();
                FamilyInstance urinal = dataElems.Where(x => x.Symbol.FamilyName.Contains("MRT_小便斗群組(SinoBIM-第1版)")).FirstOrDefault();

                if (toilet != null)
                {
                    FamilySymbol symbol = toilet.Symbol;
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, symbol, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)"))
                    {
                        Parameter counts = instance.LookupParameter("一般坐式數量");
                        counts.Set(manData.Toilet_Count);
                    }
                    manRestrooms.Add(instance);
                }
                if (washbasin != null)
                {
                    FamilySymbol symbol = washbasin.Symbol;
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, symbol, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)"))
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
                    manRestrooms.Add(instance);
                }
                if (urinal != null)
                {
                    FamilySymbol symbol = urinal.Symbol;
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, symbol, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains("MRT_小便斗群組(SinoBIM-第1版)"))
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
                    manRestrooms.Add(instance);
                }
            }
            return manRestrooms;
        }
        // 建立女廁元件
        private List<FamilyInstance> CreateWomanRestroom(Document doc, JsonData jsonData, List<LevelElevation> levelElevList, List<FamilyInstance> dataElems)
        {
            List<FamilyInstance> womanRestrooms = new List<FamilyInstance>();

            foreach (WomanData womanData in jsonData.WomanDataList)
            {
                double levelElevation = levelElevList[5].Height; // 穿堂層
                XYZ xyz = new XYZ(womanData.RestroomWoman_x, womanData.RestroomWoman_y, levelElevation);

                FamilyInstance toilet = dataElems.Where(x => x.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)")).FirstOrDefault();
                FamilyInstance washbasin = dataElems.Where(x => x.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)")).FirstOrDefault();

                if(womanData.Toilet_Count > 5)
                {
                    int toiletCount = womanData.Toilet_Count / 2;
                    for (int i = 0; i < 2; i++)
                    {
                        if (toilet != null)
                        {
                            FamilySymbol symbol = toilet.Symbol;
                            FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, symbol, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)"))
                            {
                                Parameter counts = instance.LookupParameter("一般坐式數量");
                                if (i.Equals(0)) { counts.Set(toiletCount); }
                                else { counts.Set(womanData.Toilet_Count - toiletCount); }
                            }
                            womanRestrooms.Add(instance);
                        }
                    }
                }
                if (washbasin != null)
                {
                    FamilySymbol symbol = washbasin.Symbol;
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, symbol, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)"))
                    {
                        Parameter counts = instance.LookupParameter("兒童洗面盆數量");
                        if (womanData.Washbasin_Count > 1)
                        {
                            counts.Set(1);
                            counts = instance.LookupParameter("一般洗面盆數量");
                            counts.Set(womanData.Washbasin_Count - 1);
                        }
                        else
                        {
                            counts.Set(0);
                            counts = instance.LookupParameter("一般洗面盆數量");
                            counts.Set(womanData.Washbasin_Count);
                        }
                    }
                    womanRestrooms.Add(instance);
                }
            }
            return womanRestrooms;
        }
        // 排列元件組合
        private void EditParameter(Document doc, Tuple<List<FamilyInstance>, List<FamilyInstance>> restroomElems, JsonData jsonData)
        {
            using (Transaction trans = new Transaction(doc, "組合"))
            {
                trans.Start();
                EditManRestroom(doc, restroomElems.Item1, jsonData); // 修改男廁參數
                EditWomanRestroom(doc, restroomElems.Item2, jsonData); // 修改女廁參數
                trans.Commit();
            }
        }
        // 修改男廁參數
        private void EditManRestroom(Document doc, List<FamilyInstance> manRestroomElems, JsonData jsonData)
        {
            FamilyInstance toilet = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)")).FirstOrDefault();
            FamilyInstance washbasin = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)")).FirstOrDefault();
            FamilyInstance urinal = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains("MRT_小便斗群組(SinoBIM-第1版)")).FirstOrDefault();

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
                double manRestroomLength = UnitUtils.ConvertToInternalUnits(manData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double manRestroomWidth = UnitUtils.ConvertToInternalUnits(manData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度

                if (manData.Rotate_id.Equals(1) || manData.Rotate_id.Equals(2))
                {
                    // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                    XYZ offset = new XYZ(0, -toiletTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, toilet.Id, offset);

                    // MRT_洗手台群組(SinoBIM-第3版)
                    offset = new XYZ(manRestroomLength - washbasinTotalWidth - partitionThickness, -washbasinTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);

                    // MRT_小便斗群組(SinoBIM-第1版)
                    LocationPoint lp = urinal.Location as LocationPoint;
                    Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                    double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                    ElementTransformUtils.RotateElement(doc, urinal.Id, line, rotation);
                    offset = new XYZ(urinalTotalWidth, -manRestroomWidth + urinalTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, urinal.Id, offset);
                }
                else if (manData.Rotate_id.Equals(3) || manData.Rotate_id.Equals(4))
                {
                    // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                    XYZ offset = new XYZ(manRestroomLength - toiletTotalWidth - partitionThickness, -toiletTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, toilet.Id, offset);

                    // MRT_洗手台群組(SinoBIM-第3版)
                    offset = new XYZ(0, -washbasinTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);

                    // MRT_小便斗群組(SinoBIM-第1版)
                    LocationPoint lp = urinal.Location as LocationPoint;
                    Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                    double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                    ElementTransformUtils.RotateElement(doc, urinal.Id, line, rotation);
                    offset = new XYZ(manRestroomLength, -manRestroomWidth + urinalTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, urinal.Id, offset);
                }
            }
        }
        // 修改女廁參數
        private void EditWomanRestroom(Document doc, List<FamilyInstance> womanRestroomElems, JsonData jsonData)
        {
            FamilyInstance toilet1 = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)")).FirstOrDefault();
            FamilyInstance toilet2 = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains("MRT_廁所群組(一般坐式)(SinoBIM-第2版)")).LastOrDefault();
            FamilyInstance washbasin = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains("MRT_洗手台群組(SinoBIM-第3版)")).FirstOrDefault();

            // toilet1：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
            double toiletTotalWidth1 = toilet1.LookupParameter("一般坐式總寬度").AsDouble();
            double toiletTotalDepth1 = toilet1.LookupParameter("總深度").AsDouble();
            double partitionThickness1 = toilet1.LookupParameter("隔板厚度").AsDouble() / 2;

            // toilet2：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
            double toiletTotalWidth2 = toilet2.LookupParameter("一般坐式總寬度").AsDouble();
            double toiletTotalDepth2 = toilet2.LookupParameter("總深度").AsDouble();
            double partitionThickness2 = toilet2.LookupParameter("隔板厚度").AsDouble() / 2;

            // MRT_洗手台群組(SinoBIM-第3版)
            double washbasinTotalWidth = washbasin.LookupParameter("檯面標準寬度").AsDouble();
            double washbasinTotalDepth = washbasin.LookupParameter("總深度").AsDouble();

            foreach (WomanData womanData in jsonData.WomanDataList)
            {
                double womanRestroomLength = UnitUtils.ConvertToInternalUnits(womanData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double womanRestroomWidth = UnitUtils.ConvertToInternalUnits(womanData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度

                if (womanData.Rotate_id.Equals(1) || womanData.Rotate_id.Equals(2))
                {
                    // toilet1：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                    XYZ offset = new XYZ(0, -toiletTotalDepth1, 0);
                    ElementTransformUtils.MoveElement(doc, toilet1.Id, offset);

                    // toilet2：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                    LocationPoint lp = toilet2.Location as LocationPoint;
                    Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                    double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                    ElementTransformUtils.RotateElement(doc, toilet2.Id, line, rotation);
                    offset = new XYZ(toiletTotalWidth2, -womanRestroomWidth + toiletTotalDepth2, 0);
                    ElementTransformUtils.MoveElement(doc, toilet2.Id, offset);

                    // MRT_洗手台群組(SinoBIM-第3版)
                    if (womanData.Type.Equals(1))
                    {
                        offset = new XYZ(womanRestroomLength - washbasinTotalWidth - partitionThickness1, -washbasinTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                    }
                    else if (womanData.Type.Equals(2))
                    {
                        lp = washbasin.Location as LocationPoint;
                        line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                        rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                        ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                        offset = new XYZ(womanRestroomLength - washbasinTotalDepth - partitionThickness1, 0, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                    }
                }
                else if (womanData.Rotate_id.Equals(3) || womanData.Rotate_id.Equals(4))
                {
                    // toilet1：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                    XYZ offset = new XYZ(womanRestroomLength - toiletTotalWidth1 - partitionThickness1, -toiletTotalDepth1, 0);
                    ElementTransformUtils.MoveElement(doc, toilet1.Id, offset);

                    // toilet2：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                    LocationPoint lp = toilet2.Location as LocationPoint;
                    Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                    double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                    ElementTransformUtils.RotateElement(doc, toilet2.Id, line, rotation);
                    offset = new XYZ(womanRestroomLength, -womanRestroomWidth + toiletTotalDepth2, 0);
                    ElementTransformUtils.MoveElement(doc, toilet2.Id, offset);

                    // MRT_洗手台群組(SinoBIM-第3版)
                    if (womanData.Type.Equals(1))
                    {
                        offset = new XYZ(0, -washbasinTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                    }
                    else if (womanData.Type.Equals(2))
                    {
                        lp = washbasin.Location as LocationPoint;
                        line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                        rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                        ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                        offset = new XYZ(washbasinTotalDepth, -washbasinTotalWidth, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                    }
                }
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
        // 關閉警示視窗
        private void FaliureProcessor(object sender, FailuresProcessingEventArgs e)
        {
            bool hasFailure = false;
            FailuresAccessor fas = e.GetFailuresAccessor();
            List<FailureMessageAccessor> fma = fas.GetFailureMessages().ToList();
            List<ElementId> ElemntsToDelete = new List<ElementId>();
            fas.DeleteAllWarnings();

            foreach (FailureMessageAccessor fa in fma)
            {
                try
                {
                    // 使用以下刪除警告元素
                    List<ElementId> FailingElementIds = fa.GetFailingElementIds().ToList();
                    ElementId FailingElementId = FailingElementIds[0];
                    if (!ElemntsToDelete.Contains(FailingElementId))
                    {
                        ElemntsToDelete.Add(FailingElementId);
                    }
                    hasFailure = true;
                    fas.DeleteWarning(fa);
                }
                catch (Exception)
                {

                }
            }
            if (ElemntsToDelete.Count > 0)
            {
                fas.DeleteElements(ElemntsToDelete);
            }
            // 在外部命令結束後，使用以下行禁用消息抑制器：CachedUiApp.Application.FailuresProcessing -= FaliureProcessor;
            if (hasFailure)
            {
                e.SetProcessingResult(FailureProcessingResult.ProceedWithCommit);
            }
            e.SetProcessingResult(FailureProcessingResult.Continue);
        }
        public string GetName()
        {
            return "Event handler is working now!!";
        }
    }
}
