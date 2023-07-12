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
        // 需要載入的廁所群組族群
        public class ToiletGroup
        {
            public static string toilet = "MRT_廁所群組(一般坐式)(SinoBIM-第2版)";
            public static string sink = "MRT_洗手台群組(SinoBIM-第3版)";
            public static string urinal = "MRT_小便斗群組(SinoBIM-第1版)";
            public static string familyRestroom = "MRT_親子廁所(SinoBIM-第1版)";
            public static string accessibleRestroom = "MRT_無障礙廁所(SinoBIM-第1版)";
        }
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

            TransactionGroup transGroup = new TransactionGroup(doc, "建立廁所模型");
            transGroup.Start();
            Tuple<List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>> restroomElems = CreateRestroom(doc, jsonData, levelElevList); // 建立廁所模型
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
        // 取得族群內廁所所需的Family
        private List<FamilySymbol> GetFamilySymbols(Document doc)
        {
            List<FamilySymbol> familySymbolList = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            familySymbolList = familySymbolList.Where(x => x.Family.Name.Contains(ToiletGroup.toilet) ||
                                                           x.Family.Name.Contains(ToiletGroup.sink) ||
                                                           x.Family.Name.Contains(ToiletGroup.urinal) ||
                                                           x.Family.Name.Contains(ToiletGroup.familyRestroom) ||
                                                           x.Family.Name.Contains(ToiletGroup.accessibleRestroom)).Distinct().ToList();
            return familySymbolList;
        }
        // 建立廁所模型
        private Tuple<List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>> CreateRestroom(Document doc, JsonData jsonData, List<LevelElevation> levelElevList)
        {
            List<List<FamilyInstance>> group = new List<List<FamilyInstance>>();
            List<FamilyInstance> manRestrooms = new List<FamilyInstance>();
            List<FamilyInstance> womanRestrooms = new List<FamilyInstance>();
            List<FamilyInstance> familyRestrooms = new List<FamilyInstance>();
            using (Transaction trans = new Transaction(doc, "放置元件"))
            {
                trans.Start();
                // 取得族群內廁所所需的FamilySymbol, 並啟動
                List<FamilySymbol> familySymbolList = GetFamilySymbols(doc);
                foreach (FamilySymbol familySymbol in familySymbolList)
                {
                    if (!familySymbol.IsActive)
                    {
                        familySymbol.Activate();
                    }
                }
                //manRestrooms = CreateManRestroom(doc, jsonData, levelElevList, familySymbolList, group); // 建立男廁元件
                womanRestrooms = CreateWomanRestroom(doc, jsonData, levelElevList, familySymbolList); // 建立女廁元件
                //familyRestrooms = CreateFamilyRestroom(doc, jsonData, levelElevList, familySymbolList); // 建立親子廁所元件
                trans.Commit();
            }

            return Tuple.Create(manRestrooms, womanRestrooms, familyRestrooms);
        }
        // 建立男廁元件
        private List<FamilyInstance> CreateManRestroom(Document doc, JsonData jsonData, List<LevelElevation> levelElevList, List<FamilySymbol> familySymbolList, List<List<FamilyInstance>> group)
        {
            List<FamilyInstance> manRestrooms = new List<FamilyInstance>();
            double levelElevation = levelElevList[5].Height; // 穿堂層
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(ToiletGroup.toilet)).FirstOrDefault();
            FamilySymbol washbasin = familySymbolList.Where(x => x.FamilyName.Contains(ToiletGroup.sink)).FirstOrDefault();
            FamilySymbol urinal = familySymbolList.Where(x => x.FamilyName.Contains(ToiletGroup.urinal)).FirstOrDefault();

            foreach (ManData manData in jsonData.ManDataList)
            {
                XYZ xyz = new XYZ(manData.RestroomMan_x, manData.RestroomMan_y, levelElevation);

                if (toilet != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains(ToiletGroup.toilet))
                    {
                        Parameter counts = instance.LookupParameter("一般坐式數量");
                        counts.Set(manData.Toilet_Count);
                    }
                    manRestrooms.Add(instance);
                }
                if (washbasin != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, washbasin, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains(ToiletGroup.sink))
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
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, urinal, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains(ToiletGroup.urinal))
                    {
                        Parameter counts = instance.LookupParameter("兒童(無障礙)小便斗數量");
                        if (manData.Urinal_Count > 1)
                        {
                            counts.Set(1);
                            counts = instance.LookupParameter("一般小便斗數量");
                            counts.Set(manData.Urinal_Count - 1);
                        }
                        else
                        {
                            counts.Set(0);
                            counts = instance.LookupParameter("一般小便斗數量");
                            counts.Set(manData.Urinal_Count);
                        }
                    }
                    manRestrooms.Add(instance);
                }
            }
            group.Add(manRestrooms);
            return manRestrooms;
        }
        // 建立女廁元件
        private List<FamilyInstance> CreateWomanRestroom(Document doc, JsonData jsonData, List<LevelElevation> levelElevList, List<FamilySymbol> familySymbolList)
        {
            List<FamilyInstance> womanRestrooms = new List<FamilyInstance>();
            double levelElevation = levelElevList[5].Height; // 穿堂層
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(ToiletGroup.toilet)).FirstOrDefault();
            FamilySymbol washbasin = familySymbolList.Where(x => x.FamilyName.Contains(ToiletGroup.sink)).FirstOrDefault();

            foreach (WomanData womanData in jsonData.WomanDataList)
            {
                XYZ xyz = new XYZ(womanData.RestroomWoman_x, womanData.RestroomWoman_y, levelElevation);
                if(womanData.Toilet_Count > 5)
                {
                    int toiletCount = womanData.Toilet_Count / 2;
                    for (int i = 0; i < 2; i++)
                    {
                        if (toilet != null)
                        {
                            FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains(ToiletGroup.toilet))
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
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, washbasin, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains(ToiletGroup.sink))
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
        // 建立親子廁所元件
        private List<FamilyInstance> CreateFamilyRestroom(Document doc, JsonData jsonData, List<LevelElevation> levelElevList, List<FamilySymbol> familySymbolList)
        {
            List<FamilyInstance> familyRestrooms = new List<FamilyInstance>();
            double levelElevation = levelElevList[5].Height; // 穿堂層
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(ToiletGroup.familyRestroom)).FirstOrDefault();

            foreach (FamilyData familyData in jsonData.FamilyDataList)
            {
                XYZ xyz = new XYZ(familyData.RestroomFamily_x, familyData.RestroomFamily_y, levelElevation);
                if (toilet != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                    familyRestrooms.Add(instance);
                }
            }
            return familyRestrooms;
        }
        // 建立無障礙廁所元件
        private List<FamilyInstance> CreateAccessibleRestroom(Document doc, JsonData jsonData, List<LevelElevation> levelElevList, List<FamilySymbol> familySymbolList)
        {
            List<FamilyInstance> accessibleRestrooms = new List<FamilyInstance>();
            double levelElevation = levelElevList[5].Height; // 穿堂層
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(ToiletGroup.accessibleRestroom)).FirstOrDefault();

            foreach (FamilyData familyData in jsonData.FamilyDataList)
            {
                XYZ xyz = new XYZ(familyData.RestroomFamily_x, familyData.RestroomFamily_y, levelElevation);

                if (toilet != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                    accessibleRestrooms.Add(instance);
                }
            }
            return accessibleRestrooms;
        }
        // 排列元件組合
        private void EditParameter(Document doc, Tuple<List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>> restroomElems, JsonData jsonData)
        {
            using (Transaction trans = new Transaction(doc, "組合"))
            {
                trans.Start();
                //EditManRestroom(doc, restroomElems.Item1, jsonData); // 修改男廁參數
                EditWomanRestroom(doc, restroomElems.Item2, jsonData); // 修改女廁參數
                //EditFamilyRestroom(doc, restroomElems.Item3, jsonData); // 修改親子廁所參數
                trans.Commit();
            }
        }
        // 修改男廁參數
        private void EditManRestroom(Document doc, List<FamilyInstance> manRestroomElems, JsonData jsonData)
        {
            foreach (ManData manData in jsonData.ManDataList)
            {
                FamilyInstance toilet = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains(ToiletGroup.toilet)).FirstOrDefault();
                FamilyInstance washbasin = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains(ToiletGroup.sink)).FirstOrDefault();
                FamilyInstance urinal = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains(ToiletGroup.urinal)).FirstOrDefault();

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

                double manRestroomLength = UnitUtils.ConvertToInternalUnits(manData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double manRestroomWidth = UnitUtils.ConvertToInternalUnits(manData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度

                if (manData.Rotate_id.Equals(1) || manData.Rotate_id.Equals(2))
                {
                    // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                    XYZ offset = new XYZ(0, -toiletTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, toilet.Id, offset);

                    // MRT_小便斗群組(SinoBIM-第1版)
                    LocationPoint lp = urinal.Location as LocationPoint;
                    Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                    double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                    ElementTransformUtils.RotateElement(doc, urinal.Id, line, rotation);
                    offset = new XYZ(urinalTotalWidth, -manRestroomWidth + urinalTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, urinal.Id, offset);

                    // MRT_洗手台群組(SinoBIM-第3版)
                    if (manData.Type.Equals(1))
                    {
                        offset = new XYZ(manRestroomLength - washbasinTotalWidth - partitionThickness, -washbasinTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                    }
                    else if (manData.Type.Equals(2))
                    {
                        lp = washbasin.Location as LocationPoint;
                        line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                        rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                        ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                        offset = new XYZ(manRestroomLength - washbasinTotalDepth - partitionThickness, 0, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                    }
                }
                else if (manData.Rotate_id.Equals(3) || manData.Rotate_id.Equals(4))
                {
                    // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                    XYZ offset = new XYZ(manRestroomLength - toiletTotalWidth - partitionThickness, -toiletTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, toilet.Id, offset);

                    // MRT_小便斗群組(SinoBIM-第1版)
                    LocationPoint lp = urinal.Location as LocationPoint;
                    Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                    double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                    ElementTransformUtils.RotateElement(doc, urinal.Id, line, rotation);
                    offset = new XYZ(manRestroomLength, -manRestroomWidth + urinalTotalDepth, 0);
                    ElementTransformUtils.MoveElement(doc, urinal.Id, offset);

                    // MRT_洗手台群組(SinoBIM-第3版)
                    if (manData.Type.Equals(1))
                    {
                        offset = new XYZ(0, -washbasinTotalDepth, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                    }
                    else if (manData.Type.Equals(2))
                    {
                        lp = washbasin.Location as LocationPoint;
                        line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                        rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                        ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                        offset = new XYZ(washbasinTotalDepth, -washbasinTotalWidth, 0);
                        ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                    }
                }

                // 移除修改過的族群
                manRestroomElems.Remove(toilet);
                manRestroomElems.Remove(washbasin);
                manRestroomElems.Remove(urinal);
            }
        }
        // 修改女廁參數
        private void EditWomanRestroom(Document doc, List<FamilyInstance> womanRestroomElems, JsonData jsonData)
        {
            foreach (WomanData womanData in jsonData.WomanDataList)
            {
                FamilyInstance toilet1 = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(ToiletGroup.toilet)).FirstOrDefault();
                FamilyInstance toilet2 = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(ToiletGroup.toilet) && x.Id != toilet1.Id).FirstOrDefault();
                FamilyInstance washbasin = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(ToiletGroup.sink)).FirstOrDefault();

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

                // 移除修改過的族群
                womanRestroomElems.Remove(toilet1);
                womanRestroomElems.Remove(toilet2);
                womanRestroomElems.Remove(washbasin);
            }
        }
        // 修改親子廁所參數
        private void EditFamilyRestroom(Document doc, List<FamilyInstance> familyRestroomElems, JsonData jsonData)
        {
            foreach (FamilyData familyData in jsonData.FamilyDataList)
            {
                FamilyInstance toilet = familyRestroomElems.Where(x => x.Symbol.FamilyName.Contains(ToiletGroup.familyRestroom)).FirstOrDefault();
                // toilet：MRT_親子廁所(SinoBIM-第1版)
                double toiletDepth = toilet.LookupParameter("親子廁所深度").AsDouble();

                double familyRestroomLength = UnitUtils.ConvertToInternalUnits(familyData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double familyRestroomWidth = UnitUtils.ConvertToInternalUnits(familyData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度

                // toilet：MRT_親子廁所(SinoBIM-第1版)
                XYZ offset = new XYZ(0, -toiletDepth, 0);
                ElementTransformUtils.MoveElement(doc, toilet.Id, offset);

                // 移除修改過的族群
                familyRestroomElems.Remove(toilet);
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
