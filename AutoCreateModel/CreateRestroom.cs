using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Linq;
using static AutoCreateModel.FindLevel;
using Document = Autodesk.Revit.DB.Document;

namespace AutoCreateModel
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class CreateRestroom : IExternalEventHandler
    {
        // 需要載入的廁所群組族群
        public class RestroomGroup
        {
            public static string toilet = "MRT_廁所群組(一般坐式)(SinoBIM-第";
            public static string toilet_squat = "MRT_廁所群組(蹲式)(SinoBIM-第";
            public static string toilet_advanced_age = "MRT_高齡使用者坐式馬桶間(SinoBIM-第";
            public static string mop_basin = "MRT_拖布盆間(SinoBIM-第";
            public static string washbasin = "MRT_洗手台群組(SinoBIM-第";
            public static string washbasin_accessible = "MRT_無障礙洗面盆(SinoBIM-第";
            public static string urinal = "MRT_小便斗群組(SinoBIM-第";
            public static string familyRestroom = "MRT_親子廁所(SinoBIM-第";
            public static string accessibleRestroom = "MRT_無障礙廁所(SinoBIM-第";
            public static string breastfeeding = "MRT_哺集乳室(SinoBIM-第";
            public static string janitorRoom = "MRT_清潔人員休息室(SinoBIM-第";
        }
        public List<string> noFamilySymbols = new List<string>();
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
            Tuple<List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>> restroomElems = CreateRestrooms(doc, jsonData, levelElevList); // 建立廁所模型
            EditParameter(doc, restroomElems, jsonData); // 排列元件組合
            transGroup.Assimilate();

            // 有族群未載入無法建立元件, 則提醒
            noFamilySymbols = noFamilySymbols.Distinct().ToList();
            string error = "";
            if (noFamilySymbols.Count > 0)
            {
                int i = 1;
                foreach(string noFamilySymbol in noFamilySymbols)
                {
                    error += "\n" + i + ". " + noFamilySymbol;
                    i++;
                }
            }
            if(error != "")
            {
                TaskDialog.Show("Error", "以下元件未載入無法建立：\n" + error);
            }
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
        // 取得族群內廁所所需的Family, 並優先使用已載入的最新版
        private List<FamilySymbol> GetFamilySymbols(Document doc)
        {
            List<FamilySymbol> familySymbols = new List<FamilySymbol>();
            List<string> toiletGroups = new List<string>() { RestroomGroup.toilet, RestroomGroup.toilet_squat, RestroomGroup.toilet_advanced_age, RestroomGroup.mop_basin, RestroomGroup.washbasin, RestroomGroup.washbasin_accessible, RestroomGroup.urinal, RestroomGroup.familyRestroom, RestroomGroup.accessibleRestroom, RestroomGroup.breastfeeding, RestroomGroup.janitorRoom };
            List<FamilySymbol> familySymbolList = new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>().ToList();
            // 篩選出廁所群組需要使用的族群
            List<FamilySymbol> saveFamilySymbols = new List<FamilySymbol>();
            foreach (string toiletGroup in toiletGroups)
            {
                foreach(FamilySymbol saveFamilySymbol in familySymbolList.Where(x => x.Family.Name.Contains(toiletGroup)).ToList())
                {
                    saveFamilySymbols.Add(saveFamilySymbol);
                }
            }
            saveFamilySymbols = saveFamilySymbols.OrderBy(x => x.Family.Name).ToList(); // 排序
            // 解析版次
            foreach (string toiletGroup in toiletGroups)
            {
                List<string> familyNames = saveFamilySymbols.Where(x => x.Family.Name.Contains(toiletGroup)).Select(x => x.Family.Name).ToList();
                // 取得載入的所有版本
                List<int> versions = new List<int>();
                foreach(string familyName in familyNames)
                {
                    string removeFamilyName = familyName.Remove(0, toiletGroup.Length);
                    int versionIndex = removeFamilyName.IndexOf('版');
                    try
                    {
                        int version = Convert.ToInt32(removeFamilyName.Substring(0, versionIndex));
                        versions.Add(version);
                    }
                    catch(Exception)
                    {

                    }
                }
                // 找到最新的版本並加入到familySymbols
                int newest = versions.OrderByDescending(x => x).FirstOrDefault();
                if(newest != 0)
                {
                    FamilySymbol addSymbolFamily = saveFamilySymbols.Where(x => x.Family.Name.Contains(toiletGroup + newest.ToString())).FirstOrDefault();
                    if (addSymbolFamily != null)
                    {
                        familySymbols.Add(addSymbolFamily);
                    }
                }
            }
            return familySymbols;
        }
        // 建立廁所模型
        private Tuple<List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>> CreateRestrooms(Document doc, JsonData jsonData, List<LevelElevation> levelElevList)
        {
            List<FamilyInstance> manRestrooms = new List<FamilyInstance>();
            List<FamilyInstance> womanRestrooms = new List<FamilyInstance>();
            List<FamilyInstance> familyRestrooms = new List<FamilyInstance>();
            List<FamilyInstance> accessibleRestrooms = new List<FamilyInstance>();
            List<FamilyInstance> breastfeedings = new List<FamilyInstance>();
            List<FamilyInstance> janitorRooms = new List<FamilyInstance>();
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

                double levelElevation = levelElevList[0].Height; // 建立廁所的樓層
                //manRestrooms = CreateManRestroom(doc, jsonData, levelElevation, familySymbolList); // 建立男廁元件
                womanRestrooms = CreateWomanRestroom(doc, jsonData, levelElevation, familySymbolList); // 建立女廁元件
                //familyRestrooms = CreateFamilyRestroom(doc, jsonData, levelElevation, familySymbolList); // 建立親子廁所元件
                //accessibleRestrooms = CreateAccessibleRestroom(doc, jsonData, levelElevation, familySymbolList); // 建立無障礙廁所元件
                //breastfeedings = CreateBreastfeedingRoom(doc, jsonData, levelElevation, familySymbolList); // 建立哺集乳室
                //janitorRooms = CreateJanitorRoom(doc, jsonData, levelElevation, familySymbolList); // 建立清潔人員休息室
                trans.Commit();
            }

            return Tuple.Create(manRestrooms, womanRestrooms, familyRestrooms, accessibleRestrooms, breastfeedings, janitorRooms);
        }
        // 建立男廁元件
        private List<FamilyInstance> CreateManRestroom(Document doc, JsonData jsonData, double levelElevation, List<FamilySymbol> familySymbolList)
        {
            List<FamilyInstance> manRestrooms = new List<FamilyInstance>();
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.toilet)).FirstOrDefault(); // 坐式馬桶
            FamilySymbol toilet_squat = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.toilet_squat)).FirstOrDefault(); // 蹲式馬桶
            FamilySymbol mop_basin = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.mop_basin)).FirstOrDefault(); // 拖布盆
            FamilySymbol washbasin = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.washbasin)).FirstOrDefault(); // 洗手台
            FamilySymbol washbasin_accessible = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.washbasin_accessible)).FirstOrDefault(); // 無障礙洗手台
            FamilySymbol urinal = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.urinal)).FirstOrDefault(); // 小便斗

            foreach (ManData manData in jsonData.ManDataList)
            {
                double x = UnitUtils.ConvertToInternalUnits(manData.RestroomMan_x, DisplayUnitType.DUT_METERS);
                double y = UnitUtils.ConvertToInternalUnits(manData.RestroomMan_y, DisplayUnitType.DUT_METERS);

                // Level_id要建置的樓層
                string value = string.Empty;
                FindLevel.level_id.TryGetValue(manData.Level_id, out value);
                XYZ xyz = new XYZ(x, y, levelElevation);

                if (toilet != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                    {
                        Parameter counts = instance.LookupParameter("一般坐式數量");
                        if(manData.Toilet_Count >= 3) { counts.Set(manData.Toilet_Count - 2); } // 預留蹲式馬桶+拖布盆間
                        else if (manData.Toilet_Count == 2) { counts.Set(1); } // 預留蹲式馬桶
                        else { counts.Set(manData.Toilet_Count); }                        
                    }
                    manRestrooms.Add(instance);
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.toilet.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
                if (toilet_squat != null)
                {
                    if (manData.Toilet_Count >= 2)
                    {
                        FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet_squat, StructuralType.NonStructural);
                        if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet_squat))
                        {
                            Parameter counts = instance.LookupParameter("蹲式數量");
                            if (manData.Toilet_Count >= 2) { counts.Set(1); }
                        }
                        manRestrooms.Add(instance);
                    }
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.toilet_squat.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
                if (mop_basin != null)
                {
                    if (manData.Toilet_Count >= 3)
                    {
                        FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, mop_basin, StructuralType.NonStructural);
                        manRestrooms.Add(instance);
                    }
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.mop_basin.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
                if (washbasin != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, washbasin, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains(RestroomGroup.washbasin))
                    {
                        Parameter counts = instance.LookupParameter("兒童洗面盆數量");
                        if (manData.Washbasin_Count >= 3) // 預留一個兒童洗臉盆、一個無障礙洗臉盆
                        {
                            counts.Set(1);
                            counts = instance.LookupParameter("一般洗面盆數量");
                            counts.Set(manData.Washbasin_Count - 2);
                        }
                        else if (manData.Washbasin_Count == 2) // 預留一個無障礙洗臉盆
                        {
                            counts.Set(0);
                            counts = instance.LookupParameter("一般洗面盆數量");
                            counts.Set(1);
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
                else
                {
                    string noFamilySymbol = RestroomGroup.washbasin.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
                if (washbasin_accessible != null)
                {
                    if (manData.Washbasin_Count >= 2)
                    {
                        FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, washbasin_accessible, StructuralType.NonStructural);
                        manRestrooms.Add(instance);
                    }
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.washbasin_accessible.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
                if (urinal != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, urinal, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains(RestroomGroup.urinal))
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
                else
                {
                    string noFamilySymbol = RestroomGroup.urinal.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
            }

            return manRestrooms;
        }
        // 建立女廁元件
        private List<FamilyInstance> CreateWomanRestroom(Document doc, JsonData jsonData, double levelElevation, List<FamilySymbol> familySymbolList)
        {
            List<FamilyInstance> womanRestrooms = new List<FamilyInstance>();
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.toilet)).FirstOrDefault(); // 坐式馬桶
            FamilySymbol toilet_squat = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.toilet_squat)).FirstOrDefault(); // 蹲式馬桶
            FamilySymbol mop_basin = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.mop_basin)).FirstOrDefault(); // 拖布盆
            FamilySymbol washbasin = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.washbasin)).FirstOrDefault(); // 洗手台
            FamilySymbol washbasin_accessible = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.washbasin_accessible)).FirstOrDefault(); // 無障礙洗手台

            foreach (WomanData womanData in jsonData.WomanDataList)
            {
                double x = UnitUtils.ConvertToInternalUnits(womanData.RestroomWoman_x, DisplayUnitType.DUT_METERS);
                double y = UnitUtils.ConvertToInternalUnits(womanData.RestroomWoman_y, DisplayUnitType.DUT_METERS);
                XYZ xyz = new XYZ(x, y, levelElevation);
                if (toilet != null)
                {
                    if (womanData.Type.Equals(1) || womanData.Type.Equals(2) || womanData.Type.Equals(3))
                    {
                        double toiletCount = womanData.Toilet_Count;
                        if (womanData.Toilet_Count >= 4)
                        {
                            toiletCount = womanData.Toilet_Count - 2; // 預留一個蹲式、一個拖布盆
                            FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                            {
                                Parameter counts = instance.LookupParameter("一般坐式數量");
                                counts.Set(Math.Ceiling(toiletCount / 2));
                            }
                            womanRestrooms.Add(instance);

                            instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                            {
                                Parameter counts = instance.LookupParameter("一般坐式數量");
                                counts.Set(toiletCount - Math.Ceiling(toiletCount / 2));
                            }
                            womanRestrooms.Add(instance);
                        }
                        else
                        {
                            FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                            {
                                Parameter counts = instance.LookupParameter("一般坐式數量");
                                counts.Set(Math.Ceiling(toiletCount / 2));
                            }
                            womanRestrooms.Add(instance);

                            instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                            {
                                Parameter counts = instance.LookupParameter("一般坐式數量");
                                counts.Set(toiletCount - Math.Ceiling(toiletCount / 2));
                            }
                            womanRestrooms.Add(instance);
                        }
                    }
                    else if (womanData.Type.Equals(8))
                    {
                        double toiletCount = womanData.Toilet_Count;
                        if (womanData.Toilet_Count >= 4)
                        {
                            toiletCount = womanData.Toilet_Count - 2; // 預留一個蹲式、一個拖布盆
                            FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                            {
                                Parameter counts = instance.LookupParameter("一般坐式數量");
                                counts.Set(Math.Ceiling(toiletCount * 3 / 4));
                            }
                            womanRestrooms.Add(instance);

                            instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                            {
                                Parameter counts = instance.LookupParameter("一般坐式數量");
                                counts.Set(toiletCount - Math.Ceiling(toiletCount * 3 / 4));
                            }
                            womanRestrooms.Add(instance);
                        }
                        else if(womanData.Toilet_Count >= 2)
                        {
                            FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                            {
                                Parameter counts = instance.LookupParameter("一般坐式數量");
                                counts.Set(Math.Ceiling(toiletCount / 2));
                            }
                            womanRestrooms.Add(instance);

                            instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                            if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                            {
                                Parameter counts = instance.LookupParameter("一般坐式數量");
                                counts.Set(toiletCount - Math.Ceiling(toiletCount / 2));
                            }
                            womanRestrooms.Add(instance);
                        }
                    }
                    else
                    {
                        FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                        if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet))
                        {
                            Parameter counts = instance.LookupParameter("一般坐式數量");
                            counts.Set(womanData.Toilet_Count - 2); // 預留一個蹲式、一個拖布盆
                        }
                        womanRestrooms.Add(instance);
                    }
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.toilet.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
                if(toilet_squat != null)
                {
                    if (womanData.Toilet_Count >= 3)
                    {
                        FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet_squat, StructuralType.NonStructural);
                        if (instance.Symbol.FamilyName.Contains(RestroomGroup.toilet_squat))
                        {
                            Parameter counts = instance.LookupParameter("蹲式數量");
                            if (womanData.Toilet_Count >= 3) { counts.Set(1); }
                        }
                        womanRestrooms.Add(instance);
                    }
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.toilet_squat.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
                if (mop_basin != null)
                {
                    if (womanData.Toilet_Count >= 4)
                    {
                        FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, mop_basin, StructuralType.NonStructural);
                        womanRestrooms.Add(instance);
                    }
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.mop_basin.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
                if (washbasin != null)
                {

                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, washbasin, StructuralType.NonStructural);
                    if (instance.Symbol.FamilyName.Contains(RestroomGroup.washbasin))
                    {
                        Parameter counts = instance.LookupParameter("兒童洗面盆數量");
                        if (womanData.Type.Equals(3) || womanData.Type.Equals(5))
                        {
                            counts.Set(1);
                            counts = instance.LookupParameter("一般洗面盆數量");
                            counts.Set(1);
                        }
                        else
                        {
                            if (womanData.Washbasin_Count >= 3) // 預留一個兒童洗臉盆、一個無障礙洗臉盆
                            {
                                counts.Set(1);
                                counts = instance.LookupParameter("一般洗面盆數量");
                                counts.Set(womanData.Washbasin_Count - 2);
                            }
                            else if (womanData.Washbasin_Count == 2) // 預留一個無障礙洗臉盆
                            {
                                counts.Set(0);
                                counts = instance.LookupParameter("一般洗面盆數量");
                                counts.Set(1);
                            }
                            else
                            {
                                counts.Set(0);
                                counts = instance.LookupParameter("一般洗面盆數量");
                                counts.Set(womanData.Washbasin_Count);
                            }
                        }
                    }
                    womanRestrooms.Add(instance);

                    if (womanData.Type.Equals(3) || womanData.Type.Equals(5))
                    {
                        instance = doc.Create.NewFamilyInstance(xyz, washbasin, StructuralType.NonStructural);
                        if (instance.Symbol.FamilyName.Contains(RestroomGroup.washbasin))
                        {
                            Parameter counts = instance.LookupParameter("兒童洗面盆數量");
                            counts.Set(0);
                            counts = instance.LookupParameter("一般洗面盆數量");
                            counts.Set(1);
                        }
                        womanRestrooms.Add(instance);
                    }
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.washbasin.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
                if (washbasin_accessible != null)
                {
                    if (womanData.Washbasin_Count >= 2)
                    {
                        FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, washbasin_accessible, StructuralType.NonStructural);
                        womanRestrooms.Add(instance);
                    }
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.washbasin_accessible.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
            }
            return womanRestrooms;
        }
        // 建立親子廁所元件
        private List<FamilyInstance> CreateFamilyRestroom(Document doc, JsonData jsonData, double levelElevation, List<FamilySymbol> familySymbolList)
        {
            List<FamilyInstance> familyRestrooms = new List<FamilyInstance>();
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.familyRestroom)).FirstOrDefault();

            foreach (FamilyData familyData in jsonData.FamilyDataList)
            {
                double x = UnitUtils.ConvertToInternalUnits(familyData.RestroomFamily_x, DisplayUnitType.DUT_METERS);
                double y = UnitUtils.ConvertToInternalUnits(familyData.RestroomFamily_y, DisplayUnitType.DUT_METERS);
                XYZ xyz = new XYZ(x, y, levelElevation);
                if (toilet != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                    familyRestrooms.Add(instance);
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.familyRestroom.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
            }
            return familyRestrooms;
        }
        // 建立無障礙廁所元件
        private List<FamilyInstance> CreateAccessibleRestroom(Document doc, JsonData jsonData, double levelElevation, List<FamilySymbol> familySymbolList)
        {
            List<FamilyInstance> accessibleRestrooms = new List<FamilyInstance>();
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.accessibleRestroom)).FirstOrDefault();

            foreach (AccessibleData accessibleData in jsonData.AccessibleDataList)
            {
                double x = UnitUtils.ConvertToInternalUnits(accessibleData.RestroomAccessible_x, DisplayUnitType.DUT_METERS);
                double y = UnitUtils.ConvertToInternalUnits(accessibleData.RestroomAccessible_y, DisplayUnitType.DUT_METERS);
                XYZ xyz = new XYZ(x, y, levelElevation);
                if (toilet != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                    accessibleRestrooms.Add(instance);
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.accessibleRestroom.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
            }
            return accessibleRestrooms;
        }
        // 建立哺集乳室
        private List<FamilyInstance> CreateBreastfeedingRoom(Document doc, JsonData jsonData, double levelElevation, List<FamilySymbol> familySymbolList)
        {
            List<FamilyInstance> breastfeedings = new List<FamilyInstance>();
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.breastfeeding)).FirstOrDefault();

            foreach (BreastfeedingData breastfeedingData in jsonData.BreastfeedingDataList)
            {
                double x = UnitUtils.ConvertToInternalUnits(breastfeedingData.BreastfeedingRoom_x, DisplayUnitType.DUT_METERS);
                double y = UnitUtils.ConvertToInternalUnits(breastfeedingData.BreastfeedingRoom_y, DisplayUnitType.DUT_METERS);
                XYZ xyz = new XYZ(x, y, levelElevation);
                if (toilet != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                    breastfeedings.Add(instance);
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.breastfeeding.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
            }
            return breastfeedings;
        }
        // 建立清潔人員休息室
        private List<FamilyInstance> CreateJanitorRoom(Document doc, JsonData jsonData, double levelElevation, List<FamilySymbol> familySymbolList)
        {
            List<FamilyInstance> janitorRooms = new List<FamilyInstance>();
            FamilySymbol toilet = familySymbolList.Where(x => x.FamilyName.Contains(RestroomGroup.janitorRoom)).FirstOrDefault();

            foreach (JanitorRoomData janitorRoomData in jsonData.JanitorRoomDataList)
            {
                double x = UnitUtils.ConvertToInternalUnits(janitorRoomData.JanitorRoom_x, DisplayUnitType.DUT_METERS);
                double y = UnitUtils.ConvertToInternalUnits(janitorRoomData.JanitorRoom_y, DisplayUnitType.DUT_METERS);
                XYZ xyz = new XYZ(x, y, levelElevation);
                if (toilet != null)
                {
                    FamilyInstance instance = doc.Create.NewFamilyInstance(xyz, toilet, StructuralType.NonStructural);
                    janitorRooms.Add(instance);
                }
                else
                {
                    string noFamilySymbol = RestroomGroup.janitorRoom.Replace("(SinoBIM-第", "");
                    noFamilySymbols.Add(noFamilySymbol);
                }
            }
            return janitorRooms;
        }
        // 排列元件組合
        private void EditParameter(Document doc, Tuple<List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>, List<FamilyInstance>> restroomElems, JsonData jsonData)
        {
            using (Transaction trans = new Transaction(doc, "組合"))
            {
                trans.Start();
                //EditManRestroom(doc, restroomElems.Item1, jsonData); // 修改男廁參數
                EditWomanRestroom(doc, restroomElems.Item2, jsonData); // 修改女廁參數
                //EditFamilyRestroom(doc, restroomElems.Item3, jsonData); // 修改親子廁所參數
                //EditAccessibleRestroom(doc, restroomElems.Item4, jsonData); // 修改無障礙廁所參數
                //EditBreastfeedingRoom(doc, restroomElems.Item5, jsonData); // 修改哺集乳室
                //EditJanitorRooms(doc, restroomElems.Item6, jsonData); // 修改清潔人員休息室
                trans.Commit();
            }
        }
        // 修改男廁參數
        private void EditManRestroom(Document doc, List<FamilyInstance> manRestroomElems, JsonData jsonData)
        {
            foreach (ManData manData in jsonData.ManDataList)
            {
                FamilyInstance toilet = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.toilet)).FirstOrDefault(); // 坐式馬桶
                FamilyInstance toilet_squat = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.toilet_squat)).FirstOrDefault(); // 蹲式馬桶
                FamilyInstance mop_basin = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.mop_basin)).FirstOrDefault(); // 拖布盆間
                FamilyInstance washbasin = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.washbasin)).FirstOrDefault(); // 洗手台
                FamilyInstance washbasin_accessible = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.washbasin_accessible)).FirstOrDefault(); // 無障礙洗手台
                FamilyInstance urinal = manRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.urinal)).FirstOrDefault(); // 小便斗

                // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                double toiletTotalWidth = 0.0;
                double toiletTotalDepth = 0.0;
                double partitionThickness = 0.0;
                if (toilet != null)
                {
                    toiletTotalWidth = toilet.LookupParameter("一般坐式總寬度").AsDouble();
                    toiletTotalDepth = toilet.LookupParameter("總深度").AsDouble();
                    partitionThickness = toilet.LookupParameter("隔板厚度").AsDouble() / 2;
                }

                // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                double toilet_squat_TotalWidth = 0.0;
                double toilet_squat_TotalDepth = 0.0;
                double toilet_squat_partitionThickness = 0.0;
                if (toilet_squat != null)
                {
                    toilet_squat_TotalWidth = toilet_squat.LookupParameter("蹲式寬度").AsDouble();
                    toilet_squat_TotalDepth = toilet_squat.LookupParameter("總深度").AsDouble();
                    toilet_squat_partitionThickness = toilet_squat.LookupParameter("隔板厚度").AsDouble() / 2;
                }

                // MRT_拖布盆間(SinoBIM-第2版)
                double mop_basin_TotalWidth = 0.0;
                double mop_basin_TotalDepth = 0.0;
                double mop_basin_partitionThickness = 0.0;
                if (mop_basin != null)
                {
                    mop_basin_TotalWidth = mop_basin.LookupParameter("拖布盆間寬度").AsDouble();
                    mop_basin_TotalDepth = mop_basin.LookupParameter("總深度").AsDouble();
                    mop_basin_partitionThickness = mop_basin.LookupParameter("隔板厚度").AsDouble() / 2;
                }

                // MRT_洗手台群組(SinoBIM-第3版)
                double washbasinTotalWidth = 0.0;
                double washbasinTotalDepth = 0.0;
                double washbasinNormalWidth = 0.0;
                double washbasinNormalCount = 0.0;
                if (washbasin != null)
                {
                    washbasinTotalWidth = washbasin.LookupParameter("檯面標準寬度").AsDouble();
                    washbasinTotalDepth = washbasin.LookupParameter("總深度").AsDouble();
                    washbasinNormalWidth = washbasin.LookupParameter("一般洗手台寬度").AsDouble();
                    washbasinNormalCount = washbasin.LookupParameter("一般洗面盆數量").AsInteger();
                }

                // MRT_無障礙洗面盆(SinoBIM-第1版)
                double washbasin_accessible_TotalWidth = 0.0;
                double washbasin_accessible_TotalDepth = 0.0;
                if (washbasin_accessible != null)
                {
                    washbasin_accessible_TotalWidth = washbasin_accessible.LookupParameter("設備箱預設寬度").AsDouble() +
                                                      washbasin_accessible.LookupParameter("設備箱左側延伸").AsDouble() +
                                                      washbasin_accessible.LookupParameter("設備箱右側延伸").AsDouble();
                    washbasin_accessible_TotalDepth = washbasin_accessible.LookupParameter("設備箱深度").AsDouble();
                }

                // MRT_小便斗群組(SinoBIM-第1版)
                double urinalTotalWidth = 0.0;
                double urinalTotalDepth = 0.0;
                if (urinal != null)
                {
                    urinalTotalWidth = urinal.LookupParameter("總寬度").AsDouble();
                    urinalTotalDepth = urinal.LookupParameter("總深度").AsDouble();
                }

                double manRestroomLength = UnitUtils.ConvertToInternalUnits(manData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double manRestroomWidth = UnitUtils.ConvertToInternalUnits(manData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度
                double space = UnitUtils.ConvertToInternalUnits(1.5, DisplayUnitType.DUT_METERS); // 間距

                if (manData.Type.Equals(1) || manData.Type.Equals(2) || manData.Type.Equals(3))
                {
                    if (manData.Rotate_id.Equals(1) || manData.Rotate_id.Equals(2))
                    {
                        // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet != null)
                        {
                            XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth, -toiletTotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet.Id, offset);
                            manRestroomElems.Remove(toilet); // 移除修改過的族群
                        }
                        // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                        if (toilet_squat != null)
                        {
                            XYZ offset = new XYZ(mop_basin_TotalWidth, -toilet_squat_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet_squat.Id, offset);
                            manRestroomElems.Remove(toilet_squat); // 移除修改過的族群
                        }
                        // MRT_拖布盆間(SinoBIM-第2版)
                        if (mop_basin != null)
                        {
                            XYZ offset = new XYZ(0, -mop_basin_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, mop_basin.Id, offset);
                            manRestroomElems.Remove(mop_basin); // 移除修改過的族群
                        }
                        // MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin != null)
                        {
                            if (manData.Type.Equals(1))
                            {
                                XYZ offset = new XYZ(manRestroomLength - washbasinTotalWidth - partitionThickness, -washbasinTotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            else if (manData.Type.Equals(2))
                            {
                                LocationPoint lp = washbasin.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                                XYZ offset = new XYZ(manRestroomLength - washbasinTotalDepth, -washbasin_accessible_TotalWidth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            else if (manData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                                XYZ offset = new XYZ(manRestroomLength - washbasinTotalDepth, 0, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            manRestroomElems.Remove(washbasin); // 移除修改過的族群
                        }
                        // MRT_無障礙洗面盆(SinoBIM-第1版)
                        if (washbasin_accessible != null)
                        {
                            if (manData.Type.Equals(1))
                            {
                                XYZ offset = new XYZ(manRestroomLength - washbasinTotalWidth - washbasin_accessible_TotalWidth / 2, -washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (manData.Type.Equals(2))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(manRestroomLength - washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (manData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth + toiletTotalWidth + partitionThickness + washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            manRestroomElems.Remove(washbasin_accessible); // 移除修改過的族群
                        }
                        // MRT_小便斗群組(SinoBIM-第1版)
                        if (urinal != null)
                        {
                            LocationPoint lp = urinal.Location as LocationPoint;
                            Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                            double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                            ElementTransformUtils.RotateElement(doc, urinal.Id, line, rotation);
                            XYZ offset = new XYZ(urinalTotalWidth, -manRestroomWidth + urinalTotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, urinal.Id, offset);
                            manRestroomElems.Remove(urinal); // 移除修改過的族群
                        }
                    }
                    else if (manData.Rotate_id.Equals(3) || manData.Rotate_id.Equals(4))
                    {
                        // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet != null)
                        {
                            XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth, -toiletTotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet.Id, offset);
                            manRestroomElems.Remove(toilet); // 移除修改過的族群
                        }
                        // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                        if (toilet_squat != null)
                        {
                            XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth, - toilet_squat_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet_squat.Id, offset);
                            manRestroomElems.Remove(toilet_squat); // 移除修改過的族群
                        }
                        // MRT_拖布盆間(SinoBIM-第2版)
                        if (mop_basin != null)
                        {
                            XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth, - mop_basin_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, mop_basin.Id, offset);
                            manRestroomElems.Remove(mop_basin); // 移除修改過的族群
                        }
                        // MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin != null)
                        {
                            if (manData.Type.Equals(1))
                            {
                                XYZ offset = new XYZ(0, -washbasinTotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            else if (manData.Type.Equals(2))
                            {
                                LocationPoint lp = washbasin.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                                XYZ offset = new XYZ(washbasinTotalDepth, -washbasinTotalWidth - washbasin_accessible_TotalWidth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            else if (manData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                                XYZ offset = new XYZ(washbasinTotalDepth, -washbasinTotalWidth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            manRestroomElems.Remove(washbasin); // 移除修改過的族群
                        }
                        // MRT_無障礙洗面盆(SinoBIM-第1版)
                        if (washbasin_accessible != null)
                        {
                            if (manData.Type.Equals(1))
                            {
                                XYZ offset = new XYZ(washbasinTotalWidth + washbasin_accessible_TotalWidth / 2, -washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (manData.Type.Equals(2))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (manData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth - washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            manRestroomElems.Remove(washbasin_accessible); // 移除修改過的族群
                        }
                        // MRT_小便斗群組(SinoBIM-第1版)
                        if (urinal != null)
                        {
                            LocationPoint lp = urinal.Location as LocationPoint;
                            Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                            double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                            ElementTransformUtils.RotateElement(doc, urinal.Id, line, rotation);
                            XYZ offset = new XYZ(manRestroomLength, -manRestroomWidth + urinalTotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, urinal.Id, offset);
                            manRestroomElems.Remove(urinal); // 移除修改過的族群
                        }
                    }
                }
                else if(manData.Type.Equals(4) || manData.Type.Equals(5) || manData.Type.Equals(6) || manData.Type.Equals(7))
                {
                    if (manData.Rotate_id.Equals(1) || manData.Rotate_id.Equals(2))
                    {
                        // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet != null)
                        {
                            XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth, -toiletTotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet.Id, offset);
                            manRestroomElems.Remove(toilet); // 移除修改過的族群
                        }
                        // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                        if (toilet_squat != null)
                        {
                            XYZ offset = new XYZ(mop_basin_TotalWidth, -toilet_squat_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet_squat.Id, offset);
                            manRestroomElems.Remove(toilet_squat); // 移除修改過的族群
                        }
                        // MRT_拖布盆間(SinoBIM-第2版)
                        if (mop_basin != null)
                        {
                            XYZ offset = new XYZ(0, -mop_basin_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, mop_basin.Id, offset);
                            manRestroomElems.Remove(mop_basin); // 移除修改過的族群
                        }
                        // MRT_小便斗群組(SinoBIM-第1版)
                        if (urinal != null)
                        {
                            XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth + toiletTotalWidth, -urinalTotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, urinal.Id, offset);
                            manRestroomElems.Remove(urinal); // 移除修改過的族群
                        }
                        // MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin != null)
                        {
                            if (manData.Type.Equals(4))
                            {
                                XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth + toiletTotalWidth + urinalTotalWidth + washbasin_accessible_TotalWidth, -washbasinTotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            else if (manData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                                XYZ offset = new XYZ(manRestroomLength - washbasinTotalDepth - space, 0, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            else if (manData.Type.Equals(6))
                            {
                                LocationPoint lp = washbasin.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                                XYZ offset = new XYZ(manRestroomLength - space, -manRestroomWidth + washbasinTotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            else if (manData.Type.Equals(7))
                            {
                                LocationPoint lp = washbasin.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                                double washbasinPlace = 0.0;
                                if (manData.Rotate_id.Equals(1))
                                {
                                    washbasinPlace = manRestroomLength / 2 - washbasinNormalWidth / washbasinNormalCount;
                                }
                                else if (manData.Rotate_id.Equals(2))
                                {
                                    washbasinPlace = manRestroomLength / 2 + washbasinNormalWidth / washbasinNormalCount;
                                }
                                XYZ offset = new XYZ(washbasinPlace, -manRestroomWidth + washbasinTotalDepth + space, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            manRestroomElems.Remove(washbasin); // 移除修改過的族群
                        }
                        // MRT_無障礙洗面盆(SinoBIM-第1版)
                        if (washbasin_accessible != null)
                        {
                            if (manData.Type.Equals(4))
                            {
                                XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth + toiletTotalWidth + urinalTotalWidth + washbasin_accessible_TotalWidth / 2, -washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (manData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth + toiletTotalWidth + urinalTotalWidth + washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (manData.Type.Equals(6))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(manRestroomLength - space - washbasinTotalWidth - washbasin_accessible_TotalWidth / 2, -manRestroomWidth + washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (manData.Type.Equals(7))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                double washbasin_accessible_place = 0.0;
                                if (manData.Rotate_id.Equals(1))
                                {
                                    washbasin_accessible_place = manRestroomLength / 2 - washbasinNormalWidth / washbasinNormalCount + washbasin_accessible_TotalWidth / 2;
                                }
                                else if (manData.Rotate_id.Equals(2))
                                {
                                    washbasin_accessible_place = manRestroomLength / 2 + washbasinNormalWidth / washbasinNormalCount - washbasinTotalWidth - washbasin_accessible_TotalWidth / 2;
                                }
                                XYZ offset = new XYZ(washbasin_accessible_place, -manRestroomWidth + washbasin_accessible_TotalDepth + space, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            manRestroomElems.Remove(washbasin_accessible); // 移除修改過的族群
                        }
                    }
                    else if (manData.Rotate_id.Equals(3) || manData.Rotate_id.Equals(4))
                    {
                        // MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet != null)
                        {
                            XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth, -toiletTotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet.Id, offset);
                            manRestroomElems.Remove(toilet); // 移除修改過的族群
                        }
                        // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                        if (toilet_squat != null)
                        {
                            XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth, -toilet_squat_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet_squat.Id, offset);
                            manRestroomElems.Remove(toilet_squat); // 移除修改過的族群
                        }
                        // MRT_拖布盆間(SinoBIM-第2版)
                        if (mop_basin != null)
                        {
                            XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth, -mop_basin_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, mop_basin.Id, offset);
                            manRestroomElems.Remove(mop_basin); // 移除修改過的族群
                        }
                        // MRT_小便斗群組(SinoBIM-第1版)
                        if (urinal != null)
                        {
                            XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth - urinalTotalWidth, -urinalTotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, urinal.Id, offset);
                            manRestroomElems.Remove(urinal); // 移除修改過的族群
                        }
                        // MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin != null)
                        {
                            if (manData.Type.Equals(4))
                            {
                                XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth - urinalTotalWidth - washbasin_accessible_TotalWidth - washbasinTotalWidth, -washbasinTotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            else if (manData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                                XYZ offset = new XYZ(washbasinTotalDepth + space, -washbasinTotalWidth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            else if (manData.Type.Equals(6))
                            {
                                LocationPoint lp = washbasin.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin.Id, line, rotation);
                                XYZ offset = new XYZ(space + washbasinTotalWidth, -manRestroomWidth + washbasinTotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin.Id, offset);
                            }
                            manRestroomElems.Remove(washbasin); // 移除修改過的族群
                        }
                        // MRT_無障礙洗面盆(SinoBIM-第1版)
                        if (washbasin_accessible != null)
                        {
                            if (manData.Type.Equals(4))
                            {
                                XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth - urinalTotalWidth - washbasin_accessible_TotalWidth / 2, -washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (manData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(manRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth - urinalTotalWidth - washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (manData.Type.Equals(6))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(space + washbasinTotalWidth + washbasin_accessible_TotalWidth / 2, -manRestroomWidth + washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            manRestroomElems.Remove(washbasin_accessible); // 移除修改過的族群
                        }
                    }
                }
            }
        }
        // 修改女廁參數
        private void EditWomanRestroom(Document doc, List<FamilyInstance> womanRestroomElems, JsonData jsonData)
        {
            foreach (WomanData womanData in jsonData.WomanDataList)
            {
                FamilyInstance toilet1 = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.toilet)).FirstOrDefault(); // 坐式馬桶1
                FamilyInstance toilet2 = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.toilet) && x.Id != toilet1.Id).FirstOrDefault(); // 坐式馬桶2
                // 如果toilet2不為null, 先比對toilet1、toilet2哪個數量高, 將toilet1設置為數量高的
                if(toilet2 != null)
                {
                    FamilyInstance toiletMax = toilet1;
                    FamilyInstance toiletMin = toilet2;
                    int count1 = toilet1.LookupParameter("一般坐式數量").AsInteger();
                    int count2 = toilet2.LookupParameter("一般坐式數量").AsInteger();
                    if (count2 > count1) { toiletMax = toilet2; toiletMin = toilet1; }
                    toilet1 = toiletMax;
                    toilet2 = toiletMin;
                }
                FamilyInstance toilet_squat = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.toilet_squat)).FirstOrDefault(); // 蹲式馬桶
                FamilyInstance mop_basin = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.mop_basin)).FirstOrDefault(); // 拖布盆間
                FamilyInstance washbasin1 = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.washbasin)).FirstOrDefault(); // 洗手台1
                FamilyInstance washbasin2 = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.washbasin) && x.Id != washbasin1.Id).FirstOrDefault(); // 洗手台2
                // 如果washbasin2不為null, 先比對washbasin1、washbasin2哪個數量高, 將washbasin1設置為數量高的
                if(washbasin2 != null)
                {
                    FamilyInstance washbasinMax = washbasin1;
                    FamilyInstance washbasinMin = washbasin2;
                    int count1 = washbasin1.LookupParameter("洗面盆總數量").AsInteger();
                    int count2 = washbasin2.LookupParameter("洗面盆總數量").AsInteger();
                    if (count2 > count1) { washbasinMax = washbasin2; washbasinMin = washbasin1; }
                    washbasin1 = washbasinMax;
                    washbasin2 = washbasinMin;
                }
                FamilyInstance washbasin_accessible = womanRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.washbasin_accessible)).FirstOrDefault(); // 無障礙洗手台

                // toilet1：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                double toiletTotalWidth1 = 0.0;
                double toiletTotalDepth1 = 0.0;
                double partitionThickness1 = 0.0;
                if(toilet1 != null)
                {
                    toiletTotalWidth1 = toilet1.LookupParameter("一般坐式總寬度").AsDouble();
                    toiletTotalDepth1 = toilet1.LookupParameter("總深度").AsDouble();
                    partitionThickness1 = toilet1.LookupParameter("隔板厚度").AsDouble() / 2;
                }

                // toilet2：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                double toiletTotalWidth2 = 0.0;
                double toiletTotalDepth2 = 0.0;
                double partitionThickness2 = 0.0;
                if(toilet2 != null)
                {
                    toiletTotalWidth2 = toilet2.LookupParameter("一般坐式總寬度").AsDouble();
                    toiletTotalDepth2 = toilet2.LookupParameter("總深度").AsDouble();
                    partitionThickness2 = toilet2.LookupParameter("隔板厚度").AsDouble() / 2;
                }

                // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                double toilet_squat_TotalWidth = 0.0;
                double toilet_squat_TotalDepth = 0.0;
                double toilet_squat_partitionThickness = 0.0;
                if (toilet_squat != null)
                {
                    toilet_squat_TotalWidth = toilet_squat.LookupParameter("蹲式寬度").AsDouble();
                    toilet_squat_TotalDepth = toilet_squat.LookupParameter("總深度").AsDouble();
                    toilet_squat_partitionThickness = toilet_squat.LookupParameter("隔板厚度").AsDouble() / 2;
                }

                // MRT_拖布盆間(SinoBIM-第2版)
                double mop_basin_TotalWidth = 0.0;
                double mop_basin_TotalDepth = 0.0;
                double mop_basin_partitionThickness = 0.0;
                if (mop_basin != null)
                {
                    mop_basin_TotalWidth = mop_basin.LookupParameter("拖布盆間寬度").AsDouble();
                    mop_basin_TotalDepth = mop_basin.LookupParameter("總深度").AsDouble();
                    mop_basin_partitionThickness = mop_basin.LookupParameter("隔板厚度").AsDouble() / 2;
                }

                // washbasin1：MRT_洗手台群組(SinoBIM-第3版)
                double washbasinTotalWidth1 = 0.0;
                double washbasinTotalDepth1 = 0.0;
                if(washbasin1 != null)
                {
                    washbasinTotalWidth1 = washbasin1.LookupParameter("檯面標準寬度").AsDouble();
                    washbasinTotalDepth1 = washbasin1.LookupParameter("總深度").AsDouble();
                }

                // washbasin2：MRT_洗手台群組(SinoBIM-第3版)
                double washbasinTotalWidth2 = 0.0;
                double washbasinTotalDepth2 = 0.0;
                if (washbasin2 != null)
                {
                    washbasinTotalWidth2 = washbasin2.LookupParameter("檯面標準寬度").AsDouble();
                    washbasinTotalDepth2 = washbasin2.LookupParameter("總深度").AsDouble();
                }

                // MRT_無障礙洗面盆(SinoBIM-第1版)
                double washbasin_accessible_TotalWidth = 0.0;
                double washbasin_accessible_TotalDepth = 0.0;
                if (washbasin_accessible != null)
                {
                    washbasin_accessible_TotalWidth = washbasin_accessible.LookupParameter("設備箱預設寬度").AsDouble() +
                                                      washbasin_accessible.LookupParameter("設備箱左側延伸").AsDouble() +
                                                      washbasin_accessible.LookupParameter("設備箱右側延伸").AsDouble();
                    washbasin_accessible_TotalDepth = washbasin_accessible.LookupParameter("設備箱深度").AsDouble();
                }

                double womanRestroomLength = UnitUtils.ConvertToInternalUnits(womanData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double womanRestroomWidth = UnitUtils.ConvertToInternalUnits(womanData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度
                double space = UnitUtils.ConvertToInternalUnits(1.5, DisplayUnitType.DUT_METERS); // 間距

                if (womanData.Type.Equals(1) || womanData.Type.Equals(2) || womanData.Type.Equals(3) || womanData.Type.Equals(8))
                {
                    if (womanData.Rotate_id.Equals(1) || womanData.Rotate_id.Equals(2))
                    {
                        // toilet1：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet1 != null)
                        {
                            XYZ offset = new XYZ(toilet_squat_TotalWidth, -toiletTotalDepth1, 0);
                            ElementTransformUtils.MoveElement(doc, toilet1.Id, offset);
                            womanRestroomElems.Remove(toilet1); // 移除修改過的族群
                        }
                        // toilet2：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet2 != null)
                        {
                            LocationPoint lp = toilet2.Location as LocationPoint;
                            Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                            double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                            ElementTransformUtils.RotateElement(doc, toilet2.Id, line, rotation);
                            XYZ offset = new XYZ(mop_basin_TotalWidth + toiletTotalWidth2, -womanRestroomWidth + toiletTotalDepth2, 0);
                            ElementTransformUtils.MoveElement(doc, toilet2.Id, offset);
                            womanRestroomElems.Remove(toilet2); // 移除修改過的族群
                        }
                        // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                        if (toilet_squat != null)
                        {
                            XYZ offset = new XYZ(0, -toilet_squat_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet_squat.Id, offset);
                            womanRestroomElems.Remove(toilet_squat); // 移除修改過的族群
                        }
                        // MRT_拖布盆間(SinoBIM-第2版)
                        if (mop_basin != null)
                        {
                            LocationPoint lp = mop_basin.Location as LocationPoint;
                            Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                            double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                            ElementTransformUtils.RotateElement(doc, mop_basin.Id, line, rotation);
                            XYZ offset = new XYZ(mop_basin_TotalWidth, -womanRestroomWidth + mop_basin_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, mop_basin.Id, offset);
                            womanRestroomElems.Remove(mop_basin); // 移除修改過的族群
                        }
                        // washbasin1：MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin1 != null)
                        {
                            if (womanData.Type.Equals(1))
                            {
                                XYZ offset = new XYZ(womanRestroomLength - washbasinTotalWidth1, -washbasinTotalDepth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(2))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - washbasinTotalDepth1, -washbasin_accessible_TotalWidth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(toilet_squat_TotalWidth + toiletTotalWidth1 + washbasinTotalDepth1, -washbasinTotalWidth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            womanRestroomElems.Remove(washbasin1); // 移除修改過的族群
                        }
                        // washbasin2：MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin2 != null)
                        {
                            if (womanData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin2.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin2.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength, -washbasin_accessible_TotalWidth - washbasinTotalWidth2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin2.Id, offset);
                                womanRestroomElems.Remove(washbasin2); // 移除修改過的族群
                            }
                        }
                        // MRT_無障礙洗面盆(SinoBIM-第1版)
                        if (washbasin_accessible != null)
                        {
                            if (womanData.Type.Equals(1))
                            {
                                XYZ offset = new XYZ(womanRestroomLength - washbasinTotalWidth1 - washbasin_accessible_TotalWidth / 2, -washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(2))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            womanRestroomElems.Remove(washbasin_accessible); // 移除修改過的族群
                        }
                    }
                    else if (womanData.Rotate_id.Equals(3) || womanData.Rotate_id.Equals(4))
                    {
                        // toilet1：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet1 != null)
                        {
                            XYZ offset = new XYZ(womanRestroomLength - toilet_squat_TotalWidth - toiletTotalWidth1, -toiletTotalDepth1, 0);
                            ElementTransformUtils.MoveElement(doc, toilet1.Id, offset);
                            womanRestroomElems.Remove(toilet1); // 移除修改過的族群
                        }
                        // toilet2：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet2 != null)
                        {
                            LocationPoint lp = toilet2.Location as LocationPoint;
                            Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                            double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                            ElementTransformUtils.RotateElement(doc, toilet2.Id, line, rotation);
                            XYZ offset = new XYZ(womanRestroomLength - mop_basin_TotalWidth, -womanRestroomWidth + toiletTotalDepth2, 0);
                            ElementTransformUtils.MoveElement(doc, toilet2.Id, offset);
                            womanRestroomElems.Remove(toilet2); // 移除修改過的族群
                        }
                        // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                        if (toilet_squat != null)
                        {
                            XYZ offset = new XYZ(womanRestroomLength - toilet_squat_TotalWidth, -toilet_squat_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet_squat.Id, offset);
                            womanRestroomElems.Remove(toilet_squat); // 移除修改過的族群
                        }
                        // MRT_拖布盆間(SinoBIM-第2版)
                        if (mop_basin != null)
                        {
                            LocationPoint lp = mop_basin.Location as LocationPoint;
                            Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                            double rotation = 2 * Math.PI / 360 * 180; // 轉換角度
                            ElementTransformUtils.RotateElement(doc, mop_basin.Id, line, rotation);
                            XYZ offset = new XYZ(womanRestroomLength, -womanRestroomWidth + mop_basin_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, mop_basin.Id, offset);
                            womanRestroomElems.Remove(mop_basin); // 移除修改過的族群
                        }
                        // washbasin1：MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin1 != null)
                        {
                            if (womanData.Type.Equals(1))
                            {
                                XYZ offset = new XYZ(0, -washbasinTotalDepth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(2))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(washbasinTotalDepth1, -washbasin_accessible_TotalWidth - washbasinTotalWidth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - toilet_squat_TotalWidth - toiletTotalWidth1 - washbasinTotalDepth1, 0, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            womanRestroomElems.Remove(washbasin1); // 移除修改過的族群
                        }
                        // washbasin2：MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin2 != null)
                        {
                            if (womanData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin2.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin2.Id, line, rotation);
                                XYZ offset = new XYZ(washbasinTotalDepth2, -washbasin_accessible_TotalWidth - washbasinTotalWidth2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin2.Id, offset);
                                womanRestroomElems.Remove(washbasin2); // 移除修改過的族群
                            }
                        }
                        // MRT_無障礙洗面盆(SinoBIM-第1版)
                        if (washbasin_accessible != null)
                        {
                            if (womanData.Type.Equals(1))
                            {
                                XYZ offset = new XYZ(washbasinTotalWidth1 + washbasin_accessible_TotalWidth / 2, -washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(2))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(3))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            womanRestroomElems.Remove(washbasin_accessible); // 移除修改過的族群
                        }
                    }
                }
                else if (womanData.Type.Equals(4) || womanData.Type.Equals(5) || womanData.Type.Equals(6) || womanData.Type.Equals(7))
                {
                    if (womanData.Rotate_id.Equals(1) || womanData.Rotate_id.Equals(2))
                    {
                        // toilet1：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet1 != null)
                        {
                            XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth, -toiletTotalDepth1, 0);
                            ElementTransformUtils.MoveElement(doc, toilet1.Id, offset);
                            womanRestroomElems.Remove(toilet1); // 移除修改過的族群
                        }
                        // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                        if (toilet_squat != null)
                        {
                            XYZ offset = new XYZ(mop_basin_TotalWidth, -toilet_squat_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet_squat.Id, offset);
                            womanRestroomElems.Remove(toilet_squat); // 移除修改過的族群
                        }
                        // MRT_拖布盆間(SinoBIM-第2版)
                        if (mop_basin != null)
                        {
                            XYZ offset = new XYZ(0, -mop_basin_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, mop_basin.Id, offset);
                            womanRestroomElems.Remove(mop_basin); // 移除修改過的族群
                        }
                        // washbasin1：MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin1 != null)
                        {
                            if (womanData.Type.Equals(4))
                            {
                                XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth + toiletTotalWidth1 + washbasin_accessible_TotalWidth, -washbasinTotalDepth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth + toiletTotalWidth1 + washbasinTotalDepth1, -washbasinTotalWidth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(6))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(toilet_squat_TotalWidth + toiletTotalWidth1 + washbasinTotalDepth1, -washbasinTotalWidth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(7))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(toilet_squat_TotalWidth + toiletTotalWidth1 + washbasinTotalDepth1, -washbasinTotalWidth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            womanRestroomElems.Remove(washbasin1); // 移除修改過的族群
                        }
                        // washbasin2：MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin2 != null)
                        {
                            if (womanData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin2.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin2.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - space, -washbasin_accessible_TotalWidth - washbasinTotalWidth2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin2.Id, offset);
                                womanRestroomElems.Remove(washbasin2); // 移除修改過的族群
                            }
                        }
                        // MRT_無障礙洗面盆(SinoBIM-第1版)
                        if (washbasin_accessible != null)
                        {
                            if (womanData.Type.Equals(4))
                            {
                                XYZ offset = new XYZ(mop_basin_TotalWidth + toilet_squat_TotalWidth + toiletTotalWidth1 + washbasin_accessible_TotalWidth / 2, -washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - washbasin_accessible_TotalDepth - space, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(6))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(7))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            womanRestroomElems.Remove(washbasin_accessible); // 移除修改過的族群
                        }
                    }
                    else if (womanData.Rotate_id.Equals(3) || womanData.Rotate_id.Equals(4))
                    {
                        // toilet1：MRT_廁所群組(一般坐式)(SinoBIM-第2版)
                        if (toilet1 != null)
                        {
                            XYZ offset = new XYZ(womanRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth1, -toiletTotalDepth1, 0);
                            ElementTransformUtils.MoveElement(doc, toilet1.Id, offset);
                            womanRestroomElems.Remove(toilet1); // 移除修改過的族群
                        }
                        // MRT_廁所群組(蹲式)(SinoBIM-第2版)
                        if (toilet_squat != null)
                        {
                            XYZ offset = new XYZ(womanRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth, -toilet_squat_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, toilet_squat.Id, offset);
                            womanRestroomElems.Remove(toilet_squat); // 移除修改過的族群
                        }
                        // MRT_拖布盆間(SinoBIM-第2版)
                        if (mop_basin != null)
                        {
                            XYZ offset = new XYZ(womanRestroomLength - mop_basin_TotalWidth, -mop_basin_TotalDepth, 0);
                            ElementTransformUtils.MoveElement(doc, mop_basin.Id, offset);
                            womanRestroomElems.Remove(mop_basin); // 移除修改過的族群
                        }
                        // washbasin1：MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin1 != null)
                        {
                            if (womanData.Type.Equals(4))
                            {
                                XYZ offset = new XYZ(womanRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth1 - washbasin_accessible_TotalWidth - washbasinTotalWidth1, -washbasinTotalDepth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth1 - washbasinTotalDepth1, 0, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(6))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(toilet_squat_TotalWidth + toiletTotalWidth1 + washbasinTotalDepth1, -washbasinTotalWidth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            else if (womanData.Type.Equals(7))
                            {
                                LocationPoint lp = washbasin1.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin1.Id, line, rotation);
                                XYZ offset = new XYZ(toilet_squat_TotalWidth + toiletTotalWidth1 + washbasinTotalDepth1, -washbasinTotalWidth1, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin1.Id, offset);
                            }
                            womanRestroomElems.Remove(washbasin1); // 移除修改過的族群
                        }
                        // washbasin2：MRT_洗手台群組(SinoBIM-第3版)
                        if (washbasin2 != null)
                        {
                            if (womanData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin2.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin2.Id, line, rotation);
                                XYZ offset = new XYZ(space + washbasinTotalDepth2, -washbasin_accessible_TotalWidth - washbasinTotalWidth2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin2.Id, offset);
                                womanRestroomElems.Remove(washbasin2); // 移除修改過的族群
                            }
                        }
                        // MRT_無障礙洗面盆(SinoBIM-第1版)
                        if (washbasin_accessible != null)
                        {
                            if (womanData.Type.Equals(4))
                            {
                                XYZ offset = new XYZ(womanRestroomLength - mop_basin_TotalWidth - toilet_squat_TotalWidth - toiletTotalWidth1 - washbasin_accessible_TotalWidth / 2, -washbasin_accessible_TotalDepth, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(5))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * 90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(space + washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(6))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            else if (womanData.Type.Equals(7))
                            {
                                LocationPoint lp = washbasin_accessible.Location as LocationPoint;
                                Line line = Line.CreateBound(lp.Point, new XYZ(lp.Point.X, lp.Point.Y, lp.Point.Z + 10));
                                double rotation = 2 * Math.PI / 360 * -90; // 轉換角度
                                ElementTransformUtils.RotateElement(doc, washbasin_accessible.Id, line, rotation);
                                XYZ offset = new XYZ(womanRestroomLength - washbasin_accessible_TotalDepth, -washbasin_accessible_TotalWidth / 2, 0);
                                ElementTransformUtils.MoveElement(doc, washbasin_accessible.Id, offset);
                            }
                            womanRestroomElems.Remove(washbasin_accessible); // 移除修改過的族群
                        }
                    }
                }
            }
        }
        // 修改親子廁所參數
        private void EditFamilyRestroom(Document doc, List<FamilyInstance> familyRestroomElems, JsonData jsonData)
        {
            foreach (FamilyData familyData in jsonData.FamilyDataList)
            {
                double familyRestroomLength = UnitUtils.ConvertToInternalUnits(familyData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double familyRestroomWidth = UnitUtils.ConvertToInternalUnits(familyData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度

                FamilyInstance toilet = familyRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.familyRestroom)).FirstOrDefault();
                // toilet：MRT_親子廁所(SinoBIM-第1版)
                if(toilet != null)
                {
                    double toiletDepth = toilet.LookupParameter("親子廁所深度").AsDouble();

                    // toilet：MRT_親子廁所(SinoBIM-第1版)
                    XYZ offset = new XYZ(0, -toiletDepth, 0);
                    ElementTransformUtils.MoveElement(doc, toilet.Id, offset);

                    // 移除修改過的族群
                    familyRestroomElems.Remove(toilet);
                }
            }
        }
        // 修改無障礙廁所參數
        private void EditAccessibleRestroom(Document doc, List<FamilyInstance> accessibleRestroomElems, JsonData jsonData)
        {
            foreach (AccessibleData accessibleData in jsonData.AccessibleDataList)
            {
                double accessibleRestroomLength = UnitUtils.ConvertToInternalUnits(accessibleData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double accessibleRestroomWidth = UnitUtils.ConvertToInternalUnits(accessibleData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度

                FamilyInstance toilet = accessibleRestroomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.accessibleRestroom)).FirstOrDefault();
                // toilet：MRT_無障礙廁所(SinoBIM-第1版)
                if (toilet != null)
                {
                    double toiletDepth = toilet.LookupParameter("無障礙廁所深度").AsDouble();

                    XYZ offset = new XYZ(0, -toiletDepth, 0);
                    ElementTransformUtils.MoveElement(doc, toilet.Id, offset);

                    // 移除修改過的族群
                    accessibleRestroomElems.Remove(toilet);
                }
            }
        }
        // 修改哺集乳室
        private void EditBreastfeedingRoom(Document doc, List<FamilyInstance> breastfeedingsRoomElems, JsonData jsonData)
        {
            foreach (BreastfeedingData breastfeedingData in jsonData.BreastfeedingDataList)
            {
                double breastfeedingsRoomLength = UnitUtils.ConvertToInternalUnits(breastfeedingData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double breastfeedingsRoomWidth = UnitUtils.ConvertToInternalUnits(breastfeedingData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度

                FamilyInstance breastfeeding = breastfeedingsRoomElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.breastfeeding)).FirstOrDefault();
                // toilet：MRT_哺集乳室(SinoBIM-第1版)
                if (breastfeeding != null)
                {
                    double breastfeedingDepth = breastfeeding.LookupParameter("哺集乳室深度").AsDouble();

                    XYZ offset = new XYZ(0, -breastfeedingDepth, 0);
                    ElementTransformUtils.MoveElement(doc, breastfeeding.Id, offset);

                    // 移除修改過的族群
                    breastfeedingsRoomElems.Remove(breastfeeding);
                }
            }
        }
        // 修改清潔人員休息室
        private void EditJanitorRooms(Document doc, List<FamilyInstance> janitorRoomsElems, JsonData jsonData)
        {
            foreach (JanitorRoomData janitorRoomData in jsonData.JanitorRoomDataList)
            {
                double janitorRoomLength = UnitUtils.ConvertToInternalUnits(janitorRoomData.Length, DisplayUnitType.DUT_METERS); // 廁所總長度
                double janitorRoomWidth = UnitUtils.ConvertToInternalUnits(janitorRoomData.Width, DisplayUnitType.DUT_METERS); // 廁所總寬度

                FamilyInstance janitor = janitorRoomsElems.Where(x => x.Symbol.FamilyName.Contains(RestroomGroup.janitorRoom)).FirstOrDefault();
                // toilet：MRT_清潔人員休息室(SinoBIM-第1版)
                if (janitor != null)
                {
                    double janitorDepth = janitor.LookupParameter("清潔人員休息室深度").AsDouble();

                    XYZ offset = new XYZ(0, -janitorDepth, 0);
                    ElementTransformUtils.MoveElement(doc, janitor.Id, offset);

                    // 移除修改過的族群
                    janitorRoomsElems.Remove(janitor);
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
