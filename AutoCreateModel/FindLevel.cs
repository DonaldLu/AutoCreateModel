using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoCreateModel
{
    internal class FindLevel
    {
        public class LevelElevation
        {
            public string Name { get;set; }
            public Level Level { get; set; }
            public double Height { get; set; }
            public double Elevation { get; set; }
        }
        // 找到當前視圖的Level相關資訊
        public Tuple<List<LevelElevation>, LevelElevation, double> FindDocViewLevel(Document doc)
        {
            // 查詢所有Level的高程並排序
            List<Level> levels = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().Cast<Level>().ToList();
            List<LevelElevation> levelElevList = new List<LevelElevation>();
            foreach (Level level in levels)
            {
                LevelElevation levelElevation = new LevelElevation();
                levelElevation.Name = level.Name;
                levelElevation.Level = level;
                levelElevation.Height = level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsDouble();
                levelElevation.Elevation = Convert.ToDouble(level.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString());
                levelElevList.Add(levelElevation);
            }
            levelElevList = (from x in levelElevList
                             select x).OrderBy(x => x.Elevation).ToList();
            double startElev = 0.0;
            double endElev = 0.0;
            double floorHeight = 10;
            // 找到當前樓層
            LevelElevation viewLevel = (from x in levelElevList
                                        where x.Level.Id.Equals(doc.ActiveView.GenLevel.Id)
                                        select x).FirstOrDefault();
            int leCount = levelElevList.IndexOf(viewLevel);
            // 查詢當前樓層與上一樓層的高度, 製作火源高度
            if (levelElevList.Count >= 2)
            {
                if (leCount < levelElevList.Count)
                {
                    startElev = levelElevList[leCount].Elevation;
                    endElev = levelElevList[leCount + 1].Elevation;
                    floorHeight = endElev - startElev;
                }
                else
                {
                    startElev = levelElevList[leCount].Elevation;
                    endElev = levelElevList[leCount - 1].Elevation;
                    floorHeight = startElev - endElev;
                }
            }

            Tuple<List<LevelElevation>, LevelElevation, double> multiValue = Tuple.Create(levelElevList, viewLevel, floorHeight);

            return multiValue;
        }
        // 計算每層Level的高程差
        public List<LevelElevation> LevelElevationCalcul(List<LevelElevation> levelElevList)
        {
            List<LevelElevation> newlevelElevList = new List<LevelElevation>();
            for (int i = 0; i < levelElevList.Count(); i++)
            {
                LevelElevation newlevelElev = new LevelElevation();
                double startElev = 0.0;
                double endElev = 0.0;
                double floorHeight = 9999;
                // 找到當前樓層
                LevelElevation viewLevel = (from x in levelElevList
                                            where x.Level.Id.Equals(levelElevList[i].Level.Id)
                                            select x).FirstOrDefault();
                int leCount = levelElevList.IndexOf(viewLevel);
                newlevelElev.Level = viewLevel.Level; // Level
                newlevelElev.Name = viewLevel.Name; // 名稱
                newlevelElev.Elevation = viewLevel.Elevation; // 高程
                newlevelElev.Height = floorHeight / 1000; // 與上一樓層高程差
                if (i < levelElevList.Count() - 1)
                {
                    // 查詢當前樓層與上一樓層的高度
                    if (levelElevList.Count >= 2)
                    {
                        if (leCount < levelElevList.Count)
                        {
                            startElev = levelElevList[leCount].Elevation;
                            endElev = levelElevList[leCount + 1].Elevation;
                            floorHeight = endElev - startElev;
                            newlevelElev.Height = floorHeight / 1000; // 與上一樓層高程差
                        }
                        else
                        {
                            startElev = levelElevList[leCount].Elevation;
                            endElev = levelElevList[leCount - 1].Elevation;
                            floorHeight = startElev - endElev;
                            newlevelElev.Height = floorHeight / 1000; // 與上一樓層高程差
                        }
                    }
                }
                newlevelElevList.Add(newlevelElev);
            }

            return newlevelElevList;
        }
    }
}
