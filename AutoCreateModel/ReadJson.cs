using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;

namespace AutoCreateModel
{
    public class ReadJson
    {
        public static JsonData ReadJsonFile(string folderPath)
        {
            JsonData jsonData = new JsonData();

            // 獲取文件列表
            string[] files = Directory.GetFiles(folderPath);
            foreach (string file in files)
            {
                if (File.Exists(file))
                {
                    string[] allLines = File.ReadAllLines(file);
                    foreach (string line in allLines)
                    {
                        if(line != "")
                        {
                            switch (Path.GetFileName(file))
                            {
                                case "AccessibleData.json": // 無障礙廁所
                                    List<AccessibleData> accessibleData = JsonConvert.DeserializeObject<List<AccessibleData>>(line);
                                    jsonData.AccessibleDataList.Add(accessibleData[0]);
                                    break;
                                case "BreastfeedingData.json": // 哺集乳室
                                    List<BreastfeedingData> breastfeedingData = JsonConvert.DeserializeObject<List<BreastfeedingData>>(line);
                                    jsonData.BreastfeedingDataList.Add(breastfeedingData[0]);
                                    break;
                                case "FamilyData.json": // 親子廁所
                                    List<FamilyData> familyData = JsonConvert.DeserializeObject<List<FamilyData>>(line);
                                    jsonData.FamilyDataList.Add(familyData[0]);
                                    break;
                                case "JanitorRoomData.json": // 清潔人員休息室
                                    List<JanitorRoomData> janitorRoomData = JsonConvert.DeserializeObject<List<JanitorRoomData>>(line);
                                    jsonData.JanitorRoomDataList.Add(janitorRoomData[0]);
                                    break;
                                case "ManData.json": // 男廁
                                    List<ManData> manData = JsonConvert.DeserializeObject<List<ManData>>(line);
                                    jsonData.ManDataList.Add(manData[0]);
                                    break;
                                case "WomanData.json": // 女廁
                                    List<WomanData> womanData = JsonConvert.DeserializeObject<List<WomanData>>(line);
                                    jsonData.WomanDataList.Add(womanData[0]);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }

            return jsonData;
        }
    }
}