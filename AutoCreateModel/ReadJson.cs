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
                using (StreamReader r = new StreamReader(file))
                {
                    string json = r.ReadToEnd();
                    if (File.Exists(file) && Path.GetFileName(file).Equals("AccessibleData.json"))
                    {
                        jsonData.AccessibleDataList = JsonConvert.DeserializeObject<List<AccessibleData>>(json);
                    }
                    else if (File.Exists(file) && Path.GetFileName(file).Equals("BreastfeedingData.json"))
                    {
                        jsonData.BreastfeedingDataList = JsonConvert.DeserializeObject<List<BreastfeedingData>>(json);
                    }
                    else if (File.Exists(file) && Path.GetFileName(file).Equals("FamilyData.json"))
                    {
                        jsonData.FamilyDataList = JsonConvert.DeserializeObject<List<FamilyData>>(json);
                    }
                    else if (File.Exists(file) && Path.GetFileName(file).Equals("JanitorRoomData.json"))
                    {
                        jsonData.JanitorRoomDataList = JsonConvert.DeserializeObject<List<JanitorRoomData>>(json);
                    }
                    else if (File.Exists(file) && Path.GetFileName(file).Equals("ManData.json"))
                    {
                        jsonData.ManDataList = JsonConvert.DeserializeObject<List<ManData>>(json);
                    }
                    else if (File.Exists(file) && Path.GetFileName(file).Equals("WomanData.json"))
                    {
                        jsonData.WomanDataList = JsonConvert.DeserializeObject<List<WomanData>>(json);
                    }
                }
            }

            return jsonData;
        }
    }
}