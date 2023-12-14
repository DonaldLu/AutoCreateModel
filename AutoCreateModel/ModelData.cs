using System.Collections.Generic;

namespace AutoCreateModel
{
    public class ModelData
    {
        public int id { get; set; } // 設施id
        public double Width { get; set; } // 寬度
        public double Length { get; set; } // 長度
        public int Type { get; set; } // 類型
        public int Side_id { get; set; }
        public int Level_id { get; set; } // 樓層
        public int Restroom_id { get; set; } // 總體廁所id
        public int Plan_id { get; set; } // 專案id
    }
    // 無障礙廁所
    public class AccessibleData : ModelData
    {
        public double RestroomAccessible_x { get; set; }
        public double RestroomAccessible_y { get; set; }
    }
    // 哺集乳室
    public class BreastfeedingData : ModelData
    {
        public double BreastfeedingRoom_x { get; set; }
        public double BreastfeedingRoom_y { get; set; }
    }
    // 親子廁所
    public class FamilyData : ModelData
    {
        public double RestroomFamily_x { get; set; }
        public double RestroomFamily_y { get; set; }
    }
    // 清潔人員休息室
    public class JanitorRoomData : ModelData
    {
        public double JanitorRoom_x { get; set; }
        public double JanitorRoom_y { get; set; }
    }
    // 男廁
    public class ManData : ModelData
    {
        public double Aisle_Width { get; set; } // 走道寬度
        public double Aisle_Length { get; set; } // 走道長度
        public int Mopbasin_id { get; set; } // 是否有拖布盆
        public int AccessibleWashbasin_id { get; set; } // 是否有無障礙洗面盆
        public double RestroomMan_x { get; set; }
        public double RestroomMan_y { get; set; }
        public int Toilet_Count { get; set; }
        public int Washbasin_Count { get; set; }
        public int Urinal_Count { get; set; }
        public int Rotate_id { get; set; }
    }
    // 女廁
    public class WomanData : ModelData
    {
        public double Aisle_Width { get; set; } // 走道寬度
        public double Aisle_Length { get; set; } // 走道長度
        public int Mopbasin_id { get; set; } // 是否有拖布盆
        public int AccessibleWashbasin_id { get; set; } // 是否有無障礙洗面盆
        public double RestroomWoman_x { get; set; }
        public double RestroomWoman_y { get; set; }
        public int Toilet_Count { get; set; }
        public int Washbasin_Count { get; set; }
        public int Rotate_id { get; set; }
    }
    // Json
    public class JsonData
    {
        public List<AccessibleData> AccessibleDataList = new List<AccessibleData>(); // 無障礙廁所
        public List<BreastfeedingData> BreastfeedingDataList = new List<BreastfeedingData>(); // 哺集乳室
        public List<FamilyData> FamilyDataList = new List<FamilyData>(); // 親子廁所
        public List<JanitorRoomData> JanitorRoomDataList = new List<JanitorRoomData>(); // 清潔人員休息室
        public List<ManData> ManDataList = new List<ManData>(); // 男廁
        public List<WomanData> WomanDataList = new List<WomanData>(); // 女廁
    }
}
