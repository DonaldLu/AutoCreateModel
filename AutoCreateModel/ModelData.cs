using System.Collections.Generic;

namespace AutoCreateModel
{
    public class ModelData
    {
        public double Width { get; set; }
        public double Length { get; set; }
        public int Type { get; set; }
        public int Side_id { get; set; }
        public int Level_id { get; set; }
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
        public double RestroomWoman_x { get; set; }
        public double RestroomWoman_y { get; set; }
        public int Toilet_Count { get; set; }
        public int Washbasin_Count { get; set; }
        public int Rotate_id { get; set; }
    }
    // Json
    public class JsonData
    {
        public List<AccessibleData> AccessibleDataList { get; set; } // 無障礙廁所
        public List<BreastfeedingData> BreastfeedingDataList { get; set; } // 哺集乳室
        public List<FamilyData> FamilyDataList { get; set; } // 親子廁所
        public List<JanitorRoomData> JanitorRoomDataList { get; set; } // 清潔人員休息室
        public List<ManData> ManDataList { get; set; } // 男廁
        public List<WomanData> WomanDataList { get; set; } // 女廁
    }
}
