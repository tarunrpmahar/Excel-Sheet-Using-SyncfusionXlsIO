using System.ComponentModel.DataAnnotations;

namespace ExcelSheetUsingSyncfusionXlsIO.Models
{
    public class Product
    {
        [Key]
        public string ProductCode { get; set; }
        public string Title { get; set; }
        public string Group { get; set; }
        public string Line { get; set; }
        public string Class { get; set; }
        public string UnitofMeasure { get; set; }
        public float? UnitPrice { get; set; }
        public string Plant { get; set; }
    }
}
