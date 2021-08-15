using System.Collections.Generic;
using System.Web;

namespace CreateTemplateFile.Models
{
    public class InputData
    {
        public string BulletPoint1 { get; set; }
        public string BulletPoint2 { get; set; }
        public string BulletPoint3 { get; set; }
        public string BulletPoint4 { get; set; }
        public string BulletPoint5 { get; set; }
        public string GenericKeywords { get; set; }
        public string PurchasableOffer { get; set; }
        public string ShippingTemplate { get; set; }
        public string ProductDescription { get; set; }
        public string ItemType { get; set; }
        public int Quantity { get; set; }
    }
}