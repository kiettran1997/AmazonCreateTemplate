using System.Collections.Generic;
using System.Web;

namespace CreateTemplateFile.Models
{
    public class InputData
    {
       
        public  string FeedProduct { get; set; }
        public  string RelationshipType { get; set; }
        public string BulletPoint1 { get; set; }
        public string BulletPoint2 { get; set; }
        public string BulletPoint3 { get; set; }
        public string BulletPoint4 { get; set; }
        public string BulletPoint5 { get; set; }
        public int ItemWidthSide1 { get; set; }
        public int ItemWidthSide2 { get; set; }
        public int ItemWidthSide3 { get; set; }
        public int ItemLengthHead1 { get; set; }
        public int ItemLengthHead2 { get; set; }
        public int ItemLengthHead3 { get; set; }

        public string GenericKeywords { get; set; }
        public string PurchasableOffer { get; set; }
        public string ItemShape { get; set; }
        public string ShippingTemplate { get; set; }
        public  string ProductDescription { get; set; }
        public string VariationTheme { get; set; }
        public string ItemType { get; set; }
        public int Quantity { get; set; }
    }
}