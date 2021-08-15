using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;

namespace AmazonCreateTemplateFile
{
    class Program
    {
        static void Main(string[] args)
        {

            var mapCategory = new ExcelPackage();
            mapCategory.Load(new FileStream("MapCategory.xlsx", FileMode.Open));
            var sheetMap = mapCategory.Workbook.Worksheets.FirstOrDefault();

            #region
            var excelTemplate = new ExcelPackage();
            var sheetTemplate = excelTemplate.Workbook.Worksheets.Add("Template");
            sheetTemplate.Cells["A1"].Value = "TemplateType=fptcustom";
            sheetTemplate.Cells["A2"].Value = "Product Type";
            sheetTemplate.Cells["A3"].Value = "feed_product_type";
            sheetTemplate.Cells["B1"].Value = "Version=2021.0708";
            sheetTemplate.Cells["B2"].Value = "Seller SKU";
            sheetTemplate.Cells["B3"].Value = "item_sku";
            sheetTemplate.Cells["C1"].Value = "TemplateSignature=U0hJUlQ=";
            sheetTemplate.Cells["C2"].Value = "Brand Name";
            sheetTemplate.Cells["C3"].Value = "brand_name";
            sheetTemplate.Cells["D1"].Value = "settings=contentLanguageTag=en_US&feedType=113&headerLanguageTag=en_US&metadataVersion=MatprodVkxBUHJvZC0xMDg4&primaryMarketplaceId=amzn1.mp.o.ATVPDKIKX0DER&templateIdentifier=26b31110-6e13-4d4a-8779-2da477330af7&timestamp=2021-07-08T03%3A55%3A45.513Z";
            sheetTemplate.Cells["D2"].Value = "Product Name";
            sheetTemplate.Cells["D3"].Value = "item_name";
            sheetTemplate.Cells["E1"].Value = "Use ENGLISH to fill this template.The top 3 rows are for Amazon.com use only. Do not modify or delete the top 3 rows.";
            sheetTemplate.Cells["E2"].Value = "Product ID";
            sheetTemplate.Cells["E3"].Value = "external_product_id";
            //sheet.Cells["F1"].Value = "";
            sheetTemplate.Cells["F2"].Value = "Product ID Type";
            sheetTemplate.Cells["F3"].Value = "external_product_id_type";
            //sheet.Cells["G1"].Value = "";
            sheetTemplate.Cells["G2"].Value = "Item Type Keyword";
            sheetTemplate.Cells["G3"].Value = "item_type";
            //sheet.Cells["H1"].Value = "";
            sheetTemplate.Cells["H2"].Value = "Outer Material Type";
            sheetTemplate.Cells["H3"].Value = "outer_material_type1";
            //sheet.Cells["I1"].Value = "";
            sheetTemplate.Cells["I2"].Value = "Outer Material Type";
            sheetTemplate.Cells["I3"].Value = "outer_material_type2";
            //sheet.Cells["J1"].Value = "";
            sheetTemplate.Cells["J2"].Value = "Outer Material Type";
            sheetTemplate.Cells["J3"].Value = "outer_material_type3";
            //sheetTemplate.Cells["K1"].Value = "";
            sheetTemplate.Cells["K2"].Value = "Outer Material Type";
            sheetTemplate.Cells["K3"].Value = "outer_material_type4";
            //sheet.Cells["L1"].Value = "";
            sheetTemplate.Cells["L2"].Value = "Outer Material Type";
            sheetTemplate.Cells["L3"].Value = "outer_material_type5";
            //sheet.Cells["M1"].Value = "";
            sheetTemplate.Cells["M2"].Value = "Material Composition";
            sheetTemplate.Cells["M3"].Value = "material_composition1";
            //sheet.Cells["N1"].Value = "";
            sheetTemplate.Cells["N2"].Value = "Material Composition";
            sheetTemplate.Cells["N3"].Value = "material_composition2";
            //sheet.Cells["O1"].Value = "";
            sheetTemplate.Cells["O2"].Value = "Material Composition";
            sheetTemplate.Cells["O3"].Value = "material_composition3";
            //sheet.Cells["P1"].Value = "";
            sheetTemplate.Cells["P2"].Value = "Material Composition";
            sheetTemplate.Cells["P3"].Value = "material_composition4";
            //sheet.Cells["Q1"].Value = "";
            sheetTemplate.Cells["Q2"].Value = "Material Composition";
            sheetTemplate.Cells["Q3"].Value = "material_composition5";
            //sheet.Cells["R1"].Value = "";
            sheetTemplate.Cells["R2"].Value = "Material Composition";
            sheetTemplate.Cells["R3"].Value = "material_composition6";
            //sheet.Cells["S1"].Value = "";
            sheetTemplate.Cells["S2"].Value = "Material Composition";
            sheetTemplate.Cells["S3"].Value = "material_composition7";
            //sheetTemplate.Cells["T1"].Value = "";
            sheetTemplate.Cells["T2"].Value = "Material Composition";
            sheetTemplate.Cells["T3"].Value = "material_composition8";
            //sheet.Cells["U1"].Value = "";
            sheetTemplate.Cells["U2"].Value = "Material Composition";
            sheetTemplate.Cells["U3"].Value = "material_composition9";
            //sheet.Cells["V1"].Value = "";
            sheetTemplate.Cells["V2"].Value = "Material Composition";
            sheetTemplate.Cells["V3"].Value = "material_composition10";
            //sheet.Cells["W1"].Value = "";
            sheetTemplate.Cells["W2"].Value = "Department";
            sheetTemplate.Cells["W3"].Value = "department_name";
            //sheetTemplate.Cells["X1"].Value = "";
            sheetTemplate.Cells["X2"].Value = "Is Adult Product";
            sheetTemplate.Cells["X3"].Value = "is_adult_product";
            //sheet.Cells["Y1"].Value = "";
            sheetTemplate.Cells["Y2"].Value = "Standard Price";
            sheetTemplate.Cells["Y3"].Value = "standard_price";
            //sheet.Cells["Z1"].Value = "";
            sheetTemplate.Cells["Z2"].Value = "Quantity";
            sheetTemplate.Cells["Z3"].Value = "quantity";
            //sheet.Cells["AA1"].Value = "";
            sheetTemplate.Cells["AA2"].Value = "Main Image URL";
            sheetTemplate.Cells["AA3"].Value = "main_image_url";
            //sheet.Cells["AB1"].Value = "";
            sheetTemplate.Cells["AB2"].Value = "Target Gender";
            sheetTemplate.Cells["AB3"].Value = "target_gender";
            //sheet.Cells["AC1"].Value = "";
            sheetTemplate.Cells["AC2"].Value = "Age Range Description";
            sheetTemplate.Cells["AC3"].Value = "age_range_description";
            //sheet.Cells["AD1"].Value = "";
            sheetTemplate.Cells["AD2"].Value = "Shirt Size System";
            sheetTemplate.Cells["AD3"].Value = "shirt_size_system";
            //sheetTemplate.Cells["AE1"].Value = "";
            sheetTemplate.Cells["AE2"].Value = "Shirt Size Class";
            sheetTemplate.Cells["AE3"].Value = "shirt_size_class";
            //sheetTemplate.Cells["AF1"].Value = "";
            sheetTemplate.Cells["AF2"].Value = "Shirt Size Value";
            sheetTemplate.Cells["AF3"].Value = "shirt_size";
            //sheetTemplate.Cells["AG1"].Value = "";
            sheetTemplate.Cells["AG2"].Value = "Shirt Size To Range";
            sheetTemplate.Cells["AG3"].Value = "shirt_size_to";
            //sheetTemplate.Cells["AH1"].Value = "";
            sheetTemplate.Cells["AH2"].Value = "Neck Size Value";
            sheetTemplate.Cells["AH3"].Value = "shirt_neck_size";
            //sheetTemplate.Cells["AI1"].Value = "";
            sheetTemplate.Cells["AI2"].Value = "Neck Size To Value";
            sheetTemplate.Cells["AI3"].Value = "shirt_neck_size_to";
            //sheetTemplate.Cells["AJ1"].Value = "";
            sheetTemplate.Cells["AJ2"].Value = "Sleeve Length Value";
            sheetTemplate.Cells["AJ3"].Value = "shirt_sleeve_length";
            //sheetTemplate.Cells["AK1"].Value = "";
            sheetTemplate.Cells["AK2"].Value = "Sleeve Length To Value";
            sheetTemplate.Cells["AK3"].Value = "shirt_sleeve_length_to";
            //sheetTemplate.Cells["AL1"].Value = "";
            sheetTemplate.Cells["AL2"].Value = "Shirt Body Type";
            sheetTemplate.Cells["AL3"].Value = "shirt_body_type";
            //sheetTemplate.Cells["AM1"].Value = "";
            sheetTemplate.Cells["AM2"].Value = "Shirt Height Type";
            sheetTemplate.Cells["AM3"].Value = "shirt_height_type";
            sheetTemplate.Cells["AN1"].Value = "Images";
            sheetTemplate.Cells["AN2"].Value = "Other Image URL1";
            sheetTemplate.Cells["AN3"].Value = "other_image_url1";
            //sheet.Cells["ANO1"].Value = "";
            sheetTemplate.Cells["AO2"].Value = "Other Image URL2";
            sheetTemplate.Cells["AO3"].Value = "other_image_url2";
            //sheet.Cells["AP1"].Value = "";
            sheetTemplate.Cells["AP2"].Value = "Other Image URL3";
            sheetTemplate.Cells["AP3"].Value = "other_image_url3";
            //sheet.Cells["AQ1"].Value = "";
            sheetTemplate.Cells["AQ2"].Value = "Other Image URL4";
            sheetTemplate.Cells["AQ3"].Value = "other_image_url4";
            //sheet.Cells["AR1"].Value = "";
            sheetTemplate.Cells["AR2"].Value = "Other Image URL5";
            sheetTemplate.Cells["AR3"].Value = "other_image_url5";
            //sheet.Cells["AS1"].Value = "";
            sheetTemplate.Cells["AS2"].Value = "Other Image URL6";
            sheetTemplate.Cells["AS3"].Value = "other_image_url6";
            //sheet.Cells["AT1"].Value = "";
            sheetTemplate.Cells["AT2"].Value = "Other Image URL7";
            sheetTemplate.Cells["AT3"].Value = "other_image_url7";
            //sheet.Cells["AU1"].Value = "";
            sheetTemplate.Cells["AU2"].Value = "Other Image URL8";
            sheetTemplate.Cells["AU3"].Value = "other_image_url8";
            //sheet.Cells["AV1"].Value = "";
            sheetTemplate.Cells["AV2"].Value = "Other Image URL";
            sheetTemplate.Cells["AV3"].Value = "other_image_url";
            // sheet.Cells["AW1"].Value = "";
            sheetTemplate.Cells["AW2"].Value = "Swatch Image URL";
            sheetTemplate.Cells["AW3"].Value = "swatch_image_url";
            sheetTemplate.Cells["AX1"].Value = "Variation";
            sheetTemplate.Cells["AX2"].Value = "Parentage";
            sheetTemplate.Cells["AX3"].Value = "parent_child";
            // sheet.Cells["AY1"].Value = "";
            sheetTemplate.Cells["AY2"].Value = "Parent SKU";
            sheetTemplate.Cells["AY3"].Value = "parent_sku";
            // sheet.Cells["AZ1"].Value = "";
            sheetTemplate.Cells["AZ2"].Value = "Relationship Type";
            sheetTemplate.Cells["AZ3"].Value = "relationship_type";
            //sheet.Cells["BA1"].Value = "";
            sheetTemplate.Cells["BA2"].Value = "Variation Theme";
            sheetTemplate.Cells["BA3"].Value = "variation_theme";
            sheetTemplate.Cells["BB1"].Value = "Basic";
            sheetTemplate.Cells["BB2"].Value = "Update Delete";
            sheetTemplate.Cells["BB3"].Value = "update_delete";
            //sheet.Cells["BC1"].Value = "";
            sheetTemplate.Cells["BC2"].Value = "Product Description";
            sheetTemplate.Cells["BC3"].Value = "product_description";
            //sheet.Cells["BD1"].Value = "";
            sheetTemplate.Cells["BD2"].Value = "Closure Type";
            sheetTemplate.Cells["BD3"].Value = "closure_type";
            //sheet.Cells["BE1"].Value = "";
            sheetTemplate.Cells["BE2"].Value = "Style Number";
            sheetTemplate.Cells["BE3"].Value = "model";
            //sheet.Cells["BF1"].Value = "";
            sheetTemplate.Cells["BF2"].Value = "Model Name";
            sheetTemplate.Cells["BF3"].Value = "model_name";
            //sheet.Cells["BG1"].Value = "";
            sheetTemplate.Cells["BG2"].Value = "Manufacturer Part Number";
            sheetTemplate.Cells["BG3"].Value = "part_number";
            //sheet.Cells["BH1"].Value = "";
            sheetTemplate.Cells["BH2"].Value = "Manufacturer";
            sheetTemplate.Cells["BH3"].Value = "manufacturer";
            //sheet.Cells["BI1"].Value = "";
            sheetTemplate.Cells["BI2"].Value = "Product Care Instructions";
            sheetTemplate.Cells["BI3"].Value = "care_instructions";
            sheetTemplate.Cells["BJ1"].Value = "Discovery";
            sheetTemplate.Cells["BJ2"].Value = "Key Product Features";
            sheetTemplate.Cells["BJ3"].Value = "bullet_point1";
            //sheetTemplate.Cells["BK1"].Value = "";
            sheetTemplate.Cells["BK2"].Value = "Key Product Features";
            sheetTemplate.Cells["BK3"].Value = "bullet_point2";
            //sheet.Cells["BL1"].Value = "";
            sheetTemplate.Cells["BL2"].Value = "Key Product Features";
            sheetTemplate.Cells["BL3"].Value = "bullet_point3";
            //sheet.Cells["BM1"].Value = "";
            sheetTemplate.Cells["BM2"].Value = "Key Product Features";
            sheetTemplate.Cells["BM3"].Value = "bullet_point4";
            //sheet.Cells["BN1"].Value = "";
            sheetTemplate.Cells["BN2"].Value = "Key Product Features";
            sheetTemplate.Cells["BN3"].Value = "bullet_point5";
            //sheet.Cells["BO1"].Value = "";
            sheetTemplate.Cells["BO2"].Value = "Search Terms";
            sheetTemplate.Cells["BO3"].Value = "generic_keywords";
            //sheet.Cells["BP1"].Value = "";
            sheetTemplate.Cells["BP2"].Value = "Belt Style";
            sheetTemplate.Cells["BP3"].Value = "belt_style";
            //sheet.Cells["BQ1"].Value = "";
            sheetTemplate.Cells["BQ2"].Value = "Collar Type";
            sheetTemplate.Cells["BQ3"].Value = "collar_style";
            //sheet.Cells["BR1"].Value = "";
            sheetTemplate.Cells["BR2"].Value = "Color";
            sheetTemplate.Cells["BR3"].Value = "color_name";
            //sheet.Cells["BS1"].Value = "";
            sheetTemplate.Cells["BS2"].Value = "Color Map";
            sheetTemplate.Cells["BS3"].Value = "color_map";
            //sheet.Cells["BT1"].Value = "";
            sheetTemplate.Cells["BT2"].Value = "Control Type";
            sheetTemplate.Cells["BT3"].Value = "control_type";
            //sheet.Cells["BU1"].Value = "";
            sheetTemplate.Cells["BU2"].Value = "Fit Type";
            sheetTemplate.Cells["BU3"].Value = "fit_type";
            //sheet.Cells["BV1"].Value = "";
            sheetTemplate.Cells["BV2"].Value = "Country/Region as Labeled";
            sheetTemplate.Cells["BV3"].Value = "country_as_labeled";
            // sheet.Cells["BW1"].Value = "";
            sheetTemplate.Cells["BW2"].Value = "Fur Description";
            sheetTemplate.Cells["BW3"].Value = "fur_description";
            // sheet.Cells["BX1"].Value = "";
            sheetTemplate.Cells["BX2"].Value = "NeckStyle";
            sheetTemplate.Cells["BX3"].Value = "neck_style";
            // sheet.Cells["BY1"].Value = "";
            sheetTemplate.Cells["BY2"].Value = "Pattern Style";
            sheetTemplate.Cells["BY3"].Value = "pattern_type";
            // sheet.Cells["BZ1"].Value = "";
            sheetTemplate.Cells["BZ2"].Value = "Pocket Description";
            sheetTemplate.Cells["BZ3"].Value = "pocket_description";
            //sheet.Cells["CA1"].Value = "";
            sheetTemplate.Cells["CA2"].Value = "Size";
            sheetTemplate.Cells["CA3"].Value = "size_name";
            //sheet.Cells["CB1"].Value = "";
            sheetTemplate.Cells["CB2"].Value = "Special Size Type";
            sheetTemplate.Cells["CB3"].Value = "special_size_type";
            //sheet.Cells["CC1"].Value = "";
            sheetTemplate.Cells["CC2"].Value = "Additional Features";
            sheetTemplate.Cells["CC3"].Value = "special_features1";
            //sheet.Cells["CD1"].Value = "";
            sheetTemplate.Cells["CD2"].Value = "Additional Features";
            sheetTemplate.Cells["CD3"].Value = "special_features2";
            //sheet.Cells["CE1"].Value = "";
            sheetTemplate.Cells["CE2"].Value = "Additional Features";
            sheetTemplate.Cells["CE3"].Value = "special_features3";
            //sheet.Cells["CF1"].Value = "";
            sheetTemplate.Cells["CF2"].Value = "Additional Features";
            sheetTemplate.Cells["CF3"].Value = "special_features4";
            //sheet.Cells["CG1"].Value = "";
            sheetTemplate.Cells["CG2"].Value = "Additional Features";
            sheetTemplate.Cells["CG3"].Value = "special_features5";
            //sheetTemplate.Cells["CH1"].Value = "";
            sheetTemplate.Cells["CH2"].Value = "Style";
            sheetTemplate.Cells["CH3"].Value = "style_name";
            //sheet.Cells["CI1"].Value = "";
            sheetTemplate.Cells["CI2"].Value = "theme";
            sheetTemplate.Cells["CI3"].Value = "theme";
            //sheet.Cells["CJ1"].Value = "";
            sheetTemplate.Cells["CJ2"].Value = "Top Style";
            sheetTemplate.Cells["CJ3"].Value = "top_style";
            //sheet.Cells["CK1"].Value = "";
            sheetTemplate.Cells["CK2"].Value = "Water Resistance Level";
            sheetTemplate.Cells["CK3"].Value = "water_resistance_level";
            //sheet.Cells["CL1"].Value = "";
            sheetTemplate.Cells["CL2"].Value = "Is Autographed";
            sheetTemplate.Cells["CL3"].Value = "is_autographed";
            //sheet.Cells["CM1"].Value = "";
            sheetTemplate.Cells["CM2"].Value = "Item Type Name";
            sheetTemplate.Cells["CM3"].Value = "item_type_name";
            //sheet.Cells["CN1"].Value = "";
            sheetTemplate.Cells["CN2"].Value = "Occasion Type";
            sheetTemplate.Cells["CN3"].Value = "occasion_type1";
            //sheet.Cells["CO1"].Value = "";
            sheetTemplate.Cells["CO2"].Value = "Occasion Type";
            sheetTemplate.Cells["CO3"].Value = "occasion_type2";
            //sheet.Cells["CP1"].Value = "";
            sheetTemplate.Cells["CP2"].Value = "Occasion Type";
            sheetTemplate.Cells["CP3"].Value = "occasion_type3";
            //sheetTemplate.Cells["CQ1"].Value = "";
            sheetTemplate.Cells["CQ2"].Value = "Occasion Type";
            sheetTemplate.Cells["CQ3"].Value = "occasion_type4";
            //sheet.Cells["CR1"].Value = "";
            sheetTemplate.Cells["CR2"].Value = "Occasion Type";
            sheetTemplate.Cells["CR3"].Value = "occasion_type5";
            //sheet.Cells["CS1"].Value = "";
            sheetTemplate.Cells["CS2"].Value = "Occasion Type";
            sheetTemplate.Cells["CS3"].Value = "occasion_type6";
            //sheet.Cells["CT1"].Value = "";
            sheetTemplate.Cells["CT2"].Value = "Occasion Type";
            sheetTemplate.Cells["CT3"].Value = "occasion_type7";
            //sheet.Cells["CU1"].Value = "";
            sheetTemplate.Cells["CU2"].Value = "Occasion Type";
            sheetTemplate.Cells["CU3"].Value = "occasion_type8";
            //sheet.Cells["CV1"].Value = "";
            sheetTemplate.Cells["CV2"].Value = "Occasion Type";
            sheetTemplate.Cells["CV3"].Value = "occasion_type9";
            // sheet.Cells["CW1"].Value = "";
            sheetTemplate.Cells["CW2"].Value = "Occasion Type";
            sheetTemplate.Cells["CW3"].Value = "occasion_type10";
            // sheet.Cells["CX1"].Value = "";
            sheetTemplate.Cells["CX2"].Value = "Occasion Type";
            sheetTemplate.Cells["CX3"].Value = "occasion_type11";
            // sheet.Cells["CY1"].Value = "";
            sheetTemplate.Cells["CY2"].Value = "Occasion Type";
            sheetTemplate.Cells["CY3"].Value = "occasion_type12";
            // sheet.Cells["CZ1"].Value = "";
            sheetTemplate.Cells["CZ2"].Value = "Occasion Type";
            sheetTemplate.Cells["CZ3"].Value = "occasion_type13";
            //sheet.Cells["DA1"].Value = "";
            sheetTemplate.Cells["DA2"].Value = "Occasion Type";
            sheetTemplate.Cells["DA3"].Value = "occasion_type14";
            //sheet.Cells["DB1"].Value = "";
            sheetTemplate.Cells["DB2"].Value = "Occasion Type";
            sheetTemplate.Cells["DB3"].Value = "occasion_type15";
            //sheet.Cells["DC1"].Value = "";
            sheetTemplate.Cells["DC2"].Value = "Occasion Type";
            sheetTemplate.Cells["DC3"].Value = "occasion_type15";
            //sheet.Cells["DD1"].Value = "";
            sheetTemplate.Cells["DD2"].Value = "Occasion Type";
            sheetTemplate.Cells["DD3"].Value = "occasion_type17";
            //sheet.Cells["DE1"].Value = "";
            sheetTemplate.Cells["DE2"].Value = "Occasion Type";
            sheetTemplate.Cells["DE3"].Value = "occasion_type18";
            //sheet.Cells["DF1"].Value = "";
            sheetTemplate.Cells["DF2"].Value = "Occasion Type";
            sheetTemplate.Cells["DF3"].Value = "occasion_type19";
            //sheet.Cells["DG1"].Value = "";
            sheetTemplate.Cells["DG2"].Value = "Occasion Type";
            sheetTemplate.Cells["DG3"].Value = "occasion_type20";
            //sheet.Cells["DH1"].Value = "";
            sheetTemplate.Cells["DH2"].Value = "Occasion Type";
            sheetTemplate.Cells["DH3"].Value = "occasion_type21";
            //sheet.Cells["DI1"].Value = "";
            sheetTemplate.Cells["DI2"].Value = "Occasion Type";
            sheetTemplate.Cells["DI3"].Value = "occasion_type22";
            //sheet.Cells["DJ1"].Value = "";
            sheetTemplate.Cells["DJ2"].Value = "Occasion Type";
            sheetTemplate.Cells["DJ3"].Value = "occasion_type23";
            //sheet.Cells["DK1"].Value = "";
            sheetTemplate.Cells["DK2"].Value = "Occasion Type";
            sheetTemplate.Cells["DK3"].Value = "occasion_type24";
            //sheet.Cells["DL1"].Value = "";
            sheetTemplate.Cells["DL2"].Value = "Occasion Type";
            sheetTemplate.Cells["DL3"].Value = "occasion_type25";
            //sheet.Cells["DM1"].Value = "";
            sheetTemplate.Cells["DM2"].Value = "Occasion Type";
            sheetTemplate.Cells["DM3"].Value = "occasion_type26";
            //sheet.Cells["DN1"].Value = "";
            sheetTemplate.Cells["DN2"].Value = "Occasion Type";
            sheetTemplate.Cells["DN3"].Value = "occasion_type27";
            //sheet.Cells["DO1"].Value = "";
            sheetTemplate.Cells["DO2"].Value = "Sport Type";
            sheetTemplate.Cells["DO3"].Value = "sport_type1";
            //sheet.Cells["DP1"].Value = "";
            sheetTemplate.Cells["DP2"].Value = "Sport Type";
            sheetTemplate.Cells["DP3"].Value = "sport_type2";
            //sheet.Cells["DQ1"].Value = "";
            sheetTemplate.Cells["DQ2"].Value = "Athlete";
            sheetTemplate.Cells["DQ3"].Value = "athlete";
            //sheet.Cells["DR1"].Value = "";
            sheetTemplate.Cells["DR2"].Value = "Team Name";
            sheetTemplate.Cells["DR3"].Value = "team_name";
            //sheet.Cells["DS1"].Value = "";
            sheetTemplate.Cells["DS2"].Value = "Season and collection year";
            sheetTemplate.Cells["DS3"].Value = "collection_name";
            //sheet.Cells["DT1"].Value = "";
            sheetTemplate.Cells["DT2"].Value = "Material type";
            sheetTemplate.Cells["DT3"].Value = "material_type";
            //sheet.Cells["DU1"].Value = "";
            sheetTemplate.Cells["DU2"].Value = "Occasion Lifestyle";
            sheetTemplate.Cells["DU3"].Value = "lifestyle";
            //sheet.Cells["DV1"].Value = "";
            sheetTemplate.Cells["DV2"].Value = "Weave Type";
            sheetTemplate.Cells["DV3"].Value = "weave_type";
            // sheet.Cells["DW1"].Value = "";
            sheetTemplate.Cells["DW2"].Value = "League Name";
            sheetTemplate.Cells["DW3"].Value = "league_name";
            // sheet.Cells["DX1"].Value = "";
            sheetTemplate.Cells["DX2"].Value = "Shaft Style Type";
            sheetTemplate.Cells["DX3"].Value = "shaft_style_type";
            // sheet.Cells["DY1"].Value = "";
            sheetTemplate.Cells["DY2"].Value = "Product Lifecycle Supply Type";
            sheetTemplate.Cells["DY3"].Value = "lifecycle_supply_type1";
            // sheet.Cells["DZ1"].Value = "";
            sheetTemplate.Cells["DZ2"].Value = "Product Lifecycle Supply Type";
            sheetTemplate.Cells["DZ3"].Value = "lifecycle_supply_type2";
            //sheet.Cells["EA1"].Value = "";
            sheetTemplate.Cells["EA2"].Value = "Product Lifecycle Supply Type";
            sheetTemplate.Cells["EA3"].Value = "lifecycle_supply_type3";
            //sheet.Cells["EB1"].Value = "";
            sheetTemplate.Cells["EB2"].Value = "Product Lifecycle Supply Type";
            sheetTemplate.Cells["EB3"].Value = "lifecycle_supply_type4";
            //sheet.Cells["EC1"].Value = "";
            sheetTemplate.Cells["EC2"].Value = "Product Lifecycle Supply Type";
            sheetTemplate.Cells["EC3"].Value = "lifecycle_supply_type5";
            //sheet.Cells["ED1"].Value = "";
            sheetTemplate.Cells["ED2"].Value = "Pattern";
            sheetTemplate.Cells["ED3"].Value = "pattern_name";
            //sheet.Cells["EE1"].Value = "";
            sheetTemplate.Cells["EE2"].Value = "Item Booking Date";
            sheetTemplate.Cells["EE3"].Value = "item_booking_date";
            sheetTemplate.Cells["EF1"].Value = "Product Enrichment";
            sheetTemplate.Cells["EF2"].Value = "character";
            sheetTemplate.Cells["EF3"].Value = "subject_character";
            //sheet.Cells["EG1"].Value = "";
            sheetTemplate.Cells["EG2"].Value = "Fabric Wash";
            sheetTemplate.Cells["EG3"].Value = "fabric_wash";
            //sheet.Cells["EH1"].Value = "";
            sheetTemplate.Cells["EH2"].Value = "Sleeve Type";
            sheetTemplate.Cells["EH3"].Value = "sleeve_type";
            //sheet.Cells["EI1"].Value = "";
            sheetTemplate.Cells["EI2"].Value = "Strap Type";
            sheetTemplate.Cells["EI3"].Value = "strap_type";
            sheetTemplate.Cells["EJ1"].Value = "Dimensions";
            sheetTemplate.Cells["EJ2"].Value = "Shipping Weight";
            sheetTemplate.Cells["EJ3"].Value = "website_shipping_weight";
            //sheet.Cells["EK1"].Value = "";
            sheetTemplate.Cells["EK2"].Value = "Website Shipping Weight Unit Of Measure";
            sheetTemplate.Cells["EK3"].Value = "website_shipping_weight_unit_of_measure";
            //sheet.Cells["EL1"].Value = "";
            sheetTemplate.Cells["EL2"].Value = "Chest Size";
            sheetTemplate.Cells["EL3"].Value = "chest_size";
            //sheet.Cells["EM1"].Value = "";
            sheetTemplate.Cells["EM2"].Value = "Chest Size Unit Of Measure";
            sheetTemplate.Cells["EM3"].Value = "chest_size_unit_of_measure";
            //sheet.Cells["EN1"].Value = "";
            sheetTemplate.Cells["EN2"].Value = "Cup Size";
            sheetTemplate.Cells["EN3"].Value = "cup_size";
            //sheet.Cells["EO1"].Value = "";
            sheetTemplate.Cells["EO2"].Value = "Neck Size";
            sheetTemplate.Cells["EO3"].Value = "neck_size";
            //sheet.Cells["EP1"].Value = "";
            sheetTemplate.Cells["EP2"].Value = "Neck Size Unit Of Measure";
            sheetTemplate.Cells["EP3"].Value = "neck_size_unit_of_measure";
            //sheet.Cells["EQ1"].Value = "";
            sheetTemplate.Cells["EQ2"].Value = "Size Map";
            sheetTemplate.Cells["EQ3"].Value = "size_map";
            //sheet.Cells["ER1"].Value = "";
            sheetTemplate.Cells["ER2"].Value = "SleeveLength";
            sheetTemplate.Cells["ER3"].Value = "sleeve_length";
            //sheet.Cells["ES1"].Value = "";
            sheetTemplate.Cells["ES2"].Value = "Sleeve Length Unit Of Measure";
            sheetTemplate.Cells["ES3"].Value = "sleeve_length_unit_of_measure";
            //sheet.Cells["ET1"].Value = "";
            sheetTemplate.Cells["ET2"].Value = "Item Length Unit Of Measure";
            sheetTemplate.Cells["ET3"].Value = "item_length_unit_of_measure";
            //sheet.Cells["EU1"].Value = "";
            sheetTemplate.Cells["EU2"].Value = "Item Length";
            sheetTemplate.Cells["EU3"].Value = "item_length";
            //sheet.Cells["EV1"].Value = "";
            sheetTemplate.Cells["EV2"].Value = "Item Width";
            sheetTemplate.Cells["EV3"].Value = "item_width";
            //sheetTemplate.Cells["EW1"].Value = "";
            sheetTemplate.Cells["EW2"].Value = "Item Height";
            sheetTemplate.Cells["EW3"].Value = "item_height";
            // sheet.Cells["EX1"].Value = "";
            sheetTemplate.Cells["EX2"].Value = "Item Width Unit Of Measure";
            sheetTemplate.Cells["EX3"].Value = "item_width_unit_of_measure";
            // sheet.Cells["EY1"].Value = "";
            sheetTemplate.Cells["EY2"].Value = "Item Height Unit Of Measure";
            sheetTemplate.Cells["EY3"].Value = "item_height_unit_of_measure";
            sheetTemplate.Cells["EZ1"].Value = "Compliance";
            sheetTemplate.Cells["EZ2"].Value = "Cpsia Warning";
            sheetTemplate.Cells["EZ3"].Value = "cpsia_cautionary_statement";
            //sheet.Cells["FA1"].Value = "";
            sheetTemplate.Cells["FA2"].Value = "CPSIA Warning Description";
            sheetTemplate.Cells["FA3"].Value = "cpsia_cautionary_description";
            //sheet.Cells["FB1"].Value = "";
            sheetTemplate.Cells["FB2"].Value = "Fabric Type";
            sheetTemplate.Cells["FB3"].Value = "fabric_type";
            //sheet.Cells["FC1"].Value = "";
            sheetTemplate.Cells["FC2"].Value = "Import Designation";
            sheetTemplate.Cells["FC3"].Value = "import_designation";
            //sheet.Cells["FD1"].Value = "";
            sheetTemplate.Cells["FD2"].Value = "item_weight_unit_of_measure";
            sheetTemplate.Cells["FD3"].Value = "item_weight_unit_of_measure";
            //sheet.Cells["FE1"].Value = "";
            sheetTemplate.Cells["FE2"].Value = "Item Weight";
            sheetTemplate.Cells["FE3"].Value = "item_weight";
            //sheet.Cells["EF1"].Value = "";
            sheetTemplate.Cells["FF2"].Value = "Country/Region of Origin";
            sheetTemplate.Cells["FF3"].Value = "country_of_origin";
            //sheet.Cells["FG1"].Value = "";
            sheetTemplate.Cells["FG2"].Value = "Is this product a battery or does it utilize batteries?";
            sheetTemplate.Cells["FG3"].Value = "batteries_required";
            //sheet.Cells["FH1"].Value = "";
            sheetTemplate.Cells["FH2"].Value = "Batteries are Included";
            sheetTemplate.Cells["FH3"].Value = "are_batteries_included";
            //sheet.Cells["FI1"].Value = "";
            sheetTemplate.Cells["FI2"].Value = "Battery type/size";
            sheetTemplate.Cells["FI3"].Value = "battery_type1";
            //sheetTemplate.Cells["FJ1"].Value = "";
            sheetTemplate.Cells["FJ2"].Value = "Battery type/size";
            sheetTemplate.Cells["FJ3"].Value = "battery_type2";
            //sheet.Cells["FK1"].Value = "";
            sheetTemplate.Cells["FK2"].Value = "Battery type/size";
            sheetTemplate.Cells["FK3"].Value = "battery_type3";
            //sheet.Cells["FL1"].Value = "";
            sheetTemplate.Cells["FL2"].Value = "Number of batteries";
            sheetTemplate.Cells["FL3"].Value = "number_of_batteries1";
            //sheet.Cells["FM1"].Value = "";
            sheetTemplate.Cells["FM2"].Value = "Number of batteries";
            sheetTemplate.Cells["FM3"].Value = "number_of_batteries2";
            //sheet.Cells["FN1"].Value = "";
            sheetTemplate.Cells["FN2"].Value = "Number of batteries";
            sheetTemplate.Cells["FN3"].Value = "number_of_batteries3";
            //sheet.Cells["FO1"].Value = "";
            sheetTemplate.Cells["FO2"].Value = "Number of Lithium Metal Cells";
            sheetTemplate.Cells["FO3"].Value = "number_of_lithium_metal_cells";
            //sheet.Cells["FP1"].Value = "";
            sheetTemplate.Cells["FP2"].Value = "Number of Lithium-ion Cells";
            sheetTemplate.Cells["FP3"].Value = "number_of_lithium_ion_cells";
            //sheet.Cells["FQ1"].Value = "";
            sheetTemplate.Cells["FQ2"].Value = "Lithium Battery Packaging";
            sheetTemplate.Cells["FQ3"].Value = "lithium_battery_packaging";
            //sheet.Cells["FR1"].Value = "";
            sheetTemplate.Cells["FR2"].Value = "Watt hours per battery";
            sheetTemplate.Cells["FR3"].Value = "lithium_battery_energy_content";
            //sheet.Cells["FS1"].Value = "";
            sheetTemplate.Cells["FS2"].Value = "lithium_battery_energy_content_unit_of_measure";
            sheetTemplate.Cells["FS3"].Value = "lithium_battery_energy_content_unit_of_measure";
            //sheet.Cells["FT1"].Value = "";
            sheetTemplate.Cells["FT2"].Value = "Lithium content (grams)";
            sheetTemplate.Cells["FT3"].Value = "lithium_battery_weight";
            //sheet.Cells["FU1"].Value = "";
            sheetTemplate.Cells["FU2"].Value = "lithium_battery_weight_unit_of_measure";
            sheetTemplate.Cells["FU3"].Value = "lithium_battery_weight_unit_of_measure";
            //sheet.Cells["FV1"].Value = "";
            sheetTemplate.Cells["FV2"].Value = "Volume";
            sheetTemplate.Cells["FV3"].Value = "item_volume";
            //sheet.Cells["FW1"].Value = "";
            sheetTemplate.Cells["FW2"].Value = "item_volume_unit_of_measure";
            sheetTemplate.Cells["FW3"].Value = "item_volume_unit_of_measure";
            // sheet.Cells["FX1"].Value = "";
            sheetTemplate.Cells["FX2"].Value = "Material/Fabric Regulations";
            sheetTemplate.Cells["FX3"].Value = "supplier_declared_material_regulation1";
            // sheet.Cells["FY1"].Value = "";
            sheetTemplate.Cells["FY2"].Value = "Material/Fabric Regulations";
            sheetTemplate.Cells["FY3"].Value = "supplier_declared_material_regulation2";
            // sheet.Cells["FZ1"].Value = "";
            sheetTemplate.Cells["FZ2"].Value = "Material/Fabric Regulations";
            sheetTemplate.Cells["FZ3"].Value = "supplier_declared_material_regulation3";
            //sheet.Cells["GA1"].Value = "";
            sheetTemplate.Cells["GA2"].Value = "California Proposition 65 Warning Type";
            sheetTemplate.Cells["GA3"].Value = "california_proposition_65_compliance_type";
            //sheetTemplate.Cells["GB1"].Value = "";
            sheetTemplate.Cells["GB2"].Value = "California Proposition 65 Chemical Names";
            sheetTemplate.Cells["GB3"].Value = "california_proposition_65_chemical_names1";
            //sheet.Cells["GC1"].Value = "";
            sheetTemplate.Cells["GC2"].Value = "Additional Chemical Name1";
            sheetTemplate.Cells["GC3"].Value = "california_proposition_65_chemical_names2";
            //sheet.Cells["GD1"].Value = "";
            sheetTemplate.Cells["GD2"].Value = "Additional Chemical Name2";
            sheetTemplate.Cells["GD3"].Value = "california_proposition_65_chemical_names3";
            //sheet.Cells["GE1"].Value = "";
            sheetTemplate.Cells["GE2"].Value = "Additional Chemical Name3";
            sheetTemplate.Cells["GE3"].Value = "california_proposition_65_chemical_names4";
            //sheet.Cells["GF1"].Value = "";
            sheetTemplate.Cells["GF2"].Value = "Additional Chemical Name4";
            sheetTemplate.Cells["GF3"].Value = "california_proposition_65_chemical_names5";
            //sheet.Cells["GG1"].Value = "";
            sheetTemplate.Cells["GG2"].Value = "Pesticide Marking";
            sheetTemplate.Cells["GG3"].Value = "pesticide_marking_type1";
            //sheet.Cells["GH1"].Value = "";
            sheetTemplate.Cells["GH2"].Value = "Pesticide Marking";
            sheetTemplate.Cells["GH3"].Value = "pesticide_marking_type2";
            //sheet.Cells["GI1"].Value = "";
            sheetTemplate.Cells["GI2"].Value = "Pesticide Marking";
            sheetTemplate.Cells["GI3"].Value = "pesticide_marking_type3";
            //sheet.Cells["GJ1"].Value = "";
            sheetTemplate.Cells["GJ2"].Value = "Pesticide Registration Status";
            sheetTemplate.Cells["GJ3"].Value = "pesticide_marking_registration_status1";
            //sheet.Cells["GK1"].Value = "";
            sheetTemplate.Cells["GK2"].Value = "Pesticide Registration Status";
            sheetTemplate.Cells["GK3"].Value = "pesticide_marking_registration_status2";
            //sheet.Cells["GL1"].Value = "";
            sheetTemplate.Cells["GL2"].Value = "Pesticide Registration Status";
            sheetTemplate.Cells["GL3"].Value = "pesticide_marking_registration_status3";
            //sheet.Cells["GM1"].Value = "";
            sheetTemplate.Cells["GM2"].Value = "Pesticide Certification Number";
            sheetTemplate.Cells["GM3"].Value = "pesticide_marking_certification_number1";
            //sheet.Cells["GN1"].Value = "";
            sheetTemplate.Cells["GN2"].Value = "Pesticide Certification Number";
            sheetTemplate.Cells["GN3"].Value = "pesticide_marking_certification_number2";
            //sheet.Cells["GO1"].Value = "";
            sheetTemplate.Cells["GO2"].Value = "Pesticide Certification Number";
            sheetTemplate.Cells["GO3"].Value = "pesticide_marking_certification_number3";
            sheetTemplate.Cells["GP1"].Value = "Offer";
            sheetTemplate.Cells["GP2"].Value = "Manufacturer's Suggested Retail Price";
            sheetTemplate.Cells["GP3"].Value = "list_price";
            //sheet.Cells["GQ1"].Value = "";
            sheetTemplate.Cells["GQ2"].Value = "Minimum Advertised Price";
            sheetTemplate.Cells["GQ3"].Value = "map_price";
            //sheet.Cells["GR1"].Value = "";
            sheetTemplate.Cells["GR2"].Value = "Currency";
            sheetTemplate.Cells["GR3"].Value = "currency";
            //sheet.Cells["GS1"].Value = "";
            sheetTemplate.Cells["GS2"].Value = "Item Condition";
            sheetTemplate.Cells["GS3"].Value = "condition_type";
            //sheet.Cells["GT1"].Value = "";
            sheetTemplate.Cells["GT2"].Value = "Condition Note";
            sheetTemplate.Cells["GT3"].Value = "condition_note";
            //sheet.Cells["GU1"].Value = "";
            sheetTemplate.Cells["GU2"].Value = "Product Tax Code";
            sheetTemplate.Cells["GU3"].Value = "product_tax_code";
            //sheet.Cells["GV1"].Value = "";
            sheetTemplate.Cells["GV2"].Value = "Handling Time";
            sheetTemplate.Cells["GV3"].Value = "fulfillment_latency";
            //sheet.Cells["GW1"].Value = "";
            sheetTemplate.Cells["GW2"].Value = "Launch Date";
            sheetTemplate.Cells["GW3"].Value = "product_site_launch_date";
            //sheet.Cells["GX1"].Value = "";
            sheetTemplate.Cells["GX2"].Value = "Launch Date";
            sheetTemplate.Cells["GX3"].Value = "product_site_launch_date";
            //sheet.Cells["GY1"].Value = "";
            sheetTemplate.Cells["GY2"].Value = "Restock Date";
            sheetTemplate.Cells["GT3"].Value = "restock_date";
            //sheet.Cells["GZ1"].Value = "";
            sheetTemplate.Cells["GZ2"].Value = "Sale Price";
            sheetTemplate.Cells["GZ3"].Value = "sale_price";
            //sheet.Cells["HA1"].Value = "";
            sheetTemplate.Cells["HA2"].Value = "Sale Start Date";
            sheetTemplate.Cells["HA3"].Value = "sale_from_date";
            //sheet.Cells["HB1"].Value = "";
            sheetTemplate.Cells["HB2"].Value = "Sale End Date";
            sheetTemplate.Cells["HB3"].Value = "sale_end_date";
            //sheet.Cells["HC1"].Value = "";
            sheetTemplate.Cells["HC2"].Value = "Offer End Date";
            sheetTemplate.Cells["HC3"].Value = "offering_end_date";
            //sheet.Cells["HD1"].Value = "";
            sheetTemplate.Cells["HD2"].Value = "Max Aggregate Ship Quantity";
            sheetTemplate.Cells["HD3"].Value = "max_aggregate_ship_quantity";
            //sheet.Cells["HE1"].Value = "";
            sheetTemplate.Cells["HE2"].Value = "Package Quantity";
            sheetTemplate.Cells["HE3"].Value = "item_package_quantity";
            //sheet.Cells["HF1"].Value = "";
            sheetTemplate.Cells["HF2"].Value = "Number of Items";
            sheetTemplate.Cells["HF3"].Value = "number_of_items";
            //sheet.Cells["HG1"].Value = "";
            sheetTemplate.Cells["HG2"].Value = "Offering Can Be Gift Messaged";
            sheetTemplate.Cells["HG3"].Value = "offering_can_be_gift_messaged";
            //sheet.Cells["HH1"].Value = "";
            sheetTemplate.Cells["HH2"].Value = "Is Gift Wrap Available";
            sheetTemplate.Cells["HH3"].Value = "offering_can_be_giftwrapped";
            //sheet.Cells["HI1"].Value = "";
            sheetTemplate.Cells["HI2"].Value = "Is Discontinued by Manufacturer";
            sheetTemplate.Cells["HI3"].Value = "is_discontinued_by_manufacturer";
            //sheet.Cells["HJ1"].Value = "";
            sheetTemplate.Cells["HJ2"].Value = "Max Order Quantity";
            sheetTemplate.Cells["HJ3"].Value = "max_order_quantity";
            //sheet.Cells["HK1"].Value = "";
            sheetTemplate.Cells["HK2"].Value = "Shipping-Template";
            sheetTemplate.Cells["HK3"].Value = "merchant_shipping_group_name";
            //sheet.Cells["HL1"].Value = "";
            sheetTemplate.Cells["HL2"].Value = "Offer Start Date";
            sheetTemplate.Cells["HL3"].Value = "offering_start_date";
            #endregion


            // khoi variant

            var a = "đây là dữ liệu cột A";
            var h = "đây là dữ liệu cột H ";
            var w = " đây là dữ liệu cột W ";
            var z = "đây là dữ liệu cột Z";
            var aE = "đây là dữ liệu cột AE";
            var aF = "đây là dữ liệu cột AF";
            var aG = "đây là dữ liệu cột AG";
            var aH = "đây là dữ liệu cột AH";
            var aI = "đây là dữ liệu cột AI";
            var aJ = "đây là dữ liệu cột AJ";
            var bL = "đây là dữ liệu cột BL";
            var eW = "đây là dữ liệu cột Ew";


            var excelProduct = new ExcelPackage();
            excelProduct.Load(new FileStream("zamage.com-1626882450.xlsx", FileMode.Open));
            var sheetProduct = excelProduct.Workbook.Worksheets.FirstOrDefault();
            if (sheetProduct != null)
            {
                var row = sheetProduct.Dimension.Rows;
                var rowCategory = sheetMap.Dimension.Rows;
                var rowBegin = 4;

                for (int i = 2; i <= row; i++)
                {
                    var variants = new List<Variant>();
                    for (int j = 1; j <= rowCategory; j++)
                    {
                        variants.Add(new Variant { Sku = GenSkuCode() });
                        variants.Add(new Variant { Sku = GenSkuCode() });
                        variants.Add(new Variant { Sku = GenSkuCode() });
                    }

                    // fill anh
                    var images = sheetProduct.Cells[i, 2]?.Value?.ToString().Split('|').ToList();

                    if (images != null)
                    {
                        if (images.Count > 9)
                            images = images.Take(9).ToList();

                        var name = sheetProduct.Cells[i, 1].Value;
                        for (int j = 0; j < images.Count; j++)
                        {
                            //sheetTemplate.Cells[rowBegin, 10 + j].Value = images[j];
                            for (int k = 0; k < 4; k++)
                            {
                                sheetTemplate.Cells[rowBegin +k, 10 + j].Value = images[j];

                            }

                        }

                        // diền  T
                        sheetTemplate.Cells["T" + rowBegin].Value = "Parent";
                        foreach (var item in variants)
                        {
                            //lấy sku 
                            var b = item.Sku + "-" + ThreadSafeRandom.ThisThreadsRandom.Next(10000, 99999);
                            sheetTemplate.Cells["B" + rowBegin].Value = b;
                            sheetTemplate.Cells["U" + rowBegin].Value = item.Sku;
                            sheetTemplate.Cells["D" + rowBegin].Value = name;
                            sheetTemplate.Cells["V" + rowBegin].Value = "Variation";
                            sheetTemplate.Cells["W" + rowBegin].Value = "Size";
                            for (int j = 1; j < 4; j++)
                            {
                                sheetTemplate.Cells["B" + (rowBegin + j)].Value = b;
                                sheetTemplate.Cells["U" + (rowBegin + j)].Value = item.Sku;
                                sheetTemplate.Cells["D" + (rowBegin + j)].Value = name;
                                sheetTemplate.Cells["V" + (rowBegin + j)].Value = "Variation";
                                sheetTemplate.Cells["W" + (rowBegin + j)].Value = "Size";
                                sheetTemplate.Cells["T" + (rowBegin + j)].Value = "Child";
                            }
                           


                        }

                    }
                    // fill du lieu dien tay
                    for (int j = 0; j < 4; j++)
                    {
                        sheetTemplate.Cells["A" + (rowBegin + j)].Value = a;
                        sheetTemplate.Cells["H" + (rowBegin + j)].Value = h;
                        sheetTemplate.Cells["W" + (rowBegin + j)].Value = w;
                        sheetTemplate.Cells["Z" + (rowBegin + j)].Value = z;
                        sheetTemplate.Cells["AE" + (rowBegin + j)].Value = aE;
                        sheetTemplate.Cells["AF" + (rowBegin + j)].Value = aF;
                        sheetTemplate.Cells["AG" + (rowBegin + j)].Value = aG;
                        sheetTemplate.Cells["AH" + (rowBegin + j)].Value = aH;
                        sheetTemplate.Cells["AI" + (rowBegin + j)].Value = aI;
                        sheetTemplate.Cells["AJ" + (rowBegin + j)].Value = aJ;
                        sheetTemplate.Cells["BL" + (rowBegin + j)].Value = bL;
                        sheetTemplate.Cells["Ew" + (rowBegin + j)].Value = eW;
                    }

                    for (int j = 2; j <= rowCategory; j++)
                    {
                        // lay dc size
                        var size = sheetMap.Cells[j, 2].Value;
                        sheetTemplate.Cells["AQ" + (rowBegin +1)].Value = size;

                        // lay BP
                        var bp = sheetMap.Cells[j, 3].Value;
                        sheetTemplate.Cells["BP" + (rowBegin + 1)].Value = bp;

                        // lay BQ
                        var bq = sheetMap.Cells[j, 4].Value;
                        sheetTemplate.Cells["BQ" + (rowBegin + 1)].Value = bq;
                        rowBegin++;

                    }

                    
                    rowBegin -= 8;
                }
            }

            excelTemplate.SaveAs(new FileInfo("test.xlsx"));

        }

        static string GenSkuCode()
        {

            return "SKU" + "-" + ThreadSafeRandom.ThisThreadsRandom.Next(10000, 99999);
        }
    }
}
