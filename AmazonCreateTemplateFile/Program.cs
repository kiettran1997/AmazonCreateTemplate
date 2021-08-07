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
            sheetTemplate.Cells["B1"].Value = "Version=2021.0625";
            sheetTemplate.Cells["B2"].Value = "Seller SKU";
            sheetTemplate.Cells["B3"].Value = "item_sku";
            sheetTemplate.Cells["C1"].Value = "TemplateSignature=QkxBTktFVA==";
            sheetTemplate.Cells["C2"].Value = "Brand Name";
            sheetTemplate.Cells["C3"].Value = "brand_name";
            sheetTemplate.Cells["D1"].Value = "settings=contentLanguageTag=en_US&feedType=610841&headerLanguageTag=en_US&primaryMarketplaceId=amzn1.mp.o.ATVPDKIKX0DER&templateIdentifier=deb3fee2-b6dd-4966-bd7a-3a4e12015d5a&timestamp=2021-06-25T04%3A40%3A02.301Z";
            sheetTemplate.Cells["D2"].Value = "Product Name";
            sheetTemplate.Cells["D3"].Value = "item_name";
            sheetTemplate.Cells["E1"].Value = "Use ENGLISH to fill this template.The top 3 rows are for Amazon.com use only. Do not modify or delete the top 3 rows.";
            sheetTemplate.Cells["E2"].Value = "Manufacturer";
            sheetTemplate.Cells["E3"].Value = "manufacturer";
            //sheet.Cells["F1"].Value = "";
            sheetTemplate.Cells["F2"].Value = "Product ID";
            sheetTemplate.Cells["F3"].Value = "external_product_id";
            //sheet.Cells["G1"].Value = "";
            sheetTemplate.Cells["G2"].Value = "Product ID Type";
            sheetTemplate.Cells["G3"].Value = "external_product_id_type";
            //sheet.Cells["H1"].Value = "";
            sheetTemplate.Cells["H2"].Value = "Item Type Keyword";
            sheetTemplate.Cells["H3"].Value = "item_type";
            //sheet.Cells["I1"].Value = "";
            sheetTemplate.Cells["I2"].Value = "Quantity";
            sheetTemplate.Cells["I3"].Value = "quantity";
            //sheet.Cells["J1"].Value = "";
            sheetTemplate.Cells["J2"].Value = "Main Image URL";
            sheetTemplate.Cells["J3"].Value = "main_image_url";
            sheetTemplate.Cells["K1"].Value = "Images";
            sheetTemplate.Cells["K2"].Value = "Other Image URL1";
            sheetTemplate.Cells["K3"].Value = "other_image_url1";
            //sheet.Cells["L1"].Value = "";
            sheetTemplate.Cells["L2"].Value = "Other Image URL2";
            sheetTemplate.Cells["L3"].Value = "other_image_url2";
            //sheet.Cells["M1"].Value = "";
            sheetTemplate.Cells["M2"].Value = "Other Image URL3";
            sheetTemplate.Cells["M3"].Value = "other_image_url3";
            //sheet.Cells["N1"].Value = "";
            sheetTemplate.Cells["N2"].Value = "Other Image URL4";
            sheetTemplate.Cells["N3"].Value = "other_image_url4";
            //sheet.Cells["O1"].Value = "";
            sheetTemplate.Cells["O2"].Value = "Other Image URL5";
            sheetTemplate.Cells["O3"].Value = "other_image_url5";
            //sheet.Cells["P1"].Value = "";
            sheetTemplate.Cells["P2"].Value = "Other Image URL6";
            sheetTemplate.Cells["P3"].Value = "other_image_url6";
            //sheet.Cells["Q1"].Value = "";
            sheetTemplate.Cells["Q2"].Value = "Other Image URL7";
            sheetTemplate.Cells["Q3"].Value = "Other Image URL7";
            //sheet.Cells["R1"].Value = "";
            sheetTemplate.Cells["R2"].Value = "Other Image URL8";
            sheetTemplate.Cells["R3"].Value = "other_image_url8";
            //sheet.Cells["S1"].Value = "";
            sheetTemplate.Cells["S2"].Value = "Swatch Image URL";
            sheetTemplate.Cells["S3"].Value = "swatch_image_url";
            sheetTemplate.Cells["T1"].Value = "Variation";
            sheetTemplate.Cells["T2"].Value = "Parentage";
            sheetTemplate.Cells["T3"].Value = "parent_child";
            //sheet.Cells["U1"].Value = "";
            sheetTemplate.Cells["U2"].Value = "Parent SKU";
            sheetTemplate.Cells["U3"].Value = "parent_sku";
            //sheet.Cells["V1"].Value = "";
            sheetTemplate.Cells["V2"].Value = "Relationship Type";
            sheetTemplate.Cells["V3"].Value = "relationship_type";
            //sheet.Cells["W1"].Value = "";
            sheetTemplate.Cells["W2"].Value = "Variation Theme";
            sheetTemplate.Cells["W3"].Value = "variation_theme";
            sheetTemplate.Cells["X1"].Value = "Basic";
            sheetTemplate.Cells["X2"].Value = "Update Delete";
            sheetTemplate.Cells["X3"].Value = "update_delete";
            //sheet.Cells["Y1"].Value = "";
            sheetTemplate.Cells["Y2"].Value = "Product Exemption Reason";
            sheetTemplate.Cells["Y3"].Value = "gtin_exemption_reason";
            //sheet.Cells["Z1"].Value = "";
            sheetTemplate.Cells["Z2"].Value = "Product Description";
            sheetTemplate.Cells["Z3"].Value = "product_description";
            //sheet.Cells["AA1"].Value = "";
            sheetTemplate.Cells["AA2"].Value = "Manufacturer Part Number";
            sheetTemplate.Cells["AA3"].Value = "part_number";
            //sheet.Cells["AB1"].Value = "";
            sheetTemplate.Cells["AB2"].Value = "model";
            sheetTemplate.Cells["AB3"].Value = "model";
            //sheet.Cells["AC1"].Value = "";
            sheetTemplate.Cells["AC2"].Value = "Model Name";
            sheetTemplate.Cells["AC3"].Value = "model_name";
            //sheet.Cells["AD1"].Value = "";
            sheetTemplate.Cells["AD2"].Value = "Care Instructions";
            sheetTemplate.Cells["AD3"].Value = "care_instructions";
            sheetTemplate.Cells["AE1"].Value = "Discovery";
            sheetTemplate.Cells["AE2"].Value = "Key Product Features";
            sheetTemplate.Cells["AE3"].Value = "bullet_point1";
            //sheet.Cells["AF1"].Value = "";
            sheetTemplate.Cells["AF2"].Value = "Key Product Features";
            sheetTemplate.Cells["AF3"].Value = "bullet_point2";
            //sheet.Cells["AG1"].Value = "";
            sheetTemplate.Cells["AG2"].Value = "Key Product Features";
            sheetTemplate.Cells["AG3"].Value = "bullet_point3";
            //sheet.Cells["AH1"].Value = "";
            sheetTemplate.Cells["AH2"].Value = "Key Product Features";
            sheetTemplate.Cells["AH3"].Value = "bullet_point4";
            //sheet.Cells["AI1"].Value = "";
            sheetTemplate.Cells["AI2"].Value = "Key Product Features";
            sheetTemplate.Cells["AI3"].Value = "bullet_point5";
            //sheet.Cells["AJ1"].Value = "";
            sheetTemplate.Cells["AJ2"].Value = "Search Terms";
            sheetTemplate.Cells["AJ3"].Value = "generic_keywords";
            //sheet.Cells["AK1"].Value = "";
            sheetTemplate.Cells["AK2"].Value = "Country/Region as Labeled";
            sheetTemplate.Cells["AK3"].Value = "country_as_labeled";
            //sheet.Cells["AL1"].Value = "";
            sheetTemplate.Cells["AL2"].Value = "Number of Pieces";
            sheetTemplate.Cells["AL3"].Value = "number_of_pieces";
            //sheet.Cells["AM1"].Value = "";
            sheetTemplate.Cells["AM2"].Value = "Scent";
            sheetTemplate.Cells["AM3"].Value = "scent_name";
            //sheet.Cells["AN1"].Value = "";
            sheetTemplate.Cells["AN2"].Value = "Included Components";
            sheetTemplate.Cells["AN3"].Value = "included_components";
            //sheet.Cells["ANO1"].Value = "";
            sheetTemplate.Cells["AO2"].Value = "Color";
            sheetTemplate.Cells["AO3"].Value = "color_name";
            //sheet.Cells["AP1"].Value = "";
            sheetTemplate.Cells["AP2"].Value = "Color Map";
            sheetTemplate.Cells["AP3"].Value = "color_map";
            //sheet.Cells["AQ1"].Value = "";
            sheetTemplate.Cells["AQ2"].Value = "Size";
            sheetTemplate.Cells["AQ3"].Value = "size_name";
            //sheet.Cells["AR1"].Value = "";
            sheetTemplate.Cells["AR2"].Value = "Material Type";
            sheetTemplate.Cells["AR3"].Value = "material_type";
            //sheet.Cells["AS1"].Value = "";
            sheetTemplate.Cells["AS2"].Value = "Wattage";
            sheetTemplate.Cells["AS3"].Value = "wattage";
            //sheet.Cells["AT1"].Value = "";
            sheetTemplate.Cells["AT2"].Value = "Additional Features";
            sheetTemplate.Cells["AT3"].Value = "special_features1";
            //sheet.Cells["AU1"].Value = "";
            sheetTemplate.Cells["AU2"].Value = "Additional Features";
            sheetTemplate.Cells["AU3"].Value = "special_features2";
            //sheet.Cells["AV1"].Value = "";
            sheetTemplate.Cells["AV2"].Value = "Additional Features";
            sheetTemplate.Cells["AV3"].Value = "special_features3";
            // sheet.Cells["AW1"].Value = "";
            sheetTemplate.Cells["AW2"].Value = "Additional Features";
            sheetTemplate.Cells["AW3"].Value = "special_features4";
            // sheet.Cells["AX1"].Value = "";
            sheetTemplate.Cells["AX2"].Value = "Additional Features";
            sheetTemplate.Cells["AX3"].Value = "special_features5";
            // sheet.Cells["AY1"].Value = "";
            sheetTemplate.Cells["AY2"].Value = "Pattern";
            sheetTemplate.Cells["AY3"].Value = "pattern_name";
            // sheet.Cells["AZ1"].Value = "";
            sheetTemplate.Cells["AZ2"].Value = "Season of the Product";
            sheetTemplate.Cells["AZ3"].Value = "seasons1";
            //sheet.Cells["BA1"].Value = "";
            sheetTemplate.Cells["BA2"].Value = "Season of the Product";
            sheetTemplate.Cells["BA3"].Value = "seasons2";
            //sheet.Cells["BB1"].Value = "";
            sheetTemplate.Cells["BB2"].Value = "Season of the Product";
            sheetTemplate.Cells["BB3"].Value = "seasons3";
            //sheet.Cells["BC1"].Value = "";
            sheetTemplate.Cells["BC2"].Value = "Season of the Product";
            sheetTemplate.Cells["BC3"].Value = "seasons4";
            //sheet.Cells["BD1"].Value = "";
            sheetTemplate.Cells["BD2"].Value = "Item Type";
            sheetTemplate.Cells["BD3"].Value = "item_type_name";
            //sheet.Cells["BE1"].Value = "";
            sheetTemplate.Cells["BE2"].Value = "Wattage Unit of Measure";
            sheetTemplate.Cells["BE3"].Value = "wattage_unit_of_measure";
            //sheet.Cells["BF1"].Value = "";
            sheetTemplate.Cells["BF2"].Value = "Length Range";
            sheetTemplate.Cells["BF3"].Value = "length_range";
            //sheet.Cells["BG1"].Value = "";
            sheetTemplate.Cells["BG2"].Value = "Is Assembly Required";
            sheetTemplate.Cells["BG3"].Value = "is_assembly_required";
            //sheet.Cells["BH1"].Value = "";
            sheetTemplate.Cells["BH2"].Value = "Number of Shelves";
            sheetTemplate.Cells["BH3"].Value = "number_of_shelves";
            //sheet.Cells["BI1"].Value = "";
            sheetTemplate.Cells["BI2"].Value = "Number of Boxes";
            sheetTemplate.Cells["BI3"].Value = "number_of_boxes";
            //sheet.Cells["BJ1"].Value = "";
            sheetTemplate.Cells["BJ2"].Value = "Number of Compartments";
            sheetTemplate.Cells["BJ3"].Value = "number_of_compartments";
            sheetTemplate.Cells["BK1"].Value = "Dimensions";
            sheetTemplate.Cells["BK2"].Value = "Towel Weight";
            sheetTemplate.Cells["BK3"].Value = "fabric_weight";
            //sheet.Cells["BL1"].Value = "";
            sheetTemplate.Cells["BL2"].Value = "Shape";
            sheetTemplate.Cells["BL3"].Value = "item_shape";
            //sheet.Cells["BM1"].Value = "";
            sheetTemplate.Cells["BM2"].Value = "Display Length Unit Of Measure";
            sheetTemplate.Cells["BM3"].Value = "item_display_length_unit_of_measure";
            //sheet.Cells["BN1"].Value = "";
            sheetTemplate.Cells["BN2"].Value = "Item Display Width Unit Of Measure";
            sheetTemplate.Cells["BN3"].Value = "item_display_width_unit_of_measure";
            //sheet.Cells["BO1"].Value = "";
            sheetTemplate.Cells["BO2"].Value = "Item Display Height Unit Of Measure";
            sheetTemplate.Cells["BO3"].Value = "item_display_height_unit_of_measure";
            //sheet.Cells["BP1"].Value = "";
            sheetTemplate.Cells["BP2"].Value = "Item Display Length";
            sheetTemplate.Cells["BP3"].Value = "item_display_length";
            //sheet.Cells["BQ1"].Value = "";
            sheetTemplate.Cells["BQ2"].Value = "Item Display Width";
            sheetTemplate.Cells["BQ3"].Value = "item_display_width";
            //sheet.Cells["BR1"].Value = "";
            sheetTemplate.Cells["BR2"].Value = "Item Display Height";
            sheetTemplate.Cells["BR3"].Value = "item_display_height";
            //sheet.Cells["BS1"].Value = "";
            sheetTemplate.Cells["BS2"].Value = "Item Display Weight";
            sheetTemplate.Cells["BS3"].Value = "item_display_weight";
            //sheet.Cells["BT1"].Value = "";
            sheetTemplate.Cells["BT2"].Value = "Item Display Weight Unit Of Measure";
            sheetTemplate.Cells["BT3"].Value = "item_display_weight_unit_of_measure";
            //sheet.Cells["BU1"].Value = "";
            sheetTemplate.Cells["BU2"].Value = "Item Height";
            sheetTemplate.Cells["BU3"].Value = "item_height";
            //sheet.Cells["BV1"].Value = "";
            sheetTemplate.Cells["BV2"].Value = "Item Length";
            sheetTemplate.Cells["BV3"].Value = "item_length";
            // sheet.Cells["BW1"].Value = "";
            sheetTemplate.Cells["BW2"].Value = "Item Length Unit Of Measure";
            sheetTemplate.Cells["BW3"].Value = "item_length_unit_of_measure";
            // sheet.Cells["BX1"].Value = "";
            sheetTemplate.Cells["BX2"].Value = "Item Width";
            sheetTemplate.Cells["BX3"].Value = "item_width";
            // sheet.Cells["BY1"].Value = "";
            sheetTemplate.Cells["BY2"].Value = "Unit of Measure (Per Unit Pricing)";
            sheetTemplate.Cells["BY3"].Value = "unit_count_type";
            // sheet.Cells["BZ1"].Value = "";
            sheetTemplate.Cells["BZ2"].Value = "Unit Count (Per Unit Pricing)";
            sheetTemplate.Cells["BZ3"].Value = "unit_count";
            //sheet.Cells["CA1"].Value = "";
            sheetTemplate.Cells["CA2"].Value = "Item Width Unit Of Measure";
            sheetTemplate.Cells["CA3"].Value = "item_width_unit_of_measure";
            //sheet.Cells["CB1"].Value = "";
            sheetTemplate.Cells["CB2"].Value = "Fabric Weight Unit";
            sheetTemplate.Cells["CB3"].Value = "fabric_weight_unit_of_measure";
            //sheet.Cells["CC1"].Value = "";
            sheetTemplate.Cells["CC2"].Value = "Size Map";
            sheetTemplate.Cells["CC3"].Value = "size_map";
            //sheet.Cells["CD1"].Value = "";
            sheetTemplate.Cells["CD2"].Value = "Width Range";
            sheetTemplate.Cells["CD3"].Value = "width_range";
            //sheet.Cells["CE1"].Value = "";
            sheetTemplate.Cells["CE2"].Value = "Maximum Weight Recommendation Unit Of Measure";
            sheetTemplate.Cells["CE3"].Value = "maximum_weight_recommendation_unit_of_measure";
            //sheet.Cells["CF1"].Value = "";
            sheetTemplate.Cells["CF2"].Value = "Item Height Unit Of Measure";
            sheetTemplate.Cells["CF3"].Value = "item_height_unit_of_measure";
            //sheet.Cells["CG1"].Value = "";
            sheetTemplate.Cells["CG2"].Value = "Maximum Weight Recommendation";
            sheetTemplate.Cells["CG3"].Value = "maximum_weight_recommendation";
            sheetTemplate.Cells["CH1"].Value = "Fulfillment";
            sheetTemplate.Cells["CH2"].Value = "Fulfillment Center ID";
            sheetTemplate.Cells["CH3"].Value = "fulfillment_center_id";
            //sheet.Cells["CI1"].Value = "";
            sheetTemplate.Cells["CI2"].Value = "Package Height";
            sheetTemplate.Cells["CI3"].Value = "package_height";
            //sheet.Cells["CJ1"].Value = "";
            sheetTemplate.Cells["CJ2"].Value = "Package Width";
            sheetTemplate.Cells["CJ3"].Value = "package_width";
            //sheet.Cells["CK1"].Value = "";
            sheetTemplate.Cells["CK2"].Value = "Package Length";
            sheetTemplate.Cells["CK3"].Value = "package_length";
            //sheet.Cells["CL1"].Value = "";
            sheetTemplate.Cells["CL2"].Value = "Package Length Unit Of Measure";
            sheetTemplate.Cells["CL3"].Value = "package_length_unit_of_measure";
            //sheet.Cells["CM1"].Value = "";
            sheetTemplate.Cells["CM2"].Value = "Package Weight";
            sheetTemplate.Cells["CM3"].Value = "package_weight";
            //sheet.Cells["CN1"].Value = "";
            sheetTemplate.Cells["CN2"].Value = "Package Weight Unit Of Measure";
            sheetTemplate.Cells["CN3"].Value = "package_weight_unit_of_measure";
            //sheet.Cells["CO1"].Value = "";
            sheetTemplate.Cells["CO2"].Value = "Package Height Unit Of Measure";
            sheetTemplate.Cells["CO3"].Value = "package_height_unit_of_measure";
            //sheet.Cells["CP1"].Value = "";
            sheetTemplate.Cells["CP2"].Value = "Package Width Unit Of Measure";
            sheetTemplate.Cells["CP3"].Value = "package_width_unit_of_measure";
            sheetTemplate.Cells["CQ1"].Value = "Compliance";
            sheetTemplate.Cells["CQ2"].Value = "Manufacturer Warranty Description";
            sheetTemplate.Cells["CQ3"].Value = "warranty_description";
            //sheet.Cells["CR1"].Value = "";
            sheetTemplate.Cells["CR2"].Value = "Cpsia Warning";
            sheetTemplate.Cells["CR3"].Value = "cpsia_cautionary_statement";
            //sheet.Cells["CS1"].Value = "";
            sheetTemplate.Cells["CS2"].Value = "Fabric Type";
            sheetTemplate.Cells["CS3"].Value = "fabric_type";
            //sheet.Cells["CT1"].Value = "";
            sheetTemplate.Cells["CT2"].Value = "Volume";
            sheetTemplate.Cells["CT3"].Value = "item_volume";
            //sheet.Cells["CU1"].Value = "";
            sheetTemplate.Cells["CU2"].Value = "item_volume_unit_of_measure";
            sheetTemplate.Cells["CU3"].Value = "item_volume_unit_of_measure";
            //sheet.Cells["CV1"].Value = "";
            sheetTemplate.Cells["CV2"].Value = "Country/Region of Origin";
            sheetTemplate.Cells["CV3"].Value = "country_of_origin";
            // sheet.Cells["CW1"].Value = "";
            sheetTemplate.Cells["CW2"].Value = "Batteries are Included";
            sheetTemplate.Cells["CW3"].Value = "are_batteries_included";
            // sheet.Cells["CX1"].Value = "";
            sheetTemplate.Cells["CX2"].Value = "Item Weight";
            sheetTemplate.Cells["CX3"].Value = "item_weight";
            // sheet.Cells["CY1"].Value = "";
            sheetTemplate.Cells["CY2"].Value = "Is this product a battery or does it utilize batteries?";
            sheetTemplate.Cells["CY3"].Value = "batteries_required";
            // sheet.Cells["CZ1"].Value = "";
            sheetTemplate.Cells["CZ2"].Value = "Battery type/size";
            sheetTemplate.Cells["CZ3"].Value = "battery_type1";
            //sheet.Cells["DA1"].Value = "";
            sheetTemplate.Cells["DA2"].Value = "Battery type/size";
            sheetTemplate.Cells["DA3"].Value = "battery_type2";
            //sheet.Cells["DB1"].Value = "";
            sheetTemplate.Cells["DB2"].Value = "Battery type/size";
            sheetTemplate.Cells["DB3"].Value = "battery_type3";
            //sheet.Cells["DC1"].Value = "";
            sheetTemplate.Cells["DC2"].Value = "item_weight_unit_of_measure";
            sheetTemplate.Cells["DC3"].Value = "item_weight_unit_of_measure";
            //sheet.Cells["DD1"].Value = "";
            sheetTemplate.Cells["DD2"].Value = "Number of batteries";
            sheetTemplate.Cells["DD3"].Value = "number_of_batteries1";
            //sheet.Cells["DE1"].Value = "";
            sheetTemplate.Cells["DE2"].Value = "Number of batteries";
            sheetTemplate.Cells["DE3"].Value = "number_of_batteries2";
            //sheet.Cells["DF1"].Value = "";
            sheetTemplate.Cells["DF2"].Value = "Number of batteries";
            sheetTemplate.Cells["DF3"].Value = "number_of_batteries3";
            //sheet.Cells["DG1"].Value = "";
            sheetTemplate.Cells["DG2"].Value = "Watt hours per battery";
            sheetTemplate.Cells["DG3"].Value = "lithium_battery_energy_content";
            //sheet.Cells["DH1"].Value = "";
            sheetTemplate.Cells["DH2"].Value = "Lithium Battery Packaging";
            sheetTemplate.Cells["DH3"].Value = "lithium_battery_packaging";
            //sheet.Cells["DI1"].Value = "";
            sheetTemplate.Cells["DI2"].Value = "Lithium content (grams)";
            sheetTemplate.Cells["DI3"].Value = "lithium_battery_weight";
            //sheet.Cells["DJ1"].Value = "";
            sheetTemplate.Cells["DJ2"].Value = "Number of Lithium-ion Cells";
            sheetTemplate.Cells["DJ3"].Value = "number_of_lithium_ion_cells";
            //sheet.Cells["DK1"].Value = "";
            sheetTemplate.Cells["DK2"].Value = "Number of Lithium Metal Cells";
            sheetTemplate.Cells["DK3"].Value = "number_of_lithium_metal_cells";
            //sheet.Cells["DL1"].Value = "";
            sheetTemplate.Cells["DL2"].Value = "Battery composition";
            sheetTemplate.Cells["DL3"].Value = "battery_cell_composition";
            //sheet.Cells["DM1"].Value = "";
            sheetTemplate.Cells["DM2"].Value = "Battery weight (grams)";
            sheetTemplate.Cells["DM3"].Value = "battery_weight";
            //sheet.Cells["DN1"].Value = "";
            sheetTemplate.Cells["DN2"].Value = "battery_weight_unit_of_measure";
            sheetTemplate.Cells["DN3"].Value = "battery_weight_unit_of_measure";
            //sheet.Cells["DO1"].Value = "";
            sheetTemplate.Cells["DO2"].Value = "lithium_battery_energy_content_unit_of_measure";
            sheetTemplate.Cells["DO3"].Value = "lithium_battery_energy_content_unit_of_measure";
            //sheet.Cells["DP1"].Value = "";
            sheetTemplate.Cells["DP2"].Value = "lithium_battery_weight_unit_of_measure";
            sheetTemplate.Cells["DP3"].Value = "lithium_battery_weight_unit_of_measure";
            //sheet.Cells["DQ1"].Value = "";
            sheetTemplate.Cells["DQ2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DQ3"].Value = "supplier_declared_dg_hz_regulation1";
            //sheet.Cells["DR1"].Value = "";
            sheetTemplate.Cells["DR2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DR3"].Value = "supplier_declared_dg_hz_regulation2";
            //sheet.Cells["DS1"].Value = "";
            sheetTemplate.Cells["DS2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DS3"].Value = "supplier_declared_dg_hz_regulation3";
            //sheet.Cells["DT1"].Value = "";
            sheetTemplate.Cells["DT2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DT3"].Value = "supplier_declared_dg_hz_regulation4";
            //sheet.Cells["DU1"].Value = "";
            sheetTemplate.Cells["DU2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DU3"].Value = "supplier_declared_dg_hz_regulation5";
            //sheet.Cells["DV1"].Value = "";
            sheetTemplate.Cells["DV2"].Value = "UN number";
            sheetTemplate.Cells["DV3"].Value = "hazmat_united_nations_regulatory_id";
            // sheet.Cells["DW1"].Value = "";
            sheetTemplate.Cells["DW2"].Value = "Safety Data Sheet (SDS) URL";
            sheetTemplate.Cells["DW3"].Value = "safety_data_sheet_url";
            // sheet.Cells["DX1"].Value = "";
            sheetTemplate.Cells["DX2"].Value = "Regulatory Organization Name";
            sheetTemplate.Cells["DX3"].Value = "legal_compliance_certification_regulatory_organization_name";
            // sheet.Cells["DY1"].Value = "";
            sheetTemplate.Cells["DY2"].Value = "Compliance Certification Status";
            sheetTemplate.Cells["DY3"].Value = "legal_compliance_certification_status";
            // sheet.Cells["DZ1"].Value = "";
            sheetTemplate.Cells["DZ2"].Value = "Flash point (°C)?";
            sheetTemplate.Cells["DZ3"].Value = "flash_point";
            //sheet.Cells["EA1"].Value = "";
            sheetTemplate.Cells["EA2"].Value = "Material/Fabric Regulations";
            sheetTemplate.Cells["EA3"].Value = "supplier_declared_material_regulation1";
            //sheet.Cells["EB1"].Value = "";
            sheetTemplate.Cells["EB2"].Value = "Material/Fabric Regulations";
            sheetTemplate.Cells["EB3"].Value = "supplier_declared_material_regulation2";
            //sheet.Cells["EC1"].Value = "";
            sheetTemplate.Cells["EC2"].Value = "Material/Fabric Regulations";
            sheetTemplate.Cells["EC3"].Value = "supplier_declared_material_regulation3";
            //sheet.Cells["ED1"].Value = "";
            sheetTemplate.Cells["ED2"].Value = "Legal Compliance Certification";
            sheetTemplate.Cells["ED3"].Value = "legal_compliance_certification_value";
            //sheet.Cells["EE1"].Value = "";
            sheetTemplate.Cells["EE2"].Value = "Categorization/GHS pictograms (select all that apply)";
            sheetTemplate.Cells["EE3"].Value = "ghs_classification_class1";
            //sheet.Cells["EF1"].Value = "";
            sheetTemplate.Cells["EF2"].Value = "Categorization/GHS pictograms (select all that apply)";
            sheetTemplate.Cells["EF3"].Value = "ghs_classification_class2";
            //sheet.Cells["EG1"].Value = "";
            sheetTemplate.Cells["EG2"].Value = "Categorization/GHS pictograms (select all that apply)";
            sheetTemplate.Cells["EG3"].Value = "ghs_classification_class3";
            //sheet.Cells["EH1"].Value = "";
            sheetTemplate.Cells["EH2"].Value = "California Proposition 65 Warning Type";
            sheetTemplate.Cells["EH3"].Value = "california_proposition_65_compliance_type";
            //sheet.Cells["EI1"].Value = "";
            sheetTemplate.Cells["EI2"].Value = "California Proposition 65 Chemical Names";
            sheetTemplate.Cells["EI3"].Value = "california_proposition_65_chemical_names1";
            //sheet.Cells["EJ1"].Value = "";
            sheetTemplate.Cells["EJ2"].Value = "Additional Chemical Name1";
            sheetTemplate.Cells["EJ3"].Value = "california_proposition_65_chemical_names2";
            //sheet.Cells["EK1"].Value = "";
            sheetTemplate.Cells["EK2"].Value = "Additional Chemical Name2";
            sheetTemplate.Cells["EK3"].Value = "california_proposition_65_chemical_names3";
            //sheet.Cells["EL1"].Value = "";
            sheetTemplate.Cells["EL2"].Value = "Additional Chemical Name3";
            sheetTemplate.Cells["EL3"].Value = "california_proposition_65_chemical_names4";
            //sheet.Cells["EM1"].Value = "";
            sheetTemplate.Cells["EM2"].Value = "Additional Chemical Name4";
            sheetTemplate.Cells["EM3"].Value = "california_proposition_65_chemical_names5";
            //sheet.Cells["EN1"].Value = "";
            sheetTemplate.Cells["EN2"].Value = "Pesticide Marking";
            sheetTemplate.Cells["EN3"].Value = "pesticide_marking_type1";
            //sheet.Cells["EO1"].Value = "";
            sheetTemplate.Cells["EO2"].Value = "Pesticide Marking";
            sheetTemplate.Cells["EO3"].Value = "pesticide_marking_type2";
            //sheet.Cells["EP1"].Value = "";
            sheetTemplate.Cells["EP2"].Value = "Pesticide Marking";
            sheetTemplate.Cells["EP3"].Value = "pesticide_marking_type3";
            //sheet.Cells["EQ1"].Value = "";
            sheetTemplate.Cells["EQ2"].Value = "Pesticide Registration Status";
            sheetTemplate.Cells["EQ3"].Value = "pesticide_marking_registration_status1";
            //sheet.Cells["ER1"].Value = "";
            sheetTemplate.Cells["ER2"].Value = "Pesticide Registration Status";
            sheetTemplate.Cells["ER3"].Value = "pesticide_marking_registration_status2";
            //sheet.Cells["ES1"].Value = "";
            sheetTemplate.Cells["ES2"].Value = "Pesticide Registration Status";
            sheetTemplate.Cells["ES3"].Value = "pesticide_marking_registration_status3";
            //sheet.Cells["ET1"].Value = "";
            sheetTemplate.Cells["ET2"].Value = "Pesticide Certification Number";
            sheetTemplate.Cells["ET3"].Value = "pesticide_marking_certification_number1";
            //sheet.Cells["EU1"].Value = "";
            sheetTemplate.Cells["EU2"].Value = "Pesticide Certification Number";
            sheetTemplate.Cells["EU3"].Value = "pesticide_marking_certification_number2";
            //sheet.Cells["EV1"].Value = "";
            sheetTemplate.Cells["EV2"].Value = "Pesticide Certification Number";
            sheetTemplate.Cells["EV3"].Value = "pesticide_marking_certification_number3";
            sheetTemplate.Cells["EW1"].Value = "Offer";
            sheetTemplate.Cells["EW2"].Value = "Shipping-Template";
            sheetTemplate.Cells["EW3"].Value = "merchant_shipping_group_name";
            // sheet.Cells["EX1"].Value = "";
            sheetTemplate.Cells["EX2"].Value = "Manufacturer's Suggested Retail Price";
            sheetTemplate.Cells["EX3"].Value = "list_price";
            // sheet.Cells["EY1"].Value = "";
            sheetTemplate.Cells["EY2"].Value = "Release Date";
            sheetTemplate.Cells["EY3"].Value = "merchant_release_date";
            // sheet.Cells["EZ1"].Value = "";
            sheetTemplate.Cells["EZ2"].Value = "Item Condition";
            sheetTemplate.Cells["EZ3"].Value = "condition_type";
            //sheet.Cells["FA1"].Value = "";
            sheetTemplate.Cells["FA2"].Value = "Restock Date";
            sheetTemplate.Cells["FA3"].Value = "restock_date";
            //sheet.Cells["FB1"].Value = "";
            sheetTemplate.Cells["FB2"].Value = "Handling Time";
            sheetTemplate.Cells["FB3"].Value = "fulfillment_latency";
            //sheet.Cells["FC1"].Value = "";
            sheetTemplate.Cells["FC2"].Value = "Offer Condition Note";
            sheetTemplate.Cells["FC3"].Value = "condition_note";
            //sheet.Cells["FD1"].Value = "";
            sheetTemplate.Cells["FD2"].Value = "Product Tax Code";
            sheetTemplate.Cells["FD3"].Value = "product_tax_code";
            //sheet.Cells["FE1"].Value = "";
            sheetTemplate.Cells["FE2"].Value = "Package Quantity";
            sheetTemplate.Cells["FE3"].Value = "item_package_quantity";
            //sheet.Cells["EF1"].Value = "";
            sheetTemplate.Cells["FF2"].Value = "Offering Can Be Gift Messaged";
            sheetTemplate.Cells["FF3"].Value = "offering_can_be_gift_messaged";
            //sheet.Cells["FG1"].Value = "";
            sheetTemplate.Cells["FG2"].Value = "Is Gift Wrap Available";
            sheetTemplate.Cells["FG3"].Value = "offering_can_be_giftwrapped";
            //sheet.Cells["FH1"].Value = "";
            sheetTemplate.Cells["FH2"].Value = "Max Order Quantity";
            sheetTemplate.Cells["FH3"].Value = "max_order_quantity";
            //sheet.Cells["FI1"].Value = "";
            sheetTemplate.Cells["FI2"].Value = "Number of Items";
            sheetTemplate.Cells["FI3"].Value = "number_of_items";
            sheetTemplate.Cells["FJ1"].Value = "Offer (US, CA, MX)";
            sheetTemplate.Cells["FJ2"].Value = "Sale Price USD (US)";
            sheetTemplate.Cells["FJ3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.discounted_price#1.schedule#1.value_with_tax";
            //sheet.Cells["FK1"].Value = "";
            sheetTemplate.Cells["FK2"].Value = "Sale Start Date (US)";
            sheetTemplate.Cells["FK3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.discounted_price#1.schedule#1.start_at";
            //sheet.Cells["FL1"].Value = "";
            sheetTemplate.Cells["FL2"].Value = "Sale End Date (US)";
            sheetTemplate.Cells["FL3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.discounted_price#1.schedule#1.end_at";
            //sheet.Cells["FM1"].Value = "";
            sheetTemplate.Cells["FM2"].Value = "Stop Selling Date (US)";
            sheetTemplate.Cells["FM3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.end_at.value";
            //sheet.Cells["FN1"].Value = "";
            sheetTemplate.Cells["FN2"].Value = "Your Price USD (US)";
            sheetTemplate.Cells["FN3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.our_price#1.schedule#1.value_with_tax";
            //sheet.Cells["FO1"].Value = "";
            sheetTemplate.Cells["FO2"].Value = "Offering Release Date (US)";
            sheetTemplate.Cells["FO3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.start_at.value";
            //sheet.Cells["FP1"].Value = "";
            sheetTemplate.Cells["FP2"].Value = "Sale Price CAD (CA)";
            sheetTemplate.Cells["FP3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.discounted_price#1.schedule#1.value_with_tax";
            //sheet.Cells["FQ1"].Value = "";
            sheetTemplate.Cells["FQ2"].Value = "Sale Start Date (CA)";
            sheetTemplate.Cells["FQ3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.discounted_price#1.schedule#1.start_at";
            //sheet.Cells["FR1"].Value = "";
            sheetTemplate.Cells["FR2"].Value = "Sale End Date (CA)";
            sheetTemplate.Cells["FR3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.discounted_price#1.schedule#1.end_at";
            //sheet.Cells["FS1"].Value = "";
            sheetTemplate.Cells["FS2"].Value = "Stop Selling Date (CA)";
            sheetTemplate.Cells["FS3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.end_at.value";
            //sheet.Cells["FT1"].Value = "";
            sheetTemplate.Cells["FT2"].Value = "Your Price CAD (CA)";
            sheetTemplate.Cells["FT3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.our_price#1.schedule#1.value_with_tax";
            //sheet.Cells["FU1"].Value = "";
            sheetTemplate.Cells["FU2"].Value = "Offering Release Date (CA)";
            sheetTemplate.Cells["FU3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.start_at.value";
            //sheet.Cells["FV1"].Value = "";
            sheetTemplate.Cells["FV2"].Value = "Sale Price MXN (MX)";
            sheetTemplate.Cells["FV3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.discounted_price#1.schedule#1.value_with_tax";
            //sheet.Cells["FW1"].Value = "";
            sheetTemplate.Cells["FW2"].Value = "Sale Start Date (MX)";
            sheetTemplate.Cells["FW3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.discounted_price#1.schedule#1.start_at";
            // sheet.Cells["FX1"].Value = "";
            sheetTemplate.Cells["FX2"].Value = "Sale End Date (MX)";
            sheetTemplate.Cells["FX3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.discounted_price#1.schedule#1.end_at";
            // sheet.Cells["FY1"].Value = "";
            sheetTemplate.Cells["FY2"].Value = "Stop Selling Date (MX)";
            sheetTemplate.Cells["FY3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.end_at.value";
            // sheet.Cells["FZ1"].Value = "";
            sheetTemplate.Cells["FZ2"].Value = "Your Price MXN (MX)";
            sheetTemplate.Cells["FZ3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.our_price#1.schedule#1.value_with_tax";
            //sheet.Cells["GA1"].Value = "";
            sheetTemplate.Cells["GA2"].Value = "Offering Release Date (MX)";
            sheetTemplate.Cells["GA3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.start_at.value";
            sheetTemplate.Cells["GB1"].Value = "B2B";
            sheetTemplate.Cells["GB2"].Value = "Business Price";
            sheetTemplate.Cells["GB3"].Value = "business_price";
            //sheet.Cells["GC1"].Value = "";
            sheetTemplate.Cells["GC2"].Value = "Quantity Price Type";
            sheetTemplate.Cells["GC3"].Value = "quantity_price_type";
            //sheet.Cells["GD1"].Value = "";
            sheetTemplate.Cells["GD2"].Value = "Quantity Lower Bound 1";
            sheetTemplate.Cells["GD3"].Value = "quantity_lower_bound1";
            //sheet.Cells["GE1"].Value = "";
            sheetTemplate.Cells["GE2"].Value = "Quantity Price 1";
            sheetTemplate.Cells["GE3"].Value = "quantity_price1";
            //sheet.Cells["GF1"].Value = "";
            sheetTemplate.Cells["GF2"].Value = "Quantity Lower Bound 2";
            sheetTemplate.Cells["GF3"].Value = "quantity_lower_bound2";
            //sheet.Cells["GG1"].Value = "";
            sheetTemplate.Cells["GG2"].Value = "Quantity Price 2";
            sheetTemplate.Cells["GG3"].Value = "quantity_price2";
            //sheet.Cells["GH1"].Value = "";
            sheetTemplate.Cells["GH2"].Value = "Quantity Lower Bound 3";
            sheetTemplate.Cells["GH3"].Value = "quantity_lower_bound3";
            //sheet.Cells["GI1"].Value = "";
            sheetTemplate.Cells["GI2"].Value = "Quantity Price 3";
            sheetTemplate.Cells["GI3"].Value = "quantity_price3";
            //sheet.Cells["GJ1"].Value = "";
            sheetTemplate.Cells["GJ2"].Value = "Quantity Lower Bound 4";
            sheetTemplate.Cells["GJ3"].Value = "quantity_lower_bound4";
            //sheet.Cells["GK1"].Value = "";
            sheetTemplate.Cells["GK2"].Value = "Quantity Price 4";
            sheetTemplate.Cells["GK3"].Value = "quantity_price4";
            //sheet.Cells["GL1"].Value = "";
            sheetTemplate.Cells["GL2"].Value = "Quantity Lower Bound 5";
            sheetTemplate.Cells["GL3"].Value = "quantity_lower_bound5";
            //sheet.Cells["GM1"].Value = "";
            sheetTemplate.Cells["GM2"].Value = "Quantity Price 5";
            sheetTemplate.Cells["GM3"].Value = "quantity_price5";
            //sheet.Cells["GN1"].Value = "";
            sheetTemplate.Cells["GN2"].Value = "National Stock Number";
            sheetTemplate.Cells["GN3"].Value = "national_stock_number";
            //sheet.Cells["GO1"].Value = "";
            sheetTemplate.Cells["GO2"].Value = "Progressive Discount Type";
            sheetTemplate.Cells["GO3"].Value = "progressive_discount_type";
            //sheet.Cells["GP1"].Value = "";
            sheetTemplate.Cells["GP2"].Value = "United Nations Standard Products and Services Code";
            sheetTemplate.Cells["GP3"].Value = "unspsc_code";
            //sheet.Cells["GQ1"].Value = "";
            sheetTemplate.Cells["GQ2"].Value = "Progressive Discount Lower Bound 1";
            sheetTemplate.Cells["GQ3"].Value = "progressive_discount_lower_bound1";
            //sheet.Cells["GR1"].Value = "";
            sheetTemplate.Cells["GR2"].Value = "Progressive Discount Value 1";
            sheetTemplate.Cells["GR3"].Value = "progressive_discount_value1";
            //sheet.Cells["GS1"].Value = "";
            sheetTemplate.Cells["GS2"].Value = "Pricing Action";
            sheetTemplate.Cells["GS3"].Value = "pricing_action";
            //sheet.Cells["GT1"].Value = "";
            sheetTemplate.Cells["GT2"].Value = "Progressive Discount Lower Bound 2";
            sheetTemplate.Cells["GT3"].Value = "progressive_discount_lower_bound2";
            //sheet.Cells["GU1"].Value = "";
            sheetTemplate.Cells["GU2"].Value = "Progressive Discount Value 2";
            sheetTemplate.Cells["GU3"].Value = "progressive_discount_value2";
            //sheet.Cells["GV1"].Value = "";
            sheetTemplate.Cells["GV2"].Value = "Progressive Discount Lower Bound 3";
            sheetTemplate.Cells["GV3"].Value = "progressive_discount_lower_bound3";
            //sheet.Cells["GW1"].Value = "";
            sheetTemplate.Cells["GW2"].Value = "Progressive Discount Value 3";
            sheetTemplate.Cells["GW3"].Value = "progressive_discount_value3";
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
