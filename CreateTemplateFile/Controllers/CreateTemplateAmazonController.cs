using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Mvc;
using CreateTemplateFile.Models;
using OfficeOpenXml;

namespace CreateTemplateFile.Controllers
{
    public class CreateTemplateAmazonController : Controller
    {
        // GET: CreateTemplateAmazon
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public FileResult Index(
            InputData model,
            HttpPostedFileBase category,
            HttpPostedFileBase product)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var mapCategory = new ExcelPackage();
            mapCategory.Load(category.InputStream);
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
            sheetTemplate.Cells["I2"].Value = "Item Width Unit";
            sheetTemplate.Cells["I3"].Value = "width_shorter_edge_unit_of_measure";
            //sheet.Cells["J1"].Value = "";
            sheetTemplate.Cells["J2"].Value = "Item Width Side to Side";
            sheetTemplate.Cells["J3"].Value = "width_shorter_edge";
            //sheetTemplate.Cells["K1"].Value = "";
            sheetTemplate.Cells["K2"].Value = "Item Length Unit";
            sheetTemplate.Cells["K3"].Value = "length_longer_edge_unit_of_measure";
            //sheet.Cells["L1"].Value = "";
            sheetTemplate.Cells["L2"].Value = "Item Length Head to Toe";
            sheetTemplate.Cells["L3"].Value = "length_longer_edge";
            //sheet.Cells["M1"].Value = "";
            sheetTemplate.Cells["M2"].Value = "Quantity";
            sheetTemplate.Cells["M3"].Value = "quantity";
            //sheet.Cells["N1"].Value = "";
            sheetTemplate.Cells["N2"].Value = "Main Image URL";
            sheetTemplate.Cells["N3"].Value = "main_image_url";
            sheetTemplate.Cells["O1"].Value = "Images";
            sheetTemplate.Cells["O2"].Value = "Other Image URL1";
            sheetTemplate.Cells["O3"].Value = "other_image_url1";
            //sheet.Cells["P1"].Value = "";
            sheetTemplate.Cells["P2"].Value = "Other Image URL2";
            sheetTemplate.Cells["P3"].Value = "other_image_url2";
            //sheet.Cells["Q1"].Value = "";
            sheetTemplate.Cells["Q2"].Value = "Other Image URL3";
            sheetTemplate.Cells["Q3"].Value = "other_image_url3";
            //sheet.Cells["R1"].Value = "";
            sheetTemplate.Cells["R2"].Value = "Other Image URL4";
            sheetTemplate.Cells["R3"].Value = "other_image_url4";
            //sheet.Cells["S1"].Value = "";
            sheetTemplate.Cells["S2"].Value = "Other Image URL5";
            sheetTemplate.Cells["S3"].Value = "other_image_url5";
            //sheet.Cells["T1"].Value = "";
            sheetTemplate.Cells["T2"].Value = "Other Image URL6";
            sheetTemplate.Cells["T3"].Value = "other_image_url6";
            //sheet.Cells["U1"].Value = "";
            sheetTemplate.Cells["U2"].Value = "Other Image URL7";
            sheetTemplate.Cells["U3"].Value = "Other Image URL7";
            //sheet.Cells["V1"].Value = "";
            sheetTemplate.Cells["V2"].Value = "Other Image URL8";
            sheetTemplate.Cells["V3"].Value = "other_image_url8";
            //sheet.Cells["W1"].Value = "";
            sheetTemplate.Cells["W2"].Value = "Swatch Image URL";
            sheetTemplate.Cells["W3"].Value = "swatch_image_url";
            sheetTemplate.Cells["X1"].Value = "Variation";
            sheetTemplate.Cells["X2"].Value = "Parentage";
            sheetTemplate.Cells["X3"].Value = "parent_child";
            //sheet.Cells["Y1"].Value = "";
            sheetTemplate.Cells["Y2"].Value = "Parent SKU";
            sheetTemplate.Cells["Y3"].Value = "parent_sku";
            //sheet.Cells["Z1"].Value = "";
            sheetTemplate.Cells["Z2"].Value = "Relationship Type";
            sheetTemplate.Cells["Z3"].Value = "relationship_type";
            //sheet.Cells["AA1"].Value = "";
            sheetTemplate.Cells["AA2"].Value = "Variation Theme";
            sheetTemplate.Cells["AA3"].Value = "variation_theme";
            sheetTemplate.Cells["AB1"].Value = "Basic";
            sheetTemplate.Cells["AB2"].Value = "Update Delete";
            sheetTemplate.Cells["AB3"].Value = "update_delete";
            //sheet.Cells["AC1"].Value = "";
            sheetTemplate.Cells["AC2"].Value = "Product Exemption Reason";
            sheetTemplate.Cells["AC3"].Value = "gtin_exemption_reason";
            //sheet.Cells["AD1"].Value = "";
            sheetTemplate.Cells["AD2"].Value = "Product Description";
            sheetTemplate.Cells["AD3"].Value = "product_description";
            //sheet.Cells["AE1"].Value = "";
            sheetTemplate.Cells["AE2"].Value = "Manufacturer Part Number";
            sheetTemplate.Cells["AE3"].Value = "part_number";
            //sheet.Cells["AF1"].Value = "";
            sheetTemplate.Cells["AF2"].Value = "model";
            sheetTemplate.Cells["AF3"].Value = "model";
            //sheet.Cells["AG1"].Value = "";
            sheetTemplate.Cells["AG2"].Value = "Model Name";
            sheetTemplate.Cells["AG3"].Value = "model_name";
            //sheet.Cells["AH1"].Value = "";
            sheetTemplate.Cells["AH2"].Value = "Care Instructions";
            sheetTemplate.Cells["AH3"].Value = "care_instructions";
            sheetTemplate.Cells["AI1"].Value = "Discovery";
            sheetTemplate.Cells["AI2"].Value = "Key Product Features";
            sheetTemplate.Cells["AI3"].Value = "bullet_point1";
            //sheet.Cells["AJ1"].Value = "";
            sheetTemplate.Cells["AJ2"].Value = "Key Product Features";
            sheetTemplate.Cells["AJ3"].Value = "bullet_point2";
            //sheet.Cells["AK1"].Value = "";
            sheetTemplate.Cells["AK2"].Value = "Key Product Features";
            sheetTemplate.Cells["AK3"].Value = "bullet_point3";
            //sheet.Cells["AL1"].Value = "";
            sheetTemplate.Cells["AL2"].Value = "Key Product Features";
            sheetTemplate.Cells["AL3"].Value = "bullet_point4";
            //sheet.Cells["AM1"].Value = "";
            sheetTemplate.Cells["AM2"].Value = "Key Product Features";
            sheetTemplate.Cells["AM3"].Value = "bullet_point5";
            //sheet.Cells["AN1"].Value = "";
            sheetTemplate.Cells["AN2"].Value = "Search Terms";
            sheetTemplate.Cells["AN3"].Value = "generic_keywords";
            //sheet.Cells["AO1"].Value = "";
            sheetTemplate.Cells["AO2"].Value = "Country/Region as Labeled";
            sheetTemplate.Cells["AO3"].Value = "country_as_labeled";
            //sheet.Cells["AP1"].Value = "";
            sheetTemplate.Cells["AP2"].Value = "Number of Pieces";
            sheetTemplate.Cells["AP3"].Value = "number_of_pieces";
            //sheet.Cells["AQ1"].Value = "";
            sheetTemplate.Cells["AQ2"].Value = "Scent";
            sheetTemplate.Cells["AQ3"].Value = "scent_name";
            //sheet.Cells["AR1"].Value = "";
            sheetTemplate.Cells["AR2"].Value = "Included Components";
            sheetTemplate.Cells["AR3"].Value = "included_components";
            //sheet.Cells["AS1"].Value = "";
            sheetTemplate.Cells["AS2"].Value = "Color";
            sheetTemplate.Cells["AS3"].Value = "color_name";
            //sheet.Cells["AT1"].Value = "";
            sheetTemplate.Cells["AT2"].Value = "Color Map";
            sheetTemplate.Cells["AT3"].Value = "color_map";
            //sheet.Cells["AU1"].Value = "";
            sheetTemplate.Cells["AU2"].Value = "Size";
            sheetTemplate.Cells["AU3"].Value = "size_name";
            //sheet.Cells["AV1"].Value = "";
            sheetTemplate.Cells["AV2"].Value = "Material Type";
            sheetTemplate.Cells["AV3"].Value = "material_type";
            //sheet.Cells["AW1"].Value = "";
            sheetTemplate.Cells["AW2"].Value = "Wattage";
            sheetTemplate.Cells["AW3"].Value = "wattage";
            //sheet.Cells["AX1"].Value = "";
            sheetTemplate.Cells["AX2"].Value = "Additional Features";
            sheetTemplate.Cells["AX3"].Value = "special_features1";
            //sheet.Cells["AY1"].Value = "";
            sheetTemplate.Cells["AY2"].Value = "Additional Features";
            sheetTemplate.Cells["AY3"].Value = "special_features2";
            //sheet.Cells["AZ1"].Value = "";
            sheetTemplate.Cells["AZ2"].Value = "Additional Features";
            sheetTemplate.Cells["AZ3"].Value = "special_features3";
            // sheet.Cells["BA1"].Value = "";
            sheetTemplate.Cells["BA2"].Value = "Additional Features";
            sheetTemplate.Cells["BA3"].Value = "special_features4";
            // sheet.Cells["BB1"].Value = "";
            sheetTemplate.Cells["BB2"].Value = "Additional Features";
            sheetTemplate.Cells["BB3"].Value = "special_features5";
            // sheet.Cells["BC1"].Value = "";
            sheetTemplate.Cells["BC2"].Value = "Pattern";
            sheetTemplate.Cells["BC3"].Value = "pattern_name";
            // sheet.Cells["BD1"].Value = "";
            sheetTemplate.Cells["BD2"].Value = "Season of the Product";
            sheetTemplate.Cells["BD3"].Value = "seasons1";
            //sheet.Cells["BE1"].Value = "";
            sheetTemplate.Cells["BE2"].Value = "Season of the Product";
            sheetTemplate.Cells["BE3"].Value = "seasons2";
            //sheet.Cells["BF1"].Value = "";
            sheetTemplate.Cells["BF2"].Value = "Season of the Product";
            sheetTemplate.Cells["BF3"].Value = "seasons3";
            //sheet.Cells["BG1"].Value = "";
            sheetTemplate.Cells["BG2"].Value = "Season of the Product";
            sheetTemplate.Cells["BG3"].Value = "seasons4";
            //sheet.Cells["BH1"].Value = "";
            sheetTemplate.Cells["BH2"].Value = "Item Type";
            sheetTemplate.Cells["BH3"].Value = "item_type_name";
            //sheet.Cells["BI1"].Value = "";
            sheetTemplate.Cells["BI2"].Value = "Wattage Unit of Measure";
            sheetTemplate.Cells["BI3"].Value = "wattage_unit_of_measure";
            //sheet.Cells["BJ1"].Value = "";
            sheetTemplate.Cells["BJ2"].Value = "Length Range";
            sheetTemplate.Cells["BJ3"].Value = "length_range";
            //sheet.Cells["BK1"].Value = "";
            sheetTemplate.Cells["BK2"].Value = "Is Assembly Required";
            sheetTemplate.Cells["BK3"].Value = "is_assembly_required";
            //sheet.Cells["BL1"].Value = "";
            sheetTemplate.Cells["BL2"].Value = "Number of Shelves";
            sheetTemplate.Cells["BL3"].Value = "number_of_shelves";
            //sheet.Cells["BM1"].Value = "";
            sheetTemplate.Cells["BM2"].Value = "Number of Boxes";
            sheetTemplate.Cells["BM3"].Value = "number_of_boxes";
            //sheet.Cells["BN1"].Value = "";
            sheetTemplate.Cells["BN2"].Value = "Number of Compartments";
            sheetTemplate.Cells["BN3"].Value = "number_of_compartments";
            sheetTemplate.Cells["BO1"].Value = "Dimensions";
            sheetTemplate.Cells["BO2"].Value = "Towel Weight";
            sheetTemplate.Cells["BO3"].Value = "fabric_weight";
            //sheet.Cells["BP1"].Value = "";
            sheetTemplate.Cells["BP2"].Value = "Shape";
            sheetTemplate.Cells["BP3"].Value = "item_shape";
            //sheet.Cells["BQ1"].Value = "";
            sheetTemplate.Cells["BQ2"].Value = "Display Length Unit Of Measure";
            sheetTemplate.Cells["BQ3"].Value = "item_display_length_unit_of_measure";
            //sheet.Cells["BR1"].Value = "";
            sheetTemplate.Cells["BR2"].Value = "Item Display Width Unit Of Measure";
            sheetTemplate.Cells["BR3"].Value = "item_display_width_unit_of_measure";
            //sheet.Cells["BS1"].Value = "";
            sheetTemplate.Cells["BS2"].Value = "Item Display Height Unit Of Measure";
            sheetTemplate.Cells["BS3"].Value = "item_display_height_unit_of_measure";
            //sheet.Cells["BT1"].Value = "";
            sheetTemplate.Cells["BT2"].Value = "Item Display Length";
            sheetTemplate.Cells["BT3"].Value = "item_display_length";
            //sheet.Cells["BU1"].Value = "";
            sheetTemplate.Cells["BU2"].Value = "Item Display Width";
            sheetTemplate.Cells["BU3"].Value = "item_display_width";
            //sheet.Cells["BV1"].Value = "";
            sheetTemplate.Cells["BV2"].Value = "Item Display Height";
            sheetTemplate.Cells["BV3"].Value = "item_display_height";
            //sheet.Cells["BW1"].Value = "";
            sheetTemplate.Cells["BW2"].Value = "Item Display Weight";
            sheetTemplate.Cells["BW3"].Value = "item_display_weight";
            //sheet.Cells["BX1"].Value = "";
            sheetTemplate.Cells["BX2"].Value = "Item Display Weight Unit Of Measure";
            sheetTemplate.Cells["BX3"].Value = "item_display_weight_unit_of_measure";
            //sheet.Cells["BY1"].Value = "";
            sheetTemplate.Cells["BY2"].Value = "Item Height";
            sheetTemplate.Cells["BY3"].Value = "item_height";
            //sheet.Cells["BZ1"].Value = "";
            sheetTemplate.Cells["BZ2"].Value = "Item Length";
            sheetTemplate.Cells["BZ3"].Value = "item_length";
            // sheet.Cells["CA1"].Value = "";
            sheetTemplate.Cells["CA2"].Value = "Item Length Unit Of Measure";
            sheetTemplate.Cells["CA3"].Value = "item_length_unit_of_measure";
            // sheet.Cells["CC1"].Value = "";
            sheetTemplate.Cells["CB2"].Value = "Item Width";
            sheetTemplate.Cells["CB3"].Value = "item_width";
            // sheet.Cells["CD1"].Value = "";
            sheetTemplate.Cells["CC2"].Value = "Unit of Measure (Per Unit Pricing)";
            sheetTemplate.Cells["CC3"].Value = "unit_count_type";
            // sheet.Cells["CD1"].Value = "";
            sheetTemplate.Cells["CD2"].Value = "Unit Count (Per Unit Pricing)";
            sheetTemplate.Cells["CD3"].Value = "unit_count";
            //sheet.Cells["CE1"].Value = "";
            sheetTemplate.Cells["CE2"].Value = "Item Width Unit Of Measure";
            sheetTemplate.Cells["CE3"].Value = "item_width_unit_of_measure";
            //sheet.Cells["CF1"].Value = "";
            sheetTemplate.Cells["CF2"].Value = "Fabric Weight Unit";
            sheetTemplate.Cells["CF3"].Value = "fabric_weight_unit_of_measure";
            //sheet.Cells["CG1"].Value = "";
            sheetTemplate.Cells["CG2"].Value = "Size Map";
            sheetTemplate.Cells["CG3"].Value = "size_map";
            //sheet.Cells["CH1"].Value = "";
            sheetTemplate.Cells["CH2"].Value = "Width Range";
            sheetTemplate.Cells["CH3"].Value = "width_range";
            //sheet.Cells["CI1"].Value = "";
            sheetTemplate.Cells["CI2"].Value = "Maximum Weight Recommendation Unit Of Measure";
            sheetTemplate.Cells["CI3"].Value = "maximum_weight_recommendation_unit_of_measure";
            //sheet.Cells["CJ1"].Value = "";
            sheetTemplate.Cells["CJ2"].Value = "Item Height Unit Of Measure";
            sheetTemplate.Cells["CJ3"].Value = "item_height_unit_of_measure";
            //sheet.Cells["CK1"].Value = "";
            sheetTemplate.Cells["CK2"].Value = "Maximum Weight Recommendation";
            sheetTemplate.Cells["CK3"].Value = "maximum_weight_recommendation";
            sheetTemplate.Cells["CL1"].Value = "Fulfillment";
            sheetTemplate.Cells["CL2"].Value = "Fulfillment Center ID";
            sheetTemplate.Cells["CL3"].Value = "fulfillment_center_id";
            //sheet.Cells["CM1"].Value = "";
            sheetTemplate.Cells["CM2"].Value = "Package Height";
            sheetTemplate.Cells["CM3"].Value = "package_height";
            //sheet.Cells["CN1"].Value = "";
            sheetTemplate.Cells["CN2"].Value = "Package Width";
            sheetTemplate.Cells["CN3"].Value = "package_width";
            //sheet.Cells["CO1"].Value = "";
            sheetTemplate.Cells["CO2"].Value = "Package Length";
            sheetTemplate.Cells["CO3"].Value = "package_length";
            //sheet.Cells["CP1"].Value = "";
            sheetTemplate.Cells["CP2"].Value = "Package Length Unit Of Measure";
            sheetTemplate.Cells["CP3"].Value = "package_length_unit_of_measure";
            //sheet.Cells["CQ1"].Value = "";
            sheetTemplate.Cells["CQ2"].Value = "Package Weight";
            sheetTemplate.Cells["CQ3"].Value = "package_weight";
            //sheet.Cells["CR1"].Value = "";
            sheetTemplate.Cells["CR2"].Value = "Package Weight Unit Of Measure";
            sheetTemplate.Cells["CR3"].Value = "package_weight_unit_of_measure";
            //sheet.Cells["CS1"].Value = "";
            sheetTemplate.Cells["CS2"].Value = "Package Height Unit Of Measure";
            sheetTemplate.Cells["CS3"].Value = "package_height_unit_of_measure";
            //sheet.Cells["CT1"].Value = "";
            sheetTemplate.Cells["CT2"].Value = "Package Width Unit Of Measure";
            sheetTemplate.Cells["CT3"].Value = "package_width_unit_of_measure";
            sheetTemplate.Cells["CU1"].Value = "Compliance";
            sheetTemplate.Cells["CU2"].Value = "Manufacturer Warranty Description";
            sheetTemplate.Cells["CU3"].Value = "warranty_description";
            //sheet.Cells["CV1"].Value = "";
            sheetTemplate.Cells["CV2"].Value = "Cpsia Warning";
            sheetTemplate.Cells["CV3"].Value = "cpsia_cautionary_statement";
            //sheet.Cells["CW1"].Value = "";
            sheetTemplate.Cells["CW2"].Value = "Fabric Type";
            sheetTemplate.Cells["CW3"].Value = "fabric_type";
            //sheet.Cells["CX1"].Value = "";
            sheetTemplate.Cells["CX2"].Value = "Volume";
            sheetTemplate.Cells["CX3"].Value = "item_volume";
            //sheet.Cells["CY1"].Value = "";
            sheetTemplate.Cells["CY2"].Value = "item_volume_unit_of_measure";
            sheetTemplate.Cells["CY3"].Value = "item_volume_unit_of_measure";
            //sheet.Cells["CZ1"].Value = "";
            sheetTemplate.Cells["CZ2"].Value = "Country/Region of Origin";
            sheetTemplate.Cells["CZ3"].Value = "country_of_origin";
            // sheet.Cells["DA1"].Value = "";
            sheetTemplate.Cells["DA2"].Value = "Batteries are Included";
            sheetTemplate.Cells["DA3"].Value = "are_batteries_included";
            // sheet.Cells["DB1"].Value = "";
            sheetTemplate.Cells["DB2"].Value = "Item Weight";
            sheetTemplate.Cells["DB3"].Value = "item_weight";
            // sheet.Cells["DC1"].Value = "";
            sheetTemplate.Cells["DC2"].Value = "Is this product a battery or does it utilize batteries?";
            sheetTemplate.Cells["DC3"].Value = "batteries_required";
            // sheet.Cells["DD1"].Value = "";
            sheetTemplate.Cells["DD2"].Value = "Battery type/size";
            sheetTemplate.Cells["DD3"].Value = "battery_type1";
            //sheet.Cells["DE1"].Value = "";
            sheetTemplate.Cells["DE2"].Value = "Battery type/size";
            sheetTemplate.Cells["DE3"].Value = "battery_type2";
            //sheet.Cells["DF1"].Value = "";
            sheetTemplate.Cells["DF2"].Value = "Battery type/size";
            sheetTemplate.Cells["DF3"].Value = "battery_type3";
            //sheet.Cells["DG1"].Value = "";
            sheetTemplate.Cells["DG2"].Value = "item_weight_unit_of_measure";
            sheetTemplate.Cells["DG3"].Value = "item_weight_unit_of_measure";
            //sheet.Cells["DH1"].Value = "";
            sheetTemplate.Cells["DH2"].Value = "Number of batteries";
            sheetTemplate.Cells["DH3"].Value = "number_of_batteries1";
            //sheet.Cells["DI1"].Value = "";
            sheetTemplate.Cells["DI2"].Value = "Number of batteries";
            sheetTemplate.Cells["DI3"].Value = "number_of_batteries2";
            //sheet.Cells["DJ1"].Value = "";
            sheetTemplate.Cells["DJ2"].Value = "Number of batteries";
            sheetTemplate.Cells["DJ3"].Value = "number_of_batteries3";
            //sheet.Cells["DK1"].Value = "";
            sheetTemplate.Cells["DK2"].Value = "Watt hours per battery";
            sheetTemplate.Cells["DK3"].Value = "lithium_battery_energy_content";
            //sheet.Cells["DL1"].Value = "";
            sheetTemplate.Cells["DL2"].Value = "Lithium Battery Packaging";
            sheetTemplate.Cells["DL3"].Value = "lithium_battery_packaging";
            //sheet.Cells["DM1"].Value = "";
            sheetTemplate.Cells["DM2"].Value = "Lithium content (grams)";
            sheetTemplate.Cells["DM3"].Value = "lithium_battery_weight";
            //sheet.Cells["DN1"].Value = "";
            sheetTemplate.Cells["DN2"].Value = "Number of Lithium-ion Cells";
            sheetTemplate.Cells["DN3"].Value = "number_of_lithium_ion_cells";
            //sheet.Cells["DO1"].Value = "";
            sheetTemplate.Cells["DO2"].Value = "Number of Lithium Metal Cells";
            sheetTemplate.Cells["DO3"].Value = "number_of_lithium_metal_cells";
            //sheet.Cells["DP1"].Value = "";
            sheetTemplate.Cells["DP2"].Value = "Battery composition";
            sheetTemplate.Cells["DP3"].Value = "battery_cell_composition";
            //sheet.Cells["DQ1"].Value = "";
            sheetTemplate.Cells["DQ2"].Value = "Battery weight (grams)";
            sheetTemplate.Cells["DQ3"].Value = "battery_weight";
            //sheet.Cells["DR1"].Value = "";
            sheetTemplate.Cells["DR2"].Value = "battery_weight_unit_of_measure";
            sheetTemplate.Cells["DR3"].Value = "battery_weight_unit_of_measure";
            //sheet.Cells["DS1"].Value = "";
            sheetTemplate.Cells["DS2"].Value = "lithium_battery_energy_content_unit_of_measure";
            sheetTemplate.Cells["DS3"].Value = "lithium_battery_energy_content_unit_of_measure";
            //sheet.Cells["DT1"].Value = "";
            sheetTemplate.Cells["DT2"].Value = "lithium_battery_weight_unit_of_measure";
            sheetTemplate.Cells["DT3"].Value = "lithium_battery_weight_unit_of_measure";
            //sheet.Cells["DU1"].Value = "";
            sheetTemplate.Cells["DU2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DU3"].Value = "supplier_declared_dg_hz_regulation1";
            //sheet.Cells["DV1"].Value = "";
            sheetTemplate.Cells["DV2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DV3"].Value = "supplier_declared_dg_hz_regulation2";
            //sheet.Cells["DW1"].Value = "";
            sheetTemplate.Cells["DW2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DW3"].Value = "supplier_declared_dg_hz_regulation3";
            //sheet.Cells["DX1"].Value = "";
            sheetTemplate.Cells["DX2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DX3"].Value = "supplier_declared_dg_hz_regulation4";
            //sheet.Cells["DY1"].Value = "";
            sheetTemplate.Cells["DY2"].Value = "Applicable Dangerous Goods Regulations";
            sheetTemplate.Cells["DY3"].Value = "supplier_declared_dg_hz_regulation5";
            //sheet.Cells["DZ1"].Value = "";
            sheetTemplate.Cells["DZ2"].Value = "UN number";
            sheetTemplate.Cells["DZ3"].Value = "hazmat_united_nations_regulatory_id";
            // sheet.Cells["EA1"].Value = "";
            sheetTemplate.Cells["EA2"].Value = "Safety Data Sheet (SDS) URL";
            sheetTemplate.Cells["EA3"].Value = "safety_data_sheet_url";
            // sheet.Cells["EB1"].Value = "";
            sheetTemplate.Cells["EB2"].Value = "Regulatory Organization Name";
            sheetTemplate.Cells["EB3"].Value = "legal_compliance_certification_regulatory_organization_name";
            // sheet.Cells["EC1"].Value = "";
            sheetTemplate.Cells["EC2"].Value = "Compliance Certification Status";
            sheetTemplate.Cells["EC3"].Value = "legal_compliance_certification_status";
            // sheet.Cells["ED1"].Value = "";
            sheetTemplate.Cells["ED2"].Value = "Flash point (°C)?";
            sheetTemplate.Cells["ED3"].Value = "flash_point";
            //sheet.Cells["EE1"].Value = "";
            sheetTemplate.Cells["EE2"].Value = "Material/Fabric Regulations";
            sheetTemplate.Cells["EE3"].Value = "supplier_declared_material_regulation1";
            //sheet.Cells["EF1"].Value = "";
            sheetTemplate.Cells["EF2"].Value = "Material/Fabric Regulations";
            sheetTemplate.Cells["EF3"].Value = "supplier_declared_material_regulation2";
            //sheet.Cells["EG1"].Value = "";
            sheetTemplate.Cells["EG2"].Value = "Material/Fabric Regulations";
            sheetTemplate.Cells["EG3"].Value = "supplier_declared_material_regulation3";
            //sheet.Cells["EH1"].Value = "";
            sheetTemplate.Cells["EH2"].Value = "Legal Compliance Certification";
            sheetTemplate.Cells["EH3"].Value = "legal_compliance_certification_value";
            //sheet.Cells["EI1"].Value = "";
            sheetTemplate.Cells["EI2"].Value = "Categorization/GHS pictograms (select all that apply)";
            sheetTemplate.Cells["EI3"].Value = "ghs_classification_class1";
            //sheet.Cells["EJ1"].Value = "";
            sheetTemplate.Cells["EJ2"].Value = "Categorization/GHS pictograms (select all that apply)";
            sheetTemplate.Cells["EJ3"].Value = "ghs_classification_class2";
            //sheet.Cells["EK1"].Value = "";
            sheetTemplate.Cells["EK2"].Value = "Categorization/GHS pictograms (select all that apply)";
            sheetTemplate.Cells["EK3"].Value = "ghs_classification_class3";
            //sheet.Cells["EL1"].Value = "";
            sheetTemplate.Cells["EL2"].Value = "California Proposition 65 Warning Type";
            sheetTemplate.Cells["EL3"].Value = "california_proposition_65_compliance_type";
            //sheet.Cells["EM1"].Value = "";
            sheetTemplate.Cells["EM2"].Value = "California Proposition 65 Chemical Names";
            sheetTemplate.Cells["EM3"].Value = "california_proposition_65_chemical_names1";
            //sheet.Cells["EN1"].Value = "";
            sheetTemplate.Cells["EN2"].Value = "Additional Chemical Name1";
            sheetTemplate.Cells["EN3"].Value = "california_proposition_65_chemical_names2";
            //sheet.Cells["EO1"].Value = "";
            sheetTemplate.Cells["EO2"].Value = "Additional Chemical Name2";
            sheetTemplate.Cells["EO3"].Value = "california_proposition_65_chemical_names3";
            //sheet.Cells["EP1"].Value = "";
            sheetTemplate.Cells["EP2"].Value = "Additional Chemical Name3";
            sheetTemplate.Cells["EP3"].Value = "california_proposition_65_chemical_names4";
            //sheet.Cells["EQ1"].Value = "";
            sheetTemplate.Cells["EQ2"].Value = "Additional Chemical Name4";
            sheetTemplate.Cells["EQ3"].Value = "california_proposition_65_chemical_names5";
            //sheet.Cells["ER1"].Value = "";
            sheetTemplate.Cells["ER2"].Value = "Pesticide Marking";
            sheetTemplate.Cells["ER3"].Value = "pesticide_marking_type1";
            //sheet.Cells["ES1"].Value = "";
            sheetTemplate.Cells["ES2"].Value = "Pesticide Marking";
            sheetTemplate.Cells["ES3"].Value = "pesticide_marking_type2";
            //sheet.Cells["ET1"].Value = "";
            sheetTemplate.Cells["ET2"].Value = "Pesticide Marking";
            sheetTemplate.Cells["ET3"].Value = "pesticide_marking_type3";
            //sheet.Cells["EU1"].Value = "";
            sheetTemplate.Cells["EU2"].Value = "Pesticide Registration Status";
            sheetTemplate.Cells["EU3"].Value = "pesticide_marking_registration_status1";
            //sheet.Cells["EV1"].Value = "";
            sheetTemplate.Cells["EV2"].Value = "Pesticide Registration Status";
            sheetTemplate.Cells["EV3"].Value = "pesticide_marking_registration_status2";
            //sheet.Cells["EW1"].Value = "";
            sheetTemplate.Cells["EW2"].Value = "Pesticide Registration Status";
            sheetTemplate.Cells["EW3"].Value = "pesticide_marking_registration_status3";
            //sheet.Cells["EX1"].Value = "";
            sheetTemplate.Cells["EX2"].Value = "Pesticide Certification Number";
            sheetTemplate.Cells["EX3"].Value = "pesticide_marking_certification_number1";
            //sheet.Cells["EY1"].Value = "";
            sheetTemplate.Cells["EY2"].Value = "Pesticide Certification Number";
            sheetTemplate.Cells["EY3"].Value = "pesticide_marking_certification_number2";
            //sheet.Cells["EZ1"].Value = "";
            sheetTemplate.Cells["EZ2"].Value = "Pesticide Certification Number";
            sheetTemplate.Cells["EZ3"].Value = "pesticide_marking_certification_number3";
            sheetTemplate.Cells["FA1"].Value = "Offer";
            sheetTemplate.Cells["FA2"].Value = "Shipping-Template";
            sheetTemplate.Cells["FA3"].Value = "merchant_shipping_group_name";
            // sheet.Cells["FB1"].Value = "";
            sheetTemplate.Cells["FB2"].Value = "Manufacturer's Suggested Retail Price";
            sheetTemplate.Cells["FB3"].Value = "list_price";
            // sheet.Cells["FC1"].Value = "";
            sheetTemplate.Cells["FC2"].Value = "Release Date";
            sheetTemplate.Cells["FC3"].Value = "merchant_release_date";
            // sheet.Cells["FD1"].Value = "";
            sheetTemplate.Cells["FD2"].Value = "Item Condition";
            sheetTemplate.Cells["FD3"].Value = "condition_type";
            //sheet.Cells["FE1"].Value = "";
            sheetTemplate.Cells["FE2"].Value = "Restock Date";
            sheetTemplate.Cells["FE3"].Value = "restock_date";
            //sheet.Cells["FF1"].Value = "";
            sheetTemplate.Cells["FF2"].Value = "Handling Time";
            sheetTemplate.Cells["FF3"].Value = "fulfillment_latency";
            //sheet.Cells["FG1"].Value = "";
            sheetTemplate.Cells["FG2"].Value = "Offer Condition Note";
            sheetTemplate.Cells["FG3"].Value = "condition_note";
            //sheet.Cells["FH1"].Value = "";
            sheetTemplate.Cells["FH2"].Value = "Product Tax Code";
            sheetTemplate.Cells["FH3"].Value = "product_tax_code";
            //sheet.Cells["FI1"].Value = "";
            sheetTemplate.Cells["FI2"].Value = "Package Quantity";
            sheetTemplate.Cells["FI3"].Value = "item_package_quantity";
            //sheet.Cells["EJ1"].Value = "";
            sheetTemplate.Cells["FJ2"].Value = "Offering Can Be Gift Messaged";
            sheetTemplate.Cells["FJ3"].Value = "offering_can_be_gift_messaged";
            //sheet.Cells["FK1"].Value = "";
            sheetTemplate.Cells["FK2"].Value = "Is Gift Wrap Available";
            sheetTemplate.Cells["FK3"].Value = "offering_can_be_giftwrapped";
            //sheet.Cells["FL1"].Value = "";
            sheetTemplate.Cells["FL2"].Value = "Max Order Quantity";
            sheetTemplate.Cells["FL3"].Value = "max_order_quantity";
            //sheet.Cells["FM1"].Value = "";
            sheetTemplate.Cells["FM2"].Value = "Number of Items";
            sheetTemplate.Cells["FM3"].Value = "number_of_items";
            sheetTemplate.Cells["FN1"].Value = "Offer (US, CA, MX)";
            sheetTemplate.Cells["FN2"].Value = "Sale Price USD (US)";
            sheetTemplate.Cells["FN3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.discounted_price#1.schedule#1.value_with_tax";
            //sheet.Cells["FO1"].Value = "";
            sheetTemplate.Cells["FO2"].Value = "Sale Start Date (US)";
            sheetTemplate.Cells["FO3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.discounted_price#1.schedule#1.start_at";
            //sheet.Cells["FP1"].Value = "";
            sheetTemplate.Cells["FP2"].Value = "Sale End Date (US)";
            sheetTemplate.Cells["FP3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.discounted_price#1.schedule#1.end_at";
            //sheet.Cells["FQ1"].Value = "";
            sheetTemplate.Cells["FQ2"].Value = "Stop Selling Date (US)";
            sheetTemplate.Cells["FQ3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.end_at.value";
            //sheet.Cells["FR1"].Value = "";
            sheetTemplate.Cells["FR2"].Value = "Your Price USD (US)";
            sheetTemplate.Cells["FR3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.our_price#1.schedule#1.value_with_tax";
            //sheet.Cells["FS1"].Value = "";
            sheetTemplate.Cells["FS2"].Value = "Offering Release Date (US)";
            sheetTemplate.Cells["FS3"].Value = "purchasable_offer[marketplace_id=ATVPDKIKX0DER]#1.start_at.value";
            //sheet.Cells["FT1"].Value = "";
            sheetTemplate.Cells["FT2"].Value = "Sale Price CAD (CA)";
            sheetTemplate.Cells["FT3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.discounted_price#1.schedule#1.value_with_tax";
            //sheet.Cells["FU1"].Value = "";
            sheetTemplate.Cells["FU2"].Value = "Sale Start Date (CA)";
            sheetTemplate.Cells["FU3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.discounted_price#1.schedule#1.start_at";
            //sheet.Cells["FV1"].Value = "";
            sheetTemplate.Cells["FV2"].Value = "Sale End Date (CA)";
            sheetTemplate.Cells["FV3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.discounted_price#1.schedule#1.end_at";
            //sheet.Cells["FW1"].Value = "";
            sheetTemplate.Cells["FW2"].Value = "Stop Selling Date (CA)";
            sheetTemplate.Cells["FW3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.end_at.value";
            //sheet.Cells["FX1"].Value = "";
            sheetTemplate.Cells["FX2"].Value = "Your Price CAD (CA)";
            sheetTemplate.Cells["FX3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.our_price#1.schedule#1.value_with_tax";
            //sheet.Cells["FY1"].Value = "";
            sheetTemplate.Cells["FY2"].Value = "Offering Release Date (CA)";
            sheetTemplate.Cells["FY3"].Value = "purchasable_offer[marketplace_id=A2EUQ1WTGCTBG2]#1.start_at.value";
            //sheet.Cells["FZ1"].Value = "";
            sheetTemplate.Cells["FZ2"].Value = "Sale Price MXN (MX)";
            sheetTemplate.Cells["FZ3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.discounted_price#1.schedule#1.value_with_tax";
            //sheet.Cells["GA1"].Value = "";
            sheetTemplate.Cells["GA2"].Value = "Sale Start Date (MX)";
            sheetTemplate.Cells["GA3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.discounted_price#1.schedule#1.start_at";
            // sheet.Cells["GB1"].Value = "";
            sheetTemplate.Cells["GB2"].Value = "Sale End Date (MX)";
            sheetTemplate.Cells["GB3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.discounted_price#1.schedule#1.end_at";
            // sheet.Cells["GC1"].Value = "";
            sheetTemplate.Cells["GC2"].Value = "Stop Selling Date (MX)";
            sheetTemplate.Cells["GC3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.end_at.value";
            // sheet.Cells["GD1"].Value = "";
            sheetTemplate.Cells["GD2"].Value = "Your Price MXN (MX)";
            sheetTemplate.Cells["GD3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.our_price#1.schedule#1.value_with_tax";
            //sheet.Cells["GE1"].Value = "";
            sheetTemplate.Cells["GE2"].Value = "Offering Release Date (MX)";
            sheetTemplate.Cells["GE3"].Value = "purchasable_offer[marketplace_id=A1AM78C64UM0Y8]#1.start_at.value";
            sheetTemplate.Cells["GF1"].Value = "B2B";
            sheetTemplate.Cells["GF2"].Value = "Business Price";
            sheetTemplate.Cells["GF3"].Value = "business_price";
            //sheet.Cells["GG1"].Value = "";
            sheetTemplate.Cells["GG2"].Value = "Quantity Price Type";
            sheetTemplate.Cells["GG3"].Value = "quantity_price_type";
            //sheet.Cells["GH1"].Value = "";
            sheetTemplate.Cells["GH2"].Value = "Quantity Lower Bound 1";
            sheetTemplate.Cells["GH3"].Value = "quantity_lower_bound1";
            //sheet.Cells["GI1"].Value = "";
            sheetTemplate.Cells["GI2"].Value = "Quantity Price 1";
            sheetTemplate.Cells["GI3"].Value = "quantity_price1";
            //sheet.Cells["GJ1"].Value = "";
            sheetTemplate.Cells["GJ2"].Value = "Quantity Lower Bound 2";
            sheetTemplate.Cells["GJ3"].Value = "quantity_lower_bound2";
            //sheet.Cells["GK1"].Value = "";
            sheetTemplate.Cells["GK2"].Value = "Quantity Price 2";
            sheetTemplate.Cells["GK3"].Value = "quantity_price2";
            //sheet.Cells["GL1"].Value = "";
            sheetTemplate.Cells["GL2"].Value = "Quantity Lower Bound 3";
            sheetTemplate.Cells["GL3"].Value = "quantity_lower_bound3";
            //sheet.Cells["GM1"].Value = "";
            sheetTemplate.Cells["GM2"].Value = "Quantity Price 3";
            sheetTemplate.Cells["GM3"].Value = "quantity_price3";
            //sheet.Cells["GN1"].Value = "";
            sheetTemplate.Cells["GN2"].Value = "Quantity Lower Bound 4";
            sheetTemplate.Cells["GN3"].Value = "quantity_lower_bound4";
            //sheet.Cells["GO1"].Value = "";
            sheetTemplate.Cells["GO2"].Value = "Quantity Price 4";
            sheetTemplate.Cells["GO3"].Value = "quantity_price4";
            //sheet.Cells["GP1"].Value = "";
            sheetTemplate.Cells["GP2"].Value = "Quantity Lower Bound 5";
            sheetTemplate.Cells["GP3"].Value = "quantity_lower_bound5";
            //sheet.Cells["GQ1"].Value = "";
            sheetTemplate.Cells["GQ2"].Value = "Quantity Price 5";
            sheetTemplate.Cells["GQ3"].Value = "quantity_price5";
            //sheet.Cells["GR1"].Value = "";
            sheetTemplate.Cells["GR2"].Value = "National Stock Number";
            sheetTemplate.Cells["GR3"].Value = "national_stock_number";
            //sheet.Cells["GS1"].Value = "";
            sheetTemplate.Cells["GS2"].Value = "Progressive Discount Type";
            sheetTemplate.Cells["GS3"].Value = "progressive_discount_type";
            //sheet.Cells["GT1"].Value = "";
            sheetTemplate.Cells["GT2"].Value = "United Nations Standard Products and Services Code";
            sheetTemplate.Cells["GT3"].Value = "unspsc_code";
            //sheet.Cells["GU1"].Value = "";
            sheetTemplate.Cells["GU2"].Value = "Progressive Discount Lower Bound 1";
            sheetTemplate.Cells["GU3"].Value = "progressive_discount_lower_bound1";
            //sheet.Cells["GV1"].Value = "";
            sheetTemplate.Cells["GV2"].Value = "Progressive Discount Value 1";
            sheetTemplate.Cells["GV3"].Value = "progressive_discount_value1";
            //sheet.Cells["GW1"].Value = "";
            sheetTemplate.Cells["GW2"].Value = "Pricing Action";
            sheetTemplate.Cells["GW3"].Value = "pricing_action";
            //sheet.Cells["GX1"].Value = "";
            sheetTemplate.Cells["GX2"].Value = "Progressive Discount Lower Bound 2";
            sheetTemplate.Cells["GX3"].Value = "progressive_discount_lower_bound2";
            //sheet.Cells["GY1"].Value = "";
            sheetTemplate.Cells["GY2"].Value = "Progressive Discount Value 2";
            sheetTemplate.Cells["GY3"].Value = "progressive_discount_value2";
            //sheet.Cells["GZ1"].Value = "";
            sheetTemplate.Cells["GZ2"].Value = "Progressive Discount Lower Bound 3";
            sheetTemplate.Cells["GZ3"].Value = "progressive_discount_lower_bound3";
            //sheet.Cells["HA1"].Value = "";
            sheetTemplate.Cells["HA2"].Value = "Progressive Discount Value 3";
            sheetTemplate.Cells["HA3"].Value = "progressive_discount_value3";
            #endregion

            #region dữ liệm điền tay

            var a = model.FeedProduct;
            var generic = "Generic";
            var aC = "Manufacture on Demand"; 
            var h = model.ItemType;
            var aa = model.VariationTheme;
            var aD = model.ProductDescription;
            var aJ = model.BulletPoint1;
            var aI = model.BulletPoint2;
            var aK = model.BulletPoint3;
            var aL = model.BulletPoint4;
            var aM = model.BulletPoint5;
            var aN = model.GenericKeywords;
            var displayLength = "IN";
            var itemShape = model.ItemShape;
            var fA = model.ShippingTemplate;
            var quantity = model.Quantity;
            var inches = "Inches";
           
            #endregion


            var excelProduct = new ExcelPackage();
            excelProduct.Load(product.InputStream);
            var sheetProduct = excelProduct.Workbook.Worksheets.FirstOrDefault();
            if (sheetProduct != null)
            {
                var row = sheetProduct.Dimension.Rows;
                var rowCategory = sheetMap.Dimension.Rows;
                var rowBegin = 4;

                for (int i = 2; i <= row; i++)
                {
                    var variants = new List<Variant>();
                    for (int j = 2; j <= rowCategory; j++)
                    {
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
                            for (int k = 0; k <= variants.Count; k++)
                            {
                                sheetTemplate.Cells[rowBegin + k, 14 + j].Value = images[j];

                            }

                        }

                        // diền  T
                        sheetTemplate.Cells["X" + rowBegin].Value = "Parent";
                        

                        foreach (var item in variants)
                        {
                            var id = item.Sku + ThreadSafeRandom.ThisThreadsRandom.Next(10000, 99999);
                            sheetTemplate.Cells["B" + rowBegin].Value = id;
                            sheetTemplate.Cells["D" + rowBegin].Value = name;
                            sheetTemplate.Cells["Z" + rowBegin].Value = "Variation";
                            //lấy sku 
                            for (int j = 1; j <= variants.Count; j++)
                            {
                                sheetTemplate.Cells["B" + (rowBegin + j)].Value =
                                    item.Sku + ThreadSafeRandom.ThisThreadsRandom.Next(10000, 99999); 
                                sheetTemplate.Cells["Y" + (rowBegin + j)].Value = id;
                                sheetTemplate.Cells["D" + (rowBegin + j)].Value = name;
                                sheetTemplate.Cells["Z" + (rowBegin + j)].Value = "Variation";
                                sheetTemplate.Cells["X" + (rowBegin + j)].Value = "Child";
                                sheetTemplate.Cells["M" + (rowBegin + j)].Value = quantity;
                                sheetTemplate.Cells["BQ" + (rowBegin + j)].Value = displayLength;
                                sheetTemplate.Cells["BR" + (rowBegin + j)].Value = displayLength;
                                sheetTemplate.Cells["BP" + (rowBegin + j)].Value = itemShape;
                                sheetTemplate.Cells["I" + (rowBegin + j)].Value = inches;
                                sheetTemplate.Cells["K" + (rowBegin + j)].Value = inches;


                            }
                            
                        

                        }

                    }
                    // fill du lieu dien tay
                    for (int j = 0; j <= variants.Count; j++)
                    {
                        sheetTemplate.Cells["A" + (rowBegin + j)].Value = a;
                        sheetTemplate.Cells["C" + (rowBegin + j)].Value = generic;
                        sheetTemplate.Cells["E" + (rowBegin + j)].Value = generic;
                        sheetTemplate.Cells["H" + (rowBegin + j)].Value = h;
                        sheetTemplate.Cells["AA" + (rowBegin + j)].Value = aa;
                        sheetTemplate.Cells["AC" + (rowBegin + j)].Value = aC;
                        sheetTemplate.Cells["AD" + (rowBegin + j)].Value = aD;
                        sheetTemplate.Cells["AI" + (rowBegin + j)].Value = aI;
                        sheetTemplate.Cells["AJ" + (rowBegin + j)].Value = aJ;
                        sheetTemplate.Cells["AK" + (rowBegin + j)].Value = aK;
                        sheetTemplate.Cells["AL" + (rowBegin + j)].Value = aL;
                        sheetTemplate.Cells["AM" + (rowBegin + j)].Value = aM;
                        sheetTemplate.Cells["AN" + (rowBegin + j)].Value = aN;
                        sheetTemplate.Cells["FA" + (rowBegin + j)].Value = fA;
                    }

                    for (int j = 2; j <= rowCategory; j++)
                    {

                        // Lấy FR
                        var pirce = sheetMap.Cells[j, 1].Value;
                        sheetTemplate.Cells["FR" + (rowBegin + 1)].Value = pirce;
                        // lay dc size AU
                        var size = sheetMap.Cells[j, 2].Value;
                        sheetTemplate.Cells["AU" + (rowBegin + 1)].Value = size;

                        // lay BT
                        var itemDisplayLength = sheetMap.Cells[j, 3].Value;
                        sheetTemplate.Cells["BT" + (rowBegin + 1)].Value = itemDisplayLength;
                        sheetTemplate.Cells["J" + (rowBegin + 1)].Value = itemDisplayLength;
                        // lay BU
                        var itemDisplayWidth = sheetMap.Cells[j, 4].Value;
                        sheetTemplate.Cells["BU" + (rowBegin + 1)].Value = itemDisplayWidth;
                        sheetTemplate.Cells["L" + (rowBegin + 1)].Value = itemDisplayWidth;

                        rowBegin++;


                    }

                    rowBegin += 1;
                }
            }


            return File(excelTemplate.GetAsByteArray(), "application/vnd.ms-excel", "amz.xlsx");
        }



        public string GenSkuCode()
        {

            return "SKU" + "-" + ThreadSafeRandom.ThisThreadsRandom.Next(10000, 99999) + "-";
        }

       


    }
}