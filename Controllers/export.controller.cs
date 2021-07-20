
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using cip_api.models;
using System.Linq;
using OfficeOpenXml;
using System.IO;

namespace cip_api.controllers
{

    [ApiController, Route("[controller]"), Authorize]
    public class exportController : ControllerBase
    {

        private readonly Database db;
        private IConfiguration _config;

        public exportController(Database _db, IConfiguration config)
        {
            db = _db;
            _config = config;
        }

        [HttpGet]
        public ActionResult export()
        {
            try
            {
                string deptShortname = User.FindFirst("dept")?.Value;

                if (deptShortname.ToUpper() != "ACC")
                {
                    return Unauthorized(new { success = false, message = "Permission denied." });
                }

                List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status == "finished" && item.qty == "1").ToList();
                db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "finished" && (item.newBFMorAddBFM == "NEW BFM" || item.newBFMorAddBFM.ToLower().Trim() == "newbfm")).ToList();
                return Ok(new { success = true, data });


            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }
        [HttpPost("csv")]
        public ActionResult exportCSV(cip_api.request.ExportACCrequest body)
        {
            try
            {
                // string accFile = @"\\cptfile01\Dept\2310\G. Fixed Assets\Support ICD - CIP system";

                MemoryStream stream = new MemoryStream();

                using (ExcelPackage excel = new ExcelPackage(stream))
                {
                    List<string[]> header = new List<string[]>()
                {
                    new string[] { "ASSET-NO", "SUB-ASSET-NO", "ACC-DATE", "APPROVAL-NO", "OUTLINE", "TRF-TP", "ASSET-NO(AFTER)",
                    "SUB-ASSET-NO(AFTER)", "ASSET-TP", "SECONDHAND", "Admi Acct-CD", "G Prod Biz-CD", "Segment-CD", "CC-CD",
                    "Location-CD", "Location-NM", "ALLOCATION-CD", "CLASS-CD", "ASSET-NM-CD",
                    "ASSET-NM", "ASSET-NM-K", "ACQ-DATE", "OPE-DATE", "QUANTITY", "QUANTITY-UNIT", "AREA",
                    "AREA-UNIT", "DEPRE-METHOD(Company)", "SRV-LIFE(Company)", "DEPRE-MM(Company)", "DEPRE-METHOD(Tax)",
                    "SRV-LIFE(Tax)", "DEPRE-MM(Tax)", "DEPRE-METHOD(Group)", "SRV-LIFE(Group)", "DEPRE-MONTH(Group)", "DEPRE-METHOD(USGAAP)",
                    "SRV-LIFE(USGAAP)", "DEPRE-MONTH(USGAAP)", "DEPRE-METHOD(BOOK5)", "SRV-LIFE(BOOK5)", "DEPRE-MONTH(BOOK5)",
                    "DEPRE-METHOD(BOOK6)", "SRV-LIFE(BOOK6)", "DEPRE-MONTH(BOOK6)", "RES-MONTH(Company)", "BVAL-BOY(Company)", "UN-USE",
                    "1ST-CALC-TP(Company)", "RES-MONTH(Tax)", "BVAL-BOY(Tax)", "UN-USE", "1ST-CALC-TP(Tax)", "RES-MONTH(Group)", "BVAL-BOY(Group)",
                    "UN-USE", "1ST-CALC-TP(Group)", "RES-MONTH(USGAAP)", "BVAL-BOY(USGAAP)", "UN-USE", "1ST-CALC-TP(USGAAP)", "RES-MONTH(BOOK5)",
                    "BVAL-BOY(BOOK5)", "UN-USE", "1ST-CALC-TP(BOOK5)", "RES-MONTH(BOOK6)", "BVAL-BOY(BOOK6)", "UN-USE", "1ST-CALC-TP(BOOK6)",
                    "RES-RATE(Company)", "RES-VAL(Company)", "ADD-D-RATE(Company)", "DEPRE-RES-VAL(Company)", "RES-RATE(Tax)", "RES-VAL(Tax)",
                    "ADD-D-RATE(Tax)", "DEPRE-RES-VAL(Tax)", "RES-RATE(Group)", "RES-VAL(Group)", "ADD-DEPRE-RATE(Group)", "DEPRE-RES-VAL(Group)",
                    "RES-RATE(USGAAP)", "RES-VAL(USGAAP)", "ADD-DEPRE-RATE(USGAAP)", "DEPRE-RES-VAL(USGAAP)", "RES-RATE(BOOK5)", "RES-VAL(BOOK5)",
                    "ADD-DEPRE-RATE(BOOK5)", "DEPRE-RES-VAL(BOOK5)", "RES-RATE(BOOK6)", "RES-VAL(BOOK6)", "ADD-DEPRE-RATE(BOOK6)", "DEPRE-RES-VAL(BOOK6)",
                    "RETAILER-CD", "RETAILER-NM", "BORROWER-CD", "MNG-CD", "NOTE1", "NOTE2", "APPEND-OUT-TP", "G Org.", "Prod Ctgy",
                    "Allocation", "Invoice No", "Die No", "Process", "Not-in-Use", "AcceptDate", "Model Code", "Dimension",
                    "Order", "Order Ofc", "ORDER", "BOI", "ReacqPrice", "OPTION-CD(LEDGER)16", "Currency", "Org Amount", "OPTION-CD(LEDGER)19",
                    "OPTION-CD(LEDGER)20", "JRNL-INFO1", "JRNL-INFO2", "JRNL-INFO3", "JRNL-INFO4", "JRNL-INFO5", "CONTRA-ACC-CD",
                    "CONTRA-SUBACC-CD", "REP-CITY-CD", "REP-KIND", "REP-ACQ-VAL", "REP-LIFE", "REP-ADD-DRATE", "EXEM/TAX-FREE-CD", "MIC-KIND",
                    "MIC-DETAIL", "RDC-CD", "RDC-VAL", "RDC-BLNC-BOY", "UN-USE", "RDC-ALLOW", "SPE-DEPRE-CD", "GROUP-CD",
                    "SCENARIO-CD", "MAJOR-ASSET", "RDC-RES-VAL", "CTAX", "RO-ACC-TP"
                    }
                };
                    excel.Workbook.Worksheets.Add("CSV upload");
                    string headerRange = "A1:EQ1";
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets["CSV upload"];
                    worksheet.Cells[headerRange].LoadFromArrays(header);

                    string rootFolder = Directory.GetCurrentDirectory();
                    string pathString2 = @"\API site\files\CIP-system\download\";
                    string serverPath = rootFolder.Substring(0, rootFolder.LastIndexOf(@"\")) + pathString2;

                    if (!Directory.Exists(serverPath))
                    {
                        Directory.CreateDirectory(serverPath);
                    }
                    string fileName = System.Guid.NewGuid().ToString() + "-" + DateTime.Now.ToString("yyyy-MM-dd") + ".csv";
                    excel.SaveAs(new FileInfo(serverPath + fileName));

                    stream.Position = 0;

                }

                return Ok();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return Problem(e.StackTrace);
            }
        }
    }
}