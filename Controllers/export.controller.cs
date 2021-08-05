
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

        [HttpPost("finished")]
        public ActionResult getFinishedItems(request.ExportACCrequest body)
        {
            try
            {
                string deptShortname = User.FindFirst("dept")?.Value;

                if (deptShortname.ToUpper() != "ACC")
                {
                    return Unauthorized(new { success = false, message = "Permission denied." });
                }

                List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status == "finished" && item.workType == body.workType).ToList();
                db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "finished").ToList();

                List<cipSchema> returnData = data.FindAll(item => item.cipUpdate != null);

                return Ok(new { success = true, data = returnData });


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

        [HttpPost("wrtiing"), AllowAnonymous]
        public ActionResult WriteTodept(request.exportToForm body)
        {
            try
            {
                string userfile = @"\\cptfile01\Dept\2310\G. Fixed Assets\1.Fixed Asset - Confirm CIP (Domestic+Oversea)\Master file";
                // string userfile = "C:\\Users\\013817\\Desktop";
                string sheetName = "";
                if (body.workType == "Project ENG3")
                {
                    userfile += "\\Project_ENG3.xlsx";
                    sheetName = "SUM-ACC";
                }
                else if (body.workType == "Domestic-DIE")
                {
                    userfile += "\\Domestic_DIE.xlsx";
                    sheetName = "SUM-ACC";
                }
                else if (body.workType == "Domestic")
                {
                    userfile += "\\Domestic.xlsx";
                    sheetName = "SUM-ACC";
                }
                else if (body.workType == "Oversea")
                {
                    userfile += "\\Oversea(StepB).xlsx";
                    sheetName = "STEP B";
                }
                else if (body.workType == "Project-MSC")
                {
                    userfile += "\\Project_MSC.xlsx";
                    sheetName = "SUM-ACC";
                }


                FileInfo Existfile = new FileInfo(userfile);
                List<string> errorCip = new List<string>();
                List<string> successCip = new List<string>();
                using (ExcelPackage package = new ExcelPackage(new FileInfo(userfile)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                    int rowEnd = worksheet.Dimension.End.Row;

                    List<cipSchema> updateStatus = new List<cipSchema>();
                    foreach (string id in body.id)
                    {
                        cipSchema cip = db.CIP.Find(Int32.Parse(id));
                        db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == Int32.Parse(id)).FirstOrDefault();
                        try
                        {
                            if (cip != null)
                            {
                                if (body.workType == "Project ENG3")
                                {
                                    var searchCell = from cell in worksheet.Cells["E1:E" + rowEnd.ToString()] where cell.Value.ToString() == cip.cipNo select cell.Start.Row;
                                    string rowNumber = searchCell.First().ToString();
                                    worksheet.Cells["AC" + rowNumber].Value = cip.cipUpdate.planDate;
                                    worksheet.Cells["AD" + rowNumber].Value = cip.cipUpdate.actDate;
                                    worksheet.Cells["AE" + rowNumber].Value = cip.cipUpdate.result;
                                    worksheet.Cells["AF" + rowNumber].Value = cip.cipUpdate.reasonDiff;
                                    worksheet.Cells["AG" + rowNumber].Value = cip.cipUpdate.fixedAssetCode;
                                    worksheet.Cells["AH" + rowNumber].Value = cip.cipUpdate.classFixedAsset;
                                    worksheet.Cells["AI" + rowNumber].Value = cip.cipUpdate.fixAssetName;
                                    worksheet.Cells["AJ" + rowNumber].Value = cip.cipUpdate.serialNo;
                                    worksheet.Cells["AK" + rowNumber].Value = cip.cipUpdate.partNumberDieNo;
                                    worksheet.Cells["AL" + rowNumber].Value = cip.cipUpdate.processDie;
                                    worksheet.Cells["AM" + rowNumber].Value = cip.cipUpdate.model;
                                    worksheet.Cells["AN" + rowNumber].Value = cip.cipUpdate.costCenterOfUser;
                                    worksheet.Cells["AO" + rowNumber].Value = cip.cipUpdate.tranferToSupplier;
                                    worksheet.Cells["AP" + rowNumber].Value = cip.cipUpdate.upFixAsset;
                                    worksheet.Cells["AQ" + rowNumber].Value = cip.cipUpdate.newBFMorAddBFM;
                                    worksheet.Cells["AR" + rowNumber].Value = cip.cipUpdate.reasonForDelay;
                                    worksheet.Cells["AS" + rowNumber].Value = cip.cipUpdate.addCipBfmNo;
                                    worksheet.Cells["AT" + rowNumber].Value = cip.cipUpdate.remark;
                                }
                                else if (body.workType == "Domestic-DIE")
                                {
                                    var searchCell = from cell in worksheet.Cells["C1:C" + rowEnd.ToString()] where cell.Value.ToString() == cip.cipNo select cell.Start.Row;
                                    string rowNumber = searchCell.First().ToString();

                                    worksheet.Cells["AG" + rowNumber].Value = cip.cipUpdate.planDate;
                                    worksheet.Cells["AH" + rowNumber].Value = cip.cipUpdate.actDate;
                                    worksheet.Cells["AI" + rowNumber].Value = cip.cipUpdate.result;
                                    worksheet.Cells["AJ" + rowNumber].Value = cip.cipUpdate.reasonDiff;
                                    worksheet.Cells["AK" + rowNumber].Value = cip.cipUpdate.fixedAssetCode;
                                    worksheet.Cells["AL" + rowNumber].Value = cip.cipUpdate.classFixedAsset;
                                    worksheet.Cells["AM" + rowNumber].Value = cip.cipUpdate.fixAssetName;
                                    worksheet.Cells["AN" + rowNumber].Value = cip.cipUpdate.serialNo;
                                    worksheet.Cells["AO" + rowNumber].Value = cip.cipUpdate.partNumberDieNo;
                                    worksheet.Cells["AP" + rowNumber].Value = cip.cipUpdate.processDie;
                                    worksheet.Cells["AQ" + rowNumber].Value = cip.cipUpdate.model;
                                    worksheet.Cells["AR" + rowNumber].Value = cip.cipUpdate.costCenterOfUser;
                                    worksheet.Cells["AS" + rowNumber].Value = cip.cipUpdate.tranferToSupplier;
                                    worksheet.Cells["AT" + rowNumber].Value = cip.cipUpdate.upFixAsset;
                                    worksheet.Cells["AU" + rowNumber].Value = cip.cipUpdate.newBFMorAddBFM;
                                    worksheet.Cells["AV" + rowNumber].Value = cip.cipUpdate.reasonForDelay;
                                    worksheet.Cells["AW" + rowNumber].Value = cip.cipUpdate.addCipBfmNo;
                                    worksheet.Cells["AX" + rowNumber].Value = cip.cipUpdate.remark;
                                }
                                else if (body.workType == "Oversea")
                                {
                                    var searchCell = from cell in worksheet.Cells["AI5:AI" + rowEnd.ToString()] where cell.Value.ToString() == cip.cipNo select cell.Start.Row;
                                    string rowNumber = searchCell.First().ToString();

                                    worksheet.Cells["AO" + rowNumber].Value = cip.cipUpdate.planDate;
                                    worksheet.Cells["AP" + rowNumber].Value = cip.cipUpdate.actDate;
                                    worksheet.Cells["AQ" + rowNumber].Value = cip.cipUpdate.result;
                                    worksheet.Cells["AR" + rowNumber].Value = cip.cipUpdate.reasonDiff;
                                    worksheet.Cells["AS" + rowNumber].Value = cip.cipUpdate.fixedAssetCode;
                                    worksheet.Cells["AT" + rowNumber].Value = cip.cipUpdate.classFixedAsset;
                                    worksheet.Cells["AU" + rowNumber].Value = cip.cipUpdate.fixAssetName;
                                    worksheet.Cells["AV" + rowNumber].Value = cip.cipUpdate.serialNo;
                                    worksheet.Cells["AW" + rowNumber].Value = cip.cipUpdate.partNumberDieNo;
                                    worksheet.Cells["AX" + rowNumber].Value = cip.cipUpdate.processDie;
                                    worksheet.Cells["AY" + rowNumber].Value = cip.cipUpdate.model;
                                    worksheet.Cells["AZ" + rowNumber].Value = cip.cipUpdate.costCenterOfUser;
                                    worksheet.Cells["BA" + rowNumber].Value = cip.cipUpdate.tranferToSupplier;
                                    worksheet.Cells["BB" + rowNumber].Value = cip.cipUpdate.upFixAsset;
                                    worksheet.Cells["BC" + rowNumber].Value = cip.cipUpdate.newBFMorAddBFM;
                                    worksheet.Cells["BD" + rowNumber].Value = cip.cipUpdate.reasonForDelay;
                                    // worksheet.Cells["AM" + rowNumber].Value = cip.cipUpdate.addCipBfmNo;
                                    worksheet.Cells["BE" + rowNumber].Value = cip.cipUpdate.remark;
                                }
                                else if (body.workType == "Domestic")
                                {
                                    var searchCell = from cell in worksheet.Cells["C1:C" + rowEnd.ToString()] where cell.Value.ToString() == cip.cipNo select cell.Start.Row;
                                    string rowNumber = searchCell.First().ToString();

                                    worksheet.Cells["W" + rowNumber].Value = cip.cipUpdate.planDate;
                                    worksheet.Cells["X" + rowNumber].Value = cip.cipUpdate.actDate;
                                    worksheet.Cells["Y" + rowNumber].Value = cip.cipUpdate.result;
                                    worksheet.Cells["Z" + rowNumber].Value = cip.cipUpdate.reasonDiff;
                                    worksheet.Cells["AA" + rowNumber].Value = cip.cipUpdate.fixedAssetCode;
                                    worksheet.Cells["AB" + rowNumber].Value = cip.cipUpdate.classFixedAsset;
                                    worksheet.Cells["AC" + rowNumber].Value = cip.cipUpdate.fixAssetName;
                                    worksheet.Cells["AD" + rowNumber].Value = cip.cipUpdate.serialNo;
                                    worksheet.Cells["AE" + rowNumber].Value = cip.cipUpdate.partNumberDieNo;
                                    worksheet.Cells["AF" + rowNumber].Value = cip.cipUpdate.processDie;
                                    worksheet.Cells["AG" + rowNumber].Value = cip.cipUpdate.model;
                                    worksheet.Cells["AH" + rowNumber].Value = cip.cipUpdate.costCenterOfUser;
                                    worksheet.Cells["AI" + rowNumber].Value = cip.cipUpdate.tranferToSupplier;
                                    worksheet.Cells["AJ" + rowNumber].Value = cip.cipUpdate.upFixAsset;
                                    worksheet.Cells["AK" + rowNumber].Value = cip.cipUpdate.newBFMorAddBFM;
                                    worksheet.Cells["AL" + rowNumber].Value = cip.cipUpdate.reasonForDelay;
                                    worksheet.Cells["AM" + rowNumber].Value = cip.cipUpdate.addCipBfmNo;
                                    worksheet.Cells["AN" + rowNumber].Value = cip.cipUpdate.remark;
                                }
                                else if (body.workType == "Project-MSC")
                                {
                                    var searchCell = from cell in worksheet.Cells["E2:E" + rowEnd.ToString()] where cell.Value.ToString() == cip.cipNo select cell.Start.Row;
                                    string rowNumber = searchCell.First().ToString();

                                    worksheet.Cells["AC" + rowNumber].Value = cip.cipUpdate.planDate;
                                    worksheet.Cells["AD" + rowNumber].Value = cip.cipUpdate.actDate;
                                    worksheet.Cells["AE" + rowNumber].Value = cip.cipUpdate.result;
                                    worksheet.Cells["AF" + rowNumber].Value = cip.cipUpdate.reasonDiff;
                                    worksheet.Cells["AG" + rowNumber].Value = cip.cipUpdate.fixedAssetCode;
                                    worksheet.Cells["AH" + rowNumber].Value = cip.cipUpdate.classFixedAsset;
                                    worksheet.Cells["AG" + rowNumber].Value = cip.cipUpdate.fixAssetName;
                                    worksheet.Cells["AJ" + rowNumber].Value = cip.cipUpdate.serialNo;
                                    worksheet.Cells["AK" + rowNumber].Value = cip.cipUpdate.partNumberDieNo;
                                    worksheet.Cells["AL" + rowNumber].Value = cip.cipUpdate.processDie;
                                    worksheet.Cells["AM" + rowNumber].Value = cip.cipUpdate.model;
                                    worksheet.Cells["AN" + rowNumber].Value = cip.cipUpdate.costCenterOfUser;
                                    worksheet.Cells["AO" + rowNumber].Value = cip.cipUpdate.tranferToSupplier;
                                    worksheet.Cells["AP" + rowNumber].Value = cip.cipUpdate.upFixAsset;
                                    worksheet.Cells["AQ" + rowNumber].Value = cip.cipUpdate.newBFMorAddBFM;
                                    worksheet.Cells["AR" + rowNumber].Value = cip.cipUpdate.reasonForDelay;
                                    worksheet.Cells["AS" + rowNumber].Value = cip.cipUpdate.addCipBfmNo;
                                    worksheet.Cells["AT" + rowNumber].Value = cip.cipUpdate.remark;
                                }

                                successCip.Add(cip.cipNo);
                                cip.status = "exported";
                                updateStatus.Add(cip);
                            }

                        }
                        catch (System.Exception)
                        {
                            errorCip.Add(cip.cipNo);
                        }
                        // 48992
                    }
                    db.CIP.UpdateRange(updateStatus);
                    db.SaveChanges();
                    package.Save();
                }
                return Ok(new
                {
                    success = true,
                    message = "Update CIP to dept file success.",
                    data = new
                    {
                        successItem = successCip,
                        errorItem = errorCip,
                    }
                });
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                if (e.Message.IndexOf("saving") != -1)
                {
                    return Conflict(new { success = false, message = "Please close file " + body.workType + " before export." });
                }
                return Conflict(new { success = false, message = "Have some error" });
            }
        }

        [HttpPatch("history"), AllowAnonymous]
        public ActionResult getHistory(request.getHistory body)
        {

            List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status == "exported" && item.workType == body.workType).OrderBy(x => x.id).Skip((body.page - 1) * body.perPage).Take(body.perPage).ToList();
            // List<cipSchema> all = db.CIP.Where<cipSchema>(item => true).OrderBy(x => x.id).Skip((body.page - 1) * body.perPage).Take(body.perPage).ToList();
            return Ok(new { success = true, message = "History data", data, });
        }
    }
}