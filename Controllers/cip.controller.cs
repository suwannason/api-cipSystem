
using Microsoft.AspNetCore.Mvc;
using cip_api.request;
using cip_api.models;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Authorization;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System;
using System.Drawing;

namespace cip_api.controllers
{

    [Authorize]
    [ApiController]
    [Route("[controller]")]
    public class cipController : ControllerBase
    {
        private readonly Database db;
        private IConfiguration _config;
        private readonly string ldap_auth;

        public cipController(Database _db, IConfiguration config, IEndpoint setting)
        {
            db = _db;
            _config = config;
            ldap_auth = setting.ldap_auth;
        }

        [HttpPost("upload"), Consumes("multipart/form-data")]
        public ActionResult upload([FromForm] CIPupload body)
        {
            string rootFolder = Directory.GetCurrentDirectory();
            string pathString2 = @"\API site\files\CIP-system\upload\";
            string serverPath = rootFolder.Substring(0, rootFolder.LastIndexOf(@"\")) + pathString2;

            if (!Directory.Exists(serverPath))
            {
                Directory.CreateDirectory(serverPath);
            }
            string fileName = System.Guid.NewGuid().ToString() + "-" + body.file.FileName;
            FileStream strem = System.IO.File.Create($"{serverPath}{fileName}");
            body.file.CopyTo(strem);
            strem.Close();

            string path = $"{serverPath}{fileName}";
            FileInfo Existfile = new FileInfo(path);

            List<cipSchema> excelData = new List<cipSchema>();
            string dateNow = DateTime.Now.ToString("yyyy/MM/dd");
            if (User.FindFirst("dept").Value.ToLower() == "acc")
            {
                using (ExcelPackage excel = new ExcelPackage(Existfile))
                {
                    ExcelWorkbook workbook = excel.Workbook;
                    ExcelWorksheet sheet = workbook.Worksheets[0];

                    int colCount = sheet.Dimension.End.Column;
                    int rowCount = sheet.Dimension.End.Row;

                    for (int row = 3; row < rowCount; row += 1)
                    {
                        cipSchema item = new cipSchema();

                        for (int col = 1; col < colCount; col += 1)
                        {
                            string value = sheet.Cells[row, col].Value?.ToString();
                            if (value == null)
                            {
                                value = "-";
                            }
                            switch (col)
                            {
                                case 1: item.cipNo = value; break;
                                case 2: item.subCipNo = value; break;
                                case 3: item.poNo = value; break;
                                case 4: item.vendorCode = value; break;
                                case 5: item.vendor = value; break;
                                case 6: item.acqDate = value; break;
                                case 7: item.invDate = value; break;
                                case 8: item.receivedDate = value; break;
                                case 9: item.invNo = value; break;
                                case 10: item.name = value; break;
                                case 11: item.qty = value; break;
                                case 12: item.exRate = value; break;
                                case 13: item.cur = value; break;
                                case 14: item.perUnit = value; break;
                                case 15: item.totalJpy = value; break;
                                case 16: item.totalThb = value; break;
                                case 17: item.averageFreight = value; break;
                                case 18: item.averageInsurance = value; break;
                                case 19: item.totalJpy_1 = value; break;
                                case 20: item.totalThb_1 = value; break;
                                case 21: item.perUnitThb = value; break;
                                case 22: item.cc = value; break;
                                case 23: item.totalOfCip = value; break;
                                case 24: item.budgetCode = value; break;
                                case 25: item.prDieJig = value; break;
                                case 26: item.model = value; break;
                            }
                            item.status = "open";
                            item.createDate = System.DateTime.Now.ToString("yyyy/MM/dd");
                        }
                        if (item.cipNo != "-")
                        {
                            excelData.Add(item);
                        }
                    }
                }
                db.CIP.AddRange(excelData);
                db.SaveChanges();
                return Ok(new { success = true, message = "Upload data success." });
            }

            List<cipUpdateSchema> items = new List<cipUpdateSchema>();

            List<cipSchema> updateStatus = new List<cipSchema>();

            using (ExcelPackage excel = new ExcelPackage(Existfile))
            {
                ExcelWorkbook workbook = excel.Workbook;
                ExcelWorksheet sheet = workbook.Worksheets[0];

                int colCount = sheet.Dimension.End.Column;
                int rowCount = sheet.Dimension.End.Row;

                for (int row = 3; row < rowCount; row += 1)
                {
                    cipUpdateSchema item = new cipUpdateSchema();


                    for (int col = 1; col <= colCount; col += 1)
                    {
                        string value = sheet.Cells[row, col].Value?.ToString();
                        if (value == null)
                        {
                            value = "-";
                        }
                        switch (col)
                        {
                            case 1:
                                if (value == "-")
                                {
                                    break;
                                }
                                cipSchema data = db.CIP.Where<cipSchema>(item => item.cipNo == value && item.status == "open").FirstOrDefault();
                                if (data != null)
                                {
                                    item.cipSchemaid = data.id;
                                    data.status = "save";
                                    updateStatus.Add(data);
                                }
                                break;

                            case 27: item.planDate = value; break;
                            case 28: item.actDate = value; break;
                            case 29: item.result = value; break;
                            case 30: item.reasonDiff = value; break;
                            case 31: item.fixedAssetCode = value; break;
                            case 32: item.classFixedAsset = value; break;
                            case 33: item.fixAssetName = value; break;
                            case 34: item.serialNo = value; break;
                            case 35: item.processDie = value; break;
                            case 36: item.model = value; break;
                            case 37: item.costCenterOfUser = value; break;
                            case 38: item.tranferToSupplier = value; break;
                            case 39: item.upFixAsset = value; break;
                            case 40: item.newBFMorAddBFM = value; break;
                            case 41: item.reasonForDelay = value; break;
                            case 42: item.remark = value; break;
                            case 43: item.boiType = value; break;
                        }
                        item.status = "active";
                        item.createDate = dateNow;

                    }
                    if (item.cipSchemaid != 0)
                    {
                        items.Add(item);
                    }
                }
            }
            List<ApprovalSchema> prepare = new List<ApprovalSchema>();

            string preparer = User.FindFirst("username").Value;
            foreach (cipSchema item in updateStatus)
            {
                prepare.Add(new ApprovalSchema
                {
                    cipSchemaid = item.id,
                    date = dateNow,
                    empNo = preparer,
                    onApproveStep = "save"
                });
            }

            db.APPROVAL.AddRange(prepare);
            db.CIP_UPDATE.AddRange(items);
            db.CIP.UpdateRange(updateStatus);
            db.SaveChanges();

            return Ok();
        }

        [HttpGet("list")]
        public ActionResult list()
        {
            List<cipSchema> data = null;
            if (User.FindFirst("dept").Value.ToLower() == "acc")
            {
                data = db.CIP.Where<cipSchema>(item => item.status == "open" || item.status == "reject")
               .Select(fields =>
               new cipSchema { cipNo = fields.cipNo, subCipNo = fields.subCipNo, vendor = fields.vendor, name = fields.name, qty = fields.qty, totalThb = fields.totalThb, cc = fields.cc, id = fields.id, status = fields.status })
               .ToList<cipSchema>();
                return Ok(new { success = true, data, });
            }
            string deptCode = User.FindFirst("deptCode")?.Value;

            data = db.CIP.Where<cipSchema>(item => item.status == "open" && item.cc == deptCode)
               .Select(fields =>
               new cipSchema { cipNo = fields.cipNo, subCipNo = fields.subCipNo, vendor = fields.vendor, name = fields.name, qty = fields.qty, totalThb = fields.totalThb, cc = fields.cc, id = fields.id })
               .ToList<cipSchema>();
            return Ok(new { success = true, data, });
        }
        [HttpGet("history")]
        public ActionResult history(DateRange body)
        {
            List<cipSchema> data = db.CIP.Where<cipSchema>(item => String.Compare(item.createDate, body.startDate) == 1 && String.Compare(item.createDate, body.endDate) == -1).ToList();
            return Ok(new { success = true, data, });
        }

        [HttpPatch("download")]
        public ActionResult download(cipDownlode body)
        {
            try
            {
                string deptCode = User.FindFirst("deptCode")?.Value;
                string empNo = User.FindFirst("empNo")?.Value;
                string dept = User.FindFirst("dept")?.Value;

                List<cipSchema> data = new List<cipSchema>();
                if (body.id.Length == 0)
                {
                    if (dept.ToLower() != "acc" && dept != "admin")
                    {
                        data = db.CIP.Where<cipSchema>(item => item.cc == deptCode && item.status == "open").ToList<cipSchema>();
                    }
                    else
                    {
                        Console.WriteLine("ELSE");
                        data = db.CIP.Where<cipSchema>(item => item.status == "open").ToList<cipSchema>();
                        Console.WriteLine("Case else: " + data.Count);
                    }
                }
                else
                {
                    foreach (int id in body.id)
                    {
                        cipSchema item = db.CIP.Find(id);
                        data.Add(item);
                    }
                }

                MemoryStream stream = new MemoryStream();
                using (ExcelPackage excel = new ExcelPackage(stream))
                {
                    excel.Workbook.Worksheets.Add("sheet1");

                    List<string[]> header = new List<string[]>()
                {
                    new string[] { "CIP No.", "Sub CIP No.", "PO NO.", "VENDER CODE", "VENDER", "ACQ-DATE (ETD)", "INV DATE",
                    "RECEIVED DATE", "INV NO.", "NAME (ENGLISH)", "Qty.", "EX.RATE", "CUR", "PER UNIT \n (THB/JPY/USD)",
                    "TOTAL (JPY/USD)", "TOTAL (THB)", "AVERAGE FREIGHT (JPY/USD)", "AVERAGE INSURANCE (JPY/USD)", "TOTAL (JPY/USD)",
                    "TOTAL (THB)", "PER UNIT (THB)", "CC", "TOTAL OF CIP (THB)", "Budget code", "PR.DIE/JIG", "Model",
                    "Operating Date (Plan)", "Operating Date (Act)", "Result", "Reason diff (NG) Budget&Actual", "Fixed Asset Code",
                    "CLASS FIXED ASSET", "Fix Asset Name (English only)", "Serial No.", "Process Die", "Model",
                    "Cost Center of User", "Transfer to supplier", "ให้ขึ้น Fix Asset  กี่ตัว", "New BFMor Add BFM", "Reason for Delay",
                    "REMARK (Add CIP/BFM No.)", "ITC--> BOI TYPE (Machine / Die / Sparepart / NON BOI)"
                    }
                };

                    string headerRange = "A2:AQ2";
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets["sheet1"];
                    worksheet.Cells[headerRange].LoadFromArrays(header);

                    string rootFolder = Directory.GetCurrentDirectory();
                    string pathString2 = @"\API site\files\CIP-system\download\";
                    string serverPath = rootFolder.Substring(0, rootFolder.LastIndexOf(@"\")) + pathString2;

                    if (!Directory.Exists(serverPath))
                    {
                        Directory.CreateDirectory(serverPath);
                    }
                    worksheet.Cells["A1:Z1"].Merge = true;
                    worksheet.Cells["A1"].Value = "Data for Accounting Dept.";
                    worksheet.Cells["A1"].Style.Font.Bold = true;
                    worksheet.Cells["A1"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#9FE5E6"));
                    worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A1"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    worksheet.Cells["AA1:AQ1"].Merge = true;
                    worksheet.Cells["AA1"].Value = "Data for User confirm";
                    worksheet.Cells["AA1"].Style.Font.Bold = true;
                    worksheet.Cells["AA1"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                    worksheet.Cells["AA1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["AA1"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    worksheet.Cells["AA2:AP2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F1ECB9"));
                    worksheet.Cells["AQ2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#57A868"));

                    worksheet.Column(1).Width = 5;
                    worksheet.Column(2).Width = 4;
                    worksheet.Column(6).Width = 11;
                    worksheet.Column(10).Width = 11;
                    worksheet.Column(11).Width = 4; // Qty.
                    worksheet.Column(14).Width = 13;
                    worksheet.Column(15).Width = 12;
                    worksheet.Column(17).Width = 12;
                    worksheet.Column(18).Width = 12;
                    worksheet.Column(19).Width = 12;
                    worksheet.Column(23).Width = 12;
                    worksheet.Column(25).Width = 12; // AA
                    worksheet.Column(41).Width = 12; // AO
                    worksheet.Column(43).Width = 15; // AO
                    worksheet.Row(2).Height = 80;
                    worksheet.Row(1).Height = 30;
                    worksheet.Row(2).Style.Font.Bold = true;
                    worksheet.Row(2).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    worksheet.Row(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A2:P2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));
                    worksheet.Cells["S2:Z2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));
                    worksheet.Cells["Q2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C7BDF9"));
                    worksheet.Cells["R2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#CCF9BD"));
                    worksheet.Cells["A2:AQ2"].Style.WrapText = true;
                    worksheet.Cells["A2:AQ2"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells["A2:AQ2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells["A2:AQ2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells["A2:AQ2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    int row = 3;
                    foreach (cipSchema item in data)
                    {
                        List<string[]> cellData = new List<string[]>()
                    {
                        new string [] {
                            item.cipNo, item.subCipNo, item.vendorCode, item.vendor, item.acqDate, item.invDate, item.receivedDate,
                             item.invNo, item.name, item.qty, item.exRate, item.cur, item.perUnit, item.totalJpy, item.totalThb, item.averageFreight,
                             item.averageInsurance, item.totalJpy_1, item.totalThb_1, item.perUnitThb, item.cc, item.totalOfCip, item.budgetCode, item.prDieJig,
                             item.model
                        }

                    };
                        worksheet.Cells["AC" + row].Formula = "=+IF(AF" + row + "=\"\",\"\",IF(X" + row + "=\"\",\"\",IF(AND(MID(X" + row + ",5,2)=\"09\",LEFT(AF" + row + ",2)=\"06\"),\"OK\",IF(AND(OR(MID(X" + row + ",5,2)=\"31\",MID(X" + row + ",5,2)=\"34\"),LEFT(AF" + row + ",2)=\"28\"),\"OK\",IF(MID(X" + row + ",5,2)=LEFT(AF" + row + ",2),\"OK\",\"NG\")))))";
                        worksheet.Cells["AC" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));
                        worksheet.Cells["AE" + row].Formula = "=IF(LEFT(AF" + row + ",2)=\"28\",\"SOFTWARE\",IF(LEFT(AF" + row + ",2)=\"02\",\"BUILDING\",IF(LEFT(AF" + row + ",2)=\"03\",\"STRUCTURE\",IF(LEFT(AF" + row + ",2)=\"04\",\"MACHINE\",IF(LEFT(AF" + row + ",2)=\"05\",\"VEHICLE\",IF(LEFT(AF" + row + ",2)=\"06\",\"TOOLS\",IF(LEFT(AF" + row + ",2)=\"07\",\"FURNITURE\",IF(LEFT(AF" + row + ",2)=\"08\",\"DIES\",\"\"))))))))";
                        worksheet.Cells["AE" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));

                        // set pink row space
                        worksheet.Cells["AA" + row + ":AB" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                        worksheet.Cells["AD" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                        worksheet.Cells["AF" + row + ":AP" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                        // set pink row space
                        worksheet.Cells[row, 1].LoadFromArrays(cellData);
                        worksheet.Cells["A" + row + ":AQ" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells["A" + row + ":AQ" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells["A" + row + ":AQ" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells["A" + row + ":AQ" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        row = row + 1;
                    }
                    string fileName = System.Guid.NewGuid().ToString() + "-" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                    excel.SaveAs(new FileInfo(serverPath + fileName));

                    stream.Position = 0;
                    // return Ok(new { success = true });
                    return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", serverPath + fileName);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                return Problem(e.StackTrace);
            }
        }

        [HttpGet("{id}")]
        public ActionResult getById(string id)
        {
            Int32 cipId = Int32.Parse(id);
            cipSchema cip = db.CIP.Find(cipId);

            db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == cipId).FirstOrDefault();
            db.APPROVAL.Where<ApprovalSchema>(item => item.cipSchemaid == cipId).ToList<ApprovalSchema>();

            return Ok(new
            {
                success = true,
                data = cip,
            });
        }

        [HttpGet("cipUpdate/{id}")]
        public ActionResult cipUpdate(string id) {
            try {

                cipUpdateSchema data = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == Int32.Parse(id)).FirstOrDefault();

                return Ok(new { success = true, data, });
            } catch (Exception e) {
                return Problem(e.StackTrace);
            }
        }
    }
}