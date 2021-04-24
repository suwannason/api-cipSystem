
using Microsoft.AspNetCore.Mvc;

using cip_api.request.cip;
using cip_api.models;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Authorization;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;

using System.Linq;
using System;
using System.Threading.Tasks;

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

            if (User.FindFirst("dept").Value.ToLower() == "acc" || User.FindFirst("dept").Value.ToLower() == "admin")
            {
                using (ExcelPackage excel = new ExcelPackage(Existfile))
                {
                    ExcelWorkbook workbook = excel.Workbook;
                    ExcelWorksheet sheet = workbook.Worksheets[0];

                    int colCount = sheet.Dimension.End.Column;
                    int rowCount = sheet.Dimension.End.Row;

                    for (int row = 2; row < rowCount; row += 1)
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
                            item.createDate = System.DateTime.Now.ToString("yyyyMMdd");
                        }
                        excelData.Add(item);
                    }
                }
                db.CIP.AddRange(excelData);
                db.SaveChanges();
                return Ok(new { success = true, message = "Upload data success." });
            }

            List<cipUpdateSchema> items = new List<cipUpdateSchema>();

            using (ExcelPackage excel = new ExcelPackage(Existfile))
            {
                ExcelWorkbook workbook = excel.Workbook;
                ExcelWorksheet sheet = workbook.Worksheets[0];

                int colCount = sheet.Dimension.End.Column;
                int rowCount = sheet.Dimension.End.Row;

                for (int row = 2; row < rowCount; row += 1)
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
                                cipSchema data = db.CIP.Where<cipSchema>(item => item.cipNo == value && item.status == "open").FirstOrDefault();
                                if (data != null)
                                {
                                    item.cipSchemaid = data.id;
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
                    }
                    items.Add(item);
                }
            }
            db.CIP_UPDATE.AddRange(items);
            db.SaveChanges();
            return Ok();
        }

        [HttpGet("list")]
        public ActionResult list()
        {
            List<cipSchema> data = null;
            if (User.FindFirst("dept").Value.ToLower() == "acc" || User.FindFirst("dept").Value.ToLower() == "admin")
            {
                data = db.CIP.Where<cipSchema>(item => item.status == "open")
               .Select(fields =>
               new cipSchema { cipNo = fields.cipNo, subCipNo = fields.subCipNo, vendor = fields.vendor, name = fields.name, qty = fields.qty, totalThb = fields.totalThb, cc = fields.cc, id = fields.id })
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
        public ActionResult history()
        {
            return Ok(new { success = true });
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
                        Console.WriteLine("Case else: "+data.Count);
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
                    "RECEIVED DATE", "INV NO.", "NAME (ENGLISH)", "Qty.", "EX.RATE", "CUR", "PER UNIT (THB/JPY/USD)",
                    "TOTAL (JPY/USD)", "TOTAL (THB)", "AVERAGE FREIGHT (JPY/USD)", "AVERAGE INSURANCE (JPY/USD)", "TOTAL (JPY/USD)",
                    "TOTAL (THB)", "PER UNIT (THB)", "CC", "TOTAL OF CIP (THB)", "Budget code", "PR.DIE/JIG", "Model",
                    "Operating Date (Plan)", "Operating Date (Act)", "Result", "Reason diff (NG) Budget&Actual", "Fixed Asset Code",
                    "CLASS FIXED ASSET", "Fix Asset Name (English only)", "Serial No.", "Process Die", "Model",
                    "Cost Center of User", "Transfer to supplier", "ให้ขึ้น Fix Asset  กี่ตัว", "New BFMor Add BFM", "Reason for Delay",
                    "REMARK (Add CIP/BFM No.)", "ITC--> BOI TYPE (Machine / Die / Sparepart / NON BOI)"
                    }
                };

                    string headerRange = "A1:AQ1";
                    ExcelWorksheet worksheet = excel.Workbook.Worksheets["sheet1"];
                    worksheet.Cells[headerRange].LoadFromArrays(header);

                    string rootFolder = Directory.GetCurrentDirectory();
                    string pathString2 = @"\API site\files\CIP-system\download\";
                    string serverPath = rootFolder.Substring(0, rootFolder.LastIndexOf(@"\")) + pathString2;

                    if (!Directory.Exists(serverPath))
                    {
                        Directory.CreateDirectory(serverPath);
                    }
                    int row = 2;
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
                        worksheet.Cells[row, 1].LoadFromArrays(cellData);
                        row = row + 1;
                    }
                    string fileName = System.Guid.NewGuid().ToString() + "-" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                    excel.SaveAs(new FileInfo(serverPath + fileName));

                    stream.Position = 0;
                    return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", serverPath + fileName);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                return Ok();
            }
        }
    }
}