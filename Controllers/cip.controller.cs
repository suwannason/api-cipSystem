
using Microsoft.AspNetCore.Mvc;

using cip_api.request.cip;
using cip_api.models;
using Microsoft.Extensions.Configuration;
using Microsoft.AspNetCore.Authorization;
using System.IO;
using OfficeOpenXml;
using System.Collections.Generic;

using System.Linq;

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
            FileStream strem = System.IO.File.Create($"{serverPath}{body.file.FileName}");
            body.file.CopyTo(strem);
            strem.Close();


            string path = $"{serverPath}{body.file.FileName}";
            FileInfo Existfile = new FileInfo(path);


            List<cipSchema> excelData = new List<cipSchema>();

            if (User.FindFirst("dept").Value.ToLower() == "acc")
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
                                if (data != null) {
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
            List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status == "open")
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

        [HttpGet("download")]
        public ActionResult download() {
            return Ok();
        }
    }
}