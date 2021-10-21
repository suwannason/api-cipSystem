
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
using System.Net.Http;
using System.Threading.Tasks;
using OfficeOpenXml.DataValidation.Contracts;

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
        private readonly string node_api;

        public cipController(Database _db, IConfiguration config, IEndpoint setting)
        {
            db = _db;
            _config = config;
            ldap_auth = setting.ldap_auth;
            node_api = setting.node_api;
        }

        private async void sendMail(string from, string to, string subject, string text)
        {
            using (HttpClient client = new HttpClient())
            {

                using (MultipartFormDataContent formData = new MultipartFormDataContent())
                {
                    var values = new[] {
                        new KeyValuePair<string, string>("from", from),
                        new KeyValuePair<string, string>("to", to),
                        new KeyValuePair<string, string>("subject", subject),
                        new KeyValuePair<string, string>("text", text),
                    };
                    foreach (var keyValuePair in values)
                    {
                        formData.Add(new StringContent(keyValuePair.Value),
                        String.Format("\"{0}\"", keyValuePair.Key));
                    }


                    client.Timeout = TimeSpan.FromSeconds(20);
                    HttpResponseMessage result = await client.PostAsync(node_api + "/middleware/email/sendmail", formData);
                    string input = await result.Content.ReadAsStringAsync();
                    client.Dispose();
                }
            }
        }

        [HttpPost("upload"), Consumes("multipart/form-data")]
        public ActionResult upload([FromForm] CIPupload body)
        {
            try
            {
                string rootFolder = Directory.GetCurrentDirectory();
                string pathString2 = @"\API site\files\CIP-system\upload\";
                string serverPath = rootFolder.Substring(0, rootFolder.LastIndexOf(@"\")) + pathString2;

                if (!Directory.Exists(serverPath))
                {
                    Directory.CreateDirectory(serverPath);
                }
                string deptCode = User.FindFirst("deptCode").Value;
                string username = User.FindFirst("username")?.Value;
                string fileName = System.Guid.NewGuid().ToString() + "-" + body.file.FileName;
                FileStream strem = System.IO.File.Create($"{serverPath}{fileName}");
                body.file.CopyTo(strem);
                strem.Close();

                string path = $"{serverPath}{fileName}";
                FileInfo Existfile = new FileInfo(path);

                List<cipSchema> excelData = new List<cipSchema>();
                string dateNow = DateTime.Now.ToString("yyyy/MM/dd");
                if (body.type == "Accounting")
                {
                    if (User.FindFirst("dept").Value.ToLower() != "acc")
                    {
                        return BadRequest(new { success = false, message = "Access denied." });
                    }
                    using (ExcelPackage excel = new ExcelPackage(Existfile))
                    {
                        ExcelWorkbook workbook = excel.Workbook;
                        ExcelWorksheet sheet = workbook.Worksheets[0];

                        int colCount = sheet.Dimension.End.Column;
                        int rowCount = sheet.Dimension.End.Row;

                        for (int row = 3; row <= rowCount; row += 1)
                        {
                            cipSchema item = new cipSchema();

                            for (int col = 1; col < colCount; col += 1)
                            {
                                string value = sheet.Cells[row, col].Value?.ToString();
                                if (value == null || value == "")
                                {
                                    value = "-";
                                }
                                switch (col)
                                {
                                    case 1: item.workType = value; break;
                                    case 2: item.projectNo = value; break;
                                    case 3: item.cipNo = value; break;
                                    case 4: item.subCipNo = value; break;
                                    case 5: item.poNo = value; break;
                                    case 6: item.vendorCode = value; break;
                                    case 7: item.vendor = value; break;
                                    case 8: item.acqDate = value; break;
                                    case 9: item.invDate = value; break;
                                    case 10: item.receivedDate = value; break;
                                    case 11: item.invNo = value; break;
                                    case 12: item.name = value; break;
                                    case 13: item.qty = value; break;
                                    case 14: item.exRate = value; break;
                                    case 15: item.cur = value; break;
                                    // item.perUnit = value != "-" ? Double.Parse(value).ToString("###,###,###,###.00") : "-"; break;
                                    case 16:
                                        if (value != "-")
                                        {
                                            item.perUnit = Double.Parse(value).ToString("###,###,###,###.00");
                                        }
                                        else
                                        {
                                            item.perUnit = "-";
                                        }
                                        break;
                                    case 17:
                                        if (value != "-")
                                        {
                                            item.totalJpy = Double.Parse(value).ToString("###,###,###,###.00");
                                        }
                                        else
                                        {
                                            item.totalJpy = "-";
                                        }
                                        break;
                                    case 18:
                                        if (value != "-")
                                        {
                                            item.totalThb = Double.Parse(value).ToString("###,###,###,###.00");
                                        }
                                        else
                                        {
                                            item.totalThb = "-";
                                        }
                                        break;
                                    case 19:
                                        if (value != "-")
                                        {
                                            item.averageFreight = Double.Parse(value).ToString("###,###,###,###.00");
                                        }
                                        else
                                        {
                                            item.averageFreight = "-";
                                        }
                                        break;
                                    case 20:
                                        if (value != "-")
                                        {
                                            item.averageInsurance = Double.Parse(value).ToString("###,###,###,###.00");
                                        }
                                        else
                                        {
                                            item.averageInsurance = "-";
                                        }
                                        break;
                                    case 21:
                                        if (value != "-")
                                        {
                                            item.totalJpy_1 = Double.Parse(value).ToString("###,###,###,###.00");
                                        }
                                        else
                                        {
                                            item.totalJpy_1 = "-";
                                        }
                                        break;
                                    case 22:
                                        if (value != "-")
                                        {
                                            item.totalThb_1 = Double.Parse(value).ToString("###,###,###,###.00");
                                        }
                                        else
                                        {
                                            item.totalThb_1 = "-";
                                        }
                                        break;
                                    case 23:
                                        if (value != "-")
                                        {
                                            item.perUnitThb = Double.Parse(value).ToString("###,###,###,###.00");
                                        }
                                        else
                                        {
                                            item.perUnitThb = "-";
                                        }
                                        break;
                                    case 24: item.cc = value; break;
                                    case 25:
                                        if (value != "-")
                                        {
                                            item.totalOfCip = Double.Parse(value).ToString("###,###,###,###.00");
                                        }
                                        else
                                        {
                                            item.totalOfCip = "-";
                                        }
                                        break;
                                    case 26: item.budgetCode = value; break;
                                    case 27: item.prDieJig = value; break;
                                    case 28: item.model = value; break;
                                    case 29: item.partNoDieNo = value; break;
                                }
                                item.status = "open";
                                item.createDate = System.DateTime.Now.ToString("yyyy/MM/dd");
                            }
                            if (item.cipNo != "-")
                            {
                                cipSchema cipCreate = db.CIP.Where<cipSchema>(cip => cip.cipNo == item.cipNo && cip.subCipNo == item.subCipNo).FirstOrDefault();
                                if (cipCreate == null)
                                {
                                    excelData.Add(item);
                                }
                                else
                                {
                                    Console.WriteLine("duplicated");
                                    return BadRequest(new { success = false, message = "Some data duplicatied." });
                                }
                            }
                        }
                    }
                    // return Ok(excelData);
                    db.CIP.AddRange(excelData);
                    // return Ok(excelData);
                    db.SaveChanges();
                    // SENDING MAIL
                    // PermissionSchema acc_user = db.PERMISSIONS.Where<PermissionSchema>(item => item.empNo == username).FirstOrDefault();
                    // List<string> ccDept = excelData.Select(c => c.cc).Distinct().ToList();

                    // string mailBody = "TO : All Concerned \n \n I would like to Confirm CIP-Domestic and Oversea. \n At link <WEB APP LINK> \n Please, input data pink area (data for user confirm). \n";
                    // mailBody += "\n\n\n Thank You \n Best Regards \n";
                    // mailBody += "**************************************************** \n";
                    // mailBody += "\t\t\t " + " " + User.FindFirst("name")?.Value + " \n";
                    // mailBody += "\t\t\t Accounting Dept. \n";
                    // mailBody += "\t\t Canon Prachinburi (Thailand) Ltd. \n";
                    // mailBody += "\t\t E-mail : " + acc_user.email + " \n";
                    // mailBody += "\t\t   " + "☎ : 037-284600 Ext.8114" + " \n";
                    // mailBody += "****************************************************";
                    // foreach (string cc in ccDept)
                    // {
                    //     List<PermissionSchema> ccPrepare = db.PERMISSIONS.Where<PermissionSchema>(item => item.deptCode.IndexOf(cc) != -1 && item.action == "prepare").ToList();

                    //     if (ccPrepare.Count != 0)
                    //     {
                    //         foreach (PermissionSchema userPrepare in ccPrepare)
                    //         {
                    //             sendMail(
                    //                 acc_user.email,
                    //                 userPrepare.email,
                    //                 "Confirm CIP-Domestic&Oversea" + DateTime.Now.ToString("yyyyMMdd") + " (Deadline within " + DateTime.Now.ToString("yyyy/MM/dd") + " time 16.00 pm.)",
                    //                mailBody
                    //             );
                    //         }
                    //     }
                    // }
                    // SENDING MAIL
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

                    for (int row = 3; row <= rowCount; row += 1)
                    {
                        cipUpdateSchema item = new cipUpdateSchema();

                        for (int col = 3; col <= colCount; col += 1)
                        {
                            string value = sheet.Cells[row, col].Value?.ToString();
                            if (value == null)
                            {
                                value = "-";
                            }
                            switch (col)
                            {
                                case 3:
                                    if (value == "-")
                                    {
                                        break;
                                    }
                                    cipSchema data = db.CIP.Where<cipSchema>(item => item.cipNo == value && (item.status == "open" || item.status == "reject")).FirstOrDefault();
                                    if (data != null)
                                    {
                                        item.cipSchemaid = data.id;
                                        data.status = "draft";
                                        updateStatus.Add(data);
                                    }
                                    break;

                                case 30: item.planDate = value; break;
                                case 31: item.actDate = value; break;
                                case 32: item.result = value; break;
                                case 33: item.reasonDiff = value; break;
                                case 34: item.fixedAssetCode = value; break;
                                case 35: item.classFixedAsset = value; break;
                                case 36:
                                    if (value.IndexOf(',') != -1 || value.IndexOf(':') != -1
                                        || value.IndexOf('"') != -1 || value.IndexOf('#') != -1
                                        || value.IndexOf('!') != -1 || value.IndexOf('*') != -1 || value.Length > 100)
                                    {
                                        return BadRequest(new { success = false, message = "Not allow Symbol fix asset name value." });
                                    }
                                    item.fixAssetName = value;
                                    break;
                                case 37: item.serialNo = value; break;
                                case 38: item.partNumberDieNo = value; break;
                                case 39: item.processDie = value; break;
                                case 40: item.model = value; break;
                                case 41: item.costCenterOfUser = value; break;
                                case 42: item.tranferToSupplier = value; break;
                                case 43: item.upFixAsset = value; break;
                                case 44: item.newBFMorAddBFM = value; break;
                                case 45: item.reasonForDelay = value; break;
                                case 46: item.addCipBfmNo = value; break;
                                case 47: item.remark = value; break;
                                case 48: item.boiType = value; break;
                            }
                            item.status = "active";
                            item.createDate = dateNow;
                        }
                        cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(cipUpdate => cipUpdate.cipSchemaid == item.cipSchemaid).FirstOrDefault();

                        if (cipUpdate == null)
                        {
                            items.Add(item);
                        }
                        else
                        {

                            cipUpdate.planDate = item.planDate;
                            cipUpdate.actDate = item.actDate;
                            cipUpdate.result = item.result;
                            cipUpdate.reasonDiff = item.reasonDiff;
                            cipUpdate.fixedAssetCode = item.fixedAssetCode;
                            cipUpdate.classFixedAsset = item.classFixedAsset;
                            cipUpdate.fixAssetName = item.fixAssetName;
                            cipUpdate.partNumberDieNo = item.partNumberDieNo;
                            cipUpdate.serialNo = item.serialNo;
                            cipUpdate.processDie = item.processDie;
                            cipUpdate.model = item.model;
                            cipUpdate.costCenterOfUser = item.costCenterOfUser;
                            cipUpdate.tranferToSupplier = item.tranferToSupplier;
                            cipUpdate.upFixAsset = item.upFixAsset;
                            cipUpdate.newBFMorAddBFM = item.newBFMorAddBFM;
                            cipUpdate.reasonForDelay = item.reasonForDelay;
                            cipUpdate.addCipBfmNo = item.addCipBfmNo;
                            cipUpdate.remark = item.remark;
                            cipUpdate.boiType = item.boiType;

                            db.CIP_UPDATE.Update(cipUpdate);

                            // List<ApprovalSchema> approve = db.APPROVAL.Where<ApprovalSchema>(item => item.cipSchemaid == cipUpdate.cipSchemaid).ToList();
                        }

                        // if (item.cipSchemaid != 0)
                        // {
                        //     items.Add(item);
                        // }
                    }
                }

                // return Ok(items);

                // db.APPROVAL.AddRange(prepare);
                db.CIP.UpdateRange(updateStatus);
                db.CIP_UPDATE.AddRange(items);
                db.SaveChanges();

                return Ok(items);
            }
            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);
                return Problem(e.Message);
            }
        }

        [HttpGet("list")]
        public ActionResult list()
        {
            List<cipSchema> data = new List<cipSchema>();
            if (User.FindFirst("dept").Value.ToLower() == "acc")
            {
                data = db.CIP.Where<cipSchema>(item => item.status != "finished").ToList<cipSchema>();
                return Ok(new { success = true, data, isAcc = true });
            }
            string deptCode = User.FindFirst("deptCode")?.Value;

            List<cipSchema> returnData = new List<cipSchema>();

            List<string> multidept = deptCode.Split(',').ToList();

            // Console.WriteLine(multidept.Count);
            foreach (string code in multidept)
            {
                Console.WriteLine(code);
                if (code != "55XX")
                {
                    data = db.CIP.Where<cipSchema>(item => (item.status != "finished") && (item.cc == code || item.cipUpdate.costCenterOfUser == code)).ToList<cipSchema>();
                }
                else
                { // 55XX
                    data = db.CIP.Where<cipSchema>(item => item.cc.StartsWith("55") && item.status != "finished").ToList();
                    //    db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.costCenterOfUser.IndexOf("55") == 0);
                }
                returnData.AddRange(data);
            }

            returnData = returnData.GroupBy(x => x.id).Select(x => x.First()).ToList();
            List<cipSchema> response = new List<cipSchema>();

            foreach (cipSchema cip in returnData)
            {
                cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(cipUpdate => cipUpdate.cipSchemaid == cip.id).FirstOrDefault();

                if (cip.status == "open" || cip.status == "reject" || cip.status == "draft")
                {
                    response.Add(cip);
                }
                else if (cip.status.IndexOf("cc") != -1 || cip.status == "save")
                {
                    if (cip.status.IndexOf("approve") == -1)
                    {
                        response.Add(cip);
                    }
                    else if (cip.status == "cc-approved" && deptCode.IndexOf(cipUpdate.costCenterOfUser) != -1)
                    {
                        response.Add(cip);
                    }
                    else
                    {
                        response.Add(cip);
                    }

                }
                else if (cip.status.IndexOf("cost") != -1 && deptCode.IndexOf(cipUpdate.costCenterOfUser) != -1)
                {
                    if (cip.status != "cost-approved")
                    {
                        response.Add(cip);
                    }
                }
            }
            return Ok(new { success = true, data = response, isAcc = false, });
        }
        [HttpGet("history")]
        public ActionResult history()
        {
            string lastSixMonth = DateTime.Now.AddMonths(-6).ToString("yyyy/MM/dd");
            string deptCode = User.FindFirst("deptCode")?.Value;

            List<string> multidept = deptCode.Split(',').ToList();

            List<cipSchema> requester = new List<cipSchema>();
            List<cipSchema> user = new List<cipSchema>();

            foreach (string dept in multidept)
            {
                List<cipSchema> dataRequester = new List<cipSchema>();
                List<cipSchema> dataUser = new List<cipSchema>();

                if (dept == "55XX")
                {
                    dataRequester = db.CIP.Where<cipSchema>(item => item.cc.StartsWith("55") && String.Compare(item.createDate, lastSixMonth) > 0).ToList();
                    dataUser = db.CIP.Where<cipSchema>(item => !item.cipUpdate.costCenterOfUser.StartsWith("55") && item.cipUpdate.costCenterOfUser.StartsWith("55") && String.Compare(item.createDate, lastSixMonth) > 0).ToList();
                }
                else
                {
                    dataRequester = db.CIP.Where<cipSchema>(item => item.cc == dept && String.Compare(item.createDate, lastSixMonth) > 0).ToList();
                    dataUser = db.CIP.Where<cipSchema>(item => item.cc != dept && item.cipUpdate.costCenterOfUser == dept && String.Compare(item.createDate, lastSixMonth) > 0).ToList();
                }
                requester.AddRange(dataRequester);
                user.AddRange(dataUser);
            }
            // List<cipSchema> data = db.CIP.Where<cipSchema>(item => String.Compare(item.createDate, lastSixMonth) > 0).ToList();
            return Ok(new
            {
                success = true,
                data = new
                {
                    requester,
                    user,
                },
            });


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
                // if (body.id.Length == 0)
                // {
                //     if (dept.ToLower() != "acc" && dept != "admin")
                //     {
                //         data = db.CIP.Where<cipSchema>(item => item.cc == deptCode && item.status == "open").ToList<cipSchema>();
                //         db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "active");
                //     }
                //     else
                //     {
                //         data = db.CIP.Where<cipSchema>(item => item.status == "open").ToList<cipSchema>();
                //         db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "active");
                //     }
                // }
                // else
                // {
                foreach (int id in body.id)
                {
                    cipSchema item = db.CIP.Find(id);
                    db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == id).FirstOrDefault();
                    data.Add(item);
                }
                // }

                MemoryStream stream = new MemoryStream();
                using (ExcelPackage excel = new ExcelPackage(stream))
                {
                    excel.Workbook.Worksheets.Add("sheet1");

                    List<string[]> header = new List<string[]>()
                {
                    new string[] { "Type work", "Project No.", "CIP No.", "Sub CIP No.", "PO NO.", "VENDER CODE", "VENDER", "ACQ-DATE (ETD)", "INV DATE",
                    "RECEIVED DATE", "INV NO.", "NAME (ENGLISH)", "Qty.", "EX.RATE", "CUR", "PER UNIT \n (JPY/USD)",
                    "TOTAL (JPY/USD)", "TOTAL (THB)", "AVERAGE FREIGHT (JPY/USD)", "AVERAGE INSURANCE (JPY/USD)", "TOTAL (JPY/USD)",
                    "Grand TOTAL (THB)", "PER UNIT (THB)", "CC", "TOTAL OF CIP (THB)", "Budget code", "PR.DIE/JIG", "Model", "PART No./DIE No.",
                    "Operating Date (Plan)", "Operating Date (Act)", "Result", "Reason diff (NG) Budget&Actual", "Fixed Asset Code",
                    "CLASS FIXED ASSET", "Fix Asset Name (English only)", "Serial No.", "part\nnumber\nDie No", "Process Die", "Model",
                    "Cost Center of User", "Transfer to supplier", "ให้ขึ้น Fix Asset  กี่ตัว", "New BFMor Add BFM", "Reason for Delay", "Add CIP/BFM No.",
                    "REMARK", "ITC--> BOI TYPE (Machine / Die / Sparepart / NON BOI)"
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
                    worksheet.Cells["A1:AC1"].Merge = true;
                    worksheet.Cells["A1"].Value = "Data for Accounting Dept.";
                    worksheet.Cells["A1"].Style.Font.Bold = true;
                    worksheet.Cells["A1"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#9FE5E6"));
                    worksheet.Cells["A1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A1"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    worksheet.Cells["AD1:AV1"].Merge = true;
                    worksheet.Cells["AD1"].Value = "Data for User confirm";
                    worksheet.Cells["AD1"].Style.Font.Bold = true;
                    worksheet.Cells["AD1"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                    worksheet.Cells["AD1"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["AD1"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    worksheet.Cells["AD2:AU2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F1ECB9"));
                    worksheet.Cells["AV2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#57A868"));

                    worksheet.Column(1).Width = 13;
                    worksheet.Column(2).Width = 11;
                    worksheet.Column(6).Width = 11;
                    worksheet.Column(8).Width = 11;
                    worksheet.Column(10).Width = 11;
                    worksheet.Column(11).Width = 11; // INV no.
                    worksheet.Column(12).Width = 13;
                    worksheet.Column(14).Width = 13;
                    worksheet.Column(15).Width = 12;
                    worksheet.Column(16).Width = 12;
                    worksheet.Column(17).Width = 12;
                    worksheet.Column(18).Width = 12;
                    worksheet.Column(19).Width = 12;
                    worksheet.Column(20).Width = 12;
                    worksheet.Column(21).Width = 12;
                    worksheet.Column(23).Width = 12;
                    worksheet.Column(25).Width = 12; // AA
                    worksheet.Column(27).Width = 12; // AA
                    worksheet.Column(41).Width = 12; // AO
                    worksheet.Column(43).Width = 15; // AO
                    worksheet.Row(2).Height = 80;
                    worksheet.Row(1).Height = 30;
                    worksheet.Row(2).Style.Font.Bold = true;
                    worksheet.Row(2).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    worksheet.Row(2).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A2:R2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));
                    worksheet.Cells["U2:AC2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));
                    worksheet.Cells["S2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C7BDF9"));
                    worksheet.Cells["T2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#CCF9BD"));
                    worksheet.Cells["A2:AV2"].Style.WrapText = true;
                    worksheet.Cells["A2:AV2"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells["A2:AV2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells["A2:AV2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells["A2:AV2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                    worksheet.Cells["AW2"].Value = "Requester \n prepare";
                    worksheet.Column(49).Width = 23;
                    worksheet.Cells["AX2"].Value = "Requester \n check";
                    worksheet.Column(50).Width = 23;
                    worksheet.Cells["AY2"].Value = "Requester \n approve";
                    worksheet.Column(51).Width = 23;

                    worksheet.Cells["AZ2"].Value = "User \n prepare";
                    worksheet.Column(52).Width = 23;
                    worksheet.Cells["BA2"].Value = "User \n check";
                    worksheet.Column(53).Width = 23;
                    worksheet.Cells["BB2"].Value = "User \n approve";
                    worksheet.Column(54).Width = 23;

                    worksheet.Cells["BC2"].Value = "ITC \n BOI";
                    worksheet.Column(55).Width = 23;

                    worksheet.Cells["BD2"].Value = "ACC check diff budget";
                    worksheet.Column(56).Width = 23;

                    worksheet.Cells["BE2"].Value = "ACC approve diff budget";
                    worksheet.Column(57).Width = 23;

                    int row = 3;
                    foreach (cipSchema item in data)
                    {


                        List<string[]> cellData = new List<string[]>()
                    {
                        new string [] {
                            item.workType, item.projectNo,
                            item.cipNo, item.subCipNo, item.poNo ,item.vendorCode, item.vendor, item.acqDate, item.invDate, item.receivedDate,
                             item.invNo, item.name, item.qty, item.exRate, item.cur, item.perUnit, item.totalJpy, item.totalThb, item.averageFreight,
                             item.averageInsurance, item.totalJpy_1, item.totalThb_1, item.perUnitThb, item.cc, item.totalOfCip, item.budgetCode, item.prDieJig,
                             item.model, item.partNoDieNo,
                             item.cipUpdate?.planDate, item.cipUpdate?.actDate, item.cipUpdate?.result, item.cipUpdate?.reasonDiff, item.cipUpdate?.fixedAssetCode,
                             item.cipUpdate?.classFixedAsset, item.cipUpdate?.fixAssetName, item.cipUpdate?.serialNo, item.cipUpdate?.partNumberDieNo,
                             item.cipUpdate?.processDie, item.cipUpdate?.model, item.cipUpdate?.costCenterOfUser, item.cipUpdate?.tranferToSupplier,
                             item.cipUpdate?.upFixAsset, item.cipUpdate?.newBFMorAddBFM, item.cipUpdate?.reasonForDelay,
                             (item.cipUpdate?.addCipBfmNo == "-") ? "" : item.cipUpdate?.addCipBfmNo, item.cipUpdate?.remark, (item.cipUpdate?.boiType == "-") ? "" : item.cipUpdate?.boiType
                        }

                    };
                        IExcelDataValidationList typeWork = worksheet.DataValidations.AddListValidation("A" + row);
                        typeWork.Formula.Values.Add("Domestic");
                        typeWork.Formula.Values.Add("Domestic-DIE");
                        typeWork.Formula.Values.Add("Oversea");
                        typeWork.Formula.Values.Add("Project ENG3");
                        typeWork.Formula.Values.Add("Project-MSC");

                        IExcelDataValidationList newOrAddBFM = worksheet.DataValidations.AddListValidation("AR" + row);
                        newOrAddBFM.Formula.Values.Add("Add BFM");
                        newOrAddBFM.Formula.Values.Add("NEW BFM");

                        worksheet.Cells["AF" + row].Formula = "=+IF(AI" + row + "=\"\",\"\",IF(Z" + row + "=\"\",\"\",IF(AND(MID(Z" + row + ",5,2)=\"09\",LEFT(AI" + row + ",2)=\"06\"),\"OK\",IF(AND(OR(MID(Z" + row + ",5,2)=\"31\",MID(Z" + row + ",5,2)=\"34\"),LEFT(AI" + row + ",2)=\"28\"),\"OK\",IF(MID(Z" + row + ",5,2)=LEFT(AI" + row + ",2),\"OK\",\"NG\")))))";
                        worksheet.Cells["AF" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));
                        worksheet.Cells["AH" + row].Formula = "=IF(LEFT(AI3,2)=\"28\",\"SOFTWARE\",IF(LEFT(AI" + row + ",2)=\"02\",\"BUILDING\",IF(LEFT(AI" + row + ",2)=\"03\",\"STRUCTURE\",IF(LEFT(AI" + row + ",2)=\"04\",\"MACHINE\",IF(LEFT(AI" + row + ",2)=\"05\",\"VEHICLE\",IF(LEFT(AI" + row + ",2)=\"06\",\"TOOLS\",IF(LEFT(AI" + row + ",2)=\"07\",\"FURNITURE\",IF(LEFT(AI" + row + ",2)=\"08\",\"DIES\",\"\"))))))))";
                        worksheet.Cells["AH" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));

                        // set pink row space
                        worksheet.Cells["AD" + row + ":AE" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                        worksheet.Cells["AG" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                        worksheet.Cells["AI" + row + ":AV" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                        // set pink row space
                        worksheet.Cells[row, 1].LoadFromArrays(cellData);
                        worksheet.Cells["A" + row + ":BE" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells["A" + row + ":BE" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells["A" + row + ":BE" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells["A" + row + ":BE" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                        List<ApprovalSchema> approve = db.APPROVAL.Where<ApprovalSchema>(approve => approve.cipSchemaid == item.id).ToList();

                        ApprovalSchema save = approve.Find(action => action.onApproveStep == "save");
                        if (save != null)
                        {
                            userSchema userSave = db.USERS.Find(save.empNo);
                            worksheet.Cells["AW" + row].Value = userSave.name;
                        }

                        ApprovalSchema ccCheck = approve.Find(action => action.onApproveStep == "cc-checked");

                        if (ccCheck != null)
                        {
                            userSchema userCheck = db.USERS.Find(ccCheck.empNo);
                            worksheet.Cells["AX" + row].Value = userCheck.name;
                        }

                        ApprovalSchema ccApprove = approve.Find(action => action.onApproveStep == "cc-approved");

                        if (ccApprove != null)
                        {
                            userSchema userApprove = db.USERS.Find(ccApprove.empNo);
                            worksheet.Cells["AY" + row].Value = userApprove.name;
                        }

                        ApprovalSchema costPrepare = approve.Find(action => action.onApproveStep == "cost-prepared");
                        if (costPrepare != null)
                        {
                            userSchema userCostPrepare = db.USERS.Find(costPrepare.empNo);
                            worksheet.Cells["AZ" + row].Value = userCostPrepare.name;
                        }
                        ApprovalSchema costCheck = approve.Find(action => action.onApproveStep == "cost-checked");
                        if (costCheck != null)
                        {
                            userSchema costUserCheck = db.USERS.Find(costCheck.empNo);
                            worksheet.Cells["BA" + row].Value = costUserCheck.name;
                        }
                        ApprovalSchema costApprove = approve.Find(action => action.onApproveStep == "cost-approved");
                        if (costApprove != null)
                        {
                            userSchema costUserApprove = db.USERS.Find(costApprove.empNo);
                            worksheet.Cells["BB" + row].Value = costUserApprove.name;
                        }

                        ApprovalSchema itcConfirm = approve.Find(action => action.onApproveStep == "itc-confirmed");

                        if (itcConfirm != null)
                        {
                            userSchema itcConfirmUser = db.USERS.Find(itcConfirm.empNo);
                        }
                        ApprovalSchema diffChecked = approve.Find(action => action.onApproveStep == "diff-checked");
                        if (diffChecked != null)
                        {
                            userSchema accDiffChecker = db.USERS.Find(diffChecked.empNo);
                            worksheet.Cells["BD" + row].Value = accDiffChecker.name;
                        }

                        ApprovalSchema diffApproved = approve.Find(action => action.onApproveStep == "diff-approved");
                        if (diffApproved != null)
                        {
                            userSchema diffApproveUser = db.USERS.Find(diffApproved.empNo);
                            worksheet.Cells["BE" + row].Value = diffApproveUser.name;
                        }
                        worksheet.Cells["AW2:BE2"].Style.Fill.SetBackground(ColorTranslator.FromHtml("#A9F3E1"));


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
        public ActionResult cipUpdate(string id)
        {
            try
            {

                cipUpdateSchema data = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == Int32.Parse(id)).FirstOrDefault();

                return Ok(new { success = true, data, });
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }

        [HttpPost("reject/requester")]
        public ActionResult rejectRequester(rejectCip body)
        {
            try
            {
                foreach (Int32 id in body.id)
                {
                    cipSchema data = db.CIP.Find(id);
                    if (data != null)
                    {
                        data.status = "reject";
                        data.commend = body.commend;
                        db.CIP.Update(data);

                        cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == id).FirstOrDefault();
                        if (cipUpdate != null)
                        {
                            db.CIP_UPDATE_REJECT.Add(
                                                    new cipUpdateRejectSchema
                                                    {
                                                        actDate = cipUpdate.actDate,
                                                        addCipBfmNo = cipUpdate.addCipBfmNo,
                                                        boiType = cipUpdate.boiType,
                                                        cipSchemaid = cipUpdate.cipSchemaid,
                                                        classFixedAsset = cipUpdate.classFixedAsset,
                                                        commend = body.commend,
                                                        costCenterOfUser = cipUpdate.costCenterOfUser,
                                                        createDate = DateTime.Now.ToString("yyyy/MM/dd"),
                                                        fixAssetName = cipUpdate.fixAssetName,
                                                        fixedAssetCode = cipUpdate.fixedAssetCode,
                                                        model = cipUpdate.model,
                                                        newBFMorAddBFM = cipUpdate.newBFMorAddBFM,
                                                        partNumberDieNo = cipUpdate.partNumberDieNo,
                                                        planDate = cipUpdate.planDate,
                                                        processDie = cipUpdate.processDie,
                                                        reasonDiff = cipUpdate.reasonDiff,
                                                        reasonForDelay = cipUpdate.reasonForDelay,
                                                        remark = cipUpdate.remark,
                                                        result = cipUpdate.result,
                                                        serialNo = cipUpdate.serialNo,
                                                        tranferToSupplier = cipUpdate.tranferToSupplier,
                                                        upFixAsset = cipUpdate.upFixAsset,
                                                    }
                                                );
                            // db.CIP_UPDATE.Remove(cipUpdate);
                            cipUpdate.status = "reject";
                            db.CIP_UPDATE.Update(cipUpdate);
                        }

                        List<ApprovalSchema> approve = db.APPROVAL.Where<ApprovalSchema>(item => item.cipSchemaid == id).ToList();
                        db.APPROVAL.RemoveRange(approve);
                    }
                    db.SaveChanges();
                }

                return Ok(new { success = true, message = "Reject CIP success." });
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }

        private void costCenterPrepareRejection(cipSchema cip, string commend)
        {

            cip.status = "save";
            cip.commend = commend;
            db.CIP.Update(cip);

            List<ApprovalSchema> approve = db.APPROVAL.Where<ApprovalSchema>(item => item.cipSchemaid == cip.id && (item.onApproveStep == "cc-approved" || item.onApproveStep == "cc-checked")).ToList();
            db.APPROVAL.RemoveRange(approve);

        }

        [HttpPost("reject/user")]
        public ActionResult rejectUser(rejectCostCenter body)
        {
            try
            {
                List<cipUpdateRejectSchema> rejectHistory = new List<cipUpdateRejectSchema>();
                List<cipSchema> updateCip = new List<cipSchema>();
                List<ApprovalSchema> removeApproveHistory = new List<ApprovalSchema>();
                foreach (string id in body.id)
                {
                    cipSchema cip = db.CIP.Find(Int32.Parse(id));
                    db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == Int32.Parse(id)).FirstOrDefault();

                    if (cip.status == "cc-approved")
                    {
                        costCenterPrepareRejection(cip, body.commend);
                    }
                    else
                    {
                        List<ApprovalSchema> approve = db.APPROVAL.Where<ApprovalSchema>(item =>
                               item.cipSchemaid == Int32.Parse(id) && item.onApproveStep.StartsWith("cost")).ToList();

                        removeApproveHistory.AddRange(approve);
                        cip.commend = body.commend;
                        cip.status = "cc-approved";
                        updateCip.Add(cip);

                        db.APPROVAL.RemoveRange(removeApproveHistory);
                        db.CIP.UpdateRange(updateCip);
                    }
                    // Remove approve history

                    // db.APPROVAL.RemoveRange(approve);
                    // Remove approve history

                    rejectHistory.Add(new cipUpdateRejectSchema
                    {
                        actDate = cip.cipUpdate.actDate,
                        addCipBfmNo = cip.cipUpdate.addCipBfmNo,
                        boiType = cip.cipUpdate.boiType,
                        cipSchemaid = cip.cipUpdate.cipSchemaid,
                        classFixedAsset = cip.cipUpdate.classFixedAsset,
                        commend = body.commend,
                        costCenterOfUser = cip.cipUpdate.costCenterOfUser,
                        createDate = DateTime.Now.ToString("yyyy/MM/dd"),
                        fixAssetName = cip.cipUpdate.fixAssetName,
                        fixedAssetCode = cip.cipUpdate.fixedAssetCode,
                        model = cip.cipUpdate.model,
                        newBFMorAddBFM = cip.cipUpdate.newBFMorAddBFM,
                        partNumberDieNo = cip.cipUpdate.partNumberDieNo,
                        planDate = cip.cipUpdate.planDate,
                        processDie = cip.cipUpdate.processDie,
                        reasonDiff = cip.cipUpdate.reasonDiff,
                        reasonForDelay = cip.cipUpdate.reasonForDelay,
                        remark = cip.cipUpdate.remark,
                        result = cip.cipUpdate.result,
                        serialNo = cip.cipUpdate.serialNo,
                        tranferToSupplier = cip.cipUpdate.tranferToSupplier,
                        upFixAsset = cip.cipUpdate.upFixAsset,
                    });
                }

                db.CIP_UPDATE_REJECT.AddRange(rejectHistory);
                db.SaveChanges();

                return Ok(new { success = true, message = "Reject CIP success. " });
            }
            catch (System.Exception e)
            {
                return Problem(e.StackTrace);
            }
        }

        [HttpPut]
        public ActionResult updateCip(editCip body)
        {
            cipSchema cip = db.CIP.Find(body.id);

            cip.cipNo = body.cipNo;
            cip.subCipNo = body.subCipNo;
            cip.acqDate = body.acqDate;
            cip.averageFreight = body.averageFreight;
            cip.averageInsurance = body.averageInsurance;
            cip.budgetCode = body.budgetCode;
            cip.cc = body.cc;
            cip.cur = body.cur;
            cip.exRate = body.exRate;
            cip.invDate = body.invDate;
            cip.invNo = body.invNo;
            cip.model = body.model;
            cip.name = body.name;
            cip.partNoDieNo = body.partNoDieNo;
            cip.perUnit = body.perUnit;
            cip.perUnitThb = body.perUnitThb;
            cip.poNo = body.poNo;
            cip.prDieJig = body.prDieJig;
            cip.projectNo = body.projectNo;
            cip.qty = body.qty;
            cip.receivedDate = body.receivedDate;
            cip.totalJpy = body.totalJpy;
            cip.totalJpy_1 = body.totalJpy_1;
            cip.totalOfCip = body.totalOfCip;
            cip.totalThb = body.totalThb;
            cip.totalThb_1 = body.totalThb_1;
            cip.vendor = body.vendor;
            cip.vendorCode = body.vendorCode;
            cip.workType = body.workType;
            // cip.status = "draft";
            cip.commend = null;


            db.CIP.Update(cip);
            db.SaveChanges();
            return Ok(new { success = true, message = "Update cip success." });
        }

        [HttpDelete("{id}")]
        public ActionResult deleteCip(string id)
        {
            cipSchema cip = db.CIP.Find(Int32.Parse(id));
            cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == Int32.Parse(id)).FirstOrDefault();

            db.CIP.Remove(cip);

            if (cipUpdate != null)
            {
                db.CIP_UPDATE.Remove(cipUpdate);
            }
            db.SaveChanges();

            return Ok(
                new
                {
                    success = true,
                    message = "Delete CIP success.",
                }
            );
        }


        [HttpPost("testSendMail")]
        public async Task<ActionResult> test()
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {

                    using (MultipartFormDataContent formData = new MultipartFormDataContent())
                    {
                        var values = new[] {
                        new KeyValuePair<string, string>("from", "wanpen@mail.canon"),
                        new KeyValuePair<string, string>("to", "suwannason@mail.canon"),
                        new KeyValuePair<string, string>("subject", "system-imp@email.com"),
                        new KeyValuePair<string, string>("text", "system-imp@email.com"),
                    };
                        foreach (var keyValuePair in values)
                        {
                            formData.Add(new StringContent(keyValuePair.Value),
                            String.Format("\"{0}\"", keyValuePair.Key));
                        }


                        client.Timeout = TimeSpan.FromSeconds(20);
                        HttpResponseMessage result = await client.PostAsync(node_api + "/middleware/email/sendmail", formData);
                        string input = await result.Content.ReadAsStringAsync();
                        client.Dispose();

                        return Ok(input);
                    }
                }
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }
    }
}