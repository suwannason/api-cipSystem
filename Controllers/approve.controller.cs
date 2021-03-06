
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

using cip_api.models;
using cip_api.request;
using System.Collections.Generic;
using System.Linq;
using System;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation.Contracts;
using System.Globalization;

namespace cip_api.controllers
{
    [ApiController, Route("[controller]"), Authorize]
    public class approvalController : ControllerBase
    {
        private readonly Database db;
        private IConfiguration _config;
        private readonly string ldap_auth;
        public approvalController(Database _db, IConfiguration config, IEndpoint setting)
        {
            db = _db;
            _config = config;
            ldap_auth = setting.ldap_auth;
        }

        private List<PermissionSchema> GetPermissions(string empNo)
        {
            try
            {
                return db.PERMISSIONS.Where<PermissionSchema>(item => item.empNo == empNo).ToList<PermissionSchema>();
            }
            catch (Exception e)
            {
                Console.WriteLine("GetPermissions " + e.Message);
                return null;
            }

        }

        [HttpGet("draft")]
        public ActionResult draft()
        {
            string dept = User.FindFirst("dept")?.Value;

            List<cipSchema> data = db.CIP.Where<cipSchema>
            (item => item.status == "draft" && item.cc == dept).ToList<cipSchema>();
            return Ok(
                new
                {
                    success = true,
                    message = "CIP on draft.",
                    data,
                }
            );
        }
        [HttpGet("cc")]
        public ActionResult cc(string user, string dcode)
        {
            try
            {
                string username = ""; string deptCode = "";
                if (User == null)
                {
                    username = user;
                    deptCode = dcode;
                }
                else
                {
                    username = User.FindFirst("username")?.Value;
                    deptCode = User.FindFirst("deptCode")?.Value;
                }
                List<PermissionSchema> permissions = GetPermissions(username);

                PermissionSchema preparer = permissions.Find(e => e.action == "prepare");

                PermissionSchema checker = permissions.Find(e => e.action == "checker");

                List<PermissionSchema> approver = permissions.FindAll(e => e.action == "approver");

                string message = "";
                string permissions_page = "";

                List<cipSchema> data = new List<cipSchema>();
                if (approver.Count != 0)
                {
                    permissions_page = "approve";
                    message = "CIP for data approve. (" + deptCode + " )";

                    List<string> multidept = deptCode.Split(',').ToList();
                    if (multidept.Count > 1)
                    {
                        foreach (string code in multidept)
                        {
                            if (code == "55XX")
                            {
                                data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "cc-checked" && item.cc.StartsWith("55")).ToList<cipSchema>());
                            }
                            else
                            {
                                data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "cc-checked" && item.cc.IndexOf(code) != -1).ToList<cipSchema>());
                            }

                        }
                    }
                    else
                    {
                        if (deptCode == "55XX")
                        {
                            data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "cc-checked" && item.cc.StartsWith("55")).ToList<cipSchema>());
                        }
                        else
                        {
                            data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "cc-checked" && item.cc.IndexOf(deptCode) != -1).ToList<cipSchema>());
                        }
                    }
                }
                if (checker != null && data.Count == 0)
                {
                    List<string> multidept = deptCode.Split(',').ToList();
                    message = "CIP for data check. (" + deptCode + " )";
                    permissions_page = "check";

                    if (multidept.Count > 1)
                    {
                        foreach (string code in multidept)
                        {
                            if (code == "55XX")
                            {
                                data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "save" && item.cc.StartsWith("55")).ToList<cipSchema>());
                            }
                            else
                            {
                                data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "save" && item.cc.IndexOf(code) != -1).ToList<cipSchema>());
                            }

                        }
                    }
                    else
                    {
                        if (deptCode == "55XX")
                        {
                            data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "save" && item.cc.StartsWith("55")).ToList<cipSchema>());
                        }
                        else
                        {
                            data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "save" && item.cc.IndexOf(deptCode) != -1).ToList<cipSchema>());
                        }
                    }

                }
                if (preparer != null)
                {
                    message = "CIP for confirm upload. (" + deptCode + " )";
                    permissions_page = "prepare";
                    List<string> multidept = deptCode.Split(',').ToList();

                    if (multidept.Count > 1)
                    {
                        foreach (string code in multidept)
                        {
                            if (code == "55XX")
                            {
                                data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "draft" && item.cc.StartsWith("55")).ToList<cipSchema>());
                            }
                            else
                            {
                                data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "draft" && item.cc.Contains(code)).ToList<cipSchema>());
                            }

                        }
                    }
                    else
                    {
                        if (deptCode == "55XX")
                        {
                            data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "draft" && item.cc.StartsWith("55")).ToList<cipSchema>());
                        }
                        else
                        {
                            data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "draft" && item.cc.Contains(deptCode)).ToList<cipSchema>());
                        }
                    }
                }

                /* Handle message */

                cipSchema draftStatus = data.Find(e => e.status == "draft");
                cipSchema saveStatus = data.Find(e => e.status == "save");

                if (draftStatus != null && saveStatus != null)
                {
                    permissions_page = "prepare";
                    data = data.FindAll(e => e.status == "draft");
                }
                else if (draftStatus == null && saveStatus != null)
                {
                    permissions_page = "check";
                }

                data = data.Distinct().ToList();

                if (user != null)
                {
                    return Ok(data.Count);
                }
                return Ok(
                  new
                  {
                      success = true,
                      message,
                      data,
                      permission = permissions_page,
                  }
              );
            }
            catch (Exception e)
            {
                Console.Write(e.Message);
                return Problem(e.StackTrace);
            }
        }

        [HttpGet("costCenter"),]
        public ActionResult costCenter(string user, string code)
        {
            string username = ""; string deptCode = "";
            if (User == null)
            {
                username = user;
                deptCode = code;
            }
            else
            {
                username = User.FindFirst("username")?.Value;
                deptCode = User.FindFirst("deptCode")?.Value;
            }

            List<PermissionSchema> permissions = GetPermissions(username);

            PermissionSchema checker = permissions.Find(e => e.action == "checker");
            PermissionSchema approver = permissions.Find(e => e.action == "approver");
            PermissionSchema prepare = permissions.Find(e => e.action == "prepare");

            List<cipSchema> data = new List<cipSchema>();
            string message = "";
            string permissions_page = "";

            if (prepare != null)
            {
                List<string> multidept = deptCode.Split(',').ToList();
                message = "CIP on Cost center prepare. (" + deptCode + " )";
                permissions_page = "prepare";

                foreach (string dept in multidept)
                {
                    List<cipUpdateSchema> cipUpdate = new List<cipUpdateSchema>();
                    if (dept == "55XX")
                    {
                        cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.costCenterOfUser.StartsWith("55") && item.status == "active").ToList();
                    }
                    else
                    {
                        cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.costCenterOfUser == dept && item.status == "active").ToList();
                    }

                    foreach (cipUpdateSchema cip_update in cipUpdate)
                    {
                        cipSchema dataItem = db.CIP.Find(cip_update.cipSchemaid);
                        if (dataItem.cc != cip_update.costCenterOfUser && dataItem.status == "cc-approved")
                        { // Tranfer cross dept
                            if (dataItem.cc.StartsWith("55") && cip_update.costCenterOfUser.StartsWith("55")
                                || (cip_update.costCenterOfUser.StartsWith("55") && !dataItem.cc.StartsWith("55"))
                                || (dataItem.cc == "2130" && (cip_update.costCenterOfUser == "2140" || cip_update.costCenterOfUser == "9555"))
                                || (dataItem.cc == "2140" && (cip_update.costCenterOfUser == "2130" || cip_update.costCenterOfUser == "9555"))
                                || (dataItem.cc == "9555" && (cip_update.costCenterOfUser == "2140" || cip_update.costCenterOfUser == "2130"))
                                || (dataItem.cc == "5610" && cip_update.costCenterOfUser == "5615")
                                || (dataItem.cc == "5615" && cip_update.costCenterOfUser == "5610")
                                || (dataItem.cc == "5650" && cip_update.costCenterOfUser == "9333")
                                || (dataItem.cc == "9333" && cip_update.costCenterOfUser == "5650")
                                || (dataItem.cc == "5670" && cip_update.costCenterOfUser == "9444")
                                || (dataItem.cc == "9444" && cip_update.costCenterOfUser == "5670")
                                )
                            {
                                List<PermissionSchema> mutipleUserCheck = db.PERMISSIONS.Where<PermissionSchema>(permission => permission.deptCode.Contains(dept) && permission.action == "prepare").ToList();

                                ApprovalSchema prepareDuplicatCheck = db.APPROVAL.Where<ApprovalSchema>(approve =>
                                                                      approve.cipSchemaid == dataItem.id && approve.onApproveStep == "save").FirstOrDefault();

                                PermissionSchema findUser = mutipleUserCheck.Find(user => user.empNo == prepareDuplicatCheck.empNo);

                                if (findUser == null)
                                {
                                    data.Add(dataItem);
                                }
                            }
                            else
                            {
                                Console.WriteLine("in else prepare");
                                if (dataItem.cc != cip_update.costCenterOfUser)
                                {
                                    data.Add(dataItem);
                                }
                            }

                        }
                    }
                }
            }
            if (approver != null)
            {
                List<string> multidept = deptCode.Split(',').ToList();
                message = "CIP on Cost center approve. (" + deptCode + " )";
                permissions_page = "approve";
                foreach (string dept in multidept)
                {
                    List<cipUpdateSchema> cipUpdate = new List<cipUpdateSchema>();
                    if (dept == "55XX")
                    {
                        cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.costCenterOfUser.StartsWith("55") && item.status == "active").ToList();
                    }
                    else
                    {
                        cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => dept.Contains(item.costCenterOfUser) && item.status == "active").ToList();
                    }

                    foreach (cipUpdateSchema cip_update in cipUpdate)
                    {
                        cipSchema dataItem = db.CIP.Find(cip_update.cipSchemaid);
                        if (dataItem.cc != cip_update.costCenterOfUser && dataItem.status == "cost-checked")
                        {
                            ApprovalSchema prepareDuplicatCheck = db.APPROVAL.Where<ApprovalSchema>(approve =>
                                                                  approve.cipSchemaid == dataItem.id && approve.empNo == username && approve.onApproveStep == "cc-approved").FirstOrDefault();
                            if (prepareDuplicatCheck == null)
                            {
                                data.Add(dataItem);
                            }
                        }


                    }
                }

            }

            if (checker != null)
            {
                Console.WriteLine("check");
                List<string> multidept = deptCode.Split(',').ToList();
                message = "CIP on Cost center check. (" + deptCode + " )";
                permissions_page = "check";

                foreach (string dept in multidept)
                {
                    List<cipUpdateSchema> cipUpdate = new List<cipUpdateSchema>();
                    if (dept == "55XX")
                    {
                        cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.costCenterOfUser.StartsWith("55") && item.status == "active").ToList();
                    }
                    else
                    {
                        cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => dept.Contains(item.costCenterOfUser) && item.status == "active").ToList();
                    }

                    foreach (cipUpdateSchema cip_update in cipUpdate)
                    {
                        cipSchema dataItem = db.CIP.Find(cip_update.cipSchemaid);
                        if (dataItem.cc != cip_update.costCenterOfUser && dataItem.status == "cost-prepared")
                        {
                            List<PermissionSchema> mutipleUserCheck = db.PERMISSIONS.Where<PermissionSchema>(permission => permission.deptCode.Contains(dept) && permission.action == "checker").ToList();

                            ApprovalSchema prepareDuplicatCheck = db.APPROVAL.Where<ApprovalSchema>(approve =>
                                                                  approve.cipSchemaid == dataItem.id && approve.onApproveStep == "cost-checked").FirstOrDefault();

                            if (prepareDuplicatCheck == null)
                            {
                                data.Add(dataItem);
                            }
                            else
                            {
                                PermissionSchema findUser = mutipleUserCheck.Find(user => user.empNo == prepareDuplicatCheck.empNo);

                                if (findUser == null)
                                {
                                    data.Add(dataItem);
                                }
                            }


                        }
                    }
                }

            }

            /* Handle message */

            cipSchema draftStatus = data.Find(e => e.status == "draft");
            cipSchema saveStatus = data.Find(e => e.status == "save");

            if (draftStatus != null && saveStatus != null)
            {
                permissions_page = "prepare";
                data = data.FindAll(e => e.status == "draft");
            }
            else if (draftStatus == null && saveStatus != null)
            {
                permissions_page = "check";
            }

            data = data.Distinct().ToList();
            if (user != null)
            {
                return Ok(data.Count);
            }
            return Ok(new { success = true, message, data, permission = permissions_page, });
        }

        [HttpPut("approve/cc")]
        public ActionResult approve(Approve body)
        {
            List<cipSchema> updateRange = new List<cipSchema>();
            List<ApprovalSchema> approvals = new List<ApprovalSchema>();

            string username = User.FindFirst("username")?.Value;
            string deptCode = User.FindFirst("deptCode")?.Value;

            Console.WriteLine(username);

            List<PermissionSchema> permissions = GetPermissions(username);

            PermissionSchema prepared = permissions.Find(e => e.action == "prepare");
            PermissionSchema checker = permissions.Find(e => e.action == "checker");
            List<PermissionSchema> approver = permissions.FindAll(e => e.action == "approver");

            string status = "";
            foreach (string item in body.id)
            {
                Int32 id = Int32.Parse(item);
                ApprovalSchema approve = new ApprovalSchema();
                cipSchema data = db.CIP.Find(id);

                if (data.status == "draft")
                {
                    data.status = "save";
                    status = "save";
                }
                else if (data.status == "save")
                {
                    if ((checker.deptCode.IndexOf(data.cc) != -1) && data.status != "cc-checked")
                    {
                        data.status = "cc-checked";
                        status = "cc-checked";
                    }
                    else if (checker.deptCode == "55XX" && data.status != "cc-checked" && data.cc.IndexOf("55") == 0)
                    {
                        data.status = "cc-checked";
                        status = "cc-checked";
                    }
                }
                else if (data.status == "cc-checked")
                {
                    Console.WriteLine("case approve: " + deptCode);
                    if (deptCode.IndexOf("55XX") == 0)
                    {

                        data.status = "cc-approved";
                        status = "cc-approved";

                    }
                    else // requester approve 55XX dept
                    {
                        data.status = "cc-approved";
                        status = "cc-approved";

                    }
                }
                approve.onApproveStep = status;
                approve.empNo = username;
                approve.cipSchemaid = id;
                approve.date = DateTime.Now.ToString("yyyy/MM/dd");

                updateRange.Add(data);
                if (status != "")
                {
                    approvals.Add(approve);
                }

            }

            db.CIP.UpdateRange(updateRange);
            db.APPROVAL.AddRange(approvals);

            db.SaveChanges();

            return Ok(new
            {
                success = true,
                message = "Approve CIP success"
            });
        }

        [HttpPut("approve/costCenter")]
        public ActionResult costCenter(Approve body)
        {
            List<cipSchema> updateRange = new List<cipSchema>();
            List<ApprovalSchema> approvals = new List<ApprovalSchema>();
            string username = User.FindFirst("username")?.Value;
            string deptCode = User.FindFirst("deptCode")?.Value;

            string status = "";

            foreach (string item in body.id)
            {
                Int32 id = Int32.Parse(item);
                ApprovalSchema approve = new ApprovalSchema();
                cipSchema data = db.CIP.Find(id);
                db.CIP_UPDATE.Where<cipUpdateSchema>(row => row.cipSchemaid == id).FirstOrDefault();

                if (data.status == "cost-prepared")
                {
                    List<string> multidept = deptCode.Split(',').ToList();

                    foreach (string deptCodeItem in multidept)
                    {
                        if (deptCodeItem == "55XX")
                        {
                            status = "cost-checked";
                        }
                        else
                        {
                            if (data.cipUpdate.costCenterOfUser == deptCodeItem)
                            {
                                status = "cost-checked";
                            }
                        }
                    }
                }
                else if (data.status == "cost-checked")
                {
                    List<string> multidept = deptCode.Split(',').ToList();

                    foreach (string deptCodeItem in multidept)
                    {
                        if (deptCodeItem == "55XX")
                        {
                            status = "cost-approved";
                        }
                        else
                        {
                            if (data.cipUpdate.costCenterOfUser == deptCodeItem)
                            {
                                status = "cost-approved";
                            }
                        }
                    }
                }
                else if (data.status == "cc-approved")
                {
                    List<string> multidept = deptCode.Split(',').ToList();

                    foreach (string deptCodeItem in multidept)
                    {
                        if (deptCodeItem == "55XX")
                        {
                            status = "cost-prepared";
                        }
                        else
                        {
                            if (data.cipUpdate.costCenterOfUser == deptCodeItem)
                            {
                                status = "cost-prepared";
                            }
                        }
                    }
                }
                data.status = status;
                if (status == "cost-prepared")
                {
                    data.commend = null;
                }

                approve.onApproveStep = status;
                approve.empNo = username;
                approve.cipSchemaid = id;
                approve.date = DateTime.Now.ToString("yyyy/MM/dd");

                updateRange.Add(data);
                approvals.Add(approve);
            }
            db.CIP.UpdateRange(updateRange);
            db.APPROVAL.AddRange(approvals);
            db.SaveChanges();

            return Ok(new
            {
                success = true,
                message = "Approve CIP success"
            });
        }

        [HttpGet("download")]
        public ActionResult download()
        {
            try
            {
                string deptCode = User.FindFirst("deptCode")?.Value;
                // Console.WriteLine(deptCode);
                List<cipUpdateSchema> cipUpdate = new List<cipUpdateSchema>();

                if (deptCode != "55XX")
                {
                    cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => deptCode.IndexOf(item.costCenterOfUser) != -1).ToList();
                }
                else
                {
                    cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => deptCode.IndexOf("55") != -1).ToList();
                }


                MemoryStream stream = new MemoryStream();

                using (ExcelPackage excel = new ExcelPackage(stream))
                {
                    excel.Workbook.Worksheets.Add("sheet1");

                    List<string[]> header = new List<string[]>()
                {
                    new string[] { "Type work", "Project No.", "CIP No.", "Sub CIP No.", "PO NO.", "VENDER CODE", "VENDER", "ACQ-DATE (ETD)", "INV DATE",
                    "RECEIVED DATE", "INV NO.", "NAME (ENGLISH)", "Qty.", "EX.RATE", "CUR", "PER UNIT \n (THB/JPY/USD)",
                    "TOTAL (JPY/USD)", "TOTAL (THB)", "AVERAGE FREIGHT (JPY/USD)", "AVERAGE INSURANCE (JPY/USD)", "TOTAL (JPY/USD)",
                    "TOTAL (THB)", "PER UNIT (THB)", "CC", "TOTAL OF CIP (THB)", "Budget code", "PR.DIE/JIG", "Model", "PART No./DIE No.",
                    "Operating Date (Plan)", "Operating Date (Act)", "Result", "Reason diff (NG) Budget&Actual", "Fixed Asset Code",
                    "CLASS FIXED ASSET", "Fix Asset Name (English only)", "Serial No.", "part\nnumber\nDie No", "Process Die", "Model",
                    "Cost Center of User", "Transfer to supplier", "????????????????????? Fix Asset  ??????????????????", "New BFMor Add BFM", "Reason for Delay", "Add CIP/BFM No.",
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

                    int row = 3;
                    Console.WriteLine(cipUpdate.Count);
                    foreach (cipUpdateSchema cipUpdateitem in cipUpdate)
                    {
                        cipSchema item = db.CIP.Find(cipUpdateitem.cipSchemaid);
                        List<string[]> cellData = new List<string[]>()
                    {
                        new string [] {
                            item.workType, item.projectNo,
                            item.cipNo, item.subCipNo, item.poNo ,item.vendorCode, item.vendor, item.acqDate, item.invDate, item.receivedDate,
                             item.invNo, item.name, item.qty, item.exRate, item.cur, item.perUnit, item.totalJpy, item.totalThb, item.averageFreight,
                             item.averageInsurance, item.totalJpy_1, item.totalThb_1, item.perUnitThb, item.cc, item.totalOfCip, item.budgetCode, item.prDieJig,
                             item.model, item.partNoDieNo,
                             item.cipUpdate.planDate, item.cipUpdate.actDate, item.cipUpdate.result, item.cipUpdate.reasonDiff, item.cipUpdate.fixedAssetCode,
                             item.cipUpdate.classFixedAsset, item.cipUpdate.fixAssetName, item.cipUpdate.serialNo,
                             item.cipUpdate.partNumberDieNo, item.cipUpdate.processDie, item.cipUpdate.model, item.cipUpdate.costCenterOfUser,
                             item.cipUpdate.tranferToSupplier, item.cipUpdate.upFixAsset, item.cipUpdate.newBFMorAddBFM, item.cipUpdate.reasonForDelay, (item.cipUpdate.addCipBfmNo == "-") ? "" : item.cipUpdate.addCipBfmNo,
                             item.cipUpdate.remark, (item.cipUpdate?.boiType == "-") ? "" : item.cipUpdate?.boiType
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

                        worksheet.Cells["AF" + row].Formula = "=+IF(AH" + row + "=\"\",\"\",IF(Z" + row + "=\"\",\"\",IF(AND(MID(Z" + row + ",5,2)=\"09\",LEFT(AH" + row + ",2)=\"06\"),\"OK\",IF(AND(OR(MID(Z" + row + ",5,2)=\"31\",MID(Z" + row + ",5,2)=\"34\"),LEFT(AH" + row + ",2)=\"28\"),\"OK\",IF(MID(Z" + row + ",5,2)=LEFT(AH" + row + ",2),\"OK\",\"NG\")))))";
                        worksheet.Cells["AF" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));
                        worksheet.Cells["AH" + row].Formula = "=IF(LEFT(AI" + row + ",2)=\"28\",\"SOFTWARE\",IF(LEFT(AI" + row + ",2)=\"02\",\"BUILDING\",IF(LEFT(AI" + row + ",2)=\"03\",\"STRUCTURE\",IF(LEFT(AI" + row + ",2)=\"04\",\"MACHINE\",IF(LEFT(AI" + row + ",2)=\"05\",\"VEHICLE\",IF(LEFT(AI" + row + ",2)=\"06\",\"TOOLS\",IF(LEFT(AI" + row + ",2)=\"07\",\"FURNITURE\",IF(LEFT(AI" + row + ",2)=\"08\",\"DIES\",\"\"))))))))";
                        worksheet.Cells["AH" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#C8C5C5"));

                        // set pink row space
                        worksheet.Cells["AD" + row + ":AE" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                        worksheet.Cells["AG" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                        worksheet.Cells["AI" + row + ":AV" + row].Style.Fill.SetBackground(ColorTranslator.FromHtml("#F0CDE5"));
                        // set pink row space
                        worksheet.Cells[row, 1].LoadFromArrays(cellData);
                        worksheet.Cells["A" + row + ":AV" + row].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells["A" + row + ":AV" + row].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells["A" + row + ":AV" + row].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        worksheet.Cells["A" + row + ":AV" + row].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                        row = row + 1;
                    }
                    string fileName = System.Guid.NewGuid().ToString() + "-" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                    excel.SaveAs(new FileInfo(serverPath + fileName));

                    stream.Position = 0;
                    // return Ok(new { success = true, dept = deptCode, cipUpdate, });
                    return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", serverPath + fileName);
                }

                // return Ok();
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }

        [HttpPost("costCenter/prepare"), Consumes("multipart/form-data")]
        public ActionResult uploadPrepare([FromForm] CIPupload body)
        {
            try
            {
                string username = User.FindFirst("username")?.Value;

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

                List<cipSchema> cipUpdates = new List<cipSchema>();

                using (ExcelPackage excel = new ExcelPackage(Existfile))
                {
                    ExcelWorkbook workbook = excel.Workbook;
                    ExcelWorksheet sheet = workbook.Worksheets[0];

                    int colCount = sheet.Dimension.End.Column;
                    int rowCount = sheet.Dimension.End.Row;


                    for (int row = 3; row <= rowCount; row += 1)
                    {
                        cipSchema item = new cipSchema
                        {
                            cipUpdate = new cipUpdateSchema()
                        };

                        for (int col = 1; col <= colCount; col += 1)
                        {
                            string value = sheet.Cells[row, col].Value?.ToString();
                            if (value == null)
                            {
                                value = "-";
                            }
                            switch (col)
                            {
                                case 3: item.cipNo = value; break;
                                case 4: item.subCipNo = value; break;
                                case 30:
                                    if (value != "-" && value.IndexOf(" ") != -1)
                                    {
                                        item.cipUpdate.planDate = value.Substring(0, value.IndexOf(" "));
                                    }
                                    else
                                    {
                                        item.cipUpdate.planDate = value;
                                    }
                                    break;
                                case 31:
                                    if (value != "-" && value.IndexOf(" ") != -1)
                                    {
                                        item.cipUpdate.actDate = value.Substring(0, value.IndexOf(" "));
                                    }
                                    else
                                    {
                                        item.cipUpdate.actDate = value;
                                    }
                                    break;
                                case 32: item.cipUpdate.result = value; break;
                                case 33: item.cipUpdate.reasonDiff = value; break;
                                case 34: item.cipUpdate.fixedAssetCode = value; break;
                                case 35: item.cipUpdate.classFixedAsset = value; break;
                                case 36:
                                    if (value.IndexOf(',') != -1 || value.IndexOf(':') != -1
                                        || value.IndexOf('"') != -1 || value.IndexOf('#') != -1
                                        || value.IndexOf('!') != -1 || value.IndexOf('*') != -1 || value.Length > 100)
                                    {
                                        return BadRequest(new { success = false, message = "Not allow Symbol value." });
                                    }
                                    item.cipUpdate.fixAssetName = value;
                                    break;
                                case 37: item.cipUpdate.serialNo = value; break;
                                case 38: item.cipUpdate.partNumberDieNo = value; break;
                                case 39: item.cipUpdate.processDie = value; break;
                                case 40: item.cipUpdate.model = value; break;
                                case 41: item.cipUpdate.costCenterOfUser = value; break;
                                case 42: item.cipUpdate.tranferToSupplier = value; break;
                                case 43: item.cipUpdate.upFixAsset = value; break;
                                case 44: item.cipUpdate.newBFMorAddBFM = value; break;
                                case 45: item.cipUpdate.reasonForDelay = value; break;
                                case 46: item.cipUpdate.addCipBfmNo = value; break;
                                case 47: item.cipUpdate.remark = value; break;
                                case 48: item.cipUpdate.boiType = value; break;
                            }
                        }
                        if (item.cipNo != "-")
                        {
                            cipUpdates.Add(item);
                        }
                    }

                }

                List<ApprovalSchema> approveItem = new List<ApprovalSchema>();
                List<cipUpdateSchema> cipUpdatedata = new List<cipUpdateSchema>();
                List<cipSchema> cipData = new List<cipSchema>();

                foreach (cipSchema item in cipUpdates)
                {
                    cipSchema cipTable = db.CIP.Where<cipSchema>(cip => cip.cipNo == item.cipNo && cip.subCipNo == item.subCipNo).FirstOrDefault();
                    cipUpdateSchema cipUpdateTable = db.CIP_UPDATE.Where<cipUpdateSchema>(cipUpdate => cipUpdate.cipSchemaid == cipTable.id).FirstOrDefault();

                    cipTable.status = "cost-prepared";
                    cipData.Add(cipTable);

                    cipUpdateTable.planDate = item.cipUpdate.planDate;
                    cipUpdateTable.actDate = item.cipUpdate.actDate;
                    cipUpdateTable.result = item.cipUpdate.result;
                    cipUpdateTable.reasonDiff = item.cipUpdate.reasonDiff;
                    cipUpdateTable.fixedAssetCode = item.cipUpdate.fixedAssetCode;
                    cipUpdateTable.classFixedAsset = item.cipUpdate.classFixedAsset;
                    cipUpdateTable.fixAssetName = item.cipUpdate.fixAssetName;
                    cipUpdateTable.serialNo = item.cipUpdate.serialNo;
                    cipUpdateTable.partNumberDieNo = item.cipUpdate.partNumberDieNo;
                    cipUpdateTable.processDie = item.cipUpdate.processDie;
                    cipUpdateTable.model = item.cipUpdate.model;
                    cipUpdateTable.costCenterOfUser = item.cipUpdate.costCenterOfUser;
                    cipUpdateTable.tranferToSupplier = item.cipUpdate.tranferToSupplier;
                    cipUpdateTable.upFixAsset = item.cipUpdate.upFixAsset;
                    cipUpdateTable.newBFMorAddBFM = item.cipUpdate.newBFMorAddBFM;
                    cipUpdateTable.reasonForDelay = item.cipUpdate.reasonForDelay;
                    cipUpdateTable.addCipBfmNo = item.cipUpdate.addCipBfmNo;
                    cipUpdateTable.remark = item.cipUpdate.remark;
                    cipUpdateTable.boiType = item.cipUpdate.boiType;

                    cipUpdatedata.Add(cipUpdateTable);
                    approveItem.Add(new ApprovalSchema
                    {
                        cipSchemaid = cipTable.id,
                        date = DateTime.Now.ToString("yyyy/MM/dd"),
                        empNo = username,
                        onApproveStep = "cost-prepared"
                    });
                }
                db.CIP.UpdateRange(cipData);
                db.CIP_UPDATE.UpdateRange(cipUpdatedata);
                db.APPROVAL.AddRange(approveItem);
                db.SaveChanges();

                return Ok(new { success = true, });
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
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

                foreach (int id in body.id)
                {
                    cipSchema item = db.CIP.Find(id);
                    db.CIP_UPDATE.Where<cipUpdateSchema>(cipUpdate => cipUpdate.status == "active").ToList();
                    data.Add(item);
                }


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
                    "Cost Center of User", "Transfer to supplier", "????????????????????? Fix Asset  ??????????????????", "New BFMor Add BFM", "Reason for Delay", "Add CIP/BFM No.",
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
                    worksheet.Cells["A2:BE2"].Style.WrapText = true;
                    worksheet.Cells["A2:BE2"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells["A2:BE2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells["A2:BE2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    worksheet.Cells["A2:BE2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

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
                             item.cipUpdate.planDate, item.cipUpdate.actDate,item.cipUpdate.result,
                             item.cipUpdate.reasonDiff, item.cipUpdate.fixedAssetCode, item.cipUpdate.classFixedAsset, item.cipUpdate.fixAssetName, item.cipUpdate.serialNo,
                             item.cipUpdate.partNumberDieNo, item.cipUpdate.processDie, item.cipUpdate.model, item.cipUpdate.costCenterOfUser,
                             item.cipUpdate.tranferToSupplier, item.cipUpdate.upFixAsset, item.cipUpdate.newBFMorAddBFM, item.cipUpdate.reasonForDelay,
                             (item.cipUpdate.addCipBfmNo == "-") ? "" : item.cipUpdate.addCipBfmNo, item.cipUpdate.remark, (item.cipUpdate?.boiType == "-") ? "" : item.cipUpdate?.boiType
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

                        IExcelDataValidationList boiType = worksheet.DataValidations.AddListValidation("AV" + row);
                        boiType.Formula.Values.Add("-");
                        boiType.Formula.Values.Add("BOI");
                        boiType.Formula.Values.Add("NON BOI");

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
                            DateTime dateFormat = DateTime.ParseExact(save.date.ToString(), "yyyy/MM/dd", CultureInfo.InstalledUICulture);
                            worksheet.Cells["AW" + row].AddComment(dateFormat.ToString("dd MMMM yyyy"), "system");
                        }

                        ApprovalSchema ccCheck = approve.Find(action => action.onApproveStep == "cc-checked");

                        if (ccCheck != null)
                        {
                            userSchema userCheck = db.USERS.Find(ccCheck.empNo);
                            worksheet.Cells["AX" + row].Value = userCheck.name;
                            DateTime dateFormat = DateTime.ParseExact(ccCheck.date.ToString(), "yyyy/MM/dd", CultureInfo.InstalledUICulture);
                            worksheet.Cells["AX" + row].AddComment(dateFormat.ToString("dd MMMM yyyy"), "system");
                        }

                        ApprovalSchema ccApprove = approve.Find(action => action.onApproveStep == "cc-approved");

                        if (ccApprove != null)
                        {
                            userSchema userApprove = db.USERS.Find(ccApprove.empNo);
                            worksheet.Cells["AY" + row].Value = userApprove.name;
                            DateTime dateFormat = DateTime.ParseExact(ccApprove.date.ToString(), "yyyy/MM/dd", CultureInfo.InstalledUICulture);
                            worksheet.Cells["AY" + row].AddComment(dateFormat.ToString("dd MMMM yyyy"), "system");
                        }

                        ApprovalSchema costPrepare = approve.Find(action => action.onApproveStep == "cost-prepared");
                        if (costPrepare != null)
                        {
                            userSchema userCostPrepare = db.USERS.Find(costPrepare.empNo);
                            worksheet.Cells["AZ" + row].Value = userCostPrepare.name;
                            DateTime dateFormat = DateTime.ParseExact(costPrepare.date.ToString(), "yyyy/MM/dd", CultureInfo.InstalledUICulture);
                            worksheet.Cells["AZ" + row].AddComment(dateFormat.ToString("dd MMMM yyyy"), "system");
                        }
                        ApprovalSchema costCheck = approve.Find(action => action.onApproveStep == "cost-checked");
                        if (costCheck != null)
                        {
                            userSchema costUserCheck = db.USERS.Find(costCheck.empNo);
                            worksheet.Cells["BA" + row].Value = costUserCheck.name;
                            DateTime dateFormat = DateTime.ParseExact(costCheck.date.ToString(), "yyyy/MM/dd", CultureInfo.InstalledUICulture);
                            worksheet.Cells["BA" + row].AddComment(dateFormat.ToString("dd MMMM yyyy"), "system");
                        }
                        ApprovalSchema costApprove = approve.Find(action => action.onApproveStep == "cost-approved");
                        if (costApprove != null)
                        {
                            userSchema costUserApprove = db.USERS.Find(costApprove.empNo);
                            worksheet.Cells["BB" + row].Value = costUserApprove.name;
                            DateTime dateFormat = DateTime.ParseExact(costApprove.date.ToString(), "yyyy/MM/dd", CultureInfo.InstalledUICulture);
                            worksheet.Cells["BB" + row].AddComment(dateFormat.ToString("dd MMMM yyyy"), "system");
                        }
                        ApprovalSchema itcConfirm = approve.Find(action => action.onApproveStep == "itc-confirmed");

                        if (itcConfirm != null)
                        {
                            userSchema itcConfirmUser = db.USERS.Find(itcConfirm.empNo);
                            worksheet.Cells["BC" + row].Value = itcConfirmUser.name;
                            DateTime dateFormat = DateTime.ParseExact(itcConfirm.date.ToString(), "yyyy/MM/dd", CultureInfo.InstalledUICulture);
                            worksheet.Cells["BC" + row].AddComment(dateFormat.ToString("dd MMMM yyyy"), "system");
                        }
                        ApprovalSchema diffChecked = approve.Find(action => action.onApproveStep == "diff-checked");
                        if (diffChecked != null)
                        {
                            userSchema accDiffChecker = db.USERS.Find(diffChecked.empNo);
                            worksheet.Cells["BD" + row].Value = accDiffChecker.name;
                            DateTime dateFormat = DateTime.ParseExact(diffChecked.date.ToString(), "yyyy/MM/dd", CultureInfo.InstalledUICulture);
                            worksheet.Cells["BD" + row].AddComment(dateFormat.ToString("dd MMMM yyyy"), "system");
                        }

                        ApprovalSchema diffApproved = approve.Find(action => action.onApproveStep == "diff-approved");
                        if (diffApproved != null)
                        {
                            userSchema diffApproveUser = db.USERS.Find(diffApproved.empNo);
                            worksheet.Cells["BE" + row].Value = diffApproveUser.name;
                            DateTime dateFormat = DateTime.ParseExact(diffApproved.date.ToString(), "yyyy/MM/dd", CultureInfo.InstalledUICulture);
                            worksheet.Cells["BE" + row].AddComment(dateFormat.ToString("dd MMMM yyyy"), "system");
                        }


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
                Console.WriteLine(e.Message);
                return Problem(e.StackTrace);
            }
        }
    }
}