
using System.Collections.Generic;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using cip_api.models;
using System.Linq;
using cip_api.request;
using System;

namespace cip_api.controllers
{

    [ApiController, Route("[controller]"), Authorize]
    public class accController : ControllerBase
    {
        private readonly Database db;
        private IConfiguration _config;
        private readonly string ldap_auth;
        public accController(Database _db, IConfiguration config, IEndpoint setting)
        {
            db = _db;
            _config = config;
            ldap_auth = setting.ldap_auth;
        }

        private List<PermissionSchema> GetPermissions(string empNo)
        {
            return db.PERMISSIONS.Where<PermissionSchema>(item => item.empNo == empNo).ToList<PermissionSchema>();
        }

        [HttpGet("fa")]
        public ActionResult waitingFA()
        {
            string username = User.FindFirst("username")?.Value;

            List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status == "cc-approved" || item.status == "itc-confirmed" || item.status == "cost-approved").ToList();
            db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "active").ToList();

            List<cipSchema> returnData = new List<cipSchema>();
            foreach (cipSchema item in data)
            {
                if (item.status == "cc-approved")
                {
                    if ((item.cc == item.cipUpdate.costCenterOfUser) && item.cipUpdate.result.ToLower() != "ng" && item.cipUpdate.tranferToSupplier == "-")
                    {
                        returnData.Add(item);
                    }
                }
                else if (item.status == "cost-approved")
                {
                    if (item.cipUpdate.tranferToSupplier == "-" && item.cipUpdate.result.ToLower() != "ng")
                    {
                        returnData.Add(item);
                    }
                }
                else if (item.status == "itc-confirmed")
                {
                    if (item.cipUpdate.result.ToLower() != "ng")
                    {
                        returnData.Add(item);
                    }
                }
                else if (item.status == "acc-approved")
                {
                    returnData.Add(item);
                }
            }

            return Ok(new { success = true, message = "Data for acc select.", data = returnData });
        }

        [HttpPut("fa")]
        public ActionResult accAcceptData(Approve body)
        {
            try
            {
                string dept = User.FindFirst("dept").Value.ToLower();
                if (dept != "acc")
                {
                    return BadRequest(new { success = false, message = "Permission denied." });
                }

                string currentDate = DateTime.Now.ToString("yyyy/MM/dd");
                string username = User.FindFirst("username")?.Value;

                List<cipSchema> updateItems = new List<cipSchema>();
                List<ApprovalSchema> approve = new List<ApprovalSchema>();

                foreach (string item in body.id)
                {
                    Int32 id = Int32.Parse(item);
                    cipSchema data = db.CIP.Find(id);

                    data.status = "acc-accepted";
                    updateItems.Add(data);

                    approve.Add(new ApprovalSchema
                    {
                        cipSchemaid = id,
                        date = currentDate,
                        empNo = username,
                        onApproveStep = "acc-accepted",
                    });
                }
                db.APPROVAL.AddRange(approve);
                db.CIP.UpdateRange(updateItems);
                db.SaveChanges();

                return Ok(new { success = true, message = "Accounting accepted data success." });
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }
        [HttpGet("tracking")]
        public ActionResult tracking()
        {
            try
            {
                List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status != "finished" && item.status != "exported").ToList();
                db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "active").ToList();

                List<dynamic> returnData = new List<dynamic>();

                string message = "";
                foreach (cipSchema item in data)
                {
                    if (item.status == "save")
                    {
                        message = "On user check";

                    }
                    else if (item.status == "cc-checked")
                    {
                        message = "On user approve";
                    }

                    else if (item.status == "cost-checked")
                    {
                        message = "On Cost center approve";
                    }
                    else if (item.status == "itc-confirmed")
                    {
                        if (item.cipUpdate.result == "NG")
                        {
                            message = "On ACC confirm diff";
                        }
                        else
                        {
                            message = "On confirm FA";
                        }
                    }
                    else if (item.status == "cc-approved")
                    {
                        if (item.cc == item.cipUpdate.costCenterOfUser
                         || (item.cc == "2130" && (item.cipUpdate.costCenterOfUser == "2140" || item.cipUpdate.costCenterOfUser == "9555"))
                         || (item.cc == "2410" && (item.cipUpdate.costCenterOfUser == "2130" || item.cipUpdate.costCenterOfUser == "9555"))
                         || (item.cc == "9555" && (item.cipUpdate.costCenterOfUser == "2130" || item.cipUpdate.costCenterOfUser == "2140"))
                         || (item.cc == "5610" && item.cipUpdate.costCenterOfUser == "5615")
                         || (item.cc == "5615" && item.cipUpdate.costCenterOfUser == "5610")
                         || (item.cc == "5650" && item.cipUpdate.costCenterOfUser == "9333")
                         || (item.cc == "9333" && item.cipUpdate.costCenterOfUser == "5650")
                         || (item.cc == "5670" && item.cipUpdate.costCenterOfUser == "9444")
                         || (item.cc == "9444" && item.cipUpdate.costCenterOfUser == "5670")
                         || (item.cc.StartsWith("55") && item.cipUpdate.costCenterOfUser.StartsWith("55"))
                        )
                        { // in local tranfer
                            if (item.cipUpdate.tranferToSupplier != "-" && item.cipUpdate.costCenterOfUser == "5110")
                            {
                                message = "On ITC confirm";
                            }
                            else
                            {
                                if (item.cipUpdate.result == "NG")
                                {
                                    message = "On ACC confirm diff";
                                }
                                else
                                {
                                    message = "On confirm FA";
                                }
                            }
                        }
                        else
                        { // in cross tranfer
                            message = "On Requester prepare";
                        }
                    }
                    else if (item.status == "cost-approved")
                    {
                        if (item.cipUpdate.tranferToSupplier != "-" && item.cipUpdate.costCenterOfUser == "5110")
                        {
                            message = "On ITC confirm";
                        }
                        else if (item.cipUpdate.result == "NG")
                        {
                            message = "On ACC confirm diff";
                        }
                        else
                        {
                            message = "On confirm FA";
                        }
                    }
                    else if (item.status == "acc-checked")
                    {
                        message = "On ACC approve diff";
                    }
                    else if (item.status == "acc-approved")
                    {
                        message = "On confirm FA";

                    }
                    else if (item.status == "open")
                    {
                        message = "On Requester prepare";
                    }
                    else if (item.status == "draft")
                    {
                        message = "On Requester drafted";
                    }
                    returnData.Add(
                        new
                        {
                            message = message,
                            id = item.id,
                            workType = item.workType,
                            projectNo = item.projectNo,
                            cipNo = item.cipNo,
                            cc = item.cc,
                            subCipNo = item.subCipNo,
                            poNo = item.poNo,
                            vendorCode = item.vendorCode,
                            vendor = item.vendor,
                            acqDate = item.acqDate,
                            invDate = item.invDate,
                            receivedDate = item.receivedDate,
                            invNo = item.invNo,
                            name = item.name,
                            qty = item.qty,
                            exRate = item.exRate,
                            cur = item.cur,
                            perUnit = item.perUnit,
                            totalJpy = item.totalJpy,
                            totalThb = item.totalThb,
                            averageFreight = item.averageFreight,
                            averageInsurance = item.averageInsurance,
                            totalJpy_1 = item.totalJpy_1,
                            totalThb_1 = item.totalThb_1,
                            perUnitThb = item.perUnitThb,
                            totalOfCip = item.totalOfCip,
                            budgetCode = item.budgetCode,
                            prDieJig = item.prDieJig,
                            model = item.model,
                            partNoDieNo = item.partNoDieNo,
                            createDate = item.createDate,
                        }
                    );
                }

                return Ok(new { success = true, message = "CIP tracking", data = returnData, });
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return Problem(e.Message);
            }
        }
        [HttpPut("approve/diff")]
        public ActionResult approveDiff(Approve body)
        {
            string dept = User.FindFirst("dept").Value.ToLower();
            string username = User.FindFirst("username")?.Value;

            if (dept != "acc")
            {
                return BadRequest(new { success = false, message = "Permission denied." });
            }

            List<PermissionSchema> permissions = GetPermissions(username);

            PermissionSchema checker = permissions.Find(e => e.action == "checker");
            PermissionSchema approver = permissions.Find(e => e.action == "approver");

            string status = "";
            List<int> notUpdate = new List<int>();

            List<cipSchema> updateItem = new List<cipSchema>();
            foreach (string item in body.id)
            {
                Int32 id = Int32.Parse(item);
                cipSchema cip = db.CIP.Find(id);
                db.CIP_UPDATE.Where<cipUpdateSchema>(row => row.cipSchemaid == id).FirstOrDefault();

                if (cip.status == "acc-checked" && (approver != null || checker != null))
                {
                    status = "acc-approved";
                }
                else if (cip.status == "itc-confirmed" && (approver != null || checker != null))
                {
                    status = "acc-checked";
                }
                else if (cip.status == "cc-approved" && (approver != null || checker != null))
                {
                    if (cip.cc == cip.cipUpdate.costCenterOfUser && cip.cipUpdate.tranferToSupplier == "-")
                    {
                        status = "acc-checked";
                    }
                }
                else if (cip.status == "cost-approved" && (approver != null || checker != null))
                {
                    status = "acc-checked";
                }
                else
                {
                    notUpdate.Add(id);
                    continue;
                }
                cip.status = status;
                updateItem.Add(cip);
            }
            db.CIP.UpdateRange(updateItem);
            db.SaveChanges();

            if (notUpdate.Count != 0)
            {
                return Ok(new { success = true, message = "Have " + notUpdate.Count.ToString() + " CIP can't approve Please re-check approval." });
            }
            return Ok(new { success = true, message = "Approve CIP success." });
        }

        [HttpGet("diff")]
        public ActionResult codeDiff(string user, string dcode)
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

            List<cipSchema> cipSuccess = db.CIP.Where<cipSchema>(item => item.status == "cost-approved" || item.status == "cc-approved" || item.status == "itc-confirmed").ToList();
            db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "active").ToList();
            List<cipSchema> returnData = new List<cipSchema>();

            List<PermissionSchema> permissions = GetPermissions(username);

            PermissionSchema checker = permissions.Find(e => e.action == "checker");
            PermissionSchema approver = permissions.Find(e => e.action == "approver");
            PermissionSchema preapare = permissions.Find(e => e.action == "prepare");
            // if (preapare != null)
            // {
            //     return Conflict(new { success = true, message = "Permission denied." });
            // }

            if (checker != null || preapare != null)
            {
                foreach (cipSchema item in cipSuccess)
                {

                    if (item.status == "itc-confirmed")
                    {
                        if (item.cipUpdate.result.ToLower() == "ng")
                        {
                            returnData.Add(item);
                        }
                        // returnData.Add(item);
                    }
                    else if (item.status == "cc-approved")
                    {
                        if (item.cc == item.cipUpdate.costCenterOfUser
                         || (item.cc == "2130" && (item.cipUpdate.costCenterOfUser == "2140" || item.cipUpdate.costCenterOfUser == "9555"))
                         || (item.cc == "2410" && (item.cipUpdate.costCenterOfUser == "2130" || item.cipUpdate.costCenterOfUser == "9555"))
                         || (item.cc == "9555" && (item.cipUpdate.costCenterOfUser == "2130" || item.cipUpdate.costCenterOfUser == "2140"))
                         || (item.cc == "5610" && item.cipUpdate.costCenterOfUser == "5615")
                         || (item.cc == "5615" && item.cipUpdate.costCenterOfUser == "5610")
                         || (item.cc == "5650" && item.cipUpdate.costCenterOfUser == "9333")
                         || (item.cc == "9333" && item.cipUpdate.costCenterOfUser == "5650")
                         || (item.cc == "5670" && item.cipUpdate.costCenterOfUser == "9444")
                         || (item.cc == "9444" && item.cipUpdate.costCenterOfUser == "5670")
                         || (item.cc.StartsWith("55") && item.cipUpdate.costCenterOfUser.StartsWith("55") && item.cipUpdate.result.ToLower() == "ng"))
                        {
                            returnData.Add(item);
                        }
                    }
                    else if (item.status == "cost-approved")
                    {
                        if (item.cipUpdate.result.ToLower() == "ng")
                        {
                            returnData.Add(item);
                        }
                    }
                }
            }
            if (approver != null)
            {
                returnData.AddRange(db.CIP.Where<cipSchema>(item => item.status == "acc-checked").ToList());
            }
            if (user != null)
            {
                return Ok(returnData.Count);
            }
            return Ok(new { success = true, message = "Account diff data check.", data = returnData });
        }

        [HttpGet("finish")]
        public ActionResult accFinishData(string user, string code)
        {
            try
            {
                List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status == "cc-approved" || item.status == "cost-approved" || item.status == "itc-confirmed" || item.status == "acc-approved").ToList();
                db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "active").ToList();

                List<cipSchema> returnData = new List<cipSchema>();

                foreach (cipSchema item in data)
                {

                    if (item.status == "cc-approved")
                    {
                        if (item.cc == item.cipUpdate.costCenterOfUser && item.cipUpdate.tranferToSupplier == "-" && item.cipUpdate.result.ToLower() != "ng")
                        {
                            returnData.Add(item);
                        }
                        else if (item.cc != item.cipUpdate.costCenterOfUser && item.cipUpdate.tranferToSupplier == "-" && item.cipUpdate.result.ToLower() == "ok")
                        {
                            if (item.cc.StartsWith("55") && item.cipUpdate.costCenterOfUser.StartsWith("55"))
                            {
                                returnData.Add(item);
                            }
                            else if (item.cipUpdate.costCenterOfUser == "2130" && (item.cc == "2140" || item.cc == "9555")
                            || item.cipUpdate.costCenterOfUser == "2140" && (item.cc == "2130" || item.cc == "9555")
                            || item.cipUpdate.costCenterOfUser == "9555" && (item.cc == "2130" || item.cc == "2140")
                            || item.cipUpdate.costCenterOfUser == "5610" && (item.cc == "5619")
                            || item.cipUpdate.costCenterOfUser == "5619" && (item.cc == "5610")
                            || item.cipUpdate.costCenterOfUser == "5650" && (item.cc == "5655" || item.cc == "9333")
                            || item.cipUpdate.costCenterOfUser == "5655" && (item.cc == "5650" || item.cc == "9333")
                            || item.cipUpdate.costCenterOfUser == "9333" && (item.cc == "5650" || item.cc == "5655")
                            || item.cipUpdate.costCenterOfUser == "5670" && (item.cc == "5675" || item.cc == "9444")
                            || item.cipUpdate.costCenterOfUser == "5675" && (item.cc == "5670" || item.cc == "9444")
                            || item.cipUpdate.costCenterOfUser == "9444" && (item.cc == "5670" || item.cc == "5675")
                            )
                            {
                                returnData.Add(item);
                            }
                        }
                    }
                    else if (item.status == "cost-approved")
                    {
                        if (item.cipUpdate.tranferToSupplier == "-" && item.cc != "5110" && item.cipUpdate.result != "NG")
                        {
                            returnData.Add(item);
                        }
                    }
                    else if (item.status == "itc-confirmed")
                    {
                        if (item.cipUpdate.result.ToLower() != "ng")
                        {
                            returnData.Add(item);
                        }
                    }
                    else if (item.status == "acc-approved")
                    {
                        returnData.Add(item);
                    }
                }
                if (user != null)
                {
                    return Ok(returnData.Count);
                }
                return Ok(new { success = true, message = "Accouting data check.", data = returnData, count = returnData.Count });


            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }

        [HttpPut("approve/finish")]
        public ActionResult finishData(Approve body)
        {
            try
            {
                List<cipSchema> cipUpdateItem = new List<cipSchema>();
                List<cipUpdateSchema> refCip = new List<cipUpdateSchema>();

                foreach (string item in body.id)
                {
                    Int32 id = Int32.Parse(item);
                    cipSchema cip = db.CIP.Find(id);
                    cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(row => row.cipSchemaid == id).FirstOrDefault();

                    cipUpdate.status = "finished";
                    cip.status = "finished";

                    refCip.Add(cipUpdate);
                    cipUpdateItem.Add(cip);
                }
                db.CIP.UpdateRange(cipUpdateItem);
                db.CIP_UPDATE.UpdateRange(refCip);
                db.SaveChanges();

                return Ok(new { success = true, message = "Accounting finish data." });
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }

        [HttpPost("sendback")]
        public ActionResult sendBack(ACCsendBack body)
        {
            try
            {
                foreach (string id in body.id)
                {
                    cipSchema cip = db.CIP.Find(Int32.Parse(id));
                    cip.commend = body.commend;
                    if (body.toStep == "requester-prepare")
                    {
                        cip.status = "open";
                        List<ApprovalSchema> approving = db.APPROVAL.Where<ApprovalSchema>(item => item.cipSchemaid == Int32.Parse(id)).ToList();

                        db.APPROVAL.RemoveRange(approving);
                    }
                    else if (body.toStep == "user-prepare")
                    {
                        cip.status = "cc-approved";
                        List<ApprovalSchema> approving = db.APPROVAL.Where<ApprovalSchema>(item =>
                        item.cipSchemaid == Int32.Parse(id) && item.onApproveStep.IndexOf("cost") != -1).ToList();

                        db.APPROVAL.RemoveRange(approving);
                    }
                    db.CIP.Update(cip);

                }
                db.SaveChanges();
                return Ok(new { success = true, message = "Send back CIP success. " });
            }
            catch (Exception e)
            {
                return Conflict(e.StackTrace);
            }
        }
    }
}