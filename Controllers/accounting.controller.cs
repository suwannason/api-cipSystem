
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
            List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status == "cc-approved" || item.status == "itc-confirmed" || item.status == "cost-approved" || item.status == "acc-approved").ToList();
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
            List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status != "finish").ToList();

            return Ok(new { success = true, message = "CIP tracking", data, });
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
        public ActionResult codeDiff()
        {

            List<cipSchema> cipSuccess = db.CIP.Where<cipSchema>(item => item.status == "cost-approved" || item.status == "cc-approved" || item.status == "itc-confirmed").ToList();
            db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "active").ToList();
            List<cipSchema> returnData = new List<cipSchema>();

            string username = User.FindFirst("username")?.Value;

            List<PermissionSchema> permissions = GetPermissions(username);

            PermissionSchema checker = permissions.Find(e => e.action == "checker");
            PermissionSchema approver = permissions.Find(e => e.action == "approver");

            if (checker != null)
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
                        if ((item.cc == item.cipUpdate.costCenterOfUser) && item.cipUpdate.result.ToLower() == "ng")
                        {
                            returnData.Add(item);
                        }
                    }
                    else if (item.status == "cost-approved")
                    {
                        if (item.cipUpdate.tranferToSupplier != "-" && item.cipUpdate.result.ToLower() == "ng")
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
            return Ok(new { success = true, message = "Account diff data check.", data = returnData });
        }

        [HttpGet("finish")]
        public ActionResult accFinishData()
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
                    }
                    else if (item.status == "cost-approved")
                    {
                        if (item.cipUpdate.tranferToSupplier == "-" && item.cc != "5110" && item.cipUpdate.result.ToLower() != "ng")
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

                return Ok(new { success = true, message = "Accouting data check.", data = returnData });
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

                    cipUpdate.status = "finish";
                    cipUpdate.status = "finish";

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
    }
}