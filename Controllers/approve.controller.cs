
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

using cip_api.models;
using cip_api.request;
using System.Collections.Generic;
using System.Linq;
using System;

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
            return db.PERMISSIONS.Where<PermissionSchema>(item => item.empNo == empNo).ToList<PermissionSchema>();
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
        public ActionResult cc()
        {

            string username = User.FindFirst("username")?.Value;
            List<PermissionSchema> permissions = GetPermissions(username);

            PermissionSchema checker = permissions.Find(e => e.action == "checker");
            PermissionSchema approver = permissions.Find(e => e.action == "approver");

            string message = "";
            List<cipSchema> data = new List<cipSchema>();
            if (checker != null)
            {
                List<string> multidept = checker.deptCode.Split(',').ToList();

                message = "CIP for data check.";
                if (multidept.Count > 1)
                {
                    foreach (string code in multidept)
                    {
                        data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "save" && item.cc == code).ToList<cipSchema>());
                    }
                }
                else
                {
                    data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "save" && item.cc == checker.deptCode).ToList<cipSchema>());
                }

            }
            if (approver != null)
            {
                List<string> multidept = checker.deptCode.Split(',').ToList();

                message = "CIP for data check.";
                if (multidept.Count > 1)
                {
                    foreach (string code in multidept)
                    {
                        data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "cc-checked" && item.cc == code).ToList<cipSchema>());
                    }
                }
                else
                {
                    data.AddRange(db.CIP.Where<cipSchema>(item => item.status == "cc-checked" && item.cc == checker.deptCode).ToList<cipSchema>());
                }
            }

            return Ok(
              new
              {
                  success = true,
                  message,
                  data,
              }
          );
        }

        [HttpGet("costCenter")]
        public ActionResult costCenter()
        {
            string deptCode = User.FindFirst("deptCode")?.Value;

            string username = User.FindFirst("username")?.Value;
            List<PermissionSchema> permissions = GetPermissions(username);

            PermissionSchema checker = permissions.Find(e => e.action == "checker");
            PermissionSchema approver = permissions.Find(e => e.action == "approver");

            List<cipSchema> data = new List<cipSchema>();
            if (checker != null)
            {
                data = db.CIP.Where<cipSchema>(item => item.status == "cc-approved" && item.cipUpdate.costCenterOfUser != item.cc).ToList();
            }
            if (approver != null)
            {
                data = db.CIP.Where<cipSchema>(item => item.status == "cost-checked" && item.cipUpdate.costCenterOfUser != item.cc).ToList();
            }

            return Ok(new { success = true, message = "CIP on Cost center.", data, });
        }

        [HttpPut("approve/cc")]
        public ActionResult approve(Approve body)
        {
            List<cipSchema> updateRange = new List<cipSchema>();
            List<ApprovalSchema> approvals = new List<ApprovalSchema>();
            string username = User.FindFirst("username")?.Value;

            List<PermissionSchema> permissions = GetPermissions(username);

            PermissionSchema checker = permissions.Find(e => e.action == "checker");
            PermissionSchema approver = permissions.Find(e => e.action == "approver");

            string status = "";
            foreach (string item in body.id)
            {
                Int32 id = Int32.Parse(item);
                ApprovalSchema approve = new ApprovalSchema();
                cipSchema data = db.CIP.Find(id);

                if (data.cc == checker.deptCode)
                {
                    data.status = "cc-checked";
                    status = "cc-checked";
                }
                else if (data.cc == approver.deptCode)
                {
                    data.status = "cc-approved";
                    status = "cc-approved";
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

        [HttpPut("approve/costCenter")]
        public ActionResult costCenter(Approve body)
        {
            List<cipSchema> updateRange = new List<cipSchema>();
            List<ApprovalSchema> approvals = new List<ApprovalSchema>();
            string username = User.FindFirst("username")?.Value;

            List<PermissionSchema> permissions = GetPermissions(username);

            List<PermissionSchema> checker = permissions.FindAll(e => e.action == "checker");
            List<PermissionSchema> approver = permissions.FindAll(e => e.action == "approver");

            string status = "";
            // return Ok(new { approver, checker});
            Console.WriteLine(approver.Count);
            foreach (string item in body.id)
            {
                Int32 id = Int32.Parse(item);
                ApprovalSchema approve = new ApprovalSchema();
                cipSchema data = db.CIP.Find(id);
                db.CIP_UPDATE.Where<cipUpdateSchema>(row => row.cipSchemaid == id).FirstOrDefault();

                if (checker.Count != 0)
                {
                    PermissionSchema check = checker.Find(e => e.action == "checker" && e.deptCode == data.cipUpdate.costCenterOfUser);
                    if (data.cipUpdate.costCenterOfUser == check.deptCode)
                    {
                        status = "cost-checked";
                    }
                }
                else if (approver.Count != 0)
                {
                    PermissionSchema approve_act = approver.Find(e => e.action == "approver" && e.deptCode == data.cipUpdate.costCenterOfUser);
                    if (data.cipUpdate.costCenterOfUser == approve_act.deptCode)
                    {
                        status = "cost-approved";
                    }
                }

                data.status = status;

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
    }
}