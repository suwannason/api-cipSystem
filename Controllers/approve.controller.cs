
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
        private void createNotification(string currentAction, string username, string deptCode)
        {

            List<PermissionSchema> users = db.PERMISSIONS.Where<PermissionSchema>(item => item.deptCode == deptCode).ToList();

            List<PermissionSchema> createTo = new List<PermissionSchema>();

            string message = ""; string title = "";

            if (currentAction == "save")
            { // cc prepared CIP success
                createTo = users.FindAll(e => e.action == "checker");
                message = "CIP prepared on " + DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                title = "Checking request";
            }
            else if (currentAction == "cc-checked")
            {
                createTo = users.FindAll(e => e.action == "approver");
                message = "CIP checked on " + DateTime.Now.ToString("yyyy/MM/dd HH:mm");
                title = "Approving request";
            }

            List<NotificationSchema> notifications = new List<NotificationSchema>();

            Int32 cipCount = db.CIP.Count<cipSchema>(item => deptCode.IndexOf(item.cc) != -1 && item.status == currentAction);
            foreach (PermissionSchema item in createTo)
            {
                notifications.Add(new NotificationSchema
                {
                    createDate = DateTime.Now.ToString("yyyy/MM/dd"),
                    message = message,
                    title = title,
                    status = "created",
                    userSchemaempNo = item.empNo,
                });
            }
            db.NOTIFICATIONS.AddRange(notifications);
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
            List<PermissionSchema> approver = permissions.FindAll(e => e.action == "approver");

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
            if (approver.Count != 0)
            {
                message = "CIP for data approve.";

                foreach (PermissionSchema permission in approver)
                {
                    data.AddRange((db.CIP.Where<cipSchema>(item => item.status == "cc-checked" && item.cc == permission.deptCode).ToList<cipSchema>()));
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
            PermissionSchema prepare = permissions.Find(e => e.action == "prepare");

            List<cipSchema> data = new List<cipSchema>();
            string message = "";
            if (checker != null)
            {
                message = "CIP on Cost center check";
                data = db.CIP.Where<cipSchema>(item => item.status == "cost-prepared" && item.cipUpdate.costCenterOfUser != item.cc).ToList();
            }
            if (approver != null)
            {
                message = "CIP on Cost center approve";
                data = db.CIP.Where<cipSchema>(item => item.status == "cost-checked" && item.cipUpdate.costCenterOfUser != item.cc).ToList();
            }
            if (prepare != null)
            {
                List<cipSchema> onApproved = db.CIP.Where<cipSchema>(item => item.status == "cc-approved").ToList();

                message = "CIP on Cost center prepare";
                foreach (cipSchema item in onApproved)
                {
                    cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(cip => cip.costCenterOfUser.IndexOf(prepare.deptCode) != -1).FirstOrDefault();
                    if (cipUpdate != null)
                    {
                        data.Add(item);
                    }
                }
            }

            return Ok(new { success = true, message, data, });
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

            PermissionSchema checker = permissions.Find(e => e.action == "checker");
            List<PermissionSchema> approver = permissions.FindAll(e => e.action == "approver");

            string status = "";
            foreach (string item in body.id)
            {
                Int32 id = Int32.Parse(item);
                ApprovalSchema approve = new ApprovalSchema();
                cipSchema data = db.CIP.Find(id);

                if (checker != null)
                {
                    if (data.cc == checker.deptCode)
                    {
                        data.status = "cc-checked";
                        status = "cc-checked";
                    }
                }
                else if (approver.Count != 0)
                {
                    PermissionSchema approving = approver.Find(e => e.deptCode == data.cc);
                    if (approving != null)
                    {
                        if (data.cc == approving.deptCode)
                        {
                            data.status = "cc-approved";
                            status = "cc-approved";
                        }
                    }
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

            // SEND NOTIFICATIONS
            // List<PermissionSchema> users = new List<PermissionSchema>();
            // List<PermissionSchema> createTo = new List<PermissionSchema>();
            // string message = ""; string title = ""; string type = "";

            // List<NotificationSchema> notifications = new List<NotificationSchema>();

            // if (status == "cc-checked")
            // {
            //     users = db.PERMISSIONS.Where<PermissionSchema>(item => item.deptCode == deptCode).ToList();
            //     message = updateRange.Count + " CIP prepared on " + DateTime.Now.ToString("yyyy/MM/dd HH:mm");
            //     title = "Checking request";
            //     type = "check";
            //     createTo = users.FindAll(e => e.action == "approver");
            // }
            // else if (status == "cc-approved")
            // {
            //     int diff = 0;

            //     foreach (cipSchema item in updateRange)
            //     {
            //         cipUpdateSchema cipUpdateCheck = db.CIP_UPDATE.Where<cipUpdateSchema>(cip => cip.cipSchemaid == item.id).FirstOrDefault();
            //         if (cipUpdateCheck.costCenterOfUser != item.cc)
            //         {
            //             diff += 1;
            //             // userPrepare = db.PERMISSIONS.Where<PermissionSchema>(item => item.deptCode == cipUpdateCheck.costCenterOfUser && item.action == "prepare").ToList();
            //         }
            //     }
            //     type = "approve";
            // }


            // foreach (PermissionSchema item in createTo)
            // {
            //     notifications.Add(new NotificationSchema
            //     {
            //         createDate = DateTime.Now.ToString("yyyy/MM/dd"),
            //         message = message,
            //         title = title,
            //         status = "created",
            //         type = type,
            //         userSchemaempNo = item.empNo,
            //     });
            // }
            // db.NOTIFICATIONS.AddRange(notifications);
            // SEND NOTIFICATIONS

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
            List<PermissionSchema> prepare = permissions.FindAll(e => e.action == "prepare");

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
                else if (prepare.Count != 0)
                {
                    PermissionSchema prepare_act = prepare.Find(e => e.action == "prepare" && e.deptCode == data.cipUpdate.costCenterOfUser);
                    if (data.cipUpdate.costCenterOfUser == prepare_act.deptCode)
                    {
                        status = "cost-prepared";
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

        [HttpGet("test")]
        public ActionResult test()
        {
            string deptCode = User.FindFirst("deptCode")?.Value;
            string username = User.FindFirst("username")?.Value;
            createNotification("save", username, deptCode);
            // Console.WriteLine(deptCode);
            return Ok();
        }
    }
}