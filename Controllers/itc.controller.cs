

using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using cip_api.request;
using cip_api.models;
using System.Collections.Generic;
using System.Linq;
using System;

namespace cip_api.controllers
{

    [ApiController, Route("[controller]"), Authorize]
    public class itcController : ControllerBase
    {
        private readonly Database db;
        private IConfiguration _config;
        private readonly string ldap_auth;

        public itcController(Database _db, IConfiguration config, IEndpoint setting)
        {
            db = _db;
            _config = config;
            ldap_auth = setting.ldap_auth;
        }

        private List<PermissionSchema> GetPermissions(string empNo)
        {
            return db.PERMISSIONS.Where<PermissionSchema>(item => item.empNo == empNo).ToList<PermissionSchema>();
        }

        [HttpGet("waiting")]
        public ActionResult waiting()
        {
            List<cipSchema> data = db.CIP.Where<cipSchema>(item => (item.status == "cc-approved" || item.status == "cost-approved") && item.cc == "5110").ToList<cipSchema>();
            // db.CIP_UPDATE.ToList();

            List<cipSchema> returnData = new List<cipSchema>();

            foreach (cipSchema item in data)
            {
                cipUpdateSchema cip_update = db.CIP_UPDATE.Where<cipUpdateSchema>(update => update.cipSchemaid == item.id).FirstOrDefault<cipUpdateSchema>();

                if (item.status == "cc-approved")
                {
                    if ((item.cc == cip_update.costCenterOfUser) && (cip_update.tranferToSupplier != "-" && cip_update.tranferToSupplier != null))
                    {
                        returnData.Add(item);
                    }
                }
                else if (item.status == "cost-approved")
                {
                    if (cip_update.tranferToSupplier != "-" && cip_update.tranferToSupplier != null)
                    {
                        returnData.Add(item);
                    }
                }
            }
            return Ok(new
            {
                success = true,
                message = "Data for ITC confirm.",
                data = returnData,
            });
        }

        [HttpGet("confirmed")]
        public ActionResult confirmed()
        {
            List<ApprovalSchema> data = db.APPROVAL.Where<ApprovalSchema>(item => item.onApproveStep == "itc-confirmed").ToList();
            List<cipSchema> returData = new List<cipSchema>();

            foreach (ApprovalSchema item in data) {
                returData.Add(db.CIP.Find(item.cipSchemaid));
            }
            return Ok(new { success = true, messsage = "ITC confirmed", data=returData, });
        }
        [HttpPut("confirm")]
        public ActionResult confirmData(Approve body)
        {
            string username = User.FindFirst("username")?.Value;

            List<PermissionSchema> permissions = GetPermissions(username);

            List<PermissionSchema> isBoi = permissions.FindAll(e => e.deptShortName == "ITC BOI");
            if (isBoi.Count == 0)
            {
                return BadRequest(new { success = false, message = "Permission denied." });
            }
            List<ApprovalSchema> approval = new List<ApprovalSchema>();
            List<cipSchema> updateCip = new List<cipSchema>();

            foreach (string item in body.id)
            {
                Int32 id = Int32.Parse(item);

                cipSchema cip = db.CIP.Find(id);
                cip.status = "itc-confirmed";

                ApprovalSchema approve = new ApprovalSchema
                {
                    cipSchemaid = id,
                    date = DateTime.Now.ToString("yyyy/MM/dd"),
                    empNo = username,
                    onApproveStep = "itc-confirmed",
                };
                approval.Add(approve);
                updateCip.Add(cip);
            }
            db.APPROVAL.AddRange(approval);
            db.CIP.UpdateRange(updateCip);
            db.SaveChanges();
            return Ok(new { success = true, message = "Confirm data success. " });
        }
    }
}