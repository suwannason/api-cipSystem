

using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using cip_api.request;
using cip_api.models;
using System.Collections.Generic;
using System.Linq;
using System;
using OfficeOpenXml;
using System.IO;

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
        public ActionResult waiting(string user, string dcode)
        {
            List<cipSchema> data = db.CIP.Where<cipSchema>(item => (item.status == "cc-approved" || item.status == "cost-approved") && item.cipUpdate.costCenterOfUser == "5110").ToList<cipSchema>();
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
            if (user != null)
            {
                return Ok(returnData.Count);
            }
            return Ok(new
            {
                success = true,
                message = "Data for ITC confirm.",
                data = returnData,
            });
        }

        [HttpGet("confirmed")]
        public ActionResult confirmed(string user, string dcode)
        {
            List<ApprovalSchema> data = db.APPROVAL.Where<ApprovalSchema>(item => item.onApproveStep == "itc-confirmed").ToList();
            List<cipSchema> returData = new List<cipSchema>();

            foreach (ApprovalSchema item in data)
            {
                returData.Add(db.CIP.Find(item.cipSchemaid));
            }

            if (user != null)
            {
                return Ok(returData.Count);
            }
            return Ok(new { success = true, messsage = "ITC confirmed", data = returData, });
        }
        [HttpPut("confirm")]
        public ActionResult confirmData(ITCConfirm body)
        {
            try
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
                List<cipUpdateSchema> cipUpdate_update = new List<cipUpdateSchema>();

                foreach (request.confirmBox item in body.confirm.ToArray())
                {

                    cipSchema cip = db.CIP.Find(item.id);
                    cip.status = "itc-confirmed";

                    cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(cipUpdate => cipUpdate.cipSchemaid == item.id).FirstOrDefault();
                    cipUpdate.boiType = item.boiType;
                    cipUpdate_update.Add(cipUpdate);

                    ApprovalSchema approve = new ApprovalSchema
                    {
                        cipSchemaid = item.id,
                        date = DateTime.Now.ToString("yyyy/MM/dd"),
                        empNo = username,
                        onApproveStep = "itc-confirmed",
                    };
                    approval.Add(approve);
                    updateCip.Add(cip);
                }
                db.APPROVAL.AddRange(approval);
                db.CIP.UpdateRange(updateCip);
                db.CIP_UPDATE.UpdateRange(cipUpdate_update);
                db.SaveChanges();
                return Ok(new { success = true, message = "Confirm data success. " });
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return Problem(e.Message);
            }

        }

        [HttpPost("upload"), Consumes("multipart/form-data")]
        public ActionResult confirmWithUpload(FileUpload body)
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

                string fileName = System.Guid.NewGuid().ToString() + "-" + body.file.FileName;
                FileStream strem = System.IO.File.Create($"{serverPath}{fileName}");
                body.file.CopyTo(strem);
                strem.Close();

                string path = $"{serverPath}{fileName}";
                FileInfo Existfile = new FileInfo(path);

                using (ExcelPackage excel = new ExcelPackage(Existfile))
                {
                    ExcelWorkbook workbook = excel.Workbook;
                    ExcelWorksheet sheet = workbook.Worksheets[0];

                    int colCount = sheet.Dimension.End.Column;
                    int rowCount = sheet.Dimension.End.Row;

                    for (int row = 3; row <= rowCount; row += 1)
                    {
                        for (int col = 1; col < colCount; col += 1)
                        {
                            switch (col)
                            {
                                case 1: Console.WriteLine(""); break;
                            }
                        }
                    }
                }
                return Ok(new { success = true, message = "ITC confirm success." });
            }
            catch (System.Exception e)
            {
                return Problem(e.StackTrace);
            }
        }
    }
}