

using cip_api.models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

using System;
using System.Collections.Generic;
using System.Linq;

namespace cip_api.controllers
{

    [Authorize]
    [ApiController]
    [Route("[controller]")]
    public class notificationController : ControllerBase
    {

        private readonly Database db;
        private IConfiguration _config;
        private readonly IEndpoint _setting;

        public notificationController(Database _db, IConfiguration config, IEndpoint setting)
        {
            db = _db;
            _config = config;
            _setting = setting;
        }
        private List<PermissionSchema> GetPermissions(string empNo)
        {
            return db.PERMISSIONS.Where<PermissionSchema>(item => item.empNo == empNo).ToList<PermissionSchema>();
        }

        [HttpGet]
        public ActionResult getNumber()
        {
            try
            {
                string username = User.FindFirst("username")?.Value;
                List<PermissionSchema> permissions = GetPermissions(username);

                PermissionSchema checker = permissions.Find(e => e.action == "checker");
                List<PermissionSchema> approver = permissions.FindAll(e => e.action == "approver");

                Int32 approving = 0;
                Int32 checking = 0;

                if (checker != null)
                {
                    List<string> multidept = checker.deptCode.Split(',').ToList();
                    if (multidept.Count > 1)
                    {
                        foreach (string code in multidept)
                        {
                            checking += db.CIP.Count<cipSchema>(item => item.status == "save" && item.cc == code);
                        }
                    }
                    else
                    {
                        checking += db.CIP.Count<cipSchema>(item => item.status == "save" && item.cc == checker.deptCode);
                    }

                }
                if (approver.Count != 0)
                {

                    foreach (PermissionSchema permission in approver)
                    {
                        approving += db.CIP.Count<cipSchema>(item => item.status == "cc-checked" && item.cc == permission.deptCode);
                    }
                }
                return Ok(new
                {
                    success = true,
                    message = "Notification number.",
                    data = new {
                        check = checking,
                        approve = approving,
                        sum = checking + approving
                    }
                });
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
            // return Ok(new
            // {
            //     success = true,
            //     data = new
            //     {
            //         requester = 0,
            //         userController = 0,
            //         total = 0,
            //     }
            // }
            // );
        }

    }
}