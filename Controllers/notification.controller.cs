

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
                string deptCode = User.FindFirst("deptCode")?.Value;
                string username = User.FindFirst("username")?.Value;
                List<PermissionSchema> permissions = GetPermissions(username);
                Int32 requester = 0;
                Int32 user = 0;

                PermissionSchema prepare = permissions.Find(item => item.empNo == username && item.action == "prepare");
                PermissionSchema checker = permissions.Find(item => item.empNo == username && item.action == "checker");
                PermissionSchema approver = permissions.Find(item => item.empNo == username && item.action == "approver");

                if (prepare != null)
                {
                    requester += db.CIP.Count<cipSchema>(item => item.cc == deptCode && item.status == "open");
                }
                if (checker != null)
                {
                    requester += db.CIP.Count<cipSchema>(item => item.status == "cc-prepared");
                    List<cipSchema> cip = db.CIP.Where<cipSchema>(item => item.cc == deptCode && item.status == "cost-prepared").ToList();
                                          db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "active");

                                          return Ok(cip);


                }
                if (approver != null)
                {

                }

                return Ok(new
                {
                    success = true,
                    message = "Notification number.",
                    data = new
                    {
                        requester,
                        user,
                        sum = requester + user
                    }
                });
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }

    }
}