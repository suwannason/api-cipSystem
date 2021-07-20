

using cip_api.models;
using cip_api.request.user;
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
                accController acc = new accController(db, _config, _setting);

                approvalController approval = new approvalController(db, _config, _setting);

                itcController itc = new itcController(db, _config, _setting);

                ProfileUser user = new ProfileUser {
                    deptCode = User.FindFirst("deptCode")?.Value,
                    username = User.FindFirst("username")?.Value,
                };

                OkObjectResult cc_requester = approval.cc(user.username, user.deptCode) as OkObjectResult;
                OkObjectResult cc_user = approval.costCenter(user.username, user.deptCode) as OkObjectResult;
                OkObjectResult waiting_fa = acc.accFinishData(user.username, user.deptCode) as OkObjectResult;
                OkObjectResult acc_diff = acc.codeDiff(user.username, user.deptCode) as OkObjectResult;
                OkObjectResult itc_confirm = itc.waiting(user.username, user.deptCode) as OkObjectResult;
                OkObjectResult itc_confirmed = itc.confirmed(user.username, user.deptCode) as OkObjectResult;

                string ccRequester = cc_requester == null ? "0" : cc_requester.Value.ToString();
                string ccUser = cc_user == null ? "0" : cc_user.Value.ToString();
                string waitingFA = waiting_fa == null ? "0" : waiting_fa.Value.ToString();
                string codeDiff = acc_diff == null ? "0" : acc_diff.Value.ToString();
                string itcConfirm = itc_confirm == null ? "0" : itc_confirm.Value.ToString();
                string itcConfirmed = itc_confirmed == null ? "0" : itc_confirmed.Value.ToString();

                return Ok(new
                {
                    success = true,
                    message = "Number notification",
                    data = new
                    {
                        ccRequester,
                        ccUser,
                        waitingFA,
                        codeDiff,
                        itcConfirm,
                        itcConfirmed
                    }
                });
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return Problem(e.StackTrace);
            }
        }

    }
}