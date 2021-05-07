
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

using cip_api.models;
using cip_api.request;
using System.Collections.Generic;
using System.Linq;

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
            string deptCode = User.FindFirst("deptCode")?.Value;
            string action = User.FindFirst("action")?.Value;

            System.Console.WriteLine("action: " + action);
            List<cipSchema> data = null;
            if (action == "checker")
            {
                data = db.CIP.Where<cipSchema>(item => item.status == "save" && item.cc == deptCode).ToList<cipSchema>();
            }
            else if (action == "approver")
            {
                data = db.CIP.Where<cipSchema>(item => item.status == "cc-checked" && item.cc == deptCode).ToList<cipSchema>();
            }

            return Ok(
              new
              {
                  success = true,
                  message = "CIP on save.",
                  data,
              }
          );
        }

        [HttpGet("costCenter")]
        public ActionResult costCenter()
        {
            return Ok();
        }

        [HttpPut("approve")]
        public ActionResult approve(Approve body) {

            foreach (string item in body.id) {
                
            }
            return Ok();
        }
    }
}