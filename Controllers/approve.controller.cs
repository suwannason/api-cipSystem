
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

using cip_api.models;
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
                new {
                    success = true,
                    message = "CIP on draft.",
                    data,
                }
            );
        }
        [HttpGet("save")]
        public ActionResult save() {
            string dept = User.FindFirst("dept")?.Value;

            List<cipSchema> data = db.CIP.Where<cipSchema>
            (item => item.status == "save" && item.cc == dept).ToList<cipSchema>();
              return Ok(
                new {
                    success = true,
                    message = "CIP on save.",
                    data,
                }
            );

        }
    }
}