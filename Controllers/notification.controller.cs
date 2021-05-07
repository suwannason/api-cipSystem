

using System;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Linq;

// using System;

// using cip_api.request.user;
using cip_api.models;
using System.Collections.Generic;
// using System.Net.Http;
// using System.Text;
// using Newtonsoft.Json;
// using System.Threading.Tasks;
// using Microsoft.IdentityModel.Tokens;
// using System.IdentityModel.Tokens.Jwt;


namespace cip_api.controllers
{

    [Authorize]
    [ApiController]
    [Route("[controller]")]
    public class notificationController : ControllerBase
    {

        private readonly Database db;
        private IConfiguration _config;
        private readonly string ldap_auth;
        private readonly string GLOBAL_API_ENDPOINT;

        public notificationController(Database _db, IConfiguration config, IEndpoint setting)
        {
            db = _db;
            _config = config;
            ldap_auth = setting.ldap_auth;
            GLOBAL_API_ENDPOINT = setting.node_api;
        }

        [HttpGet]
        public ActionResult getNumber(string page, string perPage)
        {

            Int32 page_ = Int32.Parse(page);
            Int32 perPage_ = Int32.Parse(perPage);

            string empNo = User.FindFirst("username")?.Value;
            Int32 num = db.NOTIFICATIONS.Count<NotificationSchema>(item => item.userSchemaempNo == empNo && item.status == "created");
            List<NotificationSchema> data = db.NOTIFICATIONS.Where<NotificationSchema>(item => item.userSchemaempNo == empNo).OrderBy(e => e.id).ThenByDescending(e => e.createDate).Skip((page_ - 1) * perPage_).Take(perPage_).ToList();
            return Ok(
                new
                {
                    success = true,
                    message = "Notification number for user",
                    number = num.ToString(),
                    data,
                }
            );
        }

    }
}