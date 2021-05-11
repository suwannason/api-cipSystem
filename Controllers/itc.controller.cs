

using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

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

        [HttpGet("waiting")]
        public ActionResult waiting()
        {
            return Ok();
        }

        [HttpGet("confirmed")]
        public ActionResult confirmed()
        {
            return Ok();
        }
    }
}