
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

namespace cip_api.controllers
{

    [ApiController, Route("[controller]"), Authorize]
    public class accController : ControllerBase
    {
        private readonly Database db;
        private IConfiguration _config;
        private readonly string ldap_auth;
        public accController(Database _db, IConfiguration config, IEndpoint setting)
        {
            db = _db;
            _config = config;
            ldap_auth = setting.ldap_auth;
        }

        [HttpGet("fa")]
        public ActionResult waitingFA() {
            return Ok();
        }

        [HttpGet("tracking")]
        public ActionResult tracking() {
            return Ok();
        }

        [HttpGet("diff")]
        public ActionResult codeDiff() {

            return Ok();
        }
    }
}