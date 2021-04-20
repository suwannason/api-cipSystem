
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

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

        [HttpGet]
        public ActionResult draft() {
            string dept = User.FindFirst("dept")?.Value;

            return Ok();
        }
    }
}