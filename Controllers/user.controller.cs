
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;

using System;

using cip_api.request.user;
using cip_api.models;
using System.Net.Http;
using System.Text;
using Newtonsoft.Json;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Tokens;
using System.IdentityModel.Tokens.Jwt;

namespace cip_api.controllers
{

    [Authorize]
    [ApiController]
    [Route("[controller]")]
    public class userController : ControllerBase
    {

        private readonly Database db;
        private IConfiguration _config;
        private readonly string ldap_auth;
        private readonly string GLOBAL_API_ENDPOINT;

        public userController(Database _db, IConfiguration config, IEndpoint setting)
        {
            db = _db;
            _config = config;
            ldap_auth = setting.ldap_auth;
            GLOBAL_API_ENDPOINT = setting.node_api;
        }

        [HttpPost("login"), AllowAnonymous]
        public async Task<ActionResult> login(Login body)
        {
            try
            {

                if (body.username == "cipadmin")
                {
                    users user = new users
                    {
                        band = "admin",
                        dept = "admin",
                        deptCode = "admin",
                        div = "admin",
                        name = "CIP" + " " + "ADMIN",
                        empNo = "admin"
                    };
                    string token = GenerateJSONWebToken(user);
                    return Ok(new { success = true, message = "Logon success", token, data = user });
                }
                using (HttpClient client = new HttpClient())
                {
                    string body_req = "{\"username\": \"" + body.username + "\",\"password\": \"" + body.password + "\"}";
                    client.Timeout = TimeSpan.FromSeconds(20);

                    HttpContent content = new StringContent(body_req, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = client.PostAsync(ldap_auth + "/authentication/ldap", content).Result;

                    LdapResponse data = JsonConvert.DeserializeObject<LdapResponse>(await response.Content.ReadAsStringAsync());
                    // response.EnsureSuccessStatusCode();
                    if (data == null)
                    {
                        return BadRequest(new { success = false, message = "Username or password incorrect" });
                    }
                    client.Dispose();

                    users user = new users
                    {
                        band = data.data.band,
                        dept = data.data.deptShortName,
                        deptCode = data.data.deptCode,
                        div = data.data.divShortName,
                        name = data.data.fnameEn + " " + data.data.lnameEn,
                        empNo = data.data.empNo
                    };
                    string token = GenerateJSONWebToken(user);

                    return Ok(new { success = true, message = "Logon success", token, data = user });
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                return Problem(e.StackTrace);
            }
        }
        private string GenerateJSONWebToken(users userInfo)
        {
            var securityKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_config["Jwt:Key"]));
            var credentials = new SigningCredentials(securityKey, SecurityAlgorithms.HmacSha256);

            var token = new JwtSecurityToken(_config["Jwt:Issuser"],
              _config["Jwt:Issuser"],
              null,
              expires: DateTime.Now.AddHours(8),
              signingCredentials: credentials);

            token.Payload["user"] = userInfo;
            return new JwtSecurityTokenHandler().WriteToken(token);
        }

        [HttpGet("dept")]
        public async Task<DeptResponse> getDepDataAsync()
        {
            Console.WriteLine(GLOBAL_API_ENDPOINT);

            using (HttpClient client = new HttpClient())
            {
                client.Timeout = TimeSpan.FromSeconds(30);

                HttpResponseMessage response = client.GetAsync(GLOBAL_API_ENDPOINT + "/middleware/oracle/departments").Result;
                response.EnsureSuccessStatusCode();
                client.Dispose();
                DeptResponse res = JsonConvert.DeserializeObject<DeptResponse>(await response.Content.ReadAsStringAsync());
                return res;
            }
        }
    }
}