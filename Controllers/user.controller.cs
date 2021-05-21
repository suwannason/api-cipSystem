
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
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace cip_api.controllers
{

    // [Authorize]
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
                userSchema userDB = db.USERS.Find(body.username);
                if (userDB == null)
                {
                    return Unauthorized(new { success = false, message = "Please contact accounting to register for " + body.username });
                }

                using (HttpClient client = new HttpClient())
                {
                    string body_req = "{\"username\": \"" + body.username + "\",\"password\": \"" + body.password + "\"}";
                    client.Timeout = TimeSpan.FromSeconds(20);

                    HttpContent content = new StringContent(body_req, Encoding.UTF8, "application/json");

                    HttpResponseMessage response = client.PostAsync(ldap_auth + "/authentication/ldap", content).Result;

                    LdapResponse data = JsonConvert.DeserializeObject<LdapResponse>(await response.Content.ReadAsStringAsync());
                    // response.EnsureSuccessStatusCode();
                    if (data.success == false)
                    {
                        return BadRequest(new { success = false, message = data.message });
                    }
                    client.Dispose();

                    users user = new users
                    {
                        band = data.data.band,
                        dept = data.data.deptShortName,
                        deptCode = data.data.deptCode,
                        div = data.data.divShortName,
                        name = data.data.fnameEn + " " + data.data.lnameEn,
                        empNo = data.data.empNo,
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


        [HttpPost("login/test"), AllowAnonymous]
        public ActionResult login_test(Login body)
        {

            userSchema data = db.USERS.Find(body.username);
            users user = new users
            {
                band = "a",
                dept = data.deptShortName,
                deptCode = data.deptCode,
                div = "div",
                name = data.name,
                empNo = data.empNo,
            };
            string token = GenerateJSONWebToken(user);

            return Ok(new { success = true, message = "Logon success", token, data = user });
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

        [HttpPost("upload"), AllowAnonymous, Consumes("multipart/form-data")]
        public async Task<ActionResult> upload([FromForm] Upload body)
        {

            string rootFolder = Directory.GetCurrentDirectory();

            string pathString2 = @"\API site\files\CIP-system\upload\";
            string serverPath = rootFolder.Substring(0, rootFolder.LastIndexOf(@"\")) + pathString2;

            Console.WriteLine(serverPath);
            if (!System.IO.Directory.Exists(serverPath))
            {
                Directory.CreateDirectory(serverPath);
            }
            string uuid = System.Guid.NewGuid().ToString();
            string filename = (serverPath + uuid + "-" + body.file.FileName).Trim();
            using (FileStream strem = System.IO.File.Create(filename))
            {
                body.file.CopyTo(strem);
            }

            FileInfo existFile = new FileInfo(filename);

            using (ExcelPackage excel = new ExcelPackage(existFile))
            {

                ExcelWorkbook workbook = excel.Workbook;
                ExcelWorksheet sheet = workbook.Worksheets[0];

                int colCount = sheet.Dimension.Columns;
                int rowCount = sheet.Dimension.Rows;

                List<userSchema> items = new List<userSchema>();
                List<PermissionSchema> permission = new List<PermissionSchema>();
                for (int row = 4; row <= rowCount; row += 1)
                {
                    userSchema approver = new userSchema();
                    PermissionSchema permiss = new PermissionSchema();
                    for (int col = 1; col <= 4; col += 1)
                    {
                        string value = "-";
                        if (sheet.Cells[row, col].Value != null)
                        {
                            value = sheet.Cells[row, col].Value.ToString();
                        }
                        switch (col)
                        {
                            case 1:
                                approver.deptCode = value;
                                permiss.deptCode = value;
                                break;
                            case 2:
                                approver.deptShortName = value;
                                permiss.deptShortName = value;
                                break;
                            case 3:
                                approver.name = value;
                                break;
                            case 4:
                                approver.empNo = value;
                                permiss.empNo = value;
                                break;
                        }
                        permiss.action = "approver";
                    }
                    if (items.FindAll(e => e.empNo == approver.empNo).Count == 0)
                    {
                        items.Add(approver);
                    }
                    permission.Add(permiss);


                    userSchema checker_2 = new userSchema();
                    PermissionSchema permiss_2 = new PermissionSchema();

                    for (int col = 1; col <= 6; col += 1)
                    {
                        string value = "-";
                        if (sheet.Cells[row, col].Value != null)
                        {
                            value = sheet.Cells[row, col].Value.ToString();
                        }
                        switch (col)
                        {
                            case 1:
                                checker_2.deptCode = value;
                                permiss_2.deptCode = value;
                                break;
                            case 2:
                                checker_2.deptShortName = value;
                                permiss_2.deptShortName = value;
                                break;
                            case 5:
                                checker_2.name = value;
                                break;
                            case 6:
                                checker_2.empNo = value;
                                permiss_2.empNo = value;
                                break;
                        }
                        permiss_2.action = "checker";
                    }
                    if (items.FindAll(e => e.empNo == checker_2.empNo).Count == 0)
                    {
                        items.Add(checker_2);
                    }
                    permission.Add(permiss_2);

                    userSchema checker_1 = new userSchema();
                    PermissionSchema permiss_3 = new PermissionSchema();
                    for (int col = 1; col <= 8; col += 1)
                    {
                        string value = "-";
                        if (sheet.Cells[row, col].Value != null)
                        {
                            value = sheet.Cells[row, col].Value.ToString();
                        }
                        switch (col)
                        {
                            case 1:
                                checker_1.deptCode = value;
                                permiss_3.deptCode = value;
                                break;
                            case 2:
                                checker_1.deptShortName = value;
                                permiss_3.deptShortName = value;
                                break;
                            case 7:
                                checker_1.name = value;
                                break;
                            case 8:
                                checker_1.empNo = value;
                                permiss_3.empNo = value;
                                break;
                        }
                        permiss_3.action = "checker";
                    }
                    if (items.FindAll(e => e.empNo == checker_1.empNo).Count == 0)
                    {
                        items.Add(checker_1);
                    }
                    permission.Add(permiss_3);

                    userSchema prepare_1 = new userSchema();
                    PermissionSchema permiss_4 = new PermissionSchema();
                    for (int col = 1; col <= 10; col += 1)
                    {
                        string value = "-";
                        if (sheet.Cells[row, col].Value != null)
                        {
                            value = sheet.Cells[row, col].Value.ToString();
                        }
                        switch (col)
                        {
                            case 1:
                                prepare_1.deptCode = value;
                                permiss_4.deptCode = value;
                                break;
                            case 2:
                                prepare_1.deptShortName = value;
                                permiss_4.deptShortName = value;
                                break;
                            case 9:
                                prepare_1.name = value;
                                break;
                            case 10:
                                prepare_1.empNo = value;
                                permiss_4.empNo = value;
                                break;
                        }
                        permiss_4.action = "prepare";
                    }
                    if (items.FindAll(e => e.empNo == prepare_1.empNo).Count == 0)
                    {
                        items.Add(prepare_1);
                    }
                    permission.Add(permiss_4);

                    userSchema prepare_2 = new userSchema();
                    PermissionSchema permiss_5 = new PermissionSchema();
                    for (int col = 1; col <= 12; col += 1)
                    {
                        string value = "-";
                        if (sheet.Cells[row, col].Value != null)
                        {
                            value = sheet.Cells[row, col].Value.ToString();
                        }
                        switch (col)
                        {
                            case 1:
                                prepare_2.deptCode = value;
                                permiss_5.deptCode = value;
                                break;
                            case 2:
                                prepare_2.deptShortName = value;
                                permiss_5.deptShortName = value;
                                break;
                            case 11:
                                prepare_2.name = value;
                                break;
                            case 12:
                                prepare_2.empNo = value;
                                permiss_5.empNo = value;
                                break;
                        }
                        permiss_5.action = "prepare";
                    }
                    if (items.FindAll(e => e.empNo == prepare_2.empNo).Count == 0)
                    {
                        items.Add(prepare_2);
                    }
                    permission.Add(permiss_5);
                }
                // return Ok(items);
                // db.USERS.Add(items[0]);
                db.PERMISSIONS.AddRange(permission);
                db.USERS.AddRange(items);
                db.SaveChanges();

                List<PermissionSchema> permissions = db.PERMISSIONS.ToList();

                List<PermissionSchema> permissUpdate = new List<PermissionSchema>();
                foreach (PermissionSchema item in permissions)
                {
                    if (item.empNo != "-")
                    {
                        using (HttpClient client = new HttpClient())
                        {

                            client.Timeout = TimeSpan.FromSeconds(20);
                            HttpResponseMessage response = client.GetAsync(ldap_auth + "/authentication/profile?empNo=" + item.empNo).Result;

                            ADprofileResponse data = JsonConvert.DeserializeObject<ADprofileResponse>(await response.Content.ReadAsStringAsync());
                            // response.EnsureSuccessStatusCode();
                            if (data.success == false)
                            {
                                Console.WriteLine(item.empNo);
                                // return BadRequest(new { success = false, message = data.message });
                            }
                            else
                            {
                                PermissionSchema user = db.PERMISSIONS.Where<PermissionSchema>(row => row.empNo == item.empNo).FirstOrDefault();
                                user.email = data.data.email;
                                permissUpdate.Add(user);
                            }

                            client.Dispose();
                        }
                    }
                }
                db.PERMISSIONS.UpdateRange(permissUpdate);
                db.SaveChanges();

                return Ok();
            }

        }
    }
}