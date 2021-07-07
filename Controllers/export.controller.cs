
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using cip_api.models;
using System.Linq;
using OfficeOpenXml;

namespace cip_api.controllers
{

    [ApiController, Route("[controller]"), Authorize]
    public class exportController : ControllerBase
    {

        private readonly Database db;
        private IConfiguration _config;

        public exportController(Database _db, IConfiguration config)
        {
            db = _db;
            _config = config;
        }

        [HttpGet]
        public ActionResult export()
        {
            try
            {
                string deptShortname = User.FindFirst("dept")?.Value;

                if (deptShortname.ToUpper() != "ACC")
                {
                    return Unauthorized(new { success = false, message = "Permission denied." });
                }

                List<cipSchema> data = db.CIP.Where<cipSchema>(item => item.status == "finished" && item.qty == "1").ToList();
                db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.status == "finished" && (item.newBFMorAddBFM == "New BFM" || item.newBFMorAddBFM.ToLower().Trim() == "newbfm")).ToList();
                return Ok(new { success = true, data });


            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }
        [HttpPost("csv")]
        public ActionResult exportCSV(cip_api.request.ExportACCrequest body)
        {
            try
            {
                using (ExcelPackage package = new ExcelPackage())
                {

                }

                return Ok();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return Problem(e.StackTrace);
            }
        }
    }
}