

using Microsoft.AspNetCore.Mvc;
using cip_api.request;
using cip_api.models;
using System.Collections.Generic;
using Microsoft.AspNetCore.Authorization;

namespace cip_api.controllers
{
    [ApiController]
    [Route("[controller]"), Authorize]
    public class cipUpdateController : ControllerBase
    {

        private Database db;

        public cipUpdateController(Database _db)
        {
            db = _db;
        }

        [HttpPost("draft")]
        public ActionResult draft(cipUpdate body)
        {
            try
            {
                List<cipUpdateSchema> cipUpdate = new List<cipUpdateSchema>();

                foreach (int item in body.cipSchemaid)
                {
                    cipUpdateSchema data = new cipUpdateSchema
                    {
                        actDate = body.actDate,
                        boiType = body.boiType,
                        cipSchemaid = item,
                        classFixedAsset = body.classFixedAsset,
                        costCenterOfUser = body.costCenterOfUser,
                        fixAssetName = body.fixAssetName,
                        fixedAssetCode = body.fixedAssetCode,
                        model = body.model,
                        newBFMorAddBFM = body.newBFMorAddBFM,
                        planDate = body.planDate,
                        processDie = body.processDie,
                        reasonDiff = body.reasonDiff,
                        reasonForDelay = body.reasonForDelay,
                        remark = body.remark,
                        result = body.result,
                        serialNo = body.serialNo,
                        tranferToSupplier = body.tranferToSupplier,
                        upFixAsset = body.upFixAsset
                    };
                    cipUpdate.Add(data);
                    cipSchema cip = db.CIP.Find(item);
                    cip.status = "draft";
                    db.CIP.Update(cip);
                }
                db.CIP_UPDATE.AddRange(cipUpdate);
                db.SaveChanges();

                return Ok(new { success = true, message = "Create draft success." });
            }
            catch (System.Exception e)
            {

                return Problem(e.StackTrace);
            }
        }

         [HttpPost("save")]
        public ActionResult OK(cipUpdate body)
        {
            try
            {
                List<cipUpdateSchema> cipUpdate = new List<cipUpdateSchema>();

                foreach (int item in body.cipSchemaid)
                {
                    cipUpdateSchema data = new cipUpdateSchema
                    {
                        actDate = body.actDate,
                        boiType = body.boiType,
                        cipSchemaid = item,
                        classFixedAsset = body.classFixedAsset,
                        costCenterOfUser = body.costCenterOfUser,
                        fixAssetName = body.fixAssetName,
                        fixedAssetCode = body.fixedAssetCode,
                        model = body.model,
                        newBFMorAddBFM = body.newBFMorAddBFM,
                        planDate = body.planDate,
                        processDie = body.processDie,
                        reasonDiff = body.reasonDiff,
                        reasonForDelay = body.reasonForDelay,
                        remark = body.remark,
                        result = body.result,
                        serialNo = body.serialNo,
                        tranferToSupplier = body.tranferToSupplier,
                        upFixAsset = body.upFixAsset
                    };
                    cipUpdate.Add(data);
                    cipSchema cip = db.CIP.Find(item);
                    cip.status = "save";
                    db.CIP.Update(cip);
                }
                db.CIP_UPDATE.AddRange(cipUpdate);
                db.SaveChanges();

                return Ok(new { success = true, message = "Create save success." });
            }
            catch (System.Exception e)
            {

                return Problem(e.StackTrace);
            }
        }
    }
}