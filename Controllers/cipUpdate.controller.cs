

using Microsoft.AspNetCore.Mvc;
using cip_api.request;
using cip_api.models;
using System.Collections.Generic;
using Microsoft.AspNetCore.Authorization;
using System.Linq;
using System;
using System.IO;
using OfficeOpenXml;

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

        [HttpPut]
        public ActionResult update(cipUpdateEdit body)
        {
            cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == Int32.Parse(body.id)).FirstOrDefault();

            if (
                body.fixAssetName == cipUpdate.fixAssetName
                && body.fixedAssetCode == cipUpdate.fixedAssetCode
                && body.actDate == cipUpdate.actDate
                && body.addCipBfmNo == cipUpdate.addCipBfmNo
                && body.boiType == cipUpdate.boiType
                && body.classFixedAsset == cipUpdate.classFixedAsset
                && body.costCenterOfUser == cipUpdate.costCenterOfUser
                && body.model == cipUpdate.model
                && body.newBFMorAddBFM == cipUpdate.newBFMorAddBFM
                && body.partNumberDieNo == cipUpdate.partNumberDieNo
                && body.planDate == cipUpdate.planDate
                && body.processDie == cipUpdate.processDie
                && body.reasonDiff == cipUpdate.reasonDiff
                && body.reasonForDelay == cipUpdate.reasonForDelay
                && body.remark == cipUpdate.remark
                && body.result == cipUpdate.result
                && body.serialNo == cipUpdate.serialNo
                && body.tranferToSupplier == cipUpdate.tranferToSupplier
                && body.upFixAsset == cipUpdate.upFixAsset
               )
            {
                return Ok(new { success = true, message = "Not update" });
            }

            cipSchema cip = db.CIP.Find(Int32.Parse(body.id));
            // cip.status = "draft";
            cip.commend = null;
            cip.status = "cc-approved";

            if (cipUpdate == null)
            {
                db.CIP_UPDATE.Add(
                    new cipUpdateSchema
                    {
                        actDate = body.actDate,
                        addCipBfmNo = body.addCipBfmNo,
                        boiType = "-",
                        cipSchemaid = Int32.Parse(body.id),
                        classFixedAsset = body.classFixedAsset,
                        costCenterOfUser = body.costCenterOfUser,
                        createDate = DateTime.Now.ToString("yyyy/MM/dd"),
                        fixAssetName = body.fixAssetName,
                        fixedAssetCode = body.fixedAssetCode,
                        model = body.model,
                        newBFMorAddBFM = body.newBFMorAddBFM,
                        partNumberDieNo = body.partNumberDieNo,
                        planDate = body.planDate,
                        processDie = body.processDie,
                        reasonDiff = body.reasonDiff,
                        reasonForDelay = body.reasonForDelay,
                        remark = body.remark,
                        result = body.result,
                        serialNo = body.serialNo,
                        status = "active",
                        tranferToSupplier = body.tranferToSupplier,
                        upFixAsset = body.upFixAsset,
                    }
                );
                db.SaveChanges();

                return Ok(new { success = true, message = "Confirm CIP success." });
            }

            if (cipUpdate.costCenterOfUser == body.costCenterOfUser)
            {
                cipUpdate.actDate = body.actDate;
                cipUpdate.addCipBfmNo = body.addCipBfmNo;
                cipUpdate.boiType = body.boiType;
                cipUpdate.classFixedAsset = body.classFixedAsset;
                cipUpdate.costCenterOfUser = body.costCenterOfUser;
                cipUpdate.fixAssetName = body.fixAssetName;
                cipUpdate.fixedAssetCode = body.fixedAssetCode;
                cipUpdate.model = body.model;
                cipUpdate.newBFMorAddBFM = body.newBFMorAddBFM;
                cipUpdate.partNumberDieNo = body.partNumberDieNo;
                cipUpdate.planDate = body.planDate;
                cipUpdate.processDie = body.processDie;
                cipUpdate.reasonDiff = body.reasonDiff;
                cipUpdate.reasonForDelay = body.reasonForDelay;
                cipUpdate.remark = body.remark;
                cipUpdate.result = body.result;
                cipUpdate.serialNo = body.serialNo;
                cipUpdate.tranferToSupplier = body.tranferToSupplier;
                cipUpdate.upFixAsset = body.upFixAsset;
                cipUpdate.status = "active";
                List<ApprovalSchema> approve = db.APPROVAL.Where<ApprovalSchema>(item => item.cipSchemaid == cipUpdate.cipSchemaid && (item.onApproveStep.StartsWith("cost"))).ToList();
                db.APPROVAL.RemoveRange(approve);
            }
            else
            {
                if (
                    body.costCenterOfUser == "2130" && (cipUpdate.costCenterOfUser == "2140" || cipUpdate.costCenterOfUser == "9555")
                    || body.costCenterOfUser == "2410" && (cipUpdate.costCenterOfUser == "2130" || cipUpdate.costCenterOfUser == "9555")
                    || body.costCenterOfUser == "9555" && (cipUpdate.costCenterOfUser == "2130" || cipUpdate.costCenterOfUser == "2140")
                    || body.costCenterOfUser.StartsWith("55") && cipUpdate.costCenterOfUser.StartsWith("55")
                    || body.costCenterOfUser == "5610" && cipUpdate.costCenterOfUser == "5619"
                    || body.costCenterOfUser == "5619" && cipUpdate.costCenterOfUser == "5610"
                    || body.costCenterOfUser == "5650" && (cipUpdate.costCenterOfUser == "5655" || cipUpdate.costCenterOfUser == "9333")
                    || body.costCenterOfUser == "5655" && (cipUpdate.costCenterOfUser == "5650" || cipUpdate.costCenterOfUser == "9333")
                    || body.costCenterOfUser == "9333" && (cipUpdate.costCenterOfUser == "5650" || cipUpdate.costCenterOfUser == "5655")
                    || body.costCenterOfUser == "5670" && (cipUpdate.costCenterOfUser == "5675" || cipUpdate.costCenterOfUser == "9444")
                    || body.costCenterOfUser == "5675" && (cipUpdate.costCenterOfUser == "5670" || cipUpdate.costCenterOfUser == "9444")
                    || body.costCenterOfUser == "9444" && (cipUpdate.costCenterOfUser == "5670" || cipUpdate.costCenterOfUser == "5675")
                    )
                {
                    cipUpdate.actDate = body.actDate;
                    cipUpdate.addCipBfmNo = body.addCipBfmNo;
                    cipUpdate.boiType = body.boiType;
                    cipUpdate.classFixedAsset = body.classFixedAsset;
                    cipUpdate.costCenterOfUser = body.costCenterOfUser;
                    cipUpdate.fixAssetName = body.fixAssetName;
                    cipUpdate.fixedAssetCode = body.fixedAssetCode;
                    cipUpdate.model = body.model;
                    cipUpdate.newBFMorAddBFM = body.newBFMorAddBFM;
                    cipUpdate.partNumberDieNo = body.partNumberDieNo;
                    cipUpdate.planDate = body.planDate;
                    cipUpdate.processDie = body.processDie;
                    cipUpdate.reasonDiff = body.reasonDiff;
                    cipUpdate.reasonForDelay = body.reasonForDelay;
                    cipUpdate.remark = body.remark;
                    cipUpdate.result = body.result;
                    cipUpdate.serialNo = body.serialNo;
                    cipUpdate.tranferToSupplier = body.tranferToSupplier;
                    cipUpdate.upFixAsset = body.upFixAsset;
                    cipUpdate.status = "active";
                }
                else
                {
                    cipUpdate.actDate = body.actDate;
                    cipUpdate.addCipBfmNo = body.addCipBfmNo;
                    cipUpdate.boiType = body.boiType;
                    cipUpdate.classFixedAsset = body.classFixedAsset;
                    cipUpdate.costCenterOfUser = body.costCenterOfUser;
                    cipUpdate.fixAssetName = body.fixAssetName;
                    cipUpdate.fixedAssetCode = body.fixedAssetCode;
                    cipUpdate.model = body.model;
                    cipUpdate.newBFMorAddBFM = body.newBFMorAddBFM;
                    cipUpdate.partNumberDieNo = body.partNumberDieNo;
                    cipUpdate.planDate = body.planDate;
                    cipUpdate.processDie = body.processDie;
                    cipUpdate.reasonDiff = body.reasonDiff;
                    cipUpdate.reasonForDelay = body.reasonForDelay;
                    cipUpdate.remark = body.remark;
                    cipUpdate.result = body.result;
                    cipUpdate.serialNo = body.serialNo;
                    cipUpdate.tranferToSupplier = body.tranferToSupplier;
                    cipUpdate.upFixAsset = body.upFixAsset;
                    cipUpdate.status = "active";

                    // cip.status = "cc-approved";
                    List<ApprovalSchema> approve = db.APPROVAL.Where<ApprovalSchema>(item =>
                    item.cipSchemaid == cipUpdate.cipSchemaid && (item.onApproveStep.StartsWith("cost"))).ToList();
                    db.APPROVAL.RemoveRange(approve);
                }
            }
            db.CIP.Update(cip);
            db.CIP_UPDATE.Update(cipUpdate);

            db.SaveChanges();
            return Ok(new { success = true, message = "Update CIP success." });
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
                        upFixAsset = body.upFixAsset,
                        createDate = DateTime.Now.ToString("yyyyMMdd")
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
                cipUpdateSchema cipUpdatedata = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == body.cipSchemaid[0]).FirstOrDefault();

                string username = User.FindFirst("username")?.Value;

                if (cipUpdatedata != null)
                {
                    cipUpdatedata.actDate = body.actDate;
                    cipUpdatedata.boiType = body.boiType;
                    cipUpdatedata.classFixedAsset = body.classFixedAsset;
                    cipUpdatedata.costCenterOfUser = body.costCenterOfUser;
                    cipUpdatedata.createDate = DateTime.Now.ToString("yyyy/MM/dd");
                    cipUpdatedata.fixAssetName = body.fixAssetName;
                    cipUpdatedata.fixedAssetCode = body.fixedAssetCode;
                    cipUpdatedata.partNumberDieNo = body.partNumberDieNo;
                    cipUpdatedata.model = body.model;
                    cipUpdatedata.newBFMorAddBFM = body.newBFMorAddBFM;
                    cipUpdatedata.planDate = body.planDate;
                    cipUpdatedata.processDie = body.processDie;
                    cipUpdatedata.reasonDiff = body.reasonDiff;
                    cipUpdatedata.reasonForDelay = body.reasonForDelay;
                    cipUpdatedata.remark = body.remark;
                    cipUpdatedata.result = body.result;
                    cipUpdatedata.serialNo = body.serialNo;
                    cipUpdatedata.status = "active";
                    cipUpdatedata.tranferToSupplier = body.tranferToSupplier;
                    cipUpdatedata.upFixAsset = body.upFixAsset;

                    db.CIP_UPDATE.Update(cipUpdatedata);
                    db.SaveChanges();

                    return Ok(new { success = true, message = "Update to User confirm CIP" });
                }
                List<cipUpdateSchema> cipUpdateIns = new List<cipUpdateSchema>();

                List<ApprovalSchema> approver = new List<ApprovalSchema>();
                foreach (int item in body.cipSchemaid)
                {
                    cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(cipUpdate => cipUpdate.cipSchemaid == item).FirstOrDefault();

                    if (cipUpdate == null)
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
                            partNumberDieNo = body.partNumberDieNo,
                            remark = body.remark,
                            result = body.result,
                            serialNo = body.serialNo,
                            tranferToSupplier = body.tranferToSupplier,
                            upFixAsset = body.upFixAsset
                        };
                        cipUpdateIns.Add(data);
                        cipSchema cip = db.CIP.Find(item);
                        cip.status = "save";
                        db.CIP.Update(cip);

                        ApprovalSchema approve = new ApprovalSchema();
                        approve.onApproveStep = "save";
                        approve.empNo = username;
                        approve.cipSchemaid = item;
                        approve.date = DateTime.Now.ToString("yyyy/MM/dd");

                        approver.Add(approve);
                    }
                    else
                    { // case update
                        cipUpdate.actDate = body.actDate;
                        cipUpdate.boiType = body.boiType;
                        cipUpdate.classFixedAsset = body.classFixedAsset;
                        cipUpdate.costCenterOfUser = body.costCenterOfUser;
                        cipUpdate.fixAssetName = body.fixAssetName;
                        cipUpdate.fixedAssetCode = body.fixedAssetCode;
                        cipUpdate.model = body.model;
                        cipUpdate.newBFMorAddBFM = body.newBFMorAddBFM;
                        cipUpdate.planDate = body.planDate;
                        cipUpdate.processDie = body.processDie;
                        cipUpdate.partNumberDieNo = body.partNumberDieNo;
                        cipUpdate.reasonDiff = body.reasonDiff;
                        cipUpdate.reasonForDelay = body.reasonForDelay;
                        cipUpdate.remark = body.remark;
                        cipUpdate.result = body.result;
                        cipUpdate.serialNo = body.serialNo;
                        cipUpdate.tranferToSupplier = body.tranferToSupplier;
                        cipUpdate.upFixAsset = body.upFixAsset;
                        db.CIP_UPDATE.Update(cipUpdate);

                        ApprovalSchema approve = new ApprovalSchema();
                        cipSchema cip = db.CIP.Find(item);
                        if (cip.cc != cipUpdate.costCenterOfUser)
                        {
                            approve.onApproveStep = "cost-prepared";
                        }
                        else
                        {
                            approve.onApproveStep = "save";
                        }
                        approve.empNo = username;
                        approve.cipSchemaid = item;
                        approve.date = DateTime.Now.ToString("yyyy/MM/dd");

                        approver.Add(approve);
                    }
                }
                if (approver.Count > 0)
                {
                    db.APPROVAL.AddRange(approver);
                }
                db.CIP_UPDATE.AddRange(cipUpdateIns);
                db.SaveChanges();

                return Ok(new { success = true, message = "Create save success." });
            }
            catch (System.Exception e)
            {

                return Problem(e.StackTrace);
            }
        }

        [HttpGet("{cipId}")]
        public ActionResult getByCIPid(int cipId)
        {

            cipUpdateSchema data = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == cipId).FirstOrDefault<cipUpdateSchema>();

            return Ok(new
            {
                success = true,
                message = "CIP update",
                data,
            });
        }

        [HttpPost("prepare/change"), Consumes("multipart/form-data")]
        public ActionResult requesterEdit([FromForm] CIPupload body)
        {
            try
            {

                string rootFolder = Directory.GetCurrentDirectory();
                string pathString2 = @"\API site\files\CIP-system\upload\";
                string serverPath = rootFolder.Substring(0, rootFolder.LastIndexOf(@"\")) + pathString2;

                if (!Directory.Exists(serverPath))
                {
                    Directory.CreateDirectory(serverPath);
                }
                string deptCode = User.FindFirst("deptCode").Value;
                string username = User.FindFirst("username")?.Value;
                string fileName = System.Guid.NewGuid().ToString() + "-" + body.file.FileName;
                FileStream strem = System.IO.File.Create($"{serverPath}{fileName}");
                body.file.CopyTo(strem);
                strem.Close();

                string path = $"{serverPath}{fileName}";
                FileInfo Existfile = new FileInfo(path);

                List<cipUpdateSchema> updateItem = new List<cipUpdateSchema>();


                using (ExcelPackage excel = new ExcelPackage(Existfile))
                {
                    ExcelWorkbook workbook = excel.Workbook;
                    ExcelWorksheet sheet = workbook.Worksheets[0];

                    int colCount = sheet.Dimension.End.Column;
                    int rowCount = sheet.Dimension.End.Row;

                    for (int row = 3; row <= rowCount; row += 1)
                    {
                        cipUpdateSchema item = new cipUpdateSchema();
                        string cipNo = sheet.Cells[row, 3].Value?.ToString();
                        string subCipNo = sheet.Cells[row, 4].Value?.ToString();

                        for (int col = 3; col <= colCount; col += 1)
                        {
                            string value = sheet.Cells[row, col].Value?.ToString();
                            if (value == null)
                            {
                                value = "-";
                            }
                            switch (col)
                            {
                                case 3:
                                    if (value == "-")
                                    {
                                        break;
                                    }
                                    break;

                                case 30: item.planDate = value; break;
                                case 31: item.actDate = value; break;
                                case 32: item.result = value; break;
                                case 33: item.reasonDiff = value; break;
                                case 34: item.fixedAssetCode = value; break;
                                case 35: item.classFixedAsset = value; break;
                                case 36:
                                    if (value.IndexOf(',') != -1 || value.IndexOf(':') != -1
                                        || value.IndexOf('"') != -1 || value.IndexOf('#') != -1
                                        || value.IndexOf('!') != -1 || value.IndexOf('*') != -1 || value.Length > 100)
                                    {
                                        return BadRequest(new { success = false, message = "Not allow Symbol fix asset name value." });
                                    }
                                    item.fixAssetName = value;
                                    break;
                                case 37: item.serialNo = value; break;
                                case 38: item.partNumberDieNo = value; break;
                                case 39: item.processDie = value; break;
                                case 40: item.model = value; break;
                                case 41: item.costCenterOfUser = value; break;
                                case 42: item.tranferToSupplier = value; break;
                                case 43: item.upFixAsset = value; break;
                                case 44: item.newBFMorAddBFM = value; break;
                                case 45: item.reasonForDelay = value; break;
                                case 46: item.addCipBfmNo = value; break;
                                case 47: item.remark = value; break;
                                case 48: item.boiType = value; break;
                            }
                        }
                        cipSchema data = db.CIP.Where<cipSchema>(item => item.cipNo == cipNo && item.subCipNo == subCipNo).FirstOrDefault();
                        cipUpdateSchema cipUpdate = db.CIP_UPDATE.Where<cipUpdateSchema>(item => item.cipSchemaid == data.id).FirstOrDefault();
                        item.status = cipUpdate.status;
                        item.createDate = cipUpdate.createDate;
                        cipUpdate = item;
                        cipUpdate.cipSchemaid = data.id;
                        updateItem.Add(cipUpdate);
                    }
                    db.CIP_UPDATE.UpdateRange(updateItem);
                    db.SaveChanges();
                }
                return Ok(new { success = true, message = "Update data confirm " + updateItem.Count + " success." });
            }
            catch (Exception e)
            {
                return Problem(e.StackTrace);
            }
        }

    }
}