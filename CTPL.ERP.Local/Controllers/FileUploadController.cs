using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using CTPL.ERP.Local.Data;
using CTPL.ERP.Local.Models;
using OfficeOpenXml;

namespace CTPL.ERP.Local.Controllers
{
    public class FileUploadController : Controller
    {
        private readonly YourDbContext _db = new YourDbContext(); 

        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            if (file != null && file.ContentLength > 0)
            {
                using (var package = new ExcelPackage(file.InputStream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var iccid1 = worksheet.Cells[row, 1].Text;

                        // Fetch existing record if available
                        var existingRecord = _db.InternalSimsActivationsModels
                            .FirstOrDefault(m => m.ICCID_1 == iccid1);

                        if (existingRecord != null)
                        {
                            // Update existing record only if new data is not null or empty
                            existingRecord.ICCID_2 = string.IsNullOrEmpty(worksheet.Cells[row, 2].Text) ? existingRecord.ICCID_2 : worksheet.Cells[row, 2].Text;
                            existingRecord.ICCID_1_Network = string.IsNullOrEmpty(worksheet.Cells[row, 3].Text) ? existingRecord.ICCID_1_Network : worksheet.Cells[row, 3].Text;
                            existingRecord.ICCID_2_Network = string.IsNullOrEmpty(worksheet.Cells[row, 4].Text) ? existingRecord.ICCID_2_Network : worksheet.Cells[row, 4].Text;

                            existingRecord.IMSI_1 = string.IsNullOrEmpty(worksheet.Cells[row, 5].Text) ? existingRecord.IMSI_1 : Convert.ToInt64(worksheet.Cells[row, 5].Text);
                            existingRecord.IMSI_2 = string.IsNullOrEmpty(worksheet.Cells[row, 6].Text) ? existingRecord.IMSI_2 : Convert.ToInt64(worksheet.Cells[row, 6].Text);

                            existingRecord.MSISDN_1 = string.IsNullOrEmpty(worksheet.Cells[row, 7].Text) ? existingRecord.MSISDN_1 : Convert.ToInt64(worksheet.Cells[row, 7].Text);
                            existingRecord.MSISDN_2 = string.IsNullOrEmpty(worksheet.Cells[row, 8].Text) ? existingRecord.MSISDN_2 : Convert.ToInt64(worksheet.Cells[row, 8].Text);

                            existingRecord.ESN = string.IsNullOrEmpty(worksheet.Cells[row, 9].Text) ? existingRecord.ESN : worksheet.Cells[row, 9].Text;

                            existingRecord.BootstrapActivationStartDate = string.IsNullOrEmpty(worksheet.Cells[row, 10].Text) ? existingRecord.BootstrapActivationStartDate : DateTime.Parse(worksheet.Cells[row, 10].Text);
                            existingRecord.BootstrapActivationEndDate = string.IsNullOrEmpty(worksheet.Cells[row, 11].Text) ? existingRecord.BootstrapActivationEndDate : DateTime.Parse(worksheet.Cells[row, 11].Text);

                            existingRecord.AllocatedToInHouseDate = string.IsNullOrEmpty(worksheet.Cells[row, 12].Text) ? existingRecord.AllocatedToInHouseDate : DateTime.Parse(worksheet.Cells[row, 12].Text);

                            existingRecord.APN_Name = string.IsNullOrEmpty(worksheet.Cells[row, 13].Text) ? existingRecord.APN_Name : worksheet.Cells[row, 13].Text;

                            existingRecord.IP_1 = string.IsNullOrEmpty(worksheet.Cells[row, 14].Text) ? existingRecord.IP_1 : worksheet.Cells[row, 14].Text;
                            existingRecord.IP_2 = string.IsNullOrEmpty(worksheet.Cells[row, 15].Text) ? existingRecord.IP_2 : worksheet.Cells[row, 15].Text;
                            existingRecord.IP_3 = string.IsNullOrEmpty(worksheet.Cells[row, 16].Text) ? existingRecord.IP_3 : worksheet.Cells[row, 16].Text;
                            existingRecord.IP_4 = string.IsNullOrEmpty(worksheet.Cells[row, 17].Text) ? existingRecord.IP_4 : worksheet.Cells[row, 17].Text;

                            existingRecord.MN_1 = string.IsNullOrEmpty(worksheet.Cells[row, 18].Text) ? existingRecord.MN_1 : worksheet.Cells[row, 18].Text;
                            existingRecord.MN_2 = string.IsNullOrEmpty(worksheet.Cells[row, 19].Text) ? existingRecord.MN_2 : worksheet.Cells[row, 19].Text;
                            existingRecord.MN_3 = string.IsNullOrEmpty(worksheet.Cells[row, 20].Text) ? existingRecord.MN_3 : worksheet.Cells[row, 20].Text;
                            existingRecord.MN_4 = string.IsNullOrEmpty(worksheet.Cells[row, 21].Text) ? existingRecord.MN_4 : worksheet.Cells[row, 21].Text;

                            existingRecord.IMEI = string.IsNullOrEmpty(worksheet.Cells[row, 22].Text) ? existingRecord.IMEI : worksheet.Cells[row, 22].Text;

                            existingRecord.For_User = string.IsNullOrEmpty(worksheet.Cells[row, 23].Text) ? existingRecord.For_User : worksheet.Cells[row, 23].Text;
                            existingRecord.For_State = string.IsNullOrEmpty(worksheet.Cells[row, 24].Text) ? existingRecord.For_State : worksheet.Cells[row, 24].Text;

                            existingRecord.Dispatch_Date = string.IsNullOrEmpty(worksheet.Cells[row, 25].Text) ? existingRecord.Dispatch_Date : DateTime.Parse(worksheet.Cells[row, 25].Text);
                            existingRecord.Dispatch_Location = string.IsNullOrEmpty(worksheet.Cells[row, 26].Text) ? existingRecord.Dispatch_Location : worksheet.Cells[row, 26].Text;
                        }
                        else
                        {

                            var model = new Internal_Sims_Activations
                            {
                                ICCID_1 = iccid1,
                                ICCID_2 = worksheet.Cells[row, 2].Text,
                                ICCID_1_Network = worksheet.Cells[row, 3].Text,
                                ICCID_2_Network = worksheet.Cells[row, 4].Text,

                                IMSI_1 = string.IsNullOrEmpty(worksheet.Cells[row, 5].Text) ? (long?)null : Convert.ToInt64(worksheet.Cells[row, 5].Text),
                                IMSI_2 = string.IsNullOrEmpty(worksheet.Cells[row, 6].Text) ? (long?)null : Convert.ToInt64(worksheet.Cells[row, 6].Text),

                                MSISDN_1 = string.IsNullOrEmpty(worksheet.Cells[row, 7].Text) ? (long?)null : Convert.ToInt64(worksheet.Cells[row, 7].Text),
                                MSISDN_2 = string.IsNullOrEmpty(worksheet.Cells[row, 8].Text) ? (long?)null : Convert.ToInt64(worksheet.Cells[row, 8].Text),

                                ESN = worksheet.Cells[row, 9].Text,

                                BootstrapActivationStartDate = string.IsNullOrEmpty(worksheet.Cells[row, 10].Text) ? (DateTime?)null : DateTime.Parse(worksheet.Cells[row, 10].Text),
                                BootstrapActivationEndDate = string.IsNullOrEmpty(worksheet.Cells[row, 11].Text) ? (DateTime?)null : DateTime.Parse(worksheet.Cells[row, 11].Text),

                                AllocatedToInHouseDate = string.IsNullOrEmpty(worksheet.Cells[row, 12].Text) ? (DateTime?)null : DateTime.Parse(worksheet.Cells[row, 12].Text),

                                APN_Name = worksheet.Cells[row, 13].Text,

                                IP_1 = worksheet.Cells[row, 14].Text,
                                IP_2 = worksheet.Cells[row, 15].Text,
                                IP_3 = worksheet.Cells[row, 16].Text,
                                IP_4 = worksheet.Cells[row, 17].Text,

                                MN_1 = worksheet.Cells[row, 18].Text,
                                MN_2 = worksheet.Cells[row, 19].Text,
                                MN_3 = worksheet.Cells[row, 20].Text,
                                MN_4 = worksheet.Cells[row, 21].Text,

                                IMEI = worksheet.Cells[row, 22].Text,

                                For_User = worksheet.Cells[row, 23].Text,
                                For_State = worksheet.Cells[row, 24].Text,

                                Dispatch_Date = string.IsNullOrEmpty(worksheet.Cells[row, 25].Text) ? (DateTime?)null : DateTime.Parse(worksheet.Cells[row, 25].Text),
                                Dispatch_Location = worksheet.Cells[row, 26].Text
                            };

                            _db.InternalSimsActivationsModels.Add(model);
                        }
                    }

                    _db.SaveChanges();
                }
            }

            return RedirectToAction("Index", "InternalSimsActivations");
        }





        public ActionResult ExportToExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Internal Sims Activations");

                // Add column headers
                worksheet.Cells[1, 1].Value = "ICCID_1";
                worksheet.Cells[1, 2].Value = "ICCID_2";
                worksheet.Cells[1, 3].Value = "ICCID_1_Network";
                worksheet.Cells[1, 4].Value = "ICCID_2_Network";
                worksheet.Cells[1, 5].Value = "IMSI_1";
                worksheet.Cells[1, 6].Value = "IMSI_2";
                worksheet.Cells[1, 7].Value = "MSISDN_1";
                worksheet.Cells[1, 8].Value = "MSISDN_2";
                worksheet.Cells[1, 9].Value = "ESN";
                worksheet.Cells[1, 10].Value = "BootstrapActivationStartDate";
                worksheet.Cells[1, 11].Value = "BootstrapActivationEndDate";
                worksheet.Cells[1, 12].Value = "AllocatedToInHouseDate";
                worksheet.Cells[1, 13].Value = "APN_Name";
                worksheet.Cells[1, 14].Value = "IP_1";
                worksheet.Cells[1, 15].Value = "IP_2";
                worksheet.Cells[1, 16].Value = "IP_3";
                worksheet.Cells[1, 17].Value = "IP_4";
                worksheet.Cells[1, 18].Value = "MN_1";
                worksheet.Cells[1, 19].Value = "MN_2";
                worksheet.Cells[1, 20].Value = "MN_3";
                worksheet.Cells[1, 21].Value = "MN_4";
                worksheet.Cells[1, 22].Value = "IMEI";
                worksheet.Cells[1, 23].Value = "For_User";
                worksheet.Cells[1, 24].Value = "For_State";
                worksheet.Cells[1, 25].Value = "Dispatch_Date";
                worksheet.Cells[1, 26].Value = "Dispatch_Location";

                var records = _db.InternalSimsActivationsModels.ToList();

                int row = 2;
                foreach (var record in records)
                {
                    worksheet.Cells[row, 1].Value = record.ICCID_1;
                    worksheet.Cells[row, 2].Value = record.ICCID_2;
                    worksheet.Cells[row, 3].Value = record.ICCID_1_Network;
                    worksheet.Cells[row, 4].Value = record.ICCID_2_Network;
                    worksheet.Cells[row, 5].Value = record.IMSI_1?.ToString(); // Convert to string to avoid scientific notation
                    worksheet.Cells[row, 6].Value = record.IMSI_2?.ToString(); // Convert to string to avoid scientific notation
                    worksheet.Cells[row, 7].Value = record.MSISDN_1?.ToString(); // Convert to string to avoid scientific notation
                    worksheet.Cells[row, 8].Value = record.MSISDN_2?.ToString(); // Convert to string to avoid scientific notation
                    worksheet.Cells[row, 9].Value = record.ESN;

                    // Format date cells
                    var bootstrapActivationStartDateCell = worksheet.Cells[row, 10];
                    if (record.BootstrapActivationStartDate.HasValue)
                    {
                        bootstrapActivationStartDateCell.Value = record.BootstrapActivationStartDate.Value;
                        bootstrapActivationStartDateCell.Style.Numberformat.Format = "yyyy/MM/dd";
                    }
                    else
                    {
                        bootstrapActivationStartDateCell.Value = DBNull.Value;
                    }

                    var bootstrapActivationEndDateCell = worksheet.Cells[row, 11];
                    if (record.BootstrapActivationEndDate.HasValue)
                    {
                        bootstrapActivationEndDateCell.Value = record.BootstrapActivationEndDate.Value;
                        bootstrapActivationEndDateCell.Style.Numberformat.Format = "yyyy/MM/dd";
                    }
                    else
                    {
                        bootstrapActivationEndDateCell.Value = DBNull.Value;
                    }

                    var allocatedToInHouseDateCell = worksheet.Cells[row, 12];
                    if (record.AllocatedToInHouseDate.HasValue)
                    {
                        allocatedToInHouseDateCell.Value = record.AllocatedToInHouseDate.Value;
                        allocatedToInHouseDateCell.Style.Numberformat.Format = "yyyy/MM/dd";
                    }
                    else
                    {
                        allocatedToInHouseDateCell.Value = DBNull.Value;
                    }

                    worksheet.Cells[row, 13].Value = record.APN_Name;
                    worksheet.Cells[row, 14].Value = record.IP_1;
                    worksheet.Cells[row, 15].Value = record.IP_2;
                    worksheet.Cells[row, 16].Value = record.IP_3;
                    worksheet.Cells[row, 17].Value = record.IP_4;
                    worksheet.Cells[row, 18].Value = record.MN_1;
                    worksheet.Cells[row, 19].Value = record.MN_2;
                    worksheet.Cells[row, 20].Value = record.MN_3;
                    worksheet.Cells[row, 21].Value = record.MN_4;
                    worksheet.Cells[row, 22].Value = record.IMEI;
                    worksheet.Cells[row, 23].Value = record.For_User;
                    worksheet.Cells[row, 24].Value = record.For_State;

                    // Format Dispatch_Date as short date
                    var dispatchDateCell = worksheet.Cells[row, 25];
                    if (record.Dispatch_Date.HasValue)
                    {
                        dispatchDateCell.Value = record.Dispatch_Date.Value;
                        dispatchDateCell.Style.Numberformat.Format = "yyyy/MM/dd";
                    }
                    else
                    {
                        dispatchDateCell.Value = DBNull.Value;
                    }

                    worksheet.Cells[row, 26].Value = record.Dispatch_Location;

                    row++;
                }

                // Save the Excel file to a MemoryStream
                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                // Return the file to the user
                string fileName = "Internal_Sims_Activations.xlsx";
                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }









    }
}