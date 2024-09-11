using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Web.Mvc;
using CTPL.ERP.Local.Data;
using System.IO;
using ExcelDataReader;

namespace CTPL.ERP.Local.Controllers
{
    public class InternalSimsActivationsController : Controller
    {
        private ERP_LocalEntities db = new ERP_LocalEntities();

        public ActionResult Index(string searchBy, string searchKeyword)
        {
            ViewBag.CurrentSearchBy = searchBy;
            ViewBag.CurrentSearchKeyword = searchKeyword;

            var records = from r in db.Internal_Sims_Activations
                          select r;

            if (!string.IsNullOrEmpty(searchKeyword))
            {
                long longSearchKeyword;
                bool isNumeric = long.TryParse(searchKeyword, out longSearchKeyword);

                switch (searchBy)
                {
                    case "ICCID_1":
                        records = records.Where(r => r.ICCID_1.Contains(searchKeyword));
                        break;

                    case "ICCID_2":
                        records = records.Where(r => r.ICCID_2.Contains(searchKeyword));
                        break;

                    case "IMSI_1":
                        if (isNumeric)
                        {
                            records = records.Where(r => r.IMSI_1 == longSearchKeyword);
                        }
                        else
                        {
                            records = Enumerable.Empty<Internal_Sims_Activations>().AsQueryable();
                        }
                        break;

                    case "MSISDN_1":
                        if (isNumeric)
                        {
                            records = records.Where(r => r.MSISDN_1 == longSearchKeyword);
                        }
                        else
                        {
                            records = Enumerable.Empty<Internal_Sims_Activations>().AsQueryable();
                        }
                        break;

                    case "ICCID_1_Network":
                        records = records.Where(r => r.ICCID_1_Network.Contains(searchKeyword));
                        break;

                    case "ICCID_2_Network":
                        records = records.Where(r => r.ICCID_2_Network.Contains(searchKeyword));
                        break;

                    case "ESN":
                        records = records.Where(r => r.ESN.Contains(searchKeyword));
                        break;

                    case "BootstrapActivationStartDate":
                        if (DateTime.TryParse(searchKeyword, out DateTime startDate))
                        {
                            records = records.Where(r => r.BootstrapActivationStartDate.HasValue && r.BootstrapActivationStartDate.Value.Date == startDate.Date);
                        }
                        else
                        {
                            records = Enumerable.Empty<Internal_Sims_Activations>().AsQueryable();
                        }
                        break;

                    case "BootstrapActivationEndDate":
                        if (DateTime.TryParse(searchKeyword, out DateTime endDate))
                        {
                            records = records.Where(r => r.BootstrapActivationEndDate.HasValue && r.BootstrapActivationEndDate.Value.Date == endDate.Date);
                        }
                        else
                        {
                            records = Enumerable.Empty<Internal_Sims_Activations>().AsQueryable();
                        }
                        break;

                    case "AllocatedToInHouseDate":
                        if (DateTime.TryParse(searchKeyword, out DateTime allocatedDate))
                        {
                            records = records.Where(r => r.AllocatedToInHouseDate.HasValue && r.AllocatedToInHouseDate.Value.Date == allocatedDate.Date);
                        }
                        else
                        {
                            records = Enumerable.Empty<Internal_Sims_Activations>().AsQueryable();
                        }
                        break;

                    case "APN_Name":
                        records = records.Where(r => r.APN_Name.Contains(searchKeyword));
                        break;

                    case "IP_1":
                        records = records.Where(r => r.IP_1.Contains(searchKeyword));
                        break;

                    case "IP_2":
                        records = records.Where(r => r.IP_2.Contains(searchKeyword));
                        break;

                    case "IP_3":
                        records = records.Where(r => r.IP_3.Contains(searchKeyword));
                        break;

                    case "IP_4":
                        records = records.Where(r => r.IP_4.Contains(searchKeyword));
                        break;

                    case "MN_1":
                        records = records.Where(r => r.MN_1.Contains(searchKeyword));
                        break;

                    case "MN_2":
                        records = records.Where(r => r.MN_2.Contains(searchKeyword));
                        break;

                    case "MN_3":
                        records = records.Where(r => r.MN_3.Contains(searchKeyword));
                        break;

                    case "MN_4":
                        records = records.Where(r => r.MN_4.Contains(searchKeyword));
                        break;

                    case "IMEI":
                        records = records.Where(r => r.IMEI.Contains(searchKeyword));
                        break;

                    case "For_User":
                        records = records.Where(r => r.For_User.Contains(searchKeyword));
                        break;

                    case "For_State":
                        records = records.Where(r => r.For_State.Contains(searchKeyword));
                        break;

                    case "Dispatch_Date":
                        if (DateTime.TryParse(searchKeyword, out DateTime dispatchDate))
                        {
                            records = records.Where(r => r.Dispatch_Date.HasValue && r.Dispatch_Date.Value.Date == dispatchDate.Date);
                        }
                        else
                        {
                            records = Enumerable.Empty<Internal_Sims_Activations>().AsQueryable();
                        }
                        break;

                    case "Dispatch_Location":
                        records = records.Where(r => r.Dispatch_Location.Contains(searchKeyword));
                        break;

                    default:
                        records = records.Where(r => r.ICCID_1.Contains(searchKeyword) ||
                                                     r.ICCID_2.Contains(searchKeyword) ||
                                                     r.ICCID_1_Network.Contains(searchKeyword) ||
                                                     r.ICCID_2_Network.Contains(searchKeyword) ||
                                                     r.ESN.Contains(searchKeyword) ||
                                                     (isNumeric ? r.MSISDN_1 == longSearchKeyword : false) ||
                                                     (isNumeric ? r.IMSI_1 == longSearchKeyword : false));
                        break;
                }
            }

            return View(records.ToList());
        }








        // GET: InternalSimsActivations/Details/5
        public async Task<ActionResult> Details(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Internal_Sims_Activations internal_Sims_Activations = await db.Internal_Sims_Activations.FindAsync(id);
            if (internal_Sims_Activations == null)
            {
                return HttpNotFound();
            }
            return View(internal_Sims_Activations);
        }

        // GET: InternalSimsActivations/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: InternalSimsActivations/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Create([Bind(Include = "ICCID_1,ICCID_2,ICCID_1_Network,ICCID_2_Network,IMSI_1,IMSI_2,MSISDN_1,MSISDN_2,ESN,BootstrapActivationStartDate,BootstrapActivationEndDate,AllocatedToInHouseDate,APN_Name,IP_1,IP_2,IP_3,IP_4,MN_1,MN_2,MN_3,MN_4,IMEI,For_User,For_State,Dispatch_Date,Dispatch_Location")] Internal_Sims_Activations internal_Sims_Activations)
        {
            if (ModelState.IsValid)
            {
                db.Internal_Sims_Activations.Add(internal_Sims_Activations);
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }

            return View(internal_Sims_Activations);
        }

        // GET: InternalSimsActivations/Edit/5
        public async Task<ActionResult> Edit(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Internal_Sims_Activations internal_Sims_Activations = await db.Internal_Sims_Activations.FindAsync(id);
            if (internal_Sims_Activations == null)
            {
                return HttpNotFound();
            }
            return View(internal_Sims_Activations);
        }

        // POST: InternalSimsActivations/Edit/5
       
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> Edit([Bind(Include = "ICCID_1,ICCID_2,ICCID_1_Network,ICCID_2_Network,IMSI_1,IMSI_2,MSISDN_1,MSISDN_2,ESN,BootstrapActivationStartDate,BootstrapActivationEndDate,AllocatedToInHouseDate,APN_Name,IP_1,IP_2,IP_3,IP_4,MN_1,MN_2,MN_3,MN_4,IMEI,For_User,For_State,Dispatch_Date,Dispatch_Location")] Internal_Sims_Activations internal_Sims_Activations)
        {
            if (ModelState.IsValid)
            {
                db.Entry(internal_Sims_Activations).State = EntityState.Modified;
                await db.SaveChangesAsync();
                return RedirectToAction("Index");
            }
            return View(internal_Sims_Activations);
        }

        // GET: InternalSimsActivations/Delete/5
        public async Task<ActionResult> Delete(string id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Internal_Sims_Activations internal_Sims_Activations = await db.Internal_Sims_Activations.FindAsync(id);
            if (internal_Sims_Activations == null)
            {
                return HttpNotFound();
            }
            return View(internal_Sims_Activations);
        }










        // POST: InternalSimsActivations/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public async Task<ActionResult> DeleteConfirmed(string id)
        {
            Internal_Sims_Activations internal_Sims_Activations = await db.Internal_Sims_Activations.FindAsync(id);
            db.Internal_Sims_Activations.Remove(internal_Sims_Activations);
            await db.SaveChangesAsync();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
