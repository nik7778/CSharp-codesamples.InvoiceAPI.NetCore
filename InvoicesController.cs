using GoWebApp.Common.Enums;
using Invoicing.Api.Infrastructure.Filters;
using Invoicing.Api.Infrastructure.Models;
using Invoicing.Infrastructure.Enums;
using Invoicing.Infrastructure.Models;
using Invoicing.Infrastructure.Services;
using Microsoft.AspNetCore.Mvc;
using MvcCore.Common;
using MvcCore.Common.Controllers;
using MvcCore.Common.Filters;
using System;
using System.IO;
using System.Linq;

namespace Invoicing.Api.Controllers
{
    [Route("invoicing/[controller]")]
    public class InvoicesController : BaseController
    {
        private InvoiceService _invoiceService;

        public InvoicesController(InvoiceService service)
        {
            _invoiceService = service;
        }

        private InvoiceService Invoices
        {
            get
            {
                return (InvoiceService)_invoiceService.SetDependencies(AppAccess.SelectedCompany, AppAccess.UserId, AppAccess.TenantId);
            }
        }

        private IQueryable<DetailsInvoiceModel> FilterInvoices(FilterInvoiceModel filter = null)
        {
            IQueryable<DetailsInvoiceModel> result = null;
            if (AppAccess.IsTpcAdministrator)
            {
                result = Invoices.GetAll();
            }
            else
            {
                var userId = filter.OnlyMyInvoices ? AppAccess.UserId : null;
                result = Invoices.GetAll(userId, AppAccess.SelectedCompany);
            }
            if (filter != null)
            {
                result = result.Where(i =>
                i.Date.Date >= filter.StartDate.Date
                &&
                i.Date.Date <= filter.EndDate.Date);
            }
            return result;
        }

        [HttpGet]
        public IActionResult Get([FromQuery]FilterInvoiceModel model = null)
        {
            return MyInvoices(model);
        }

        [HttpGet]
        [Route("MyInvoices")]
        public IActionResult MyInvoices([FromQuery]FilterInvoiceModel model = null)
        {
            return new ApiResponseResult(FilterInvoices(model));
        }

        [HttpPost("ExportInvoicesToExcel")]
        public ActionResult ExportFilteredInvoicesToExcel([FromBody] FIlteredInvoicesModel value)
        {
            var result = _invoiceService.ExportToExcel(value.InvoiceList);
            var excelResult = new ExcelResult((MemoryStream)result.Data, "InvoiceList-" + DateTime.Now.ToString("yyyy-MM-dd"));

            return excelResult;
        }

        [HttpGet]
        [Route("ForClients/{clients}")]
        public IActionResult ForClient(string clients)
        {
            var clientsSplit = clients.Split(',', System.StringSplitOptions.RemoveEmptyEntries);
            var result = Invoices.GetAllForClients(clientsSplit);
            return new ApiResponseResult(result);
        }

        [HttpGet("Statuses")]
        public IActionResult Statuses()
        {
            return new ApiResponseResult(Invoices.GetInvoiceStatuses());
        }

        [HttpGet("{id}")]
        public IActionResult Get(string id)
        {
            return new ApiResponseResult(Invoices.FindById(id) ?? new CreateInvoiceModel());
        }

        [HttpPost]
        [ValidateModelActionFilter]
        [CheckActionLimitationFilter(UserActionEnum.Create_Invoice)]
        public IActionResult Post([FromBody]CreateInvoiceModel value)
        {
            return new ApiResponseResult(Invoices.CreateInvoice(value));
        }

        [HttpPut]
        [ValidateModelActionFilter]
        public IActionResult Put([FromBody]UpdateInvoiceModel value)
        {
            return new ApiResponseResult(Invoices.UpdateInvoice(value));
        }

        [HttpPost("{id}/Activate")]
        public IActionResult Activate([FromRoute]string id)
        {
            return new ApiResponseResult(Invoices.ActivateInvoice(id));
        }

        [HttpPost("{id}/RepetitiveData")]
        public IActionResult RepetitiveData([FromRoute]string id, [FromBody] RepetitiveDataModel repetitiveData)
        {
            return new ApiResponseResult(Invoices.SaveRepetitiveData(id, repetitiveData));
        }


        [HttpPost("{id}/SetStatus/{status}")]
        public IActionResult SetStatus([FromRoute]string id, [FromRoute]int status, [FromBody]dynamic extraData)
        {
            return new ApiResponseResult(Invoices.UpdateStatus(id, status, extraData));
        }

        [HttpDelete("{id}")]
        public IActionResult Delete(string id)
        {
            return new ApiResponseResult(Invoices.DeleteInvoice(id));
        }

        [HttpGet("Logs")]
        public IActionResult Logs([FromQuery]int count = 10)
        {
            return new ApiResponseResult(Invoices.GetLogs(count));
        }
    }
}
