using AutoMapper;
using AutoMapper.QueryableExtensions;
using Common.Extensions;
using Common.Models;
using Common.Services;
using Common.ServicesResult;
using GoWebApp.Common.Enums;
using GoWebApp.Common.Services;
using Invoicing.Data.Entities;
using Invoicing.Infrastructure.Data.Repositories;
using Invoicing.Infrastructure.Enums;
using Invoicing.Infrastructure.Models;
using Logging.Infrastructure;
using Logging.Infrastructure.Models;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace Invoicing.Infrastructure.Services
{
    public class InvoiceService : BaseService
    {
        private InvoiceRepository _invoiceRepo;
        private ILogService _logService;

        public InvoiceService(InvoiceRepository invoiceRepo, ILogService logService)
        {
            _invoiceRepo = invoiceRepo;
            _logService = logService;
        }

        public DetailsInvoiceModel FindById(string id)
        {
            return Mapper.Map<DetailsInvoiceModel>(GetById(id));
        }

        public IQueryable<DetailsInvoiceModel> GetAllForTenant(string tenantId)
        {
            return GetAllActive(i => i.TenantId == tenantId).AsQueryable().ProjectTo<DetailsInvoiceModel>();
        }

        public IQueryable<DetailsInvoiceModel> GetAllForClients(string[] clients = null)
        {
            return GetAllActive(i => clients == null || clients.Contains(i.ClientId)).AsQueryable().ProjectTo<DetailsInvoiceModel>();
        }

        public IQueryable<DetailsInvoiceModel> GetAllForUsers(string[] users = null)
        {
            return GetAllActive(i => users == null || users.Contains(i.CreatedBy)).AsQueryable().ProjectTo<DetailsInvoiceModel>();
        }

        public IQueryable<DetailsInvoiceModel> GetAll(string userId = null, string companyId = null)
        {
            return GetAllActive(i => (userId == null || i.CreatedBy == userId) && (companyId == null || i.CompanyId == companyId)).AsQueryable().ProjectTo<DetailsInvoiceModel>();
        }

        public IQueryable<DetailsInvoiceModel> GetMyInvoices()
        {
            return GetAllActive(i => i.CreatedBy == currentUserId && (companyId == null || i.CompanyId == companyId)).AsQueryable().ProjectTo<DetailsInvoiceModel>();
        }

        public List<KeyValuePair<int, string>> GetInvoiceStatuses()
        {
            //return Enum.GetNames(typeof(InvoiceStatus)).Select(i => new KeyValuePair<string, string>(i, i.ToReadableText())).ToList();
            return Enum.GetValues(typeof(InvoiceStatus)).Cast<InvoiceStatus>().Select(i => new KeyValuePair<int, string>((int)i, i.ToString().ToReadableText())).ToList();
        }

        public int GetNextInvoiceNumber(string clientId)
        {
            var lastInvoice = GetAllForClients(new string[] { clientId }).Where(i => i.Status.ToLower() != InvoiceStatus.Draft.ToString().ToLower()).OrderByDescending(i => i.Number).FirstOrDefault();
            return lastInvoice == null ? 1 : lastInvoice.Number + 1;
        }

        public IEnumerable<DetailsLogModel> GetLogs(int count = 10)
        {
            return _logService.GetAll(null, ModulesEnum.Invoicing.ToString(), currentUserId, null, null, null, count);
        }

        public IOperationResult CreateInvoice(CreateInvoiceModel invoiceModel)
        {
            var result = new OperationResult(ValidateModel(invoiceModel));
            if (!result.IsSuccess)
            {
                return result;
            }

            invoiceModel.CreatedBy = currentUserId;
            invoiceModel.TenantId = tenantId;
            invoiceModel.CompanyId = companyId;
            // save invoice
            var invoice = Mapper.Map<Invoice>(invoiceModel);
            invoice.Status = InvoiceStatus.Draft.Value();
            _invoiceRepo.Add(invoice);
            _logService.LogSuccess(ModulesEnum.Invoicing.ToString(), UserActionEnum.Create_Invoice.ToString(), invoiceModel.CreatedBy, MessagesService.CreateSuccessMessage("Invoice"), null, invoice);
            return result.HasSucceeded(Mapper.Map<DetailsInvoiceModel>(invoice), MessagesService.CreateSuccessMessage("Invoice"));
        }

        public IOperationResult UpdateInvoice(UpdateInvoiceModel invoiceModel)
        {
            var result = new OperationResult(ValidateModel(invoiceModel));
            if (!result.IsSuccess)
            {
                return result;
            }

            var dbInvoice = GetById(invoiceModel.Id);
            if (dbInvoice == null)
            {
                return result.HasNotSucceeded(MessagesService.NotFoundMessage("Invoice"));
            }

            if (dbInvoice.Status > InvoiceStatus.Draft.Value())
            {
                return result.HasNotSucceeded(MessagesService.InvalidActionMessage("Update"));
            }

            dbInvoice = Mapper.Map(invoiceModel, dbInvoice);
            return UpdateInvoice(dbInvoice);
        }

        public IOperationResult UpdateStatus(string invoiceId, InvoiceStatus status, dynamic extraData = null)
        {
            return UpdateStatus(invoiceId, status.Value(), extraData);
        }             

        public IOperationResult UpdateStatus(string invoiceId, int status, dynamic extraData)
        {
            var dbInvoice = GetById(invoiceId);
            if (dbInvoice == null)
            {
                return OperationResult.NotSucceeded(MessagesService.NotFoundMessage("Invoice"));
            }
            if (dbInvoice.Status == status)
            {
                return OperationResult.Succeeded(Mapper.Map<DetailsInvoiceModel>(dbInvoice), MessagesService.UpdateSuccessMessage("Invoice Status"));
            }
            if (status == (int)InvoiceStatus.Storno)
            {
                return MakeStornoInvoiceFrom(dbInvoice);
            }
            if (status == (int)InvoiceStatus.PartialStorno)
            {
                var items = extraData.stornedItems.ToObject<List<string>>();
                if (items == null)
                    return OperationResult.NotSucceeded("Products not selected!");

                return MakeStornoInvoiceFrom(dbInvoice, items);
            }

            dbInvoice.Status = status;
            return UpdateInvoice(dbInvoice);
        }

        public IOperationResult ActivateInvoice(string invoiceId)
        {
            var dbInvoice = GetById(invoiceId);
            if (dbInvoice == null)
            {
                return OperationResult.NotSucceeded(MessagesService.NotFoundMessage("Invoice"));
            }

            var lastInvoice = GetAll(null, dbInvoice.CompanyId).Where(i => i.Status != InvoiceStatus.Draft.ToString()).OrderByDescending(i => i.Number).FirstOrDefault();
            dbInvoice.Number = lastInvoice == null ? 1 : lastInvoice.Number + 1;
            dbInvoice.Status = InvoiceStatus.Active.Value();
            return UpdateInvoice(dbInvoice);
        }

        public IOperationResult SaveRepetitiveData(string invoiceId, RepetitiveDataModel repetitiveData)
        {
            var dbInvoice = GetById(invoiceId);
            if (dbInvoice == null)
            {
                return OperationResult.NotSucceeded(MessagesService.NotFoundMessage("Invoice"));
            }

            if (dbInvoice.CompanyId != companyId) return OperationResult.NotSucceeded(MessagesService.NotAuthorizedFor());

            dbInvoice.RepetitiveData = Mapper.Map<RepetitiveData>(repetitiveData);
            return UpdateInvoice(dbInvoice, MessagesService.UpdateSuccessMessage($"Repetitive data for Invoice ({dbInvoice.Number})"));
        }

        public IOperationResult DeleteInvoice(string id)
        {
            var invoice = GetById(id);
            if (invoice != null)
            {
                _invoiceRepo.Delete(invoice);
                _logService.LogSuccess(ModulesEnum.Invoicing.ToString(), UserActionEnum.Delete_Invoice.ToString(), invoice.CreatedBy, MessagesService.DeleteSuccessMessage("Invoice"));
            }
            return OperationResult.Succeeded(MessagesService.DeleteSuccessMessage("Invoice"));
        }

        public IOperationResult ExportToExcel(List<DetailsInvoiceModel> model)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                excel.Workbook.Properties.Author = "Go Web App";
                excel.Workbook.Properties.Title = "Invoices List";
                var sheet = excel.Workbook.Worksheets.Add("Invoices List");

                var headerColor = System.Drawing.Color.FromArgb(34, 112, 161);
                var cellsColumns = sheet.Cells["A2:G2"];
                cellsColumns.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cellsColumns.Style.Fill.BackgroundColor.SetColor(headerColor);
                cellsColumns.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.White);
                cellsColumns.Style.Font.Color.SetColor(System.Drawing.Color.White);
                cellsColumns.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                cellsColumns.Style.Border.Left.Color.SetColor(System.Drawing.Color.White);
                cellsColumns.Style.Font.Bold = true;
                cellsColumns.Style.Font.Size = 13;
                cellsColumns.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                cellsColumns.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                sheet.Cells[1, 1].Value = "Invoices List " + DateTime.Now.ToString("yyyy-MM-dd");
                sheet.Cells[1, 1].Style.Font.Bold = true;
                sheet.Cells[1, 1].Style.Font.Size = 15;

                var rowDefaultIndex = 2;
                var colDefaultIndex = 1;
                var col = colDefaultIndex;
                var counter = 0;
                sheet.Cells[rowDefaultIndex, col++].Value = "No.";
                sheet.Cells[rowDefaultIndex, col++].Value = "Invoice";
                sheet.Cells[rowDefaultIndex, col++].Value = "Client";
                sheet.Cells[rowDefaultIndex, col++].Value = "Amount";
                sheet.Cells[rowDefaultIndex, col++].Value = "Date";
                sheet.Cells[rowDefaultIndex, col++].Value = "Due Date";
                sheet.Cells[rowDefaultIndex, col++].Value = "Status";

                foreach (var item in model)
                {
                    rowDefaultIndex++;
                    col = colDefaultIndex;
                    sheet.Cells[rowDefaultIndex, col++].Value = ++counter;
                    sheet.Cells[rowDefaultIndex, col++].Value = item.Name +" -" + item.Serie + " #" + item.Number;
                    sheet.Cells[rowDefaultIndex, col++].Value = item.Client.Name;
                    sheet.Cells[rowDefaultIndex, col++].Value = string.Format("{0:0.00} {1}", item.TotalAmount, item.Currency.BaseCurrency);
                    sheet.Cells[rowDefaultIndex, col++].Value = item.Date.ToString("yyyy-MM-dd");
                    sheet.Cells[rowDefaultIndex, col++].Value = item.DueDate.ToString("yyyy-MM-dd");
                    sheet.Cells[rowDefaultIndex, col++].Value = item.StatusText;
                }

                //sheet.Cells.AutoFitColumns();
                MemoryStream excelStream = new MemoryStream();
                excel.SaveAs(excelStream);
                return new OperationResult().HasSucceeded(excelStream, "Succesful exported to excel");
            }
        }

        #region Private Api
              
        private IOperationResult MakeStornoInvoiceFrom(Invoice dbInvoice, List<string> items = null)
        {
            var isPartialStorno = items != null;
            // differantiate between storno or partial storno
            if (isPartialStorno)
            {
                var stornedItems = dbInvoice.Items.Where(i => items.Contains(i.ItemId));
                dbInvoice.Items = stornedItems;
            }

            var stornoCreationResult = CreateStornoInvoice(dbInvoice, isPartialStorno);
            if (!stornoCreationResult.IsSuccess) return stornoCreationResult;

            var createdInvoice = (DetailsInvoiceModel)stornoCreationResult.Data;

            // update current invoice
            dbInvoice.Status = isPartialStorno ? InvoiceStatus.Active.Value() : InvoiceStatus.Paid.Value();
            dbInvoice.RelatedToInvoiceId = createdInvoice.Id;
            dbInvoice.RelatedToInvoiceMentions = $"Was storned in '{DateTime.Now.ToShortDateString()}'. Storno invoice number: #{createdInvoice.Number}";

            return UpdateInvoice(dbInvoice);
        }
             

        private IOperationResult CreateStornoInvoice(Invoice invoice, bool isPartial = false)
        {
            // verify time
            if (invoice.Status != InvoiceStatus.Active.Value()) return OperationResult.NotSucceeded("Only active invoices can be storned");

            // process storno invoice
            var newInvoice = Mapper.Map<CreateInvoiceModel>(invoice);
            newInvoice.Items = newInvoice.Items.Select(i =>
            {
                i.Quantity = i.Quantity * -1;
                return i;
            }).ToList();

            newInvoice.Date = DateTime.Now;
            newInvoice.Name += isPartial ? " - Partial Storno" : " - Storno";

            var createResult = CreateInvoice(newInvoice);
            if (!createResult.IsSuccess)
            {
                return createResult;
            }
            var createdInvoice = ((DetailsInvoiceModel)createResult.Data);
            // activate invoice
            createdInvoice = (DetailsInvoiceModel)ActivateInvoice(createdInvoice.Id).Data;
            var toUpdateInvoice = Mapper.Map<Invoice>(createdInvoice);

            toUpdateInvoice.Status = isPartial ? InvoiceStatus.PartialStorno.Value() : InvoiceStatus.Storno.Value();
            toUpdateInvoice.RelatedToInvoiceId = invoice.Id;
            toUpdateInvoice.RelatedToInvoiceMentions = isPartial ? "Partial storno for invoice #" + invoice.Number : "Storno for invoice #" + invoice.Number;

            return UpdateInvoice(toUpdateInvoice);
        }

        private IEnumerable<Invoice> GetAllActive(Expression<Func<Invoice, bool>> expression = null)
        {
            var result = _invoiceRepo.Where(u => u.IsDeleted != true);
            var results = result.ToList().AsQueryable();
            if (expression != null)
            {
                results = results.Where(expression);
            }

            return results;
        }
        private Invoice GetById(string id)
        {
            return GetAllActive(c => c.Id == id).FirstOrDefault();
        }

        private IOperationResult UpdateInvoice(Invoice invoice, string logMessageDetails = null)
        {
            _invoiceRepo.Update(invoice);
            _logService.LogSuccess(ModulesEnum.Invoicing.ToString(), UserActionEnum.Update_Invoice.ToString(), currentUserId, MessagesService.UpdateSuccessMessage("Invoice"), null, invoice);
            return new OperationResult().HasSucceeded(Mapper.Map<DetailsInvoiceModel>(invoice), MessagesService.UpdateSuccessMessage(logMessageDetails ?? "Invoice"));
        }

        public override ICollection<ValidationResult> ValidateModel(BaseModel @object)
        {
            CreateInvoiceModel model = (CreateInvoiceModel)@object;
            var results = base.ValidateModel(model);

            if (model.Company.Id != companyId)
            {
                results.Add(new ValidationResult("Company of invoice differs from selected company!"));
            }

            if (model.Items == null || model.Items.Count() <= 0)
            {
                results.Add(new ValidationResult("Invoice cannot be created without items."));
            }

            return results;
        }
        #endregion
    }
}
