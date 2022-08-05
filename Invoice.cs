using Common.Models;
using GoWebApp.Common.Data;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Invoicing.Data.Entities
{
    [BsonIgnoreExtraElements]
    public class Invoice : TenantEntity
    {
        public string Name { get; set; }
        public string CompanyId { get; set; }
        public CompanyData Company { get; set; }
        public string ClientId { get; set; }
        public CompanyData Client { get; set; }
        public string ClientBankAccount { get; set; }
        public string Description { get; set; }
        public string Serie { get; set; }
        public int Number { get; set; }
        public DateTime Date { get; set; }
        public DateTime DueDate
        {
            get
            {
                int paymentTerm = 1;
                int.TryParse(PaymentTerm, out paymentTerm);
                return Date.AddDays(paymentTerm);
            }
        }
        public string Language { get; set; }
        public CurrencyDetails Currency { get; set; }
        public string PaymentTerm { get; set; }
        public string Note { get; set; }
        public string Type { get; set; }
        public int Status { get; set; }

        public decimal SubTotalAmount
        {
            get
            {
                return Items.Sum(i => i.TotalAmount);
            }
        }
        public decimal TotalAmount
        {
            get
            {
                var taxes = new
                {
                    percentage = Taxes.Where(i => i.IsPercentage == true).Sum(i => i.Value) / 100,
                    value = Taxes.Where(i => i.IsPercentage != true).Sum(i => i.Value)
                };
                var discounts = new
                {
                    percentage = Discounts.Where(i => i.IsPercentage == true).Sum(i => i.Value) / 100,
                    value = Discounts.Where(i => i.IsPercentage != true).Sum(i => i.Value)
                };

                return SubTotalAmount * (1 + taxes.percentage) * (1 - discounts.percentage) + taxes.value - discounts.value;
            }
        }

        public InvoiceExtendedInfo[] Taxes { get; set; } = new InvoiceExtendedInfo[] { };
        public InvoiceExtendedInfo[] Discounts { get; set; } = new InvoiceExtendedInfo[] { };

        public IEnumerable<InvoiceItem> Items { get; set; } = new List<InvoiceItem>();

        public string Template { get; set; }

        public string RelatedToInvoiceId { get; set; }
        public string RelatedToInvoiceMentions { get; set; }

        // repetitive invoice
        public RepetitiveData RepetitiveData { get; set; }

        public string SplitVatBankAccounts { get; set; }
        public bool SplitVat { get; set; }
        public string ClientLegalRepresentative { get; set; }
        public string UserLegalRepresentative { get; set; }
        public DateTime DeliveryDate { get; set; }
        public string MeansOfTransportations { get; set; }
        public DateTime AdvancePaymentDate { get; set; }
        public int AdvancePaymentAmount { get; set; }
        public bool ReverseTaxes { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class RepetitiveData
    {
        public bool IsActive { get; set; }
        public int Days { get; set; }
        public DateTime StartOn { get; set; } = DateTime.Now;
        public DateTime EndOn { get; set; } = DateTime.Now;
    }

    [BsonIgnoreExtraElements]
    public class InvoiceItem
    {
        public int NrCrt { get; set; }
        public string ItemId { get; set; }
        public decimal VATTaxValue { get; set; } = 0;
        public decimal OtherTaxValue { get; set; } = 0;
        public decimal DiscountValue { get; set; } = 0;
        public string Description { get; set; }
        public int Quantity { get; set; } = 0;
        public decimal Price { get; set; } = 0;
        public decimal Amount { get { return Quantity * Price; } }
        public decimal VATAmount { get { return Amount * VATTaxValue / 100; } }
        public decimal DiscountAmount { get { return Amount * DiscountValue / 100; } }
        public decimal OtherTaxesAmount { get { return Amount * OtherTaxValue / 100; } }
        public decimal TotalAmount { get { return Amount * (1 + VATTaxValue / 100) * (1 + OtherTaxValue / 100) * (1 - DiscountValue / 100); } }
    }

    public class InvoiceExtendedInfo
    {
        public string Name { get; set; }
        public decimal Value { get; set; }
        public bool IsPercentage { get; set; }
    }

    public class CompanyData : TenantEntity
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string PostalAddress { get; set; }
        public string ZipCode { get; set; }
        public string City { get; set; }
        public string AreaOrState { get; set; }
        public string Country { get; set; }
        public string Phone { get; set; }
        public string Website { get; set; }

        // Company Details
        public string Cui { get; set; }
        public string NrRegCom { get; set; }
        public bool IsIndividual { get; set; }
        public string VatCode { get; set; }
        public string EuVatCode { get; set; }
        public bool UsingVAT { get; set; }
        public bool IsFaded { get; set; }
        public string Status { get; set; }
        public string AuthorityData { get; set; }
        public string AuthorityBalances { get; set; }
        public IEnumerable<VatDetails> VatAtPayment { get; set; }
        public bool IsUsingVatAtPayment { get; set; }

        public string SubscribedRevenue { get; set; }
        public string SubscribedRevenueCurrency { get; set; }
        public string CapitalizedRevenue { get; set; }
        public string CapitalizedRevenueCurrency { get; set; }

        public IEnumerable<BankDetails> BankAccounts { get; set; } = new List<BankDetails>();
    }

    public class BankDetails
    {
        public string BankName { get; set; }
        public string BankAccount { get; set; }
        public string BankAccountCurrency { get; set; }
        public string SwiftBicCode { get; set; }
    }

    public class VatDetails
    {
        public string Tip_Descriere { get; set; }
        public string Tip { get; set; }
        public string Publicare { get; set; }
        public string Actualizare { get; set; }
        public string De_La { get; set; }
        public string Pana_La { get; set; }
    }

    public class CurrencyDetails
    {
        private string _selected;
        public string Selected
        {
            get
            {
                return _selected ?? BaseCurrency;
            }
            set
            {
                _selected = value;
            }
        }

        public string BaseCurrency { get; set; } = "";
        public double ExchangeRate { get; set; } // 1 selected = xxx base currency
    }
}
