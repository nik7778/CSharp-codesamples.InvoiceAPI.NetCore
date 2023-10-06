using MongoDB.Driver;
using Invoicing.Data;
using Invoicing.Data.Entities;

namespace Invoicing.Infrastructure.Data.Repositories
{
    public class InvoiceRepository : BaseMongoDbRepository<Invoice>
    {
        private const string CollectionName = "Invoices";
        public InvoiceRepository(MongoDbContext context) : base(context)
        {
        }
        protected override IMongoCollection<Invoice> Collection => _context.MongoDatabase.GetCollection<Invoice>(CollectionName);
    }
}
