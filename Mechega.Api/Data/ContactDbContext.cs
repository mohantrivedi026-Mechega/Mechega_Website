using Microsoft.EntityFrameworkCore;
using Mechega.Api.Models;

namespace Mechega.Api.Data
{
    public class ContactDbContext : DbContext
    {
        public ContactDbContext(DbContextOptions<ContactDbContext> options) : base(options) { }

        public DbSet<ContactMessage> ContactMessages => Set<ContactMessage>();
    }
}
