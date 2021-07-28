using Microsoft.EntityFrameworkCore;

namespace ExcelSheetUsingSyncfusionXlsIO.Models
{
    public class ExcelSheetDBContext : DbContext
    {
        public ExcelSheetDBContext(DbContextOptions<ExcelSheetDBContext> options)
            : base(options)
        {
        }

        //entities
        public DbSet<Product> Products { get; set; }
    }
}
