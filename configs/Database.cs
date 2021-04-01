
using Microsoft.EntityFrameworkCore;
using cip_api.models;
public class Database : DbContext {
    public Database(DbContextOptions<Database> options) : base(options) { }

    public DbSet<cipSchema> CIP { get; set; }
}