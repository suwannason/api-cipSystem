
using Microsoft.EntityFrameworkCore;
using cip_api.models;
public class Database : DbContext {
    public Database(DbContextOptions<Database> options) : base(options) { }

    public DbSet<cipSchema> CIP { get; set; }
    public DbSet<cipUpdateSchema> CIP_UPDATE { get; set; }
    public DbSet<ApprovalSchema> APPROVAL { get; set; }
    public DbSet<userSchema> USERS { get; set; }
    public DbSet<NotificationSchema> NOTIFICATIONS { get; set; }
}