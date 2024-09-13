using Microsoft.EntityFrameworkCore;

public class SecurityDbContext : DbContext
{
    public DbSet<UserRole> Security { get; set; }

    public SecurityDbContext(DbContextOptions<SecurityDbContext> options) : base(options) { }
}

public class UserRole
{
    public int Id { get; set; }
    public string UserName { get; set; }
    public string Role { get; set; }
}
