using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using OpsAccountingWF.DataModel;

public static class SeedRoles
{
    public static void Initialize(IServiceProvider serviceProvider)
    {
        using (var context = new ApplicationDbContext(serviceProvider.GetRequiredService<DbContextOptions<ApplicationDbContext>>()))    
        {
            string[] roles = new string[] { "Administrator", "Manager","User"};

            var newrolelist = new List<IdentityRole>();
            foreach (string role in roles)
            {
                if (!context.Roles.Any(r => r.Name == role))
                {
                    newrolelist.Add(new IdentityRole(role));
                }
            }
            context.Roles.AddRange(newrolelist);
            context.SaveChanges();



        }
    }
}