using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;

public class CustomAuthorizationAttribute : Attribute, IAuthorizationFilter
{
    private readonly string _role;

    public CustomAuthorizationAttribute(string role)
    {
        _role = role;
    }

    public void OnAuthorization(AuthorizationFilterContext context)
    {
        var userName = context.HttpContext.User.Identity?.Name;

        if (string.IsNullOrEmpty(userName))
        {
            context.Result = new UnauthorizedResult();
            return;
        }

        var dbContext = context.HttpContext.RequestServices.GetService<SecurityDbContext>();

        var user = dbContext.Security.FirstOrDefault(u => u.UserName == userName);

        if (user == null || user.Role != _role)
        {
            context.Result = new ForbidResult();
        }
    }
}
