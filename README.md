var builder = WebApplication.CreateBuilder(args);

// Enable CORS to allow requests from localhost:3000
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowSpecificOrigins", policy =>
    {
        policy.WithOrigins("http://localhost:3000")  // Replace with your React app's origin
              .AllowAnyMethod()                      // Allows all HTTP methods (GET, POST, etc.)
              .AllowAnyHeader()                      // Allows custom headers (e.g., Content-Type)
              .AllowCredentials();                   // Allows credentials (cookies, authentication)
    });
});

builder.Services.AddControllers().AddJsonOptions(options =>
{
    options.JsonSerializerOptions.PropertyNamingPolicy = null;
});

var app = builder.Build();

// Use CORS with the configured policy
app.UseCors("AllowSpecificOrigins");

app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();

app.Run();

