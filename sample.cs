using Serilog;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Host.UseSerilog((context, config) =>
{
    config.WriteTo.File(
        path: @"C:\inetpub\logs\app\log-.txt",  // Directory on IIS server
        rollingInterval: RollingInterval.Day,   // Logs rotated daily
        restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Information, // Minimum log level
        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level}] {Message}{NewLine}{Exception}" // Log format
    );
});

// Other middleware configurations
builder.Services.AddControllers();

var app = builder.Build();

// Application middleware
app.UseSerilogRequestLogging();  // Enable request logging for HTTP requests

app.UseRouting();
app.UseEndpoints(endpoints =>
{
    endpoints.MapControllers();
});

app.Run();


#### 3. Set IIS Folder Permissions

- Ensure that the IIS app pool user has **write permissions** to the folder where logs will be saved. For instance, in the above code, the path is `C:\inetpub\logs\app\`.
- Right-click on the folder → Properties → Security → Edit permissions to include the app pool identity (e.g., `IIS_IUSRS`).

#### 4. Enable IIS Logs for Diagnostics (Optional)

IIS also provides its own logs. To enable IIS logging:
1. Open **IIS Manager**.
2. Select your site.
3. In the **Logging** section, configure the location and format of the IIS logs (default is `%SystemDrive%\inetpub\logs\LogFiles`).

#### 5. Adjust Logging Levels in `appsettings.json`

If you want more control over logging levels, configure Serilog settings in `appsettings.json`:

```json
{
  "Serilog": {
    "MinimumLevel": {
      "Default": "Information",
      "Override": {
        "Microsoft": "Warning",
        "System": "Warning"
      }
    },
    "WriteTo": [
      {
        "Name": "File",
        "Args": {
          "path": "C:\\inetpub\\logs\\app\\log-.txt",
          "rollingInterval": "Day"
        }
      }
    ]
  }
}
```

#### 6. Viewing Logs

The logs will be stored in the specified folder (`C:\inetpub\logs\app\`) with a filename like `log-2024-09-13.txt` (rotated daily). These logs will capture important details such as timestamps, log levels, messages, and exceptions.

### Summary
- Use **Serilog** to write application logs to a file.
- Ensure IIS folder permissions are configured correctly.
- Optionally, enable **IIS logs** for additional diagnostics.


