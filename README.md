private readonly AppSettings _appSettings;

public SPHelper(IOptions<AppSettings> appSettings)
{
    _appSettings = appSettings.Value;
}
