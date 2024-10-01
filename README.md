// Create a new HttpClientHandler to configure the proxy
var httpClientHandler = new HttpClientHandler
{
    Proxy = new WebProxy("http://yourproxyaddress:port")
    {
        Credentials = new NetworkCredential("proxyUsername", "proxyPassword")
    },
    UseProxy = true
};

// Pass the handler to the RestSharp client
var client = new RestClient(new HttpClient(httpClientHandler))
{
    BaseUrl = new Uri(apiUrl)
};

var request = new RestRequest(apiUrl, Method.Get);

// Add headers (adjust the authentication method as needed)
request.AddHeader("Accept", "application/json;odata=nometadata");
request.AddHeader("Authorization", "Bearer " + authToken.access_token); // Adjust for your authentication

// Execute the request and get the response
var response = await client.ExecuteAsync(request);

if (response.IsSuccessful)
{
    Console.WriteLine("SP User Group retrieved successfully:");
    var resultSet = JsonConvert.DeserializeObject<SPUserDetails>(response.Content);
    var result = resultSet?.value?.Where(item => item?.Email?.ToLower() == email.ToLower()).First();
    Console.WriteLine("SP User Group :" + result.Id);
    spUserSiteID = result.Id;
}
else
{
    Console.WriteLine($"Failed to retrieve SP Group: {response?.StatusCode} - {response?.ErrorMessage}");
}
