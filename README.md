 // Initialize RestSharp client
 var client = new RestClient(apiUrl);
 var request = new RestRequest(apiUrl, Method.Get);

 // Add headers (adjust the authentication method as needed)
 request.AddHeader("Accept", "application/json;odata=nometadata");
 request.AddHeader("Authorization", "Bearer " + authToken.access_token); // Adjust for your authentication

 // Execute the request and get the response
 var response = client.Execute(request);

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
