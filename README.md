System.InvalidOperationException: The exception handler configured on ExceptionHandlerOptions produced a 404 status response. This InvalidOperationException containing the original exception was thrown since this is often due to a misconfigured ExceptionHandlingPath. If the exception handler is expected to return 404 status responses then set AllowStatusCode404Response to true.
 ---> System.FormatException: The format of value 'application/json;odata=verbose' is invalid.
   at System.Net.Http.Headers.MediaTypeHeaderValue.CheckMediaTypeFormat(String mediaType, String parameterName)
   at System.Net.Http.StringContent..ctor(String content, Encoding encoding, String mediaType)
   at RestSharp.RequestContent.Serialize(BodyParameter body)
   at RestSharp.RequestContent.AddBody(Boolean hasPostParameters, BodyParameter bodyParameter)
   at RestSharp.RequestContent.BuildContent()
   at RestSharp.RestClient.ExecuteRequestAsync(RestRequest request, CancellationToken cancellationToken)
   at RestSharp.RestClient.ExecuteAsync(RestRequest request, CancellationToken cancellationToken)
   at RestSharp.AsyncHelpers.<>c__DisplayClass1_0`1.<<RunSync>b__0>d.MoveNext()
