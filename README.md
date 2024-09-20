const stringValue = "Hello, .NET 8.0!";

fetch('https://localhost:5001/api/mycontroller/post-string', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
  },
  body: JSON.stringify(stringValue)  // String is sent as a JSON string
})
  .then(response => response.json())
  .then(data => console.log(data))
  .catch(error => console.error('Error:', error));
