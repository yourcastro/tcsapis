const https = require('https');
const agent = new https.Agent({
  rejectUnauthorized: false // Disable SSL verification
});
