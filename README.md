
module.exports = function(app) {
  app.use(
    '/api',
    createProxyMiddleware({
      target: 'http://localhost:5000',  // Your API backend
      changeOrigin: true,
      secure: false,  // Disable SSL verification
    })
  );
};
