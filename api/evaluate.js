const app = require('../server');

// Vercel serverless handler – delegate to Express app
module.exports = (req, res) => app(req, res);


