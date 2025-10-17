const express = require('express');
const multer = require('multer');
const handler = require('../server');

// Vercel will call default export as a serverless function
module.exports = handler;


