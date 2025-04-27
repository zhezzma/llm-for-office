// Test file to verify that environment variables are accessible
require('dotenv').config();

console.log('Testing environment variables from .env file:');
console.log('PRODUCTION_URL:', process.env.PRODUCTION_URL);
