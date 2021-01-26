const { config } = require('dotenv');
const result = config();

if (!result.error) {
  Object.keys(result.parsed).forEach((key) => {
    const value = result.parsed[key];
    if (value) {
      process.env[key] = value;
    }
  });
}
