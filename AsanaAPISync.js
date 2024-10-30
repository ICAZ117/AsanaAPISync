require('dotenv').config();

function main() {
    // Get API key from local .env file
    const apiKey = process.env.API_KEY;
    console.log(apiKey);
}

main();