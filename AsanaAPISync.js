require('dotenv').config();

const doDebug = true;

function debug(text, object) {
    if (doDebug) {
        console.log(text || "", object || "");
    }
}

async function createComment(task_gid, text, apiKey) {
    const axios = require('axios');
    let data = JSON.stringify({
        "data": {
            "text": text
        }
    });

    let config = {
        method: 'post',
        maxBodyLength: Infinity,
        url: `https://app.asana.com/api/1.0/tasks/${task_gid}/stories`,
        headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`
        },
        data: data
    };

    axios.request(config)
        .then((response) => {
            debug(JSON.stringify(response.data));
        })
        .catch((error) => {
            debug(error);
        });
}

async function main() {
    // Get API key from local .env file
    const apiKey = process.env.API_KEY;
    debug("API Key", apiKey);

    await createComment(1207196044602694, "AsanaAPISync_v0.0.1alpha: API Test", apiKey);
}

main();