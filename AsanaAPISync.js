require('dotenv').config();
const reader = require('xlsx');
const file = reader.readFile('Workbook.xlsx');

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

    const targetSheet = "SheetB";
    const sheets = file.SheetNames;

    // Iterate through all sheets in the workbook
    for (let i = 0; i < sheets.length; i++) {
        // Check if the current sheet is the target sheet
        if (sheets[i] === targetSheet) {
            debug("Sheet", sheets[i]);
            const temp = reader.utils.sheet_to_json(
                file.Sheets[file.SheetNames[i]]
            );

            // Iterate through each row
            for (let j = 0; j < temp.length; j++) {
                debug("Row", temp[j]);

                // Get project, date, and notes from the row
                const project = temp[j].Project;
                const date = temp[j].Date;
                const notes = temp[j].Notes;

                debug("Project", project);
                debug("Date", date);
                debug("Notes", notes);
                debug("NEXT ROW\n\n");
            }
        }
    }
}

main();