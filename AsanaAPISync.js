const reader = require('xlsx');

const doDebug = true;
const doTest = true;

// doc: asana
const taskConversion = {
    "Meeting: Intake": "Discovery Phase",
    "Meeting: Methods/Ideas": "Protocol Development",
    "Analysis": "Statsitical Analysis",
    "Products": "Publication",
    "Review/Revise Package": "IRB Package Preparation Phase",
    "SAP": "IRB Package Preparation Phase",
    "DRR": "IRB Package Preparation Phase",
    "Prep Work": "Statistical Analysis"
}

function debug(text, object, override = false) {
    if (doDebug || override) {
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

async function getAllProjects(apiKey) {
    const axios = require('axios');

    var res;

    let config = {
        method: 'get',
        maxBodyLength: Infinity,
        url: 'https://app.asana.com/api/1.0/workspaces/47107657393912/projects',
        headers: {
            'Authorization': `Bearer ${apiKey}`
        }
    };

    const axiosResponse = axios.request(config).catch((error) => {
        console.log(error);
    });

    await axiosResponse.then((response) => {
        res = response.data.data;
    });

    return res;
}

async function getTasksInProject(project_gid, apiKey) {
    const axios = require('axios');

    var res;

    let config = {
        method: 'get',
        maxBodyLength: Infinity,
        url: `https://app.asana.com/api/1.0/projects/${project_gid}/tasks`,
        headers: {
            'Authorization': `Bearer ${apiKey}`
        }
    };

    const axiosResponse = axios.request(config).catch((error) => {
        console.log(error);
    });

    await axiosResponse.then((response) => {
        res = response.data.data;
    });

    return res;
}

function printDetails(text) {
    console.log(`String: "${text}"`);

    // Basic details
    console.log(`Length: ${text.length}`);
    console.log(`Type: ${typeof text}`);

    // Unicode code points
    const unicodePoints = Array.from(text).map(char => char.codePointAt(0).toString(16).toUpperCase());
    console.log(`Unicode Code Points (Hex): ${unicodePoints.join(' ')}`);

    // Hexadecimal encoding
    const hexEncoding = Array.from(text).map(char => char.charCodeAt(0).toString(16).toUpperCase());
    console.log(`Hexadecimal Encoding: ${hexEncoding.join(' ')}`);

    // Escape sequences
    const escapeSequences = Array.from(text).map(char => `\\u${char.charCodeAt(0).toString(16).padStart(4, '0')}`);
    console.log(`Escape Sequences: ${escapeSequences.join(' ')}`);

    // Whitespace or blank characters analysis
    const whitespaceAnalysis = Array.from(text).map(char => {
        if (/\s/.test(char)) return `[Whitespace: "${char}"]`;
        return `[Char: "${char}"]`;
    });
    console.log(`Character Analysis: ${whitespaceAnalysis.join(' ')}`);

    // Normalize forms
    console.log(`NFC: ${text.normalize('NFC')}`);
    console.log(`NFD: ${text.normalize('NFD')}`);
    console.log(`NFKC: ${text.normalize('NFKC')}`);
    console.log(`NFKD: ${text.normalize('NFKD')}`);

    // Binary comparison of characters
    console.log(`Binary Representation: ${Array.from(text).map(char => char.charCodeAt(0).toString(2).padStart(16, '0')).join(' ')}`);
}

function cleanWhitespace(text) {
    return text ? text.replace(/\s/g, " ").trim() : text;
}

exports.asanaAPISync = async function (apiKey, file, targetSheet, startingRow = 2) {
    // async function asanaAPISync(apiKey, file, targetSheet, startingRow = 2) {
    debug("--------------- LAUNCHING ASANA API SYNC ---------------", "", doTest);
    debug("API Key:      ", apiKey, doTest);
    debug("Target Sheet: ", targetSheet, doTest);
    debug("Starting Row: ", startingRow, doTest);
    debug("--------------------------------------------------------", "", doTest);

    // Get all sheet names from the workbook
    const sheets = file.SheetNames;
    debug("\r\nSheets", sheets);

    // Get a list of all projects from Asana
    const projectList = await getAllProjects(apiKey);
    projectList.forEach((project) => {
        project.name = cleanWhitespace(project.name);
    });
    debug("\r\nProject List", projectList);

    debug("\r\nSearching for target sheet...", "", doTest);
    // Iterate through all sheets in the workbook
    for (let i = 0; i < sheets.length; i++) {
        // Check if the current sheet is the target sheet
        if (sheets[i] === targetSheet) {
            debug("Sheet found: ", sheets[i], doTest);
            const rows = reader.utils.sheet_to_json(
                file.Sheets[file.SheetNames[i]]
            );

            // Iterate through each row
            debug("\r\nIterating through rows...", "", doTest);
            for (let j = startingRow - 2; j < rows.length; j++) {
                debug(`\r\n----------- Row ${j + 2} -----------`, "", doTest);
                debug("Raw data: ", rows[j], doTest);

                // Get project, date, and notes from the row
                var project = rows[j].Project;

                if (project) {
                    project = cleanWhitespace(project);
                }

                var date = rows[j].Date;

                // Date is currently stored as the number of days since January 1, 1900. Convert it to a date string in the form MM/dd/YYYY.
                if (date) {
                    const epochDate = new Date(1899, 11, 30);
                    const dateObject = new Date(epochDate.getTime() + (date * 24 * 60 * 60 * 1000));
                    date = `${dateObject.getMonth() + 1}/${dateObject.getDate()}/${dateObject.getFullYear()}`;
                }

                const notes = `${date} - ${rows[j].Notes}`;

                // Convert the task name to the Asana task name
                var task = cleanWhitespace(rows[j].Task);
                if (taskConversion[task]) {
                    task = cleanWhitespace(taskConversion[task]);
                }
                else {
                    debug("Invalid task!", task, doTest);
                    continue;
                }

                debug("\r\nCleaned data: {", "", doTest);
                debug("\tProject", project, doTest);
                debug("\tDate", date, doTest);
                debug("\tTask", task, doTest);
                debug("\tNotes", notes, doTest);
                debug("}", "", doTest);

                var projectExists = false;

                // Loop through projectList to find the project (if it exists)
                debug("\r\nSearching for project...", "", doTest);
                for (let k = 0; k < projectList.length; k++) {
                    // debug(`\r\nComparing ${project} to ${projectList[k].name}`, { value: projectList[k].name == project }, true);

                    if (projectList[k].name == project) {
                        projectExists = true;
                        debug("Project found: ", projectList[k], doTest);

                        // Get the project's GID
                        const project_gid = projectList[k].gid;

                        // Get the project's tasks if they haven't been fetched yet
                        if (!projectList[k].tasks) {
                            projectList[k].tasks = await getTasksInProject(project_gid, apiKey);
                        }
                        // debug("\r\nTasks", projectList[k].tasks, doTest);

                        var taskExists = false;

                        // Loop through the project's tasks to find the task (if it exists)
                        debug("\r\nSearching for task...", "", doTest);
                        for (let l = 0; l < projectList[k].tasks.length; l++) {
                            projectList[k].tasks[l].name = cleanWhitespace(projectList[k].tasks[l].name);
                            if (projectList[k].tasks[l].name === task) {
                                taskExists = true;
                                debug("Task found: ", projectList[k].tasks[l], doTest);

                                // Get the task's GID
                                const task_gid = projectList[k].tasks[l].gid;

                                // Create a comment on the task with the notes
                                debug("\r\nCreating comment...", "", doTest);
                                await createComment(task_gid, notes, apiKey);
                                debug("Comment created!", "", doTest);
                                break;
                            }
                        }

                        // If the task doesn't exist, print an error message
                        if (!taskExists) {
                            console.log("Task not found in project", task, doTest);
                        }

                        break;
                    }
                }

                // If the project doesn't exist, print an error message
                if (!projectExists) {
                    console.log("Project not found: ", project, doTest);
                }
            }

            break;
        }
    }
}

async function excelAsanaAPISync(apiKey, file, targetSheet, startingRow = 2, endingRow = 3) {
// exports.excelAsanaAPISync = async function (apiKey, file, targetSheet, startingRow = 2) {
    // async function asanaAPISync(apiKey, file, targetSheet, startingRow = 2) {
    debug("--------------- LAUNCHING ASANA API SYNC ---------------", "", doTest);
    debug("API Key:      ", apiKey, doTest);
    debug("Target Sheet: ", targetSheet, doTest);
    debug("Starting Row: ", startingRow, doTest);
    debug("Ending Row:   ", endingRow, doTest);
    debug("--------------------------------------------------------", "", doTest);

    // Get all sheet names from the workbook
    const sheets = file.SheetNames;
    debug("\r\nSheets", sheets);

    // Get a list of all projects from Asana
    const projectList = await getAllProjects(apiKey);
    projectList.forEach((project) => {
        project.name = cleanWhitespace(project.name);
    });
    debug("\r\nProject List", projectList);

    debug("\r\nSearching for target sheet...", "", doTest);
    // Iterate through all sheets in the workbook
    for (let i = 0; i < sheets.length; i++) {
        // Check if the current sheet is the target sheet
        if (sheets[i] === targetSheet) {
            debug("Sheet found: ", sheets[i], doTest);
            const rows = reader.utils.sheet_to_json(
                file.Sheets[file.SheetNames[i]]
            );

            // Iterate through each row
            debug("\r\nIterating through rows...", "", doTest);
            for (let j = startingRow - 2; j < endingRow - 2; j++) {
                debug(`\r\n----------- Row ${j + 2} -----------`, "", doTest);
                debug("Raw data: ", rows[j], doTest);

                // Get project, date, and notes from the row
                var project = rows[j].Project;

                if (project) {
                    project = cleanWhitespace(project);
                }

                var date = rows[j].Date;

                // Date is currently stored as the number of days since January 1, 1900. Convert it to a date string in the form MM/dd/YYYY.
                if (date) {
                    const epochDate = new Date(1899, 11, 30);
                    const dateObject = new Date(epochDate.getTime() + (date * 24 * 60 * 60 * 1000));
                    date = `${dateObject.getMonth() + 1}/${dateObject.getDate()}/${dateObject.getFullYear()}`;
                }

                const notes = `${date} - ${rows[j].Notes}`;

                // Convert the task name to the Asana task name
                var task = cleanWhitespace(rows[j].Task);
                if (taskConversion[task]) {
                    task = cleanWhitespace(taskConversion[task]);
                }
                else {
                    debug("Invalid task!", task, doTest);
                    continue;
                }

                debug("\r\nCleaned data: {", "", doTest);
                debug("\tProject", project, doTest);
                debug("\tDate", date, doTest);
                debug("\tTask", task, doTest);
                debug("\tNotes", notes, doTest);
                debug("}", "", doTest);

                var projectExists = false;

                // Loop through projectList to find the project (if it exists)
                debug("\r\nSearching for project...", "", doTest);
                for (let k = 0; k < projectList.length; k++) {
                    // debug(`\r\nComparing ${project} to ${projectList[k].name}`, { value: projectList[k].name == project }, true);

                    if (projectList[k].name == project) {
                        projectExists = true;
                        debug("Project found: ", projectList[k], doTest);

                        // Get the project's GID
                        const project_gid = projectList[k].gid;

                        // Get the project's tasks if they haven't been fetched yet
                        if (!projectList[k].tasks) {
                            projectList[k].tasks = await getTasksInProject(project_gid, apiKey);
                        }
                        // debug("\r\nTasks", projectList[k].tasks, doTest);

                        var taskExists = false;

                        // Loop through the project's tasks to find the task (if it exists)
                        debug("\r\nSearching for task...", "", doTest);
                        for (let l = 0; l < projectList[k].tasks.length; l++) {
                            projectList[k].tasks[l].name = cleanWhitespace(projectList[k].tasks[l].name);
                            if (projectList[k].tasks[l].name === task) {
                                taskExists = true;
                                debug("Task found: ", projectList[k].tasks[l], doTest);

                                // Get the task's GID
                                const task_gid = projectList[k].tasks[l].gid;

                                // Create a comment on the task with the notes
                                debug("\r\nCreating comment...", "", doTest);
                                await createComment(task_gid, notes, apiKey);
                                debug("Comment created!", "", doTest);
                                break;
                            }
                        }

                        // If the task doesn't exist, print an error message
                        if (!taskExists) {
                            console.log("Task not found in project", task, doTest);
                        }

                        break;
                    }
                }

                // If the project doesn't exist, print an error message
                if (!projectExists) {
                    console.log("Project not found: ", project, doTest);
                }
            }

            break;
        }
    }
}

if (doTest) {
    console.log("Testing is enabled");

    // Read in the Excel file
    const file = reader.readFile('WorkbookEden.xlsx');

    // asanaAPISync("", file, "E.JAN2025", 14);
    excelAsanaAPISync("", file, "E.JAN2025", 14);
}