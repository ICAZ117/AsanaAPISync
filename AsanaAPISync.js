const reader = require('xlsx');

const doDebug = false;

// doc: asana
const taskConversion = {
    "Meeting: Intake": "Discovery Phase",
    "Meeting: Methods/Ideas": "Protocol Development",
    "Analysis": "Statsitical Analysis",
    "Products": "Publication",
    "Review/Revise Package": "IRB Package Preparation Phase",
    "SAP": "IRB Package Preparation Phase",
    "DRR": "IRB Package Preparation Phase",
    "PrepWork": "Statistical Analysis"
}

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

exports.asanaAPISync = async function (apiKey, file, targetSheet, startingRow = 0) {
    // Get all sheet names from the workbook
    const sheets = file.SheetNames;

    // Get a list of all projects from Asana
    const projectList = await getAllProjects(apiKey);
    console.log(projectList);

    // Iterate through all sheets in the workbook
    for (let i = 0; i < sheets.length; i++) {
        // Check if the current sheet is the target sheet
        if (sheets[i] === targetSheet) {
            debug("Sheet", sheets[i]);
            const row = reader.utils.sheet_to_json(
                file.Sheets[file.SheetNames[i]]
            );

            // Iterate through each row
            for (let j = startingRow; j < row.length; j++) {
                debug("Row", row[j]);

                // Get project, date, and notes from the row
                const project = row[j].Project;
                const date = row[j].Date;
                const notes = `${date} - ${row[j].Notes}`;

                // Convert the task name to the Asana task name
                var task = row[j].Task;
                if (taskConversion[task]) {
                    task = taskConversion[task];
                }
                else {
                    console.log("Invalid task!", task);
                    continue;
                }

                debug("Project", project);
                debug("Date", date);
                debug("Task", task);
                debug("Notes", notes);
                debug("NEXT ROW\n\n");

                var projectExists = false;

                // Loop through projectList to find the project (if it exists)
                for (let k = 0; k < projectList.length; k++) {
                    if (projectList[k].name === project) {
                        projectExists = true;
                        debug("Project found", projectList[k]);

                        // Get the project's GID
                        const project_gid = projectList[k].gid;

                        // Get the project's tasks if they haven't been fetched yet
                        if (!projectList[k].tasks) {
                            projectList[k].tasks = await getTasksInProject(project_gid, apiKey);
                        }
                        debug("Tasks", projectList[k].tasks);

                        var taskExists = false;

                        // Loop through the project's tasks to find the task (if it exists)
                        for (let l = 0; l < projectList[k].tasks.length; l++) {
                            if (projectList[k].tasks[l].name === task) {
                                taskExists = true;
                                debug("Task found", projectList[k].tasks[l]);

                                // Get the task's GID
                                const task_gid = projectList[k].tasks[l].gid;

                                // Create a comment on the task with the notes
                                await createComment(task_gid, notes, apiKey);
                                break;
                            }
                        }

                        // If the task doesn't exist, print an error message
                        if (!taskExists) {
                            console.log("Task not found in project", task);
                        }

                        break;
                    }
                }

                // If the project doesn't exist, print an error message
                if (!projectExists) {
                    console.log("Project not found", project);
                }
            }

            break;
        }
    }
}