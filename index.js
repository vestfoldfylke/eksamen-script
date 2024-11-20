// Hent filen 
// Finn alle som hadde eksamen i dag (i dag)
// Finn alle som skal ha eksamen i morgen (i dag + 1)
// Fjern alle som hadde eksamen i dag om de er elev hos oss og finnes i gruppen
// Legg til alle som skal ha eksamen i morgen om de er elev hos oss
// Send en e-post til de som ønsker informasjon, (Hvor mange ble meldt ut og hvor mange ble meldt inn og hvor mange var ikke elever hos oss)
// Sjekk om det er noen fremtidige datoer i filen, om det ikke er det flytt filen til en "finished" mappe


(async () => {
    const { graphRequest } = require('./graph/call-graph')
    const xlsx = require('xlsx')
    const fs = require('fs')
    const { misc, email } = require('./config')
    const { default: axios } = require('axios')
    const { logger } = require('@vtfk/logger')

    let isTest = false
    const logPrefix = 'checkIfPersonIsStudent'

    // Get todays date on this format: DD.MM.YYYY
    let today = new Date().toLocaleDateString('nb-NO')
    // Get tomorrows date on this format: DD.MM.YYYY
    let tomorrow = new Date(new Date().setDate(new Date().getDate() + 1)).toLocaleDateString('nb-NO')

    // Add a zero in front of the day if the day is less than 10
    const addZero = (i) => {
        if (i < 10) {
            i = "0" + i
        }
        return i
    }

    today = today.split('.').map(addZero).join('.')
    tomorrow = tomorrow.split('.').map(addZero).join('.')

    // Read passing arguments
    const args = process.argv.slice(2)
    if(args.length < 1) {
        isTest = false
    } else if(args[0] === 'prod') {
        isTest = false
    } else if(args[0] === 'test') {
        isTest = true
    } 
    else {
        logger('error', [logPrefix, `Invalid argument. Please provide a env value [prod/test]`])
        // Write the error to a file 
        fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ error: 'Invalid argument. Please provide a env value [prod/test]' }, null, 2), { flag: 'a' })
        return
    }

    // Find all xlsx and xls files in the folder
    const files = fs.readdirSync(misc.serverPath).filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'))
    // If no files are found, exit the function
    if(files.length === 0) {
        logger('info', [logPrefix, `No files found in the folder`])
        return
    }

    let numberOfStudentsRemoved = 0
    let numberOfStudentsAdded = 0
    let numberOfStudentsWithInvalidSSN = 0
    let numberOfStudnetsNotStudents = 0
    let totalNumberOfStudents = 0
    let totlaNumberOfErrors = 0
    let studentsArray = []
    let studentsErrorArray = []
    let isAnyDateInFuture = false

    
    // Create an object for the students with errors
    let studentErrorObj = {
        Fødselsnummer: undefined, // 11 digits
        BrukerId: undefined, // User ID
        Navn: undefined, // Navn på Kandidaten
        Type: undefined, // Error Type
        Eksamensparti: undefined, // Eksamensparti
        Error: undefined // Error message
    }

    // Functions for removing the last 5 digits of the SSN and replacing them with '*****'
    const removeSSN = (ssn) => {
        return ssn.slice(0, -5) + '*****'
    }

    // Function for querying the graph API
    const getGraphData = async (ssn) => {
        let url = `https://graph.microsoft.com/v1.0/users/?$count=true&$select=id,displayName,userPrincipalName,customSecurityAttributes&$filter=customSecurityAttributes/IDM/SSN eq '${ssn}'`
        const data = await graphRequest(url, 'GET', 'null', 'eventual')
        return data
    }

    // Function for removing students from the group
    const removeMember = async (id, user) => {
        const url = `https://graph.microsoft.com/v1.0/groups/${misc.groupID}/members/${id}/$ref`
        try {
            // Wait a few MS before sending the request
            await new Promise(resolve => setTimeout(resolve, 500))
            const request = await graphRequest(url, 'DELETE', 'null', null)
        } catch (error) {
            studentErrorObj.Fødselsnummer = removeSSN(user['Fødselsnummer'])
            studentErrorObj.BrukerId = id
            studentErrorObj.Navn = `${user['Fornavn']} ${user['Etternavn']}`
            studentErrorObj.Eksamensparti = user['Eksamensparti']
            studentErrorObj.Type = 'Klarte ikke å fjerne elev fra gruppen'
            studentErrorObj.Error = `Error while trying to remove member with id: ${id} from the group`
            
            studentsErrorArray.push(studentErrorObj)

            // Write the error to a file
            totlaNumberOfErrors++ 
            fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: `Error while trying to remove member with id: ${id} from the group`, error: error }, null, 2), { flag: 'a' })
            logger('error', [logPrefix, `Error while trying to remove member with id: ${id} from the group`, error])
        }

    }
    // Function for adding students to the group
    const addMember = async (id, user) => {
        const url = `https://graph.microsoft.com/v1.0/groups/${misc.groupID}/members/$ref`
        try {
            // Wait a few MS before sending the request
            await new Promise(resolve => setTimeout(resolve, 500))
            const request = await graphRequest(url, 'POST', { '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${id}` }, null)
        } catch (error) {
            studentErrorObj.Fødselsnummer = removeSSN(user['Fødselsnummer'])
            studentErrorObj.BrukerId = id
            studentErrorObj.Navn = `${user['Fornavn']} ${user['Etternavn']}`
            studentErrorObj.Eksamensparti = user['Eksamensparti']
            studentErrorObj.Type = 'Klarte ikke legge til elev i gruppen'
            studentErrorObj.Error = `Error while trying to add member with id: ${id} to the group`
            
            studentsErrorArray.push(studentErrorObj)

            // Write the error to a file
            totlaNumberOfErrors++ 
            fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: `Error while trying to add member with id: ${id} to the group`, error: error }, null, 2), { flag: 'a' })
            logger('error', [logPrefix, `Error while trying to add member with id: ${id} to the group`, error])
        }
    }
    // Function for checking for future dates
    const isFuture = (date) => {
        // Split the date into day, month and year
        let [day, month, year] = date.split('.')
        // Convert the day, month and year to integers
        day = parseInt(day)
        month = parseInt(month)
        year = parseInt(year)
        // Create a new date object with the given date
        let givenDate = new Date(`${year}.${month}.${day}`)
        // Create a new date object with todays date
        const todayDate = new Date(today.split('.').reverse().join('-'))
        // Check if the given date is in the future
        if(givenDate > todayDate) {
            isAnyDateInFuture = true
        }
    }

    const sendEmail = async (studentsArray) => {
        const emailBody = {
            to: email.to,
            from: email.from,
            subject: isTest ? `TEST! - ${email.subject} - ${today} / ${tomorrow}` : `${email.subject} - ${today} / ${tomorrow}`,
            text: `Hei, i dag ble det meldt ut ${numberOfStudentsRemoved} elever og inn ${numberOfStudentsAdded} elever. ${numberOfStudnetsNotStudents} elever var ikke elever hos oss.`,
            html: 
            `
            <p>I dag ble <strong>${totalNumberOfStudents}</strong> kandidater sjekket</p>
            <p>Det ble meldt ut <strong>${numberOfStudentsRemoved}</strong> elever</p>
            <p>Det ble meldt inn <strong>${numberOfStudentsAdded}</strong> elever</p>
            <p><strong>${numberOfStudnetsNotStudents}</strong> elever var ikke elever hos oss</p>
            <p><strong>${totlaNumberOfErrors}</strong> feil skjedde under kjøringen av scriptet, sjekk log filen på server eller papertrail!</p> <br>
            <table border="1" cellpadding="5" cellspacing="0">
            <thead>
                <tr>
                <th>Fødselsnummer</th>
                <th>Navn</th>
                <th>Type</th>
                <th>Eksamensparti</th>
                </tr>
            </thead>
            <tbody>
                ${studentsArray.map(student => `
                <tr style="background-color: ${student.Type === 'Lagt til gruppen' ? 'green' : student.Type === 'Fjernet fra gruppen' ? 'red' : 'white'};">
                    <td>${student.Fødselsnummer}</td>
                    <td>${student.Navn}</td>
                    <td>${student.Type}</td>
                    <td>${student.Eksamensparti}</td>
                </tr>
                `).join('')}
            </tbody>
            </table>`
        }
        try {
            await axios.post(`${email.api_url}/mail`, emailBody, { headers: { 'x-functions-key': `${email.api_key}` } })
        } catch (error) {
            // Write the error to a file
            totlaNumberOfErrors++
            fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: 'Error while trying to send email', error: error }, null, 2), { flag: 'a' })
            logger('error', [logPrefix, `Error while trying to send email`, error])
        }
    }

    const sendErrorEmail = async (studentsErrorArray) => {
        const emailBody = {
            to: email.to,
            from: email.from,
            subject: isTest ? `TEST! - ${email.subject} - ${today} / ${tomorrow}` : `${email.subject} - ${today} / ${tomorrow}`,
            text: `Hei, i dag skjedde det ${totlaNumberOfErrors} feil under kjøringen av scriptet. Mer informasjon finner du i log filen på server eller papertrail!`,
            html: 
            `
            <p> Det skjedde <strong>${totlaNumberOfErrors}</strong> feil under kjøringen av scriptet!</p>
            <p> Disse feilene må håndteres manuelt. </p> <br>
            <p> Gruppen elevene ble forsøkt lagt til og fjernet fra er: <strong>${misc.groupID}</strong> </p> <br>
            <p> I tabellen under finner du informasjon om feilene som skjedde under kjøringen av scriptet. </p> <br>
            <table border="1" cellpadding="5" cellspacing="0">
            <thead>
                <tr>
                <th>Fødselsnummer</th>
                <th>BrukerId</th>
                <th>Navn</th>
                <th>Type</th>
                <th>Eksamensparti</th>
                <th>Error</th>
                </tr>
            </thead>
            <tbody>
                ${studentsErrorArray.map(student => `
                <tr style="background-color: white;">
                    <td>${student.Fødselsnummer}</td>
                    <td>${student.BrukerId}</td>
                    <td>${student.Navn}</td>
                    <td>${student.Type}</td>
                    <td>${student.Eksamensparti}</td>
                    <td>${student.Error}</td>
                </tr>
                `).join('')}
            </tbody>
            </table>`
        }
        try {
            await axios.post(`${email.api_url}/mail`, emailBody, { headers: { 'x-functions-key': `${email.api_key}` } })
        } catch (error) {
            // Write the error to a file
            totlaNumberOfErrors++
            fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: 'Error while trying to send email', error: error }, null, 2), { flag: 'a' })
            logger('error', [logPrefix, `Error while trying to send email`, error])
        }
    }

    for (const file of files) {
        // Read the file
        const readFile = xlsx.readFile(`${misc.serverPath}/${file}`)
        const sheets = readFile.SheetNames
        // Reset the state for each file!
        isAnyDateInFuture = false
        for (const sheet of sheets) {
            const temp = xlsx.utils.sheet_to_json(readFile.Sheets[sheet], { raw: false })
            // Total number of students temp.length
            // In the array of objects, in the object find the key 'Fødslesnummer' or 'SSN or 'FNR' and push the value to the data array
            for (const obj of temp) {
                totalNumberOfStudents++
                // Check if the SSN is valid
                logger('info', [logPrefix, `Checking SSN`, removeSSN(obj['Fødselsnummer'])])
                if (obj['Fødselsnummer'].length !== 11) {
                    logger('warn', [logPrefix, `The SSN is invalid`, removeSSN(obj['Fødselsnummer'])])
                    numberOfStudentsWithInvalidSSN++
                    return // Exit the function if SSN is invalid
                }
                logger('info', [logPrefix, `The SSN is valid`, removeSSN(obj['Fødselsnummer'])])
                // Create an object for the student
                let studentObj = {
                    Fødselsnummer: removeSSN(obj['Fødselsnummer']), // 11 digits
                    Navn: undefined, // Navn på Kandidaten
                    Type: undefined, // Lagt til, Fjernet, Ikke elev
                    Eksamensparti: undefined, // Eksamensparti               
                }

                // For each row find the key 'Eksamensdato' and check if it is today or tomorrow
                let providedDate = obj['Eksamensdato']
                // Check if prividedDate is a valid date, is not empty
                if(providedDate === '' || providedDate === undefined || providedDate === null) {
                    logger('warn', [logPrefix, `Date is empty`, obj['Eksamensdato']])
                    return // Exit the function if the date is empty
                }
                // Check if the date is on the correct format DD.MM.YYYY

                // If the date is provided in this format MM/DD/YYYY, split the date and convert it to DD.MM.YYYY
                const dateArray = providedDate.split('/')
                if(dateArray.length === 3) {
                    // Add a zero in front of the day if the day is less than 10
                    if(dateArray[1] < 10) {
                        dateArray[1] = `0${dateArray[1]}`
                    }
                    // Make sure that the year is on the correct format
                    if(dateArray[2].length === 2) {
                        dateArray[2] = `20${dateArray[2]}`
                    }
                    providedDate = `${dateArray[1]}.${dateArray[0]}.${dateArray[2]}`
                }
                
                try {
                    const dateArray = providedDate.split('.')
                    if(dateArray.length !== 3) {
                        logger('warn', [logPrefix, `Date is not on the correct format`, obj['Eksamensdato']])
                        return // Exit the function if the date is not on the correct format
                    }
                } catch (error) {
                    totlaNumberOfErrors++
                    logger('error', [logPrefix, `Error while trying to split the date`, error])
                }
               
                if (providedDate === today) {
                    logger('info', [logPrefix, `The exam was today, checking if ${removeSSN(obj['Fødselsnummer'])} is a student`])
                    // Remove the students from the group if they are in the group
                    // Check if the person is a student or privatist and add the result to the object
                    try {
                        const result = await getGraphData(obj['Fødselsnummer'])
                        if(result.value.length > 0) {
                            if(isTest === false) {
                                try {
                                    logger('info', [logPrefix, `Person is a student, removing ${removeSSN(obj['Fødselsnummer'])} from the group`])
                                    await removeMember(result.value[0].id, obj)
                                } catch (error) {
                                    // Write the error to a file
                                    totlaNumberOfErrors++ 
                                    fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: 'Error while trying to remove member from the group', error: error }, null, 2), { flag: 'a' })
                                    logger('error', [logPrefix, `Error while trying to remove member from the group`, error])
                                }
                            } else {
                                logger('info', [logPrefix, `TEST - Person is a student, removing ${removeSSN(obj['Fødselsnummer'])} from the group`])
                            }
                            studentObj.Eksamensparti = obj['Eksamensparti']
                            studentObj.Navn = `${obj['Fornavn']} ${obj['Etternavn']}`
                            studentObj.Type = 'Fjernet fra gruppen'
                            numberOfStudentsRemoved++
                            // Push the object to the array
                            studentsArray.push(studentObj)
                        } else {
                            if(isTest === false) {
                                logger('info', [logPrefix, `Person is not a student`, removeSSN(obj['Fødselsnummer'])])
                            } else {
                                logger('info', [logPrefix, `TEST - Person is not a student`, removeSSN(obj['Fødselsnummer'])])
                            }
                            studentObj.Eksamensparti = obj['Eksamensparti']
                            studentObj.Navn = `${obj['Fornavn']} ${obj['Etternavn']}`
                            studentObj.Type = 'Ikke elev hos oss'
                            numberOfStudnetsNotStudents++
                            // Push the object to the array
                            studentsArray.push(studentObj)
                        }
                    } catch (error) {
                        // Write the error to a file
                        fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: 'Error while trying to get data from the graph API', error: error }, null, 2), { flag: 'a' })
                        logger('error', [logPrefix, `Error while trying to get data from the graph API`, error])
                    } 
                } else if (providedDate === tomorrow) {
                    // Add the students to the group if they are not in the group
                    logger('info', [logPrefix, `The exam is tomorrow, checking if ${removeSSN(obj['Fødselsnummer'])} is a student`])
                    try {
                        const result = await getGraphData(obj['Fødselsnummer'])
                        if(result.value.length > 0) {
                            if(isTest === false) {
                                try {
                                    logger('info', [logPrefix, `Person is a student, adding ${removeSSN(obj['Fødselsnummer'])} to the group`])
                                    await addMember(result.value[0].id, obj)
                                } catch (error) {
                                    // Write the error to a file
                                    totlaNumberOfErrors++
                                    fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: 'Error while trying to add member to the group', error: error }, null, 2), { flag: 'a' })
                                    logger('error', [logPrefix, `Error while trying to add member to the group`, error])
                                }
                            } else {
                                logger('info', [logPrefix, `TEST - Person is a student, adding ${removeSSN(obj['Fødselsnummer'])} to the group`])
                            }
                            studentObj.Eksamensparti = obj['Eksamensparti']
                            studentObj.Navn = `${obj['Fornavn']} ${obj['Etternavn']}`
                            studentObj.Type = 'Lagt til gruppen'
                            numberOfStudentsAdded++
                            // Push the object to the array
                            studentsArray.push(studentObj)
                        } else {
                            if(isTest === false) {
                                logger('info', [logPrefix, `Person is not a student`, removeSSN(obj['Fødselsnummer'])])
                            } else {
                                logger('info', [logPrefix, `TEST - Person is not a student`, removeSSN(obj['Fødselsnummer'])])
                            }
                            studentObj.Eksamensparti = obj['Eksamensparti']
                            studentObj.Navn = `${obj['Fornavn']} ${obj['Etternavn']}`
                            studentObj.Type = 'Ikke elev hos oss'
                            numberOfStudnetsNotStudents++
                            // Push the object to the array
                            studentsArray.push(studentObj)
                        }
                    } catch (error) {
                        // Write the error to a file
                        totlaNumberOfErrors++ 
                        fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: 'Error while trying to get data from the graph API', error: error }, null, 2), { flag: 'a' })
                        logger('error', [logPrefix, `Error while trying to get data from the graph API`, error])
                    }
                }
                // Check if the date is in the future
                isFuture(providedDate)
            }
        }
        if(!isAnyDateInFuture) {
            // Move the file to the finished folder
            logger('info', [logPrefix, `No future dates found in the file, moving the file to the finished folder.`])
            try {
                logger('info', [logPrefix, `Moving ${file} to finished folder`])
                fs.renameSync(`${misc.serverPath}/${file}`, `${misc.serverPath}/finished/${file}`)
            } catch (error) {
                // Write the error to a file
                totlaNumberOfErrors++
                fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: 'Error while trying to move file to finished folder', error: error }, null, 2), { flag: 'a' })
                logger('error', [logPrefix, `Error while trying to move file to finished folder`, error])
            }
        }
    }
    try {
        // Write studentsArray to a file if the array is not empty
        if(studentsArray.length === 0) {
            logger('info', [logPrefix, `No students found, skip writing logs to file`])
        } else {
            logger('info', [logPrefix, `Writing logs to file`])
            fs.writeFileSync(`${misc.serverPath}/logs/student-logs-${today}-${tomorrow}.json`, JSON.stringify(studentsArray, null, 2))
        }
    } catch (error) {
        // Write the error to a file
        totlaNumberOfErrors++
        fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: 'Error while trying to write to file', error: error }, null, 2), { flag: 'a' })
        logger('error', [logPrefix, `Error while trying to write to file`, error])
    }
    
    // Send email
    try {
        // Send mail only if the array is not empty
        if(studentsArray.length === 0) {
            logger('info', [logPrefix, `No students found, skip sending email`])
        } else {
            logger('info', [logPrefix, `Sending email`])
            sendEmail(studentsArray)
            // Send error email if there are any errors
            if(studentsErrorArray.length > 0) {
                sendErrorEmail(studentsErrorArray)
            }
        }
    } catch (error) {
        // Write the error to a file
        totlaNumberOfErrors++
        fs.writeFileSync(`${misc.serverPath}/logs/error-${today}-${tomorrow}.json`, JSON.stringify({ errorMsg: 'Error while trying to send email' ,error: error }, null, 2), { flag: 'a' })
        logger('error', [logPrefix, `Error while trying to send email`, error])
    }
})()