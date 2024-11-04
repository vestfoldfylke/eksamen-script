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
       process.exit(1)
    }

    // Find all xlsx and xls files in the folder
    const files = fs.readdirSync(misc.serverPath).filter(file => file.endsWith('.xlsx') || file.endsWith('.xls'))
 
    let numberOfStudentsRemoved = 0
    let numberOfStudentsAdded = 0
    let numberOfStudentsWithInvalidSSN = 0
    let numberOfStudnetsNotStudents = 0
    let totalNumberOfStudents = 0
    let studentsArray = []
    let isAnyDateInFuture = false

    // Get todays date on this format: DD.MM.YYYY
    const today = new Date().toLocaleDateString('nb-NO')
    // Get tomorrows date on this format: DD.MM.YYYY
    const tomorrow = new Date(new Date().setDate(new Date().getDate() + 1)).toLocaleDateString('nb-NO')

    // Functions for removing the last 5 digits of the SSN and replacing them with '*****'
    const removeSSN = (ssn) => {
        return ssn.slice(0, -5) + '*****'
    }

    // Function for querying the graph API
    const getGraphData = async (ssn) => {
        let url = `https://graph.microsoft.com/v1.0/users/?$count=true&$select=id,displayName,userPrincipalName,customSecurityAttributes&$filter=customSecurityAttributes/IDM/SSN eq '${ssn}'`
        const data = await graphRequest(url, 'GET', 'null', 'eventual')
        // console.log(data)
        return data
    }

    // Function for removing students from the group
    const removeMember = async (id) => {
        const url = `https://graph.microsoft.com/v1.0/groups/${misc.groupID}/members/${id}/$ref`
        try {
            const request = await graphRequest(url, 'DELETE', 'null', 'eventual')
        } catch (error) {
            console.log(error)
        }

    }
    // Function for adding students to the group
    const addMember = async (id) => {
        const url = `https://graph.microsoft.com/v1.0/groups/${misc.groupID}/members/$ref`
        try {
            const request = await graphRequest(url, 'POST', { '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${id}` }, 'eventual')
        } catch (error) {
            console.log(error)
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
        // Add +1 to the day
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
            <p><strong>${numberOfStudnetsNotStudents}</strong> elever var ikke elever hos oss</p> <br>
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
            logger('error', [logPrefix, `Error while trying to send email`, error])
            console.log(error)
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
                const dateConverted = new Date(obj['Eksamensdato']).toLocaleDateString('nb-NO')
                if (dateConverted === today) {
                    logger('info', [logPrefix, `The exam was today, checking if ${removeSSN(obj['Fødselsnummer'])} is a student`])
                    // Remove the students from the group if they are in the group
                    // Check if the person is a student or privatist and add the result to the object
                    try {
                        const result = await getGraphData(obj['Fødselsnummer'])
                        if(result.value.length > 0) {
                            if(isTest === false) {
                                logger('info', [logPrefix, `Person is a student, removing ${removeSSN(obj['Fødselsnummer'])} from the group`])
                                console.log(result.value[0].id) // Bruk dette til å fjerne personen fra gruppen
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
                        console.log(error)
                        logger('error', [logPrefix, `Error while trying to get data from the graph API`, error])
                    } 
                } else if (dateConverted === tomorrow) {
                    // Add the students to the group if they are not in the group
                    logger('info', [logPrefix, `The exam is tomorrow, checking if ${removeSSN(obj['Fødselsnummer'])} is a student`])
                    try {
                        const result = await getGraphData(obj['Fødselsnummer'])
                        if(result.value.length > 0) {
                            if(isTest === false) {
                                logger('info', [logPrefix, `Person is a student, adding ${removeSSN(obj['Fødselsnummer'])} to the group`])
                                console.log(result.value[0].id) // Bruk dette til å legge til personen i gruppen
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
                        console.log(error)
                        logger('error', [logPrefix, `Error while trying to get data from the graph API`, error])
                    }
                }
                // Check if the date is in the future
                isFuture(dateConverted)
            }
        }
        if(!isAnyDateInFuture) {
            // Move the file to the finished folder
            logger('info', [logPrefix, `No future dates found in the file, moving the file to the finished folder.`])
            try {
                logger('info', [logPrefix, `Moving ${file} to finished folder`])
                fs.renameSync(`${misc.serverPath}/${file}`, `${misc.serverPath}/finished/${file}`)
            } catch (error) {
                logger('error', [logPrefix, `Error while trying to move file to finished folder`, error])
            }
        }
    }
    try {
        // Write studentsArray to a file
        logger('info', [logPrefix, `Writing logs to file`])
        fs.writeFileSync(`${misc.serverPath}/logs/student-logs-${today}-${tomorrow}.json`, JSON.stringify(studentsArray, null, 2))
    } catch (error) {
        logger('error', [logPrefix, `Error while trying to write to file`, error])
    }
    
    // Send email
    try {
        logger('info', [logPrefix, `Sending email`])
        sendEmail(studentsArray)
    } catch (error) {
        logger('error', [logPrefix, `Error while trying to send email`, error])
    }
})()