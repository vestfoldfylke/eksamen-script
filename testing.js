
(async() => {
    // Get todays date on this format: DD.MM.YYYY
    const today = new Date().toLocaleDateString('nb-NO')
    // Get tomorrows date on this format: DD.MM.YYYY
    const tomorrow = new Date(new Date().setDate(new Date().getDate() + 1)).toLocaleDateString('nb-NO')

    let isAnyDateInFuture = false

    // Check if the given date is in the future or not
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
        console.log(new Date(givenDate).toISOString())
        console.log(new Date (todayDate).toISOString())
        
        if(givenDate > todayDate) {
            isAnyDateInFuture = true
        }
    }


    // Array of dates to check
    const dates = ['01.11.2021', '01.11.2022', '01.11.2024', '01.11.2023', '01.12.2023']

    for (const date of dates) {
        isFuture(date)
    }
    
    console.log(isAnyDateInFuture)
})()