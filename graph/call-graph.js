const axios = require('axios').default
const getMsalToken = require('./get-endtraid-token.js')
// Calls the MSgraph API and returns the data
/**
 * 
 * @param {string} url 
 * @param {string} method 
 * @param {object} data 
 * @param {string} consistencyLevel 
 * @returns {Promise<any>}
 */
const graphRequest = async (url, method, data, consistencyLevel) => {
    // Get access token
    const accessToken = await getMsalToken('https://graph.microsoft.com/.default')
    // Build the request with data from the call
    const options = {
        method: method,
        url: url,
        headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${accessToken}`,
        }
    }
    // Add data to the request if it exists
    if (data) options.data = data
    // Add consistency level to the request if it exists
    if(consistencyLevel) options.headers['ConsistencyLevel'] = 'eventual'
    // Make the request
    const response = await axios(options)
    // Return the data
    return response.data
}

module.exports = { graphRequest }
