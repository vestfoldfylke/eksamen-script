require('dotenv').config()
module.exports = {
    appReg: {
        tenantId: process.env.AZURE_APP_TENANT_ID,
        clientId: process.env.AZURE_APP_ID,
        clientSecret: process.env.AZURE_APP_SECRET,
        scope: process.env.AZURE_APP_SCOPE || 'https://graph.microsoft.com/.default',
        grantType: process.env.AZURE_APP_GRANT_TYPE || 'client_credentials'
    },
    email: {
        to: (process.env.MAIL_TO && process.env.MAIL_TO.split(',') || []),
        from: process.env.MAIL_FROM,
        subject: process.env.MAIL_SUBJECT,
        api_url: process.env.MAIL_URL,
        api_key: process.env.MAIL_KEY
    },
    misc: {
        groupID: process.env.NETTPSERRE_EKSAMEN_GROUP_ID
    }
}