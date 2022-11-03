const axios = require('axios');

module.exports = {
    async getSelf(accessToken) {
        let resp = await callGraph('/me', accessToken)
        if (resp) {
            let data = await resp.json()
            return data
        }
    },
    async getPhoto() {
        let resp = await callGraph('/me/photos/240x240/$value')
        if (resp) {
            let blob = await resp.blob()
            return URL.createObjectURL(blob)
        }
    },
    async searchUsers(searchString, max = 50) {
        let resp = await callGraph(
            `/users?$filter=startswith(displayName, '${searchString}') or startswith(userPrincipalName, '${searchString}')&$top=${max}`
        )
        if (resp) {
            let data = await resp.json()
            return data
        }
    },
    async getAccessToken() {
        const headers = {
            'Accept': 'application/json',
            "Content-Type": "application/x-www-form-urlencoded",
        }
        //const url = `https://login.microsoftonline.com/f7dce5d8-272f-42d1-a49e-436ea515a324/oauth2/token`
        const url = `https://login.microsoftonline.com/c619bd99-fb23-4286-a720-d05f272a3a64/oauth2/token`
        const postData = {
            client_id: `c8dda09f-ac44-46f1-ab66-3a16f12c171f`, //`20e121f5-af03-4e79-912e-c975dc889259`,
            scope: 'https://graph.microsoft.com/.default',
            client_secret: `47c8Q~g0v1PRSS1Q~2PwrlB8KeBzf8lkQLDwubzV`,
            grant_type: 'client_credentials'
        };
        const response = await axios.post(url, postData, {
            headers: headers
        }).then((response) => {
            return response.data.access_token;
        }).catch((error) => {
            console.log(error);
        })
        return response;
    }
}

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0'
const GRAPH_SCOPES = ['user.read', 'user.readbasic.all']

async function callGraph(apiPath, accessToken) {
    let response = await axios.get(`${GRAPH_BASE}${apiPath}`, {
        headers: { authorization: `bearer ${accessToken}` }
    }).then((response) => {
        console.log(response);
    }).catch((error) => {
        console.log(error);
    });
}
