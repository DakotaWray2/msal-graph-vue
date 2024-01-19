// ----------------------------------------------------------------------------
// Copyright (c) Ben Coleman, 2020
// Licensed under the MIT License.
//
// Set of methods to call the beta Microsoft Graph API, using REST and fetch
// Requires auth.js
// ----------------------------------------------------------------------------

import auth from './auth'

const GRAPH_BASE = 'https://api.high.powerbigov.us/v1.0/myorg'
const GRAPH_SCOPES = [
  'https://high.analysis.usgovcloudapi.net/powerbi/api/App.Read.All',
  'https://high.analysis.usgovcloudapi.net/powerbi/api/Dashboard.Read.All',
  'https://high.analysis.usgovcloudapi.net/powerbi/api/Report.Read.ALl',
  'https://high.analysis.usgovcloudapi.net/powerbi/api/Workspace.Read.All',
  'https://high.analysis.usgovcloudapi.net/powerbi/api/Dataset.Read.All'
]

let accessToken

export default {
  //
  // Get details of user, and return as JSON
  // https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http#response-1
  //
  async getSelf() {
    let resp = await callGraph('/datasets/6ee70ff8-7b70-4b2a-b630-11405d2a9a85')
    if (resp) {
      let data = await resp.json()
      return data
    }
  },

  //
  // Get user's photo and return as a blob object URL
  // https://developer.mozilla.org/en-US/docs/Web/API/URL/createObjectURL
  //
  // async getPhoto() {
  //   let resp = await callGraph('/datasets/6ee70ff8-7b70-4b2a-b630-11405d2a9a85')
  //   if (resp) {
  //     let blob = await resp.blob()
  //     return URL.createObjectURL(blob)
  //   }
  // },

  //
  // Search for users
  // https://developer.mozilla.org/en-US/docs/Web/API/URL/createObjectURL
  //
  // async searchUsers() {
  //   let resp = await callGraph(
  //     `/datasets/6ee70ff8-7b70-4b2a-b630-11405d2a9a85`
  //   )
  //   if (resp) {
  //     let data = await resp.json()
  //     return data
  //   }
  // },

  //
  // Accessor for access token, only included for demo purposes
  //
  getAccessToken() {
    return accessToken
  }
}

//
// Common fetch wrapper (private)
//
async function callGraph(apiPath) {
  // Acquire an access token to call APIs (like Graph)
  // Safe to call repeatedly as MSAL caches tokens locally
  accessToken = await auth.acquireToken(GRAPH_SCOPES)

  let resp = await fetch(`${GRAPH_BASE}${apiPath}`, {
    headers: {
      'Access-Control-Allow-Origin': '*',
      'Content-Type': 'application/json',
      authorization: `Bearer ${accessToken}`
    }
  })

  if (!resp.ok) {
    throw new Error(`Call to ${GRAPH_BASE}${apiPath} failed: ${resp.statusText}`)
  }

  return resp
}
