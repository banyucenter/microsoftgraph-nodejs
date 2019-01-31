var graph = require('@microsoft/microsoft-graph-client');

module.exports = {
  getUserDetails: async function (accessToken) {
    const client = getAuthenticatedClient(accessToken);

    const user = await client.api('/me').get();
    return user;
  },

  getEvents: async function (accessToken) {
    const client = getAuthenticatedClient(accessToken);

    const events = await client
      .api('/me/events')
      .select('subject,organizer,start,end,location')
      .orderby('createdDateTime DESC')
      .get();
    console.log(events);
    return events;

  },

  getDrive: async function (accessToken) {
    const client = getAuthenticatedClient(accessToken);
    const driven = await client
      .api('/me/drive/recent')
      .select('name,size,webUrl')
      .orderby('name')
      .get();
    console.log(driven);
    return driven;
  },

  getCalendar: async function (accessToken) {
    const client = getAuthenticatedClient(accessToken);
    const calends = await client.api('/me/calendars').get();
    console.log(calends);
    return calends;
  }

};

function getAuthenticatedClient(accessToken) {
  // Initialize Graph client
  const client = graph.Client.init({
    // Use the provided access token to authenticate
    // requests
    authProvider: (done) => {
      done(null, accessToken);
    }
  });

  return client;
}