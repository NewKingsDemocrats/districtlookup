/**
 * A custom function that takes an address and returns the Assembly District.
 *
 * @param {String} address The street address
 * @return {Number} The Assembly District.
 */
function getAssemblyDistrict(address) {
  // The API to look up.
  var url = 'https://ccsunlight.org/api/v1/address/' + encodeURIComponent(address);
  // Fetch the full response.
  var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  
  var json = response.getContentText();
  var data = JSON.parse(json);
  // Extract the Assembly District from the response.
  // To see what else is in the response, use Logger.log(response).
  return data.ad;
}

/**
 * A custom function that takes an address and returns the Election District.
 *
 * @param {String} address The street address
 * @return {Number} The Election District.
 */
function getElectionDistrict(address) {
  // The API to look up.
  var url = 'https://ccsunlight.org/api/v1/address/' + encodeURIComponent(address);
  // Fetch the full response.
  var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
  
  var json = response.getContentText();
  var data = JSON.parse(json);
  // Extract the Election District from the response.
  // To see what else is in the response, use Logger.log(response).
  return data.ed;
}
