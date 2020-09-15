const chalk = require("chalk");
const excel = require('exceljs');
const fs = require('fs');
const XLSX = require('xlsx');
const objectNames = require('./JsonConfig/ObjectNames.json');
const sheetsConfig = require('./JsonConfig/Sheetnames.json');
const headerConfig = require('./JsonConfig/ChangelogHeaders.json');
const appConfig = require('./JsonConfig/Config.json');

async function createQueries(tickets, date, fileName) {
    var data = [];
    let epochDate = convertDateString(date);
    let jiraTickets = convertJiraTicketsString(tickets);

    try {
         data = await getExcelData();
     } catch (error) {
         console.error(`getExcelData: Error occurred:  ${error}`);
     }

    processExcelData(data, jiraTickets, epochDate, fileName);
}

/*
    Retrives Excel data synchronously
*/

async function getExcelData() {
    var wb = new excel.Workbook();
    //TODO: move to config
    var filePath = appConfig.changelogFilePath;
    var excelData = [];
    await wb.xlsx.readFile(filePath);
    var sheetNames = sheetsConfig.sheetNames;
    sheetNames.forEach(name => {
        var sh = wb.getWorksheet(name);
        excelData.push(sh);
    });
    return excelData;
}

const processExcelData = (data, jiraTickets, epochDate, fileName) => {
    // Create XML
    let metadataData = [];
    // Create DeployQueries.txt
    let configData = [];
    // Create DeprecationQueries.txt
    let deprecatedData = [];

    // iterate through data var. This contains all sheets and their data
    data.forEach(sheet => {
       console.log('SHEET: ', sheet.name);

        switch(sheet.name) {
            case sheetsConfig.ConfigData :
                console.log('generating queries for config data');
                configData = createConfigDataObjArray(sheet, configData, jiraTickets, epochDate);
                break;
            case sheetsConfig.BugFixes :
                console.log('generating queries for bug fixes');
                configData = createConfigDataObjArray(sheet, configData, jiraTickets, epochDate);
                break;
            // case sheetsConfig.Metadata :
            //     console.log('creating metadata');
            //     break;
            // case sheetsConfig.Deprecated :
            //     console.log('generating queries for Deprecated');
            //     break;
        }
    })

    // create queries
   createConfigDeloyQueries(configData, fileName);
}

const createConfigDataObjArray = (sheet, configData, jiraTickets, date ) => {

    let headerKeys = createHeaders(sheet);
    let missingObjNames = [];

    sheet.eachRow(function(row, rowNumber) {
        let rowObj = {};

        if(rowNumber !== 1) {
            // create row object
            headerKeys.forEach((key, i) =>  {
            // set row number attribute to highlight row after queries are created
                rowObj.rowNumber = rowNumber;
                rowObj.sheetName = sheet.name;

                let objectTypeKey = '';

                // Use map to check object type. Some sheets have different header names.
                // TODO: Come up with better way to handle Header Names :,(
                // Use array of possible values in json
    
                if(key === 'Object' || key === 'ObjectType') {
                    key = 'ObjectType';
                    objectTypeKey = key;
                } 

                rowObj[key] = row.values[i]
                
                // Sometimes Dates with leading zeros (i.e. 09/02/2020) will be in excel doc
                // These are treated as strings, so must be converted to a date
                
                if(key === 'Date') {
                    if(rowObj[key] && typeof rowObj[key] === 'string' && rowObj[key].indexOf('/') > -1) {
                        rowObj[key] = convertDateString(rowObj[key]);
                    }
                }

                // Takes the Object Type entered in changelog and creates a Key Value
                // This key value is used to lookup the api name of the obj
                // This ensures entries like Custom Object and Custom_Object__c yield the same value
                if(objectTypeKey !== '' && rowObj[objectTypeKey]) {
                    let objectTypeString = rowObj[objectTypeKey].toLowerCase();
                    objectTypeString = removeSpaces(objectTypeString);

                    if(objectNames[objectTypeString]) {
                        rowObj[objectTypeKey] = objectNames[objectTypeString].objName;
                        rowObj.objOrder = objectNames[objectTypeString].order;
                    } else {
                        missingObjNames.push(objectTypeString);
                    }
                }
            });

            // Filter by Jira Task and Date
            if(rowObj[headerConfig.Date] >= date &&  jiraTickets.includes(rowObj[headerConfig.JiraTask])) {                
                configData.push(rowObj);
            }

        } 
    });

    /*
        If User enters an obj name in the excel doc that does not exist as a key in ObjectNames.json, the row will be missed
        If data is missing, you can debug by uncommenting the lines below.
        This prints each object name from the excel file that does not exist in the ObjectNames.json
        Add the name as a key if appliciable
    */

    checkForMissingObjNames(missingObjNames);

    return configData;
}

const createConfigDeloyQueries = (configData, fileName) => {
    let queryDict = {}

    configData.forEach(row => {

        let objectNameKey = row[headerConfig.ObjectType].toLowerCase();
        let mapKey = row[headerConfig.ObjectType].toLowerCase();
        let childObjLookupKey = row.ExternalID ? row.ExternalID : row.Name

        if(queryDict[mapKey]) {
            let queryDictObj = queryDict[mapKey];
            if(row[headerConfig.ObjectType]) {
                //TODO move to function
                if(row.ExternalID) {
                    queryDictObj[objectNameKey].ExternalIds.push(row.ExternalID);
                } else {
                    queryDictObj[objectNameKey].Names.push(row.Name);
                }
                //TODO move to function
                if(row.ChildObjects) {
                    if(queryDictObj[objectNameKey].ChildObjectData[childObjLookupKey]) {
                        queryDictObj[objectNameKey].ChildObjectData[childObjLookupKey].childObjInfo.push(row.ChildObjects);
                    } else {
                        queryDictObj[objectNameKey].ChildObjectData = {
                            [childObjLookupKey] : {
                                childObjInfo: [row.ChildObjects]
                            }
                        }
                    }
                }
            }
        } else {

            ///////// Handle new Object Type

            let rowObj = {
                [objectNameKey] : {
                    ObjectApiName: '',
                    ExternalIds: [],
                    Names: [],
                    DeployOrder: 0,
                    ChildObjectData: {}
                }
            }

            rowObj[objectNameKey].ObjectApiName = row.ObjectType;
            rowObj[objectNameKey].DeployOrder = row.objOrder;

            //TODO move to function
            if(row.ExternalID) {
                rowObj[objectNameKey].ExternalIds.push(row.ExternalID);   
            } else {
                rowObj[objectNameKey].Names.push(row.Name);
            }

            //TODO move to function
            if(row.ChildObjects) {
                rowObj[objectNameKey].ChildObjectData = {
                    [childObjLookupKey] : {
                        childObjInfo: [row.ChildObjects]
                    }
                }
            }

            queryDict[objectNameKey] = rowObj;
        }
    });

    sortedDictionaryKeys = sortQueryDict(queryDict);

    createDeployQueriesTextFile(queryDict, fileName, sortedDictionaryKeys);
}


async function createDeployQueriesTextFile(queryDict, fileName, sortedDictionaryKeys) {
    try {
        await createTextFile(queryDict, fileName, sortedDictionaryKeys);
    } catch(error) {
        console.log(`createDeployQueriesTextFile ${error}`);
    }
}

async function createTextFile(queryDict, fileName, sortedDictionaryKeys) {
    //TODO: Use current directory
    var filePath = 'C:/Users/jclappsy/Desktop'
    //let queryString = '';
    let queryStringArray = [];
    let divider1 = '_______________________________________'
    let divider2 = '-------------------------'

    sortedDictionaryKeys.forEach(key => {
        let queryString = '';
        let value = queryDict[key];

        if(objectNames[key]) {
            queryString += `${divider1}\n\n${divider2}\n${value[key].ObjectApiName}\n${divider2}\n\n`;

            const hasIdsAndNames = value[key].ExternalIds.length !== 0 && value[key].Names.length !== 0;
            const hasIdsOnly = value[key].ExternalIds.length !== 0 && value[key].Names.length === 0
            const hasNamesOnly = value[key].ExternalIds.length === 0 && value[key].Names.length !== 0
            const hasChildObjectInfo = Object.keys(value[key].ChildObjectData).length !== 0;

            if(hasChildObjectInfo) {
                queryString += 'Child Object Information:\n';
                queryString += creatChildObjectsList(value[key].ChildObjectData);
            }
            
            if(hasIdsAndNames) {
                queryString += 'Conversion_Ref_Id__c IN ('
                queryString = createWhereClause(queryString, value[key].ExternalIds, divider1, divider2);
                queryString += ' OR Name IN ('
                queryString = createWhereClause(queryString, value[key].Names, divider1, divider2);
            } 
            else if(hasIdsOnly) {
                queryString += 'Conversion_Ref_Id__c IN ('
                queryString = createWhereClause(queryString, value[key].ExternalIds, divider1, divider2);
            }
            else if(hasNamesOnly) {
                queryString += 'Name IN ('
                queryString = createWhereClause(queryString, value[key].Names, divider1, divider2);
            }
        }

        queryString += '\n\n';
        queryStringArray.push(queryString);
        
    });

    finalQueryString = '';

    queryStringArray.forEach(string => {
      finalQueryString += string;
    });

    fs.writeFileSync(appConfig.deployQueriesFilePath, finalQueryString);
    console.log('Text File Created.');
}

const createWhereClause = (queryString, data, divider1, divider2) => {

    data.forEach(function(value, index) {
        queryString += index + 1 === data.length ?  `'${value}'` :  `'${value}',\n`
    })

    queryString += `) \n\n${divider2}\n${divider2}\n\n${divider1}`

    return queryString;
}

const creatChildObjectsList = (childObjDataObj) => {
    let childObjInfoString = '';
    let childObjDataObjKeyLength = Object.entries(childObjDataObj).length;
    let keyCount = 0;

    for(let key in childObjDataObj) {
        childObjInfoString += `${key}: `
        childObjDataObj[key].childObjInfo.forEach(function (entry, i) {
            childObjInfoString += i !== childObjDataObj[key].childObjInfo.length -1 ? `${entry}, ` : `${entry}`
        });

       keyCount += 1;

       if(keyCount === childObjDataObjKeyLength) {
           childObjInfoString += '\n'
       }
    }

    childObjInfoString += '\n';

    return childObjInfoString;
}


/*
    Helpers
*/

// Removes spaces from Excel Column Labels. Used to create Row object attributes
const createHeaders = (sheet) => {
   let headers = sheet._rows[0].values;
   let headerValues = [];
   headers.forEach(h => {
        if(h.indexOf(' ') >= 0); {
            h = h.replace(/ /g,'');
        }

        headerValues.push(h);
   });

   // Add a fake string to header values array to account blank number column. This will ensure header keys line up with data.
   headerValues.unshift('FirstEmptyCol')
   return headerValues;
}

const removeSpaces = (string) => {
    if(string.indexOf(' ') >= 0); {
        return string.replace(/ /g,'');
    }
}

const convertDateString = (date) => {
    // remove leading 0s
    let dateStringArray = date.split('/');
    let newDateStringArray = [];

    for(let i = 0; i < dateStringArray.length; i++) {
        if(dateStringArray[i].charAt(0) == 0) {
            dateStringArray[i] = dateStringArray[i].substring(1);
        }
        newDateStringArray.push(dateStringArray[i]);
    }

    let newDateString = newDateStringArray.join('/');

    return new Date(newDateString).getTime();
}

const convertJiraTicketsString = (tickets) => {
    let ticketsArray = tickets.split(',');
    return ticketsArray;
}

const checkForMissingObjNames = (objNames) => {
    console.log('...Checking for missing object names...')
    if(objNames.length > 0) {
        objNames.forEach(name => {
            console.log(name);
        })
    } else {
        console.log('No missing object names! :-)');
    }
}

const sortQueryDict = (queryDict) => {
    let sortedDictKeys = Object.keys(queryDict).sort(function(a,b) {
       return queryDict[a][a].DeployOrder - queryDict[b][b].DeployOrder;;
    });

    return sortedDictKeys;
}

module.exports = {createQueries};