const chalk = require("chalk");
const excel = require('exceljs');
const fs = require('fs');
const XLSX = require('xlsx');
const objectNames = require('./JsonConfig/ObjectNames.json');
const sheetsConfig = require('./JsonConfig/Sheetnames.json');
const headerConfig = require('./JsonConfig/ChangelogHeaders.json');
//var fs = Promise.promisifyAll(require('fs'));


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
    Retrives Excel data synchronous  
*/

async function getExcelData() {
    var wb = new excel.Workbook();
    var filePath = 'C:/Users/jclappsy/Desktop/TestQueries.xlsx'
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
       // console.log('SHEET: ', sheet);

        switch(sheet.name) {
            case sheetsConfig.ConfigData :
                console.log('generating queries for config data');
                configData = createConfigDataObjArray(sheet, configData, jiraTickets, epochDate);
                break;
            case sheetsConfig.BugFixes :
                console.log('generating queries for bug fixes');
                configData = createConfigDataObjArray(sheet, configData, jiraTickets, epochDate);
                // console.log(configData);
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

    sheet.eachRow(function(row, rowNumber) {
        let rowObj = {};

        if(rowNumber !== 1) {
            // create row object
            headerKeys.forEach((headerKeys, i) =>  {
            // set row number attribute to highlight row after queries are created
                rowObj.rowNumber = rowNumber;
                rowObj.sheetName = sheet.name;
                rowObj[headerKeys] = row.values[i]
            });

            // Filter by Jira Task and Date
            if(rowObj[headerConfig.Date] >= date &&  jiraTickets.includes(rowObj[headerConfig.JiraTask])) {
                configData.push(rowObj);
            }

        } 
    });

    return configData;
}

const createConfigDeloyQueries = (configData, fileName) => {
    let queryDict = {}

    configData.forEach(row => {
        let objectNameKey = row[headerConfig.ObjectType];

        if(queryDict[row[headerConfig.ObjectType]]) {
            let queryDictObj = queryDict[row[headerConfig.ObjectType]];
            if(row[headerConfig.ExternalId]) {
                //TODO move to function
                // Use JSON Map to get value for obejct type/name
                // For now, just get it to work!! :-)

                if(row[headerConfig.ExternalId]) {
                    queryDictObj[objectNameKey].ExternalIds.push(row[headerConfig.ExternalId]);
                } else {
                    queryDictObj[objectNameKey].Names.push(row[headerConfig.Name]);
                }
            }
        } else {

            let rowObj = {
                [objectNameKey] : {
                    ObjectApiName: '',
                    ExternalIds: [],
                    Names: []
                }
            }

            //TODO move to function
            // Use JSON Map to get value for obejct type/name

            rowObj[objectNameKey].ObjectApiName = row[headerConfig.ObjectType];

            if(row[headerConfig.ExternalId]) {
                rowObj[objectNameKey].ExternalIds.push(row[headerConfig.ExternalId]);
            } else {
                rowObj[objectNameKey].Names.push(row[headerConfig.Name]);
            }

            queryDict[objectNameKey] = rowObj;

        }
    });

    createDeployQueriesTextFile(queryDict, fileName);

}


async function createDeployQueriesTextFile(queryDict, fileName) {
    try {
        await createTextFile(queryDict, fileName);
    } catch(error) {
        console.log(`createDeployQueriesTextFile ${error}`);
    }
}

async function createTextFile(queryDict, fileName) {
    //TODO: Use current directory
    var filePath = 'C:/Users/jclappsy/Desktop'
    //let queryString = '';
    let queryStringArray = [];
    let divider1 = '_______________________________________'
    let divider2 = '-------------------------'


    console.log('CREATING TEXT FILE')

    for (const [key, value] of Object.entries(queryDict)) {
        let queryString = '';
        //console.log(key, value);
        if(objectNames[key]) {
            queryString += `${divider1}\n\n${divider2}\n${value[key].ObjectApiName}\n${divider2}\n\n`;

            const hasIdsAndNames = value[key].ExternalIds.length !== 0 && value[key].Names.length !== 0;
            const hasIdsOnly = value[key].ExternalIds.length !== 0 && value[key].Names.length === 0
            const hasNamesOnly = value[key].ExternalIds.length === 0 && value[key].Names.length !== 0
            
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
      }

      finalQueryString = '';

      queryStringArray.forEach(string => {
        finalQueryString += string;
    })

    fs.writeFileSync('C:/Users/jclappsy/Desktop/deployqueries.txt', finalQueryString);





}

const createWhereClause = (queryString, data, divider1, divider2) => {

    data.forEach(function(value, index) {

        queryString += index + 1 === data.length ?  `'${value}'` :  `'${value}',\n`


    })

    // data.forEach(value => {
    // })

    queryString += `) \n\n${divider2}\n${divider2}\n\n${divider1}`

    return queryString;
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

   headerValues.unshift('FirstEmptyCol')
   return headerValues;
}


const convertDateString = (date) => {
    return new Date(date).getTime();
}

const convertJiraTicketsString = (tickets) => {
    let ticketsArray = tickets.split(',');
    return ticketsArray;
}


module.exports = {createQueries};