const activeUser = Session.getActiveUser();
const userProperties = PropertiesService.getUserProperties();
const sheetProperties = PropertiesService.getDocumentProperties();

function onOpen(e) {
    loadUI();
}

function loadUI() {
    let ss = SpreadsheetApp.getActive();
    if (sheetProperties.getProperty('installed') !== null) {
        if (ss.getSheetByName('TDSMaker') === null || ss.getSheetByName('Mappings') === null) {
            loadPage('templates');
        } else {
            loadPage('send');
        }
    } else {
        if (userProperties.getProperty('token') !== null) {
            loadPage('templates');
        } else {
            loadPage('welcome');
        }
    }
}

function parseSpreadsheetData(): any {
    let wrapperArray = [];
    try {
        let mappingSheet = getSheetById(sheetProperties.getProperty('mappingSheet'));
        let dataSheet = getSheetById(sheetProperties.getProperty('dataSheet'));
        let columnLimit = mappingSheet.getLastColumn();
        let mappingValues = mappingSheet.getRange(1, 1, 1, columnLimit).getValues()[0];
        let dataRange = dataSheet.getLastRow();

        for (let i = 3; i <= dataRange; i++) {
            let dataValues = dataSheet.getRange(i, 1, 1, columnLimit).getValues()[0];
            Logger.log(dataRange);
            let wrapper = {};
            for (let j = 0; j < columnLimit; j++) {
                if (dataValues[j].toString() !== "") {
                    wrapper[mappingValues[j].toString()] = dataValues[j];
                } else {
                    continue;
                }
            }

            wrapperArray.push(wrapper);
        }
    } catch (e) {
        Log.send("failed to parse data: %s", e.message);
        throw new Error("failed to parse data, please try again later.");
    }

    return new TemplateData((JSON.parse(sheetProperties.getProperty('templateData'))).templateId, wrapperArray);
}

function prepareSpreadsheet(templateData: any): boolean {
    try {
        let ss = SpreadsheetApp.getActive();

        if (ss.getSheetByName('TDSMaker') === null) {
            ss.insertSheet('TDSMaker', 0);
            fillHeaders(0, templateData);
        }

        if (ss.getSheetByName('Mappings') === null) {
            ss.insertSheet('Mappings', 1);
            fillMappings(1, templateData);
        }

        let dataSheet = ss.getSheetByName('TDSMaker');
        let mappingSheet = ss.getSheetByName('Mappings');

        mappingSheet.hideSheet();

        sheetProperties.setProperties({
            'dataSheet': dataSheet.getSheetId().toString(),
            'mappingSheet': mappingSheet.getSheetId().toString(),
            'templateData': JSON.stringify(templateData)
        });
    } catch (e) {
        Log.send("error occurred while preparing necessary sheets: %s", e.message);
        return false;
    }

    return true;
}

function apiLogin(email: string, password: string) {
    const url = "/token";
    const data = {
        "email": email,
        "password": password
    };

    const response = fetchPost(url, data);
    if (response.status === false) {
        Log.send("%s tried to login with data: %s", activeUser.getEmail(), JSON.stringify(data));

        throw new Error("Your login credentials are incorrect or your account is no longer valid.");
    } else if (response.status === true) {
        Log.send("%s did login successfully with credentials %s. Token: %s", activeUser.getEmail(), JSON.stringify(data), response.authData.token);

        userProperties.setProperties({
            'token': response.authData.token,
            '_id': response.authData._id
        });

        if (getApiMapList() === true) {
            loadPage("templates");
        } else {
            throw new Error("An error occurred while getting sheets from server, please try again.");
        }
    }
}

function getApiMapList(): boolean {
    const url = "/form/mappings"

    const userToken = userProperties.getProperty("token");

    try {
        const response = fetchGet(url, userToken);

        userProperties.setProperty('mappings', JSON.stringify(response));
    } catch (err) {
        Log.send("failed to load map list: %s", err.message);
        return false;
    }

    return true;
}

function getSheetById(id: string) {
    let intID = parseInt(id);
    return SpreadsheetApp.getActive().getSheets().filter((s) => {
        return s.getSheetId() === intID;
    })[0];
}

function findMappingById(id: string) {
    const mappings = JSON.parse(userProperties.getProperty('mappings'));

    let templateData = null;

    for (let i = 0; i < mappings.length; i++) {
        if (mappings[i].templateId === id) {
            templateData = mappings[i];
            break;
        }
    }

    if (templateData === null) {
        throw new Error("given document id is not valid!");
    }

    let sheetsReady = prepareSpreadsheet(templateData);

    if (sheetsReady) {
        sheetProperties.setProperty('installed', 'true');
        loadPage('send');
    } else {
        throw new Error("failed to create/prepare necessary sheets, please try again later.")
    }
}

function getFormattedMappings(): any {
    const mappings = JSON.parse(userProperties.getProperty('mappings'));
    let wrapper = [];

    mappings.forEach(element => {
        wrapper.push({
            key: element.templateId,
            value: element.name
        })
    });

    return wrapper;
}


function fillHeaders(index: number, data: any) {
    let s = SpreadsheetApp.getActive().getSheets()[index];
    let dataCount = Object.keys(data.data).length;
    let range = s.getRange(2, 1, 1, dataCount);

    let headerData = mapObject("value", data.data);

    range.setValues([
        [...headerData]
    ]);

    s.autoResizeColumns(1, dataCount); // auto resize columns to visual beauty

    range.protect().setWarningOnly(true);
}

function fillMappings(index: number, data: any) {
    let s = SpreadsheetApp.getActive().getSheets()[index];
    let dataCount = Object.keys(data.data).length;
    let range = s.getRange(1, 1, 1, dataCount);

    let mappingData = mapObject("key", data.data);

    range.setValues([
        [...mappingData]
    ]);

    range.protect().setWarningOnly(true);
}

function mapObject(key: string, data: any) {
    let dataWrapper = [];
    for (let dataKey in data) {
        if (Object.prototype.hasOwnProperty.call(data, dataKey)) {
            dataWrapper.push(data[dataKey][key]);
        }
    }

    return dataWrapper;
}

// helpers

function fetchPost(url: string, data: object, token?: string): any {
    const baseurl = "https://privateapi.tdsmaker.com/api/v2";

    const response = UrlFetchApp.fetch(`${baseurl}${url}`, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(data),
        headers: token === null || typeof token === "undefined" ? {} : {
            Authorization: 'Bearer ' + token.trim()
        }
    });

    return JSON.parse(response.getContentText());
}

function fetchGet(url: string, token?: string): any {
    const baseurl = "https://privateapi.tdsmaker.com/api/v2";

    const response = UrlFetchApp.fetch(`${baseurl}${url}`, {
        method: 'get',
        headers: token === null || typeof token === "undefined" ? {} : {
            Authorization: 'Bearer ' + token.trim()
        }
    });

    return JSON.parse(response.getContentText());
}

function include(filename: string) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

function loadPage(filename: string) {
    var ui = SpreadsheetApp.getUi();

    var sidebarContent = HtmlService.createTemplateFromFile(filename)
        .evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle("TDSMaker")
        .setWidth(300);

    ui.showSidebar(sidebarContent);
}

namespace Log {
    class Log {
        private format: string;
        private parameters: Array<any>;

        constructor(format: string, values: Array<any>) {
            this.format = format;
            this.parameters = values;
        }
    }

    export function send(format: string, ...values: Array<any>): Log {
        let log = (new Log(format, values));
        Logger.log(format, ...values);
        // send request to API
        return log;
    }
}

class TemplateData {
    private templateID: string;
    private rows: Array<any>;

    constructor(templateID: string, data: Array<any>) {
        this.templateID = templateID;
        this.rows = data;
    }
}