import * as React from 'react';
import { useEffect } from 'react';
import { Checkbox, Stack, Label, PrimaryButton,  ThemeSettingName, Text } from '@fluentui/react'
import { TextField } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MyMessageBar } from './MyMessageBar';
import { CellBinders } from './CellBinders';

var myConfig = require('../../../config.json'); //TODO Multilanguage Support




const MainFormular = () => {
    const [ustID, setustID] = React.useState("");
    const [enableUStID, setEnableUStID] = React.useState(false)
    const [isLoading, setIsLoading] = React.useState(false)
    const [returnMessage, setReturnMessage] = React.useState("");
    const [successMessage, setSuccessMessage] = React.useState("");
    const [VatRange, setVatRange] = React.useState("");
    const [CitiesRange, setCitiesRange] = React.useState("");
    const [AreaCodeRange, setAreaCodeRange] = React.useState("");
    const [CompanyNames, setCompanyNameRange] = React.useState("");
    const [CompanyTypes, setCompanyTypeRange] = React.useState("");
    

    const handleSubmit = (event) => { //unused.
        event.preventDefault();
    }


    const handleButtonClick = async () => {
        try {
            setReturnMessage("");
            setSuccessMessage("");
            setIsLoading(true);
            if (enableUStID) {
                await callAPIandFillExcelQualified(ustID);
            } else {
                await callAPIandFillExcel(ustID);
            }
            
            setIsLoading(false);
            setReturnMessage("");
            setSuccessMessage(myConfig.SuccessMessage)
        } catch (error){
            console.log(error); 
            setReturnMessage(error.message)
        } finally {
            setIsLoading(false);
        }
        
    }

    const handleCheckboxChange = () => {
        setEnableUStID(!enableUStID)
    }

    const handleMessageBarDismiss = () => {
        setReturnMessage("");
        setSuccessMessage("");
    };

    
    return (
        <Stack> 
            
            <Checkbox label={myConfig.CheckboxLabel} value = {enableUStID} onChange= { handleCheckboxChange }/> 
            <Label> 
                    {myConfig.TextFieldLabelText}
            <TextField 
                prefix={myConfig.TextFieldPrefix}
                disabled={ !enableUStID }
                onChange = { (e) => setustID(e.target.value) }
                value = { ustID }
            />
            </Label>
            <CellBinders EnableBindings= { enableUStID } VatRange={ VatRange} setVatRange = { setVatRange} CitiesRange = {CitiesRange} 
                setCitiesRange = {setCitiesRange} AreaCodeRange = {AreaCodeRange} setAreaCodeRange = {setAreaCodeRange} CompanyNames = {CompanyNames} setCompanyNameRange = { setCompanyNameRange}
                CompanyTypes = {CompanyTypes} setCompanyTypeRange= {setCompanyTypeRange}></CellBinders>
            <PrimaryButton text = {myConfig.SendButtonText} onClick= { handleButtonClick } disabled= { isLoading }/>
            { isLoading ? <Spinner label= {myConfig.SpinnerInitialText} size={ SpinnerSize.medium  } /> : null }
            { returnMessage ? <MyMessageBar message = { returnMessage } messageType = "Error" handleMessageBarDismiss= {handleMessageBarDismiss}/> : null }
            { successMessage ? <MyMessageBar message = { successMessage } messageType = "Success" handleMessageBarDismiss= {handleMessageBarDismiss}/> : null }
        </Stack>
    )
}

const callAPIandFillExcel = async (requesterVATID) => { 
    await Excel.run(async(myExcelInstance) => { 
        let vatRange
        vatRange = myExcelInstance.workbook.bindings.getItem("bindingVatIdsRange").getRange();
        vatRange.load("text");
        const worksheets = myExcelInstance.workbook.worksheets; //used later to determine name of new sheet
        worksheets.load("items/name");
        try {
            await myExcelInstance.sync();
        } catch (error) {
            if (error.code == "ItemNotFound") {
                throw new Error ("Item not found - Please select a Vat IDs Range.")
            } else {
                throw error;
            }
        }

        const selectedVatIDs = vatRange.text;

    //create the javascript object to post
        const ownAPIJsons = [];
        for ( let i = 0; i< selectedVatIDs.length; i++) {
            if (selectedVatIDs[i] !== "") {
                ownAPIJsons.push({
                    vatID: selectedVatIDs[i][0],
                    traderName: "",
                    traderCompanyType: "",
                    traderStreet: "",
                    traderPostcode: "",
                    traderCity: "",
                    requestervatID : requesterVATID
                })
            } 
        }
        //break it up in Chunks
        
        const maxAPICalls = myConfig.MaxApiCalls
        if (ownAPIJsons.length > maxAPICalls){
            let maxAPICallError = new Error("You can send a Maximum of " + String(maxAPICalls) + " Vat IDs over the App at once. Pls send them in several parts or refer to the developer.");
            throw maxAPICallError;
        }
        const chunkSize = myConfig.chunkSize;
        const apiReturnPromises = [];
        for (let i = 0; i < ownAPIJsons.length; i += chunkSize) {
            const chunk = ownAPIJsons.slice(i, i + chunkSize);
            apiReturnPromises.push(makeTheAPICall(chunk));
        }; //TODO Test what happends if ChunkSize > lenght.
        const apiResponses = await Promise.all(apiReturnPromises);
        
        //check if one gave back wrong status
        for (const thisApiResponse of apiResponses) { //foreach() does not work with Async Await as expected.
            if (thisApiResponse.status != 200 && thisApiResponse.status != 201) {
                let statusResponseText = await thisApiResponse.text();
                console.log(thisApiResponse)
                console.log({statusResponseText})
                let apiCallResponseError = new Error("The API returned an Error with Status Code " + String(thisApiResponse.status)+ "-"+statusResponseText+"- Please refer to the Developer if you cannot resolve this issue.")
                throw apiCallResponseError //TODO Who catches this Error?
            }
        };
        
        //grep all the JSONS
        const apiJSONResponses = await Promise.all(apiResponses.map(responseObject => responseObject.json()))
        const apiAllJSONResponses = apiJSONResponses.reduce( //add together all the JSONs from the different Responses
            (previousValue, currentValue) => {
                return previousValue.concat(currentValue);
            }, []
        );
        
        
        
        
        //now write down the result on a new sheet:
        
        let ws
        let wsCreated = false;
        let numberOfRuns = 0;
        const worksheetNames = [];
        try { 
            
            worksheets.items.forEach((worksheet => {
                worksheetNames.push(worksheet.name);
            }))
            for (let i = 0; i < 10; i++) {
                if (! worksheetNames.includes(myConfig.NewSheetsNamePrefix + String(i))) {
                    ws = myExcelInstance.workbook.worksheets.add(myConfig.NewSheetsNamePrefix + String(i));
                    numberOfRuns = String(i);
                    wsCreated = true;
                    break;
                };
            }
            if (wsCreated == false) { 
                ws = myExcelInstance.workbook.worksheets.add() //Excel determines the name if more than 10 sheets
            };
           
        } catch (error) {
            console.error(error);
            throw (error);
        };

        //create a table for this.
        let returnTable = ws.tables.add("A1:F1", true); //any name
        
        returnTable.getHeaderRowRange().values = [["VatID", "Country", "isValid", "TraderName", "Address", "ConstelationID"]];
        
        //insert the values
        let lengthOfReturn = apiAllJSONResponses.length;
        console.log(ownAPIJsons[1])
        for (let i = 0; i < lengthOfReturn; i++) {
            let thisRowValues = []
            if(apiAllJSONResponses[i]) {
                console.log(apiAllJSONResponses[i])
                thisRowValues = [[apiAllJSONResponses[i].countryCode + apiAllJSONResponses[i].vatNumber, apiAllJSONResponses[i].countryCode, apiAllJSONResponses[i].valid, apiAllJSONResponses[i].traderName, apiAllJSONResponses[i].traderAddress, apiAllJSONResponses[i].requestIdentifier]];
            } else {
                console.log("test")
                thisRowValues = [ownAPIJsons[i].vatID, "", "not a VatID", "","", ""];
                console.log({thisRowValues})
            };
            returnTable.rows.add(null, thisRowValues);
        };
        
        //format the table
        ws.getUsedRange().format.autofitColumns();
        ws.getUsedRange().format.autofitRows();
        await myExcelInstance.sync();
    }).catch(function (error) {
        console.log("Error: " + error +  " +++++ " + error.stack)
        throw (error);
    });
    
    return "done"; 
};
        
    

const callAPIandFillExcelQualified = async (requesterVATID) => { 
    await Excel.run(async(myExcelInstance) => { 
        let vatRange
        let citiesRange 
        let AreaCodesRange
        let companyNamesRange
        let companyTypeRange
        try{
            vatRange = myExcelInstance.workbook.bindings.getItem("bindingVatIdsRange").getRange()
            citiesRange = myExcelInstance.workbook.bindings.getItem("CitiesRange").getRange()
            AreaCodesRange = myExcelInstance.workbook.bindings.getItem("AreaCodesRange").getRange()
            companyNamesRange = myExcelInstance.workbook.bindings.getItem("CompanyNames").getRange()
            companyTypeRange = myExcelInstance.workbook.bindings.getItem("CompanyTypes").getRange()
        } catch (e) {
            console.log(e)
        }
        vatRange.load("text");
        citiesRange.load("text");
        AreaCodesRange.load("text");
        companyNamesRange.load("text");
        companyTypeRange.load("text");
        const worksheets = myExcelInstance.workbook.worksheets; //used later to determine name of new sheet
        worksheets.load("items/name");
        await myExcelInstance.sync();
        const selectedVatIDs = vatRange.text;
        const selectedCities = citiesRange.text;
        const selectedAreaCodes = AreaCodesRange.text;
        const selectedNames = companyNamesRange.text;
        const selectedTypes = companyTypeRange.text;
        
    //create the javascript object to post
        const ownAPIJsons = [];
        for ( let i = 0; i++; i< selectedVatIDs.length) {
            ownAPIJsons.push({
                vatID: selectedVatIDs[i],
                traderName: selectedNames[i],
                traderCompanyType: selectedTypes[i],
                traderStreet: "", //TODO fehlt noch
                traderPostcode: selectedAreaCodes[i],
                traderCity: selectedCities[i],
                requestervatID : requesterVATID
            }) 
        }


        //break it up in Chunks
        
        const maxAPICalls = myConfig.MaxApiCalls
        if (ownAPIJsons.length > maxAPICalls){
            let maxAPICallError = new Error("You can send a Maximum of " + String(maxAPICalls) + " Vat IDs over the App at once. Pls send them in several parts or refer to the developer.");
            throw maxAPICallError;
        }
        const chunkSize = myConfig.chunkSize;
        const apiReturnPromises = [];
        for (let i = 0; i < ownAPIJsons.length; i += chunkSize) {
            const chunk = ownAPIJsons.slice(i, i + chunkSize);
            apiReturnPromises.push(makeTheAPICall(chunk));
        };
        const apiResponses = await Promise.all(apiReturnPromises);
        
        //check if one gave back wrong status
        for (const thisApiResponse of apiResponses) { //foreach() does not work with Async Await as expected.
            if (thisApiResponse.status != 200 && thisApiResponse.status != 201) {
                let statusResponseText = await thisApiResponse.text();
                console.log(thisApiResponse)
                console.log({statusResponseText})
                let apiCallResponseError = new Error("The API returned an Error with Status Code " + String(thisApiResponse.status)+ "-"+statusResponseText+"- Please refer to the Developer if you cannot resolve this issue.")
                throw apiCallResponseError
            }
        };
        
        //grep all the JSONS
        const apiJSONResponses = await Promise.all(apiResponses.map(responseObject => responseObject.json()))
        const apiAllJSONResponses = apiJSONResponses.reduce( //add together all the JSONs from the different Responses
            (previousValue, currentValue) => {
                return previousValue.concat(currentValue);
            }, []
        );
        
        
        
        
        //now write down the result on a new sheet:
        
        let ws
        let wsCreated = false;
        let numberOfRuns = 0;
        const worksheetNames = [];
        try { 
            
            worksheets.items.forEach((worksheet => {
                worksheetNames.push(worksheet.name);
            }))
            for (let i = 0; i < 10; i++) {
                if (! worksheetNames.includes(myConfig.NewSheetsNamePrefix + String(i))) {
                    ws = myExcelInstance.workbook.worksheets.add(myConfig.NewSheetsNamePrefix + String(i));
                    numberOfRuns = String(i);
                    wsCreated = true;
                    break;
                };
            }
            if (wsCreated == false) { 
                ws = myExcelInstance.workbook.worksheets.add() //Excel determines the name if more than 10 sheets
            };
           
        } catch (error) {
            console.error(error);
            throw (error);
        };

        //create a table for this.
        let returnTable = ws.tables.add("A1:F1", true); //any name
        
        returnTable.getHeaderRowRange().values = [["VatID", "Country", "isValid", "TraderName", "Address", "ConstelationID"]];
        
        //insert the values
        let lengthOfReturn = apiAllJSONResponses.length;
        console.log(ownAPIJsons[1])
        for (let i = 0; i < lengthOfReturn; i++) {
            let thisRowValues = []
            if(apiAllJSONResponses[i]) {
                console.log(apiAllJSONResponses[i])
                thisRowValues = [[apiAllJSONResponses[i].countryCode + apiAllJSONResponses[i].vatNumber, apiAllJSONResponses[i].countryCode, apiAllJSONResponses[i].valid, apiAllJSONResponses[i].traderName, apiAllJSONResponses[i].traderAddress, apiAllJSONResponses[i].requestIdentifier]];
            } else {
                thisRowValues = [ownAPIJsons[i].vatID, "", "not a VatID", "","", ""];
                
            };
            returnTable.rows.add(null, thisRowValues);
        };
        
        //format the table
        ws.getUsedRange().format.autofitColumns();
        ws.getUsedRange().format.autofitRows();
        await myExcelInstance.sync();
    }).catch(function (error) {
        console.log("Error: " + error +  " +++++ " + error.stack)
        throw (error);
    });
    
    return "done"; 
};

async function makeTheAPICall(apiJSON) {
    console.log(JSON.stringify(apiJSON))
    try {
        let response = await fetch (myConfig.APIAdress, {
            method: "POST",
            header:{ 'Content-Type': 'application/json' },
            mode: "cors", //i could make this better and safer, not using cors but a backend to call the api
            body: JSON.stringify(apiJSON) //sends the javascript object as  JSON String
        });
        return response; //this just returns the promise, you have to await it to use.
        
    } catch(error) {
        throw (error);
    };
        
    
}

export { MainFormular }