import * as React from 'react';
import { useEffect } from 'react';
import * as fluentUI from '@fluentui/react';
import { Checkbox, Stack, Label, PrimaryButton, ThemeSettingName, Text } from '@fluentui/react'
import { TextField } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MyMessageBar } from './MyMessageBar';

var myConfig = require('../../../config.json');




const MainFormular = () => {
    const [ustID, setustID] = React.useState("");
    const [enableUStID, setEnableUStID] = React.useState(false)
    const [isLoading, setIsLoading] = React.useState(false)
    const [returnMessage, setReturnMessage] = React.useState("");
    const [successMessage, setSuccessMessage] = React.useState("");
    

    const handleSubmit = (event) => {
        event.preventDefault();
    }


    const handleButtonClick = async () => {
        try {
            setReturnMessage("");
            setSuccessMessage("");
            setIsLoading(true);
            await callAPIandFillExcel(ustID);
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
    const handleBindingButton= async() => {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Matrix,
            { id: 'bindingVatIdsRange', promptText: 'Select the given Vat IDs:' }
            //just create the binding via prompt over the common API 2013
        )
        await Excel.run(async (context) => {
            try{
                let bindingVatIdsRange = context.workbook.bindings.getItem("bindingVatIdsRange") //get the binding via the excel API 2016 to get an Excel.Binding object which has the getRange() method
                let range = bindingVatIdsRange.getRange();
                range.load("address");
                range.select();
                await context.sync();
                console.log(range.address)
            } catch (error) {
                console.log(error.message)
            }
            
        })
    }
    const handleBindingButtonOfficeAPI = () => {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Matrix,
            { id: 'MyBinding', promptText: 'Select text to bind to.' },
            function (asyncResult) {
                console.log(asyncResult.status); //Todo find the values in the MAtrix
                console.log(asyncResult.getRange())
            }
        )
    }
    
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
            <PrimaryButton text = {myConfig.SendButtonText} onClick= { handleButtonClick } disabled= { isLoading }/>
            <PrimaryButton text='Test Binding' onClick={ handleBindingButton }/>
            <Text id='message'/>
            { isLoading ? <Spinner label= {myConfig.SpinnerInitialText} size={ SpinnerSize.medium  } /> : null }
            { returnMessage ? <MyMessageBar message = { returnMessage } messageType = "Error" handleMessageBarDismiss= {handleMessageBarDismiss}/> : null }
            { successMessage ? <MyMessageBar message = { successMessage } messageType = "Success" handleMessageBarDismiss= {handleMessageBarDismiss}/> : null }
        </Stack>
    )
}

const callAPIandFillExcel = async (requesterVATID) => { 
    await Excel.run(async(myExcelInstance) => { 
        let range = myExcelInstance.workbook.getSelectedRange(); //throws an InvalidSelection Error if multiple selects
        range.load("text");
        const worksheets = myExcelInstance.workbook.worksheets; //used later to determine name of new sheet
        worksheets.load("items/name");
        await myExcelInstance.sync();
        const selectedVatIDs = range.text
        
        
    //create the json to post
        const ownAPIJsons = [];
        
        selectedVatIDs.forEach((vatID) => {
            if (vatID[0] != "") {
            ownAPIJsons.push({
                vatID: vatID[0],
                traderName: "",
                traderCompanyType: "",
                traderStreet: "",
                traderPostcode: "",
                traderCity: "",
                requestervatID : requesterVATID
            })//creates a JavaScriptObject
        };
        });

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

async function makeTheAPICall(apiJSON) {
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