import * as React from 'react';
import { useEffect } from 'react';
import * as fluentUI from '@fluentui/react';
import { Checkbox, Stack, Label, PrimaryButton, ThemeSettingName } from '@fluentui/react'
import { TextField } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MyMessageBar } from './MyMessageBar';




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
            setIsLoading(true);
            await callAPIandFillExcel(ustID);
            setIsLoading(false);
            setReturnMessage("");
            setSuccessMessage("Erfolgreich!")
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
            
            <Checkbox label="Qualifiziert Prüfung" value = {enableUStID} onChange= { handleCheckboxChange }/> 
            <Label> 
                    Identifikation
            <TextField 
                prefix='Eigene USt-ID'
                disabled={ !enableUStID }
                onChange = { (e) => setustID(e.target.value) }
                value = { ustID }
            />
            </Label>
            <PrimaryButton text = "Prüfen" onClick= { handleButtonClick } disabled= { isLoading }/>
            { isLoading ? <Spinner label='Checking VAT IDs' size={ SpinnerSize.medium  } /> : null }
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
            ownAPIJsons.push(JSON.stringify({
                vatID: vatID[0],
                traderName: "",
                traderCompanyType: "",
                traderStreet: "",
                traderPostcode: "",
                traderCity: "",
                requestervatID : requesterVATID
            }))
        };
        });
        ///const apiJSONResponse = await Promise.all( /// this is pretty cool, but i dont want to send every line by itself.
         //   ownAPIJsons.map(makeTheAPICall)
        //)

        //break it up in Chunks of 40
        
        const maxAPICalls = 2000
        if (ownAPIJsons.length > maxAPICalls){
            let maxAPICallError = new Error("You can send a Maximum of " + String(maxAPICalls) + " Vat IDs over the App at once. Pls send them in Chunks");
            throw maxAPICallError;
        }
        const chunkSize = 2;
        const apiReturnPromises = [];
        for (let i = 0; i < ownAPIJsons.length; i += chunkSize) {
            const chunk = ownAPIJsons.slice(i, i + chunkSize);
            apiReturnPromises.push(makeTheAPICall(chunk));
        };
        const apiResponses = await Promise.all(apiReturnPromises);
        //check if one gave back wrong status
        apiResponses.forEach(apiResponse => {
            if (apiResponse.status != 200 && apiResponse.status != 201) {
                let apiCallResponseError = new Error("The API returned an Error with Status Code " + String(apiStatusResponse))
                throw apiCallResponseError
            }
        });
        
        //grep all the JSONS
        const apiJSONResponses = await Promise.all(apiResponses.map(responseObject => responseObject.json()))
        const apiAllJSONResponses = apiJSONResponses.reduce(
            (previousValue, currentValue) => {
                console.log({previousValue})
                console.log({currentValue})
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
                if (! worksheetNames.includes("VAT_IDs_by_Heegs_" + String(i))) {
                    ws = myExcelInstance.workbook.worksheets.add("VAT_IDs_by_Heegs_" + String(i));
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
        for (let i = 0; i < lengthOfReturn; i++) {
            let thisRowValues = [[apiAllJSONResponses[i].countryCode + apiAllJSONResponses[i].vatNumber, apiAllJSONResponses[i].countryCode, apiAllJSONResponses[i].valid, apiAllJSONResponses[i].traderName, apiAllJSONResponses[i].traderAddress, apiAllJSONResponses[i].requestIdentifier]];
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
        let response = await fetch ("https://checkvatfirst.azurewebsites.net/api/httpTriggerOne", {
            method: "POST",
            header:{ 'Content-Type': 'application/json' },
            mode: "cors", //i could make this better and safer, not using cors but a backend to call the api
            body: JSON.stringify(apiJSON)
        });
        return response; //this just returns the promise, you have to await it to use.
        
    } catch(error) {
        throw (error);
    };
        
    
}

export { MainFormular }