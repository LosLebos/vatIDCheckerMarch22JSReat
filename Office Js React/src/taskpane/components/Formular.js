import * as React from 'react';
import { useEffect } from 'react';
import * as fluentUI from '@fluentui/react';
import { Checkbox, Stack, Label, PrimaryButton, ThemeSettingName } from '@fluentui/react'
import { TextField } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MessageBar, MessageBarType } from '@fluentui/react';



const MainFormular = () => {
    const [ustID, setustID] = React.useState("");
    const [enableUStID, setEnableUStID] = React.useState(false)
    const [isLoading, setIsLoading] = React.useState(false)
    const [errorMessage, setErrorMessage] = React.useState("");
    

    const handleSubmit = (event) => {
        event.preventDefault();
    }


    const handleButtonClick = async () => {
        try {
            setIsLoading(true);
            await callAPIandFillExcel(ustID);
            setIsLoading(false);
            setErrorMessage("");
        } catch (error){
            console.log(error); 
            debugger;
            setErrorMessage(error)
        } finally {
            setIsLoading(false);
        }
        
    }

    const handleCheckboxChange = () => {
        setEnableUStID(!enableUStID)
    }

    const handleMessageBarDismiss = () => {
        setErrorMessage("");
    }
    const MyMessageBar = (props) => ( //without {} means return value directly
    //i would rather have it as its own component file but then iwould have to use setErrorMessage()
    //TODO untested due to no Error
    <MessageBar
      messageBarType={MessageBarType.error}
      isMultiline={true}
      onDismiss= { handleMessageBarDismiss }
      dismissButtonAriaLabel="Close"
    >
      { props.message }
    </MessageBar>
  );

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
            { errorMessage ? <MyMessageBar message = { errorMessage }/> : null }
        </Stack>
    )
}

const callAPIandFillExcel = async (requesterVATID) => { 
    await Excel.run(async(myExcelInstance) => { //this probably shit to execute all the code within this await???
        let range = myExcelInstance.workbook.getSelectedRange(); //throws an InvalidSelection Error if multiple selects
        range.load("text");
        const worksheets = myExcelInstance.workbook.worksheets; //used later to determine name of new sheet
        worksheets.load("items/name");
        await myExcelInstance.sync();
        console.log("rangetext", range.text)
        for (let i=0; i < range.text.length; i++){
            if (range.text[i] == ""){
                throw new Error("Please do not select empty Cells.");
            }
        }
        const selectedVatIDs = range.text
    //create the json to post
        const ownAPIJsons = [];
        
        selectedVatIDs.forEach((vatID) => {
            
            ownAPIJsons.push(JSON.stringify({
                vatID: vatID[0],
                traderName: "",
                traderCompanyType: "",
                traderStreet: "",
                traderPostcode: "",
                traderCity: "",
                requestervatID : requesterVATID
            }));
        });
        ///const apiJSONResponse = await Promise.all( /// this is pretty cool, but i dont want to send every line by itself.
         //   ownAPIJsons.map(makeTheAPICall)
        //)

        
        const apiResponse = await makeTheAPICall(ownAPIJsons)
        const apiStatusResponse = await apiResponse.status;
        let apiJSONResponse;
        if (apiStatusResponse == 200 || apiStatusResponse == 201) {
            apiJSONResponse = await apiResponse.json();
        } else {
           //apiJSONResponse = await apiResponse.json(); //TODO i still get the response here, because iam an idiot and do not know how to throw an error
            throw "The API returned an Error with STatus Code " + String(apiStatusResponse)
        }
        
        
        
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
            if (wsCreated = false) { //TODO does not work, never fires.
                ws = myExcelInstance.workbook.worksheets.add() //Excel determines the name.
            }
            }
        } catch (error) {
            console.error(error);
            throw (error);
        };

        //create a table for this.
        let returnTable = ws.tables.add("A1:F1", true);
        if (wsCreated) { //you cannot name two talbes the same.
            returnTable.name = "VAT_IDs_by_Heegs" + numberOfRuns;
        }
        
        returnTable.getHeaderRowRange().values = [["VatID", "Country", "isValid", "TraderName", "Address", "ConstelationID"]];
        //returnTable.getHeaderRowRange().format.fill.color = "grey";
        
        let lengthOfReturn = apiJSONResponse.length;
        for (let i = 0; i < lengthOfReturn; i++) {
            let thisRowValues = [[apiJSONResponse[i].countryCode + apiJSONResponse[i].vatNumber, apiJSONResponse[i].countryCode, apiJSONResponse[i].valid, apiJSONResponse[i].traderName, apiJSONResponse[i].traderAddress, apiJSONResponse[i].requestIdentifier]];
            returnTable.rows.add(null, thisRowValues);
            
        };
        
        ws.getUsedRange().format.autofitColumns();
        ws.getUsedRange().format.autofitRows();
        await myExcelInstance.sync();
    }).catch(function (error) {
        console.log("Error: " + error +  " +++++ " + error.stack)
        throw (error);
    });
    console.log("end")
    
    return "done"; 
};

async function makeTheAPICall(apiJSON) {
    try {
        let response = await fetch ("https://checkvatfirst.azurewebsites.net/api/httpTriggerOne", {
            method: "POST",
            header:{ 'Content-Type': 'application/json' },
            mode: "cors", //i could make this better and safer, not using cors but another backend to call the api
            body: JSON.stringify(apiJSON)
        });
        return response; //this just returns the promise, you have to await it to use.
        
    } catch(error) {
        if (error instanceof InvalidSelection) { //getselectedRange return multiple ranges
            //todo alert 
            console.error(error);
            throw (error);
        }
        console.error(error);
        throw (error);
    };
}

export { MainFormular }