import * as React from 'react';
import * as fluentUI from '@fluentui/react';
import { Checkbox, Stack, Label, PrimaryButton, ThemeSettingName } from '@fluentui/react'
import { TextField } from '@fluentui/react';




const MainFormular = () => {
    const [ustID, setustID] = React.useState("");
    const [enableUStID, setEnableUStID] = React.useState(false)

    const handleSubmit = (event) => {
        event.preventDefault();
    }

    const handleButtonClick = (event) => {
        callAPIandFillExcel(ustID)
    }

    const handleCheckboxChange = () => {
        setEnableUStID(!enableUStID)
    }
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
            <PrimaryButton text = "Prüfen" onClick= { handleButtonClick } />
            
        </Stack>
    )
}

const callAPIandFillExcel = async (requesterVATID) => { // TODO 16.03. FIX this async Await. Look at the book for guidance.
    await Excel.run(async(myExcelInstance) => {
        let range = myExcelInstance.workbook.getSelectedRange();
        range.load("text");
        const worksheets = myExcelInstance.workbook.worksheets; //used later to determine name of new sheet
        worksheets.load("items/name");
        await myExcelInstance.sync();
        console.log("Range Text: " + range.text)
        const selectedVatIDs = [];
        range.text.forEach((vatIDFromCell) => {
            console.log("logging-3")
            console.log(vatIDFromCell);
            selectedVatIDs.push(vatIDFromCell)
        });
        console.log(selectedVatIDs)
    //create the json to post
        const ownAPIJsons = [];
        selectedVatIDs.forEach((vatID) => {
            ownAPIJsons.push(JSON.stringify({
                vatID: vatID,
                traderName: "",
                traderCompanyType: "",
                traderStreet: "",
                traderPostcode: "",
                traderCity: "",
                requestervatID : requesterVATID
            }));
        });
        console.log(ownAPIJsons)
        const apiJSONResponse = await Promise.all(
            ownAPIJsons.map(makeTheAPICall)
        )
        console.log(apiJSONResponse[0])

        //now write down the result on a new sheet:
        let ws
        const worksheetNames = [];
        try { 
            worksheets.items.forEach((worksheet => {
                worksheetNames.push(worksheet.name);
            }))
            console.log(worksheetNames)
            for (let i = 0; i < 10; i++) {
                if (! worksheetNames.includes("VAT_IDs_by_Heegs_" + String(i))) {
                    ws = myExcelInstance.workbook.worksheets.add("VAT_IDs_by_Heegs_" + String(i));
                    break;
                };
            }
            

        } catch (error) {
            console.error(error);
            throw (error);
        };
        let vatIDTableHead = ws.getRange("A1:F1");
        vatIDTableHead.values = [["VatID", "Country", "isValid", "belongsToCompany", "Adress", "ConstelationID"]];
        vatIDTableHead.format.fill.color = "grey";
        let lengthOfReturn = apiJSONResponse.length
        console.log(lengthOfReturn)
        //await myExcelInstance.sync() //TODO maybe unnessesary
        for (let i = 0; i <= lengthOfReturn; i++) {
            let currentRow = i + 2
            let vatIDTableBodyRow = ws.getRange("A" + String(currentRow) + ":G" + String(currentRow)); 
            vatIDTableBodyRow.values = [apiJSONResponse[i]]
        }
        

    })//.catch(function (error) {
        //console.log("Error: " + error)
        //throw (error);
    //});
    
    //vatIDTableBody.values = [[selectedVatIDs[0], result.countryCode, result.valid, result.traderName, result.traderAddress, result.requestIdentifier]];
    return await myExcelInstance.sync();
};

async function makeTheAPICall(apiJSON) {
    try {
        let response = await fetch ("https://checkvatfirst.azurewebsites.net/api/httpTriggerOne", {
            method: "POST",
            header:{ 'Content-Type': 'application/json' },
            mode: "cors",
            body: apiJSON
        });
        return await response.json();
    } catch(error) {
        console.error(error);
        throw (error);
    };
}

export { MainFormular }