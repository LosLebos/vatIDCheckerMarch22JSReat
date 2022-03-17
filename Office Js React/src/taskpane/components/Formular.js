import * as React from 'react';
import * as fluentUI from '@fluentui/react';
import { Checkbox, Stack, Label, PrimaryButton } from '@fluentui/react'
import { TextField } from '@fluentui/react';




const MainFormular = () => {
    const [ustID, setustID] = React.useState("");
    const [enableUStID, setEnableUStID] = React.useState(false)

    const handleSubmit = (event) => {
        event.preventDefault();
        //callAPIandFillExcel(ustID);
    }

    const handleButtonClick = (event) => {
        console.log("why")
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
    console.log("logging-2")
    await Excel.run(async(myExcelInstance) => {
        let range = myExcelInstance.workbook.getSelectedRange();
        range.load("text");
        await myExcelInstance.sync();
    });
    console.log("logging-1")
    const selectedVatIDs = [];
    console.log(range.text)
    range.text.forEach((cell) => {
        console.log("logging-3")
        console.log(cell);
        selectedVatIDs.push(cell.text)
    });
    console.log("logging-5")
    console.log(selectedVatIDs)
    console.log(typeof selectedVatIDs)
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
        

    //now write down the result on a new sheet:
    let ws
    try { 
        ws = myExcelInstance.workbook.worksheets.add("VAT_IDs_by_Heegs");
    } catch (error) {
        console.error(error);
        throw (error);
    };
    let vatIDTableHead = ws.getRange("A1:F1");
    vatIDTableHead.values = [["VatID", "Country", "isValid", "belongsToCompany", "Adress", "ConstelationID"]];
    vatIDTableHead.format.fill.color = "grey";
    let vatIDTableBody = ws.getRange("A2:F2"); //TODO wrong range
    let i = 0
    vatIDTableBody.forEach((cell) => {
        cell.text = apiJSONResponse[i];
        i++;
    });
    //vatIDTableBody.values = [[selectedVatIDs[0], result.countryCode, result.valid, result.traderName, result.traderAddress, result.requestIdentifier]];
    return myExcelInstance.sync();
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