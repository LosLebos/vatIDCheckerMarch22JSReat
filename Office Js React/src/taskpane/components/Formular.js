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

const callAPIandFillExcel = async ( requesterVATID) => {
    await Excel.run(async(myExcelInstance) => {
        let range = myExcelInstance.workbook.getSelectedRange();
        range.load("text");
        await myExcelInstance.sync();
        let result;
        let selectedVatID = range.text[0];
        console.log(selectedVatID[0])
        console.log(typeof selectedVatID)
    //create the json to post
        const ownAPIJson = JSON.stringify({
            vatID: selectedVatID[0],
            traderName: "",
            traderCompanyType: "",
            traderStreet: "",
            traderPostcode: "",
            traderCity: "",
            requestervatID : requesterVATID
        });
        console.log(ownAPIJson)
        try {
            const response = await fetch ("https://checkvatfirst.azurewebsites.net/api/httpTriggerOne", {
                method: "POST",
                header:{ 'Content-Type': 'application/json' },
                mode: "cors",
                body: ownAPIJson
            });
            result = await response.json();
            console.log(result);
        } catch(error) {
            console.error(error);
            throw (error);
        };

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
        vatIDTableBody.values = [[selectedVatID[0], result.countryCode, result.valid, result.traderName, result.traderAddress, result.requestIdentifier]];
        return myExcelInstance.sync();
    });  
}

export { MainFormular }