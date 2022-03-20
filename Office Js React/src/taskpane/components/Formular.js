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
        console.log("rangetext", range.text)
        const selectedVatIDs = range.text
    //create the json to post
        const ownAPIJsons = [];
         // 17.03. wird irgendwie nicht angezeigt. Ausserdem ist VatID nächste Zeile irgendwie nen Tupel und die API mag das nicht.
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
        console.log("ownJSONAPI", ownAPIJsons);
        console.log("the other way: ",JSON.stringify(ownAPIJsons));
        ///const apiJSONResponse = await Promise.all( /// this is pretty cool, but i dont want to send every line by itself.
         //   ownAPIJsons.map(makeTheAPICall)
        //)
        const apiJSONResponse = await makeTheAPICall(ownAPIJsons)
        
        console.log("first apiJson response", apiJSONResponse[0])
        //now write down the result on a new sheet:
        let ws
        const worksheetNames = [];
        try { 
            worksheets.items.forEach((worksheet => {
                worksheetNames.push(worksheet.name);
            }))
            for (let i = 0; i < 10; i++) {
                if (! worksheetNames.includes("VAT_IDs_by_Heegs_" + String(i))) {
                    ws = myExcelInstance.workbook.worksheets.add("VAT_IDs_by_Heegs_" + String(i));
                    break;
                };
            //TODO warning more than 10 sheets
            }
            

        } catch (error) {
            console.error(error);
            throw (error);
        };
        let vatIDTableHead = ws.getRange("A1:F1");
        vatIDTableHead.values = [["VatID", "Country", "isValid", "TraderName", "Address", "ConstelationID"]];
        vatIDTableHead.format.fill.color = "grey";
        let lengthOfReturn = apiJSONResponse.length
        for (let i = 0; i < lengthOfReturn; i++) {
            console.log({i})
            let currentRow = i + 2
            let vatIDTableBodyRow = ws.getRange("A" + String(currentRow) + ":F" + String(currentRow)); 
            let thisRowValues = [[apiJSONResponse[i].countryCode + apiJSONResponse[i].vatNumber, apiJSONResponse[i].countryCode, apiJSONResponse[i].valid, apiJSONResponse[i].traderName, apiJSONResponse[i].traderAddress, apiJSONResponse[i].requestIdentifier]];
            vatIDTableBodyRow.values = thisRowValues
            vatIDTableBodyRow.format.fill.color = "green";
            console.log("thisRowValues" , thisRowValues);
            await myExcelInstance.sync();
        }
        

    }).catch(function (error) {
        console.log("Error: " + error)
        throw (error);
    });
    
    return await myExcelInstance.sync();
};

async function makeTheAPICall(apiJSON) {
    console.log({apiJSON})
    console.log(typeof apiJSON)
    try {
        let response = await fetch ("https://checkvatfirst.azurewebsites.net/api/httpTriggerOne", {
            method: "POST",
            header:{ 'Content-Type': 'application/json' },
            mode: "cors",
            body: JSON.stringify(apiJSON)
        });
        //console.log("response: " , await response.json())
        return await response.json();
        
    } catch(error) {
        console.error(error);
        throw (error);
    };
}

export { MainFormular }