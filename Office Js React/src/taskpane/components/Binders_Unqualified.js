import * as React from 'react';
import { useEffect } from 'react';
import * as fluentUI from '@fluentui/react';
import { Stack, PrimaryButton, Text, DefaultButton, StackItem } from '@fluentui/react'
import { TextField } from '@fluentui/react';
import { MyMessageBar } from './MyMessageBar';
var myConfig = require('../../../config.json');

const CellBinders_Unqualified = (props) => {
    // props:
    // VatRange, setVatRange
    // CitiesRange, setCitiesRange
    // AreaCodeRange, setAreaCodeRange
    // CompanyNames, setCompanyNameRange
    // CompanyTypes , setCompanyTypeRange
    
    const [errorMessage, setErrorMessage] = React.useState("");

    ///useEffect(() => { //Sets all Props to "" whenever EnableBindings changes, with the exception of the VatRange // I took this out might not be wanted.
        //props.setVatRange("");
        //props.setCitiesRange("");
        //props.setAreaCodeRange("");
        //props.setCompanyNameRange("");
        //rops.setCompanyTypeRange("");
      //}, [props.EnableBindings]);
    
    const handleSelectionChangeOfBindings = async (binding) => {
        //this is unused, I left it in so I know how to use a handler if wanted.
        //console.log(binding.id)
    }
    const handleBindingButton= async(rangeType, setterOfTheRangeToChange) => {
        let selectedRange;
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Matrix,
            { id: rangeType, promptText: 'Select the given ' + rangeType }
            //just create the binding via prompt over the common API 2013
            // the Office API is older. I nevertheless use is to create the binding. The Excel API is newer and I later use it to work with the Bindings.
        );

        await Excel.run(async (context) => {
            try{
                let binding = context.workbook.bindings.getItem(rangeType) //get the binding via the excel API 2016 to get an Excel.Binding object which has the getRange() method
                let range = binding.getRange();
                binding.onSelectionChanged.add(handleSelectionChangeOfBindings);
                range.load("address");
                range.select();
                await context.sync();
                selectedRange =  range.address;
            } catch (error) {
                console.log(error.message);
                selectedRange = "";
            } finally {
                setterOfTheRangeToChange(selectedRange); //in JS you can give Functions in Variables. The setterOfTheRangeToChange represents the function to change the correct State.

            }
            
        })

    }
    
    const handleMessageBarDismiss = () => {
        setErrorMessage("");
    };

    return (
        <>
        <Stack vertical gap = { 5 } style = {{ marginBottom: 10 }}>
            <Stack.Item horizontal horizontalAlign='left' style={{}}>
                
                <DefaultButton
                    text='Vat IDs Range' onClick={ () => handleBindingButton("VatIDs", props.setVatRange) }
        	        />
                <TextField
                //   prefix="Vat ID Range"
                    disabled = {true}
                    onChange = { (e) => props.setVatRange(e.target.value) }
                    value = { props.VatRange }>
                </TextField>
                
            </Stack.Item>
        </Stack>
            { errorMessage ? <MyMessageBar message = { errorMessage} messageType = "Error" handleMessageBarDismiss= { handleMessageBarDismiss } /> : null }
        </>
    )
}


export { CellBinders_Unqualified }