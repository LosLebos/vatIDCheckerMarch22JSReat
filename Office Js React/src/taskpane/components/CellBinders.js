import * as React from 'react';
import { useEffect } from 'react';
import * as fluentUI from '@fluentui/react';
import { Stack, PrimaryButton, Text, DefaultButton } from '@fluentui/react'
import { TextField } from '@fluentui/react';
import { MyMessageBar } from './MyMessageBar';
var myConfig = require('../../../config.json');

const CellBinders = (props) => {
    // props:
    // EnableBindings Boolean
    // VatRange, setVatRange
    // CitiesRange, setCitiesRange
    // AreaCodeRange, setAreaCodeRange
    const [errorMessage, setErrorMessage] = React.useState("");

    useEffect(() => { 
        props.setVatRange("");
        props.setCitiesRange("");
        props.setAreaCodeRange("");
      }, [props.EnableBindings]);
    
    const handleBindingButtonVat= async() => {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Matrix,
            { id: 'bindingVatIdsRange', promptText: 'Select the given Vat Ids:' }
            //just create the binding via prompt over the common API 2013
        )

        await Excel.run(async (context) => {
            try{
                let bindingVatIdsRange = context.workbook.bindings.getItem("bindingVatIdsRange") //get the binding via the excel API 2016 to get an Excel.Binding object which has the getRange() method
                let range = bindingVatIdsRange.getRange();
                range.load("address");
                range.select();
                await context.sync();
                props.setVatRange(range.address);
            } catch (error) {
                console.log(error.message)
            }
            
        })
    }
    const handleBindingButtonCities= async() => {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Matrix,
            { id: 'bindingCitiesRange', promptText: 'Select the given Cities:' }
        )

        await Excel.run(async (context) => {
            try{
                let bindingCitiesRange = context.workbook.bindings.getItem("bindingCitiesRange") //get the binding via the excel API 2016 to get an Excel.Binding object which has the getRange() method
                let range = bindingCitiesRange.getRange();
                range.load("address");
                range.select();
                await context.sync();
                props.setCitiesRange(range.address);
            } catch (error) {
                console.log(error.message)
            }
            
        })
    }
    const handleBindingButtonAreaCodes= async() => {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Matrix,
            { id: 'bindingAreaCodeRange', promptText: 'Select the given AreaCode:' }
            //just create the binding via prompt over the common API 2013
        )

        await Excel.run(async (context) => {
            try{
                let bindingAreaCodeRange = await context.workbook.bindings.getItem("bindingAreaCodeRange") //get the binding via the excel API 2016 to get an Excel.Binding object which has the getRange() method
                let range = bindingAreaCodeRange.getRange();
                range.load("address");
                range.select();
                await context.sync();
                props.setAreaCodeRange(range.address);
            } catch (error) {
                setErrorMessage = error.message;
            }
            
        })
    }
    const handleMessageBarDismiss = () => {
        setErrorMessage("");
    };

    return (
        <>
            <Stack horizontal horizontalAlign='start'>
                <DefaultButton
                
                text='Vat IDs Binding' onClick={ () => handleBindingButtonVat() }
        	    />
                <TextField
                prefix="Vat ID Range"
                
                onChange = { (e) => props.setVatRange(e.target.value) }
                value = { props.VatRange }>
                </TextField>
            </Stack>
            <Stack horizontal>
            <DefaultButton
                disabled={ !props.EnableBindings}
                text='Cities Binding' onClick={ () =>handleBindingButtonCities() }
        	/>
            <TextField
                prefix="Cities Range"
                disabled={ !props.EnableBindings}
                onChange = { (e) => props.setCitiesRange(e.target.value) }
                value = { props.CitiesRange }>
            </TextField>
            </Stack>
            <Stack horizontal>
            <DefaultButton
                disabled={ !props.EnableBindings}
                text='Area Code Binding' onClick={ () =>handleBindingButtonAreaCodes() }
        	/>
            <TextField 
            prefix="Area Code Range"
            disabled={ !props.EnableBindings}
            onChange = { (e) => props.setAreaCodeRange(e.target.value) }
            value = { props.AreaCodeRange }>
            </TextField>
            </Stack>
            { errorMessage ? <MyMessageBar message = { errorMessage} messageType = "Error" handleMessageBarDismiss= { handleMessageBarDismiss } /> : null }
        </>
    )
}


export { CellBinders }