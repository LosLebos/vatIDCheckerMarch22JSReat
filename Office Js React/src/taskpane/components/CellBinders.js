import * as React from 'react';
import { useEffect } from 'react';
import * as fluentUI from '@fluentui/react';
import { Stack, PrimaryButton, Text, DefaultButton, StackItem } from '@fluentui/react'
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
        props.setCompanyNameRange("");
        props.setCompanyTypeRange("");
      }, [props.EnableBindings]);
    
    const handleBindingButton= async(rangeType) => {
        Office.context.document.bindings.addFromPromptAsync(
            Office.BindingType.Matrix,
            { id: rangeType, promptText: 'Select the given ' & rangeType & ':' }
            //just create the binding via prompt over the common API 2013
            // in hindsight , i dont remember the difference between this API 2013 call and the later API 2016 call.
        )

        await Excel.run(async (context) => {
            try{
                let bindingRange = context.workbook.bindings.getItem(rangeType) //get the binding via the excel API 2016 to get an Excel.Binding object which has the getRange() method
                let range = bindingRange.getRange();
                range.load("address");
                range.select();
                await context.sync();
                props.setVatRange(range.address);
            } catch (error) {
                console.log(error.message)
            }
            
        })
    }
    
    const handleMessageBarDismiss = () => {
        setErrorMessage("");
    };

    return (
        <>
        <Stack vertical gap = { 5 } style = {{ marginBottom: 10 }}>
            <Stack horizontal horizontalAlign='center' style={{}}>
                <Stack.Item>
                    <DefaultButton
                    text='Vat IDs Range' onClick={ () => handleBindingButton("bindingVatIdsRange") }
        	        />
                </Stack.Item>
                <StackItem>
                    <TextField
                    //   prefix="Vat ID Range"
                        disabled = {true}
                        onChange = { (e) => props.setVatRange(e.target.value) }
                        value = { props.VatRange }>
                        </TextField>
                </StackItem>
                
            </Stack>
            <Stack horizontal style={{}} horizontalAlign = "left">
            <DefaultButton
                disabled={ !props.EnableBindings }
                text='Cities Range' onClick={ () =>handleBindingButton('CitiesRange') }
        	/>
            <TextField
             //   prefix="Cities Range"
                disabled={true}
                onChange = { (e) => props.setCitiesRange(e.target.value) }
                value = { props.CitiesRange }>
            </TextField>
            </Stack>
            <Stack horizontal style={{}} horizontalAlign = "left">
            <DefaultButton
                disabled={ !props.EnableBindings}
                text='Area Code Range' onClick={ () =>handleBindingButton('AreaCodesRange') }
        	/>
            <TextField 
            //prefix="Area Code Range"
            disabled={ true }
            onChange = { (e) => props.setAreaCodeRange(e.target.value) }
            value = { props.AreaCodeRange }>
            </TextField>
            </Stack>
            <Stack horizontal style={{}} horizontalAlign = "left">
            <DefaultButton
                disabled={ !props.EnableBindings}
                text='Company Name Range' onClick={ () =>handleBindingButton('CompanyNames') }
        	/>
            <TextField 
            //prefix="Area Code Range"
            disabled={ true }
            onChange = { (e) => props.setCompanyNameRange(e.target.value) }
            value = { props.AreaCodeRange }>
            </TextField>
            </Stack>
            <Stack horizontal style={{}} horizontalAlign = "left">
            <DefaultButton
                disabled={ !props.EnableBindings}
                text='Company Type Range' onClick={ () =>handleBindingButton('CompanyTypes') }
        	/>
            <TextField 
            //prefix="Area Code Range"
            disabled={ true }
            onChange = { (e) => props.setCompanyTypeRange(e.target.value) }
            value = { props.AreaCodeRange }>
            </TextField>
            </Stack>
        </Stack>
            { errorMessage ? <MyMessageBar message = { errorMessage} messageType = "Error" handleMessageBarDismiss= { handleMessageBarDismiss } /> : null }
        </>
    )
}


export { CellBinders }