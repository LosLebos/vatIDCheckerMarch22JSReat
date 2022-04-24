import * as React from 'react';
import { TextField } from '@fluentui/react';
import { Checkbox, Stack, Label, PrimaryButton,  ThemeSettingName, Text } from '@fluentui/react'

var myConfig = require('../../../config.json');

const WelcomePage = () => {
    return (
        <>
            Welcome 

            On the Vat ID Check Page determine the given Range and the Programm will create a new Excel Sheet with the results.
        </>
    )
}


export  { WelcomePage}