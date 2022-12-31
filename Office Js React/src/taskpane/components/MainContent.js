import * as React from 'react';
import { MainFormular } from './Formular';
import { MainFormular_Qualified } from './Formular_Qualified';
import { Pivot, PivotItem } from '@fluentui/react';
import { FeedbackTab } from './Feedback';
import { WelcomePage } from './WelcomePage';

var myConfig = require('../../../config.json'); //TODO Multilanguage Support
// this component holds the State of the Formulars.

const MainContent_Window = () => {

    //Version control:
    const showQualifiedOption = myConfig.showQualifiedContent;

    // the props for the Feedback Tab
    const [feedbackText, setFeedbackText] = React.useState("");
    const propsFeedbackTab = { 
        feedbackText: feedbackText,
        setFeedbackText: setFeedbackText
    }

    const propsWelcomeTab = {
        //no props for the welcome tab yet. 
    };

    //Props and state for the MainFormular. This is the uplifted State.
    const [ustID, setustID] = React.useState("");
    const [VatRange, setVatRange] = React.useState("");
    const [VatRangeQualified, setVatRangeQualified] = React.useState("");
    const [CitiesRange, setCitiesRange] = React.useState("");
    const [AreaCodeRange, setAreaCodeRange] = React.useState("");
    const [CompanyNames, setCompanyNameRange] = React.useState("");
    const [CompanyTypes, setCompanyTypeRange] = React.useState("");
    const propsMainFormularTab = { //this later gets destructured. This way we do not clutter the HTML with a mass of props
        ustID: ustID,
        setustID: setustID,
        VatRange: VatRange,
        setVatRange: setVatRange,
        VatRangeQualified: VatRangeQualified,
        setVatRangeQualified: setVatRangeQualified,
        CitiesRange: CitiesRange,
        setCitiesRange: setCitiesRange,
        AreaCodeRange: AreaCodeRange,
        setAreaCodeRange: setAreaCodeRange,
        CompanyNames: CompanyNames,
        setCompanyNameRange: setCompanyNameRange,
        CompanyTypes: CompanyTypes,
        setCompanyTypeRange: setCompanyTypeRange
    };

    return (
        <Pivot>
            <PivotItem headerText= "Welcome" style={{margin:5}}>
              <WelcomePage/>
            </PivotItem>
            <PivotItem headerText = "Vat Id Check" style={{ margin: 5}}>
              <MainFormular props = {propsMainFormularTab }></MainFormular>
            </PivotItem>
            { showQualifiedOption ? <PivotItem headerText = "Qualified Check" style = {{ margin:5}}>
              <MainFormular_Qualified props = { propsMainFormularTab } > </MainFormular_Qualified>
            </PivotItem>: null }
            <PivotItem headerText="Feedback" style={{ margin :5}}>
              <FeedbackTab feedbackText = { feedbackText} setFeedbackText = { setFeedbackText} ></FeedbackTab>
            </PivotItem>
          </Pivot>
    )
}

export { MainContent_Window }