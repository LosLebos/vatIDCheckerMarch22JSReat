
import * as React from 'react';
import { TextField } from '@fluentui/react';
import { Spinner, SpinnerSize } from '@fluentui/react';
import { MyMessageBar } from './MyMessageBar';
import { Stack, Label, PrimaryButton } from '@fluentui/react'
var myConfig = require('../../../config.json');

const FeedbackTab = (props) => {
    const { feedbackText, setFeedbackText } = props; //destructure the Props to simple Variable
    const [successMessage, setSuccessMessage] = React.useState("");
    const [isLoading, setIsLoading] = React.useState(false);


    const handleSubmitButton = async() => {
        try {
            setSuccessMessage("");
            setIsLoading(true);
            let random = "LinuxTimestampforRandom:"+String(Date.now())
            let feedbackJSON = { //js object, later stringify to json
                message: feedbackText,
                user: random,
                auth: "",
                others: ""
            };
            const feedbackAPIReturn = await makeTheFeedbackAPICall(feedbackJSON);
            setIsLoading(false);
            setSuccessMessage(await feedbackAPIReturn.text())
        } catch (error) {
            console.log(error + error.Stack);
            setSuccessMessage(error.message);
        } finally {
            setIsLoading(false);
        }
    }
    const handleMessageBarDismiss = () => {
        setSuccessMessage("");
    };

    return (
        <Stack>
            <Label>
                Feedback
            <TextField multiline rows={7} value = { feedbackText } onChange = {(e) => setFeedbackText(e.target.value)}>
            </TextField>
            </Label>
            <PrimaryButton text = "Send" onClick= { handleSubmitButton} disabled = { !feedbackText} />
            {isLoading ? <Spinner label = "Sending Feedback" size= {SpinnerSize.medium} /> : null}
            {successMessage ? <MyMessageBar message = {successMessage} messageType = "Success" handleMessageBarDismiss = {handleMessageBarDismiss}/> : null}
        </Stack>
    )
}

async function makeTheFeedbackAPICall(apiJSON) {
    let myHeaders = new Headers();
    myHeaders.append("Content-Type", "application/json");
    try {
        let response = await fetch (myConfig.FeedbackAPIAdress, {
            method: "POST",
            headers:myHeaders,
            mode: "cors", //i could make this better and safer, not using cors but a backend to call the api
            body: JSON.stringify(apiJSON) //sends the javascript object as  JSON String
        });
        console.log(response)
        return response; //this just returns the promise, you have to await it to use, idealy returns "Feedback received."
        
    } catch(error) {
        throw (error);
    };
};

export {FeedbackTab}