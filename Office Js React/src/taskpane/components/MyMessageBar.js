import { MessageBar, MessageBarType } from '@fluentui/react';
import * as React from 'react';


const MyMessageBar = (props) => {
    //i would rather have it as its own component file but then i would have to use setErrorMessage() on a parent object?
    
    let currentMessageType = 0;
    if (props.messageType) {
        if (props.messageType === "Error") {
            currentMessageType = MessageBarType.error; // 1
        } else if (props.messageType === "Success") {
            currentMessageType = MessageBarType.success;
        } else {
            currentMessageType = MessageBarType.info; //0
        }
    }
    return(
    <>
    <MessageBar
      messageBarType={currentMessageType}
      isMultiline={true}
      onDismiss= { props.handleMessageBarDismiss }
      dismissButtonAriaLabel="Close"
    >
      { props.message }
    </MessageBar>
    </>
  )};

  export { MyMessageBar}