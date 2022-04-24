import * as React from "react";
import PropTypes from "prop-types";
import Progress from "./Progress";
import { MainFormular } from "./Formular";
import { FeedbackTab } from "./Feedback";
import { Label, Pivot, PivotItem } from '@fluentui/react';
import { WelcomePage } from "./WelcomePage";
/* global console, Excel, require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      
    });
  }

  click = async () => {
    try {
      await Excel.run(async (context) => {
        /**
         * Insert your Excel code here
         */
        
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          message="Bitte in Excel öffnen."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <div className = "Header"> 
          Header
          LU26375245
        </div>
        <div className = "mainBody">
          <Pivot>
            <PivotItem headerText= "Welcome" style={{margin:5}}>
              <WelcomePage/>
            </PivotItem>
            <PivotItem headerText = "Vat Id Check" style={{ margin: 5}}>
              <MainFormular></MainFormular>
            </PivotItem>
            <PivotItem headerText="Feedback" style={{ margin :5}}>
              <FeedbackTab></FeedbackTab>
            </PivotItem>
            
          </Pivot>
          
        </div>

      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
