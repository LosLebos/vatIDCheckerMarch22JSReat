import * as React from "react";
import PropTypes from "prop-types";
import Progress from "./Progress";
import { MainFormular } from "./Formular";
import { FeedbackTab } from "./Feedback";
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
          <MainFormular></MainFormular>
          <FeedbackTab></FeedbackTab>
        </div>

      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
