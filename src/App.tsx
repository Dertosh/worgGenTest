import React from 'react';
import logo from './logo.svg';
import './App.css';
import DocumentGanaration from './doc/DocumentGenerationBase';

class App extends React.Component<any, any> {

  doc = new DocumentGanaration();

  constructor(props: any) {
    super(props);
    this.state = { isToggleOn: true };

    this.handleClick = this.handleClick.bind(this);
  }

  handleClick() {
    this.doc.writeData();
  }

  render() {
    return (
      <div>
        <button onClick={this.handleClick}>
          Generate file
        </button>
      </div>

    );
  }
}



export default App;
