import React, { Component } from 'react';

class LargeObjectExample extends Component {
  constructor(props) {
    super(props);
    this.state = {
      user: {
        name: 'John',
        age: 30,
        address: {
          city: 'New York',
          zip: '10001',
        },
      },
      settings: {
        theme: 'light',
        notifications: true,
      },
    };
  }

  updateCity = (newCity) => {
    this.setState((prevState) => ({
      user: {
        ...prevState.user,
        address: {
          ...prevState.user.address,
          city: newCity,
        },
      },
    }));
  };

  updateTheme = (newTheme) => {
    this.setState((prevState) => ({
      settings: {
        ...prevState.settings,
        theme: newTheme,
      },
    }));
  };

  render() {
    return (
      <div>
        <h1>{this.state.user.name}</h1>
        <p>City: {this.state.user.address.city}</p>
        <p>Theme: {this.state.settings.theme}</p>
        <button onClick={() => this.updateCity('Los Angeles')}>Change City</button>
        <button onClick={() => this.updateTheme('dark')}>Change Theme</button>
      </div>
    );
  }
}

export default LargeObjectExample;
