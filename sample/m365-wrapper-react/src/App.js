import M365Wrapper from '../node_modules/m365-wrapper/lib/index';
import React from 'react';
import './App.css';
import { LocalSettings } from './local-settings';

class App extends React.Component {
  constructor(props) {
    super(props);
    this.state = { userInfo: undefined, m365wrapper: new M365Wrapper(LocalSettings.APPLICATION_ID) };
  }

  async componentDidMount() {
    let result = await this.state.m365wrapper.user.loginPopup();

    if (result.isSuccess) {
      this.updateUserInfo();
    }
  }

  updateUserInfo() {
    let infos = this.state.m365wrapper.user.getMyAccountInfo();

    if (infos.isSuccess) {
      this.setState({ userInfo: infos.data });
    }
    else {
      this.setState({ userInfo: undefined });
    }
  }

  async loginPopup() {
    await this.state.m365wrapper.user.loginPopup();
    this.updateUserInfo();
  }

  async logoutPopup() {
    await this.state.m365wrapper.user.logoutPopup();
    this.updateUserInfo();
  }

  async getMyApps() {
    let apps = await this.state.m365wrapper.user.getMyApps();

    if (apps.isSuccess) {
      let appContainer = document.getElementById('office-apps');
      for (let i = 0; i < apps.data.length; i++) {
        let app = apps.data[i];
        let img = document.createElement('img');
        img.src = app.icon;
        img.alt = app.name;
        img.width = 50;
        img.height = 50;
        img.style.margin = '15px';
        let a = document.createElement('a');
        a.href = app.link;
        a.target = '_blank';
        a.appendChild(img);
        appContainer.appendChild(a);
      }
    }
  }

  async getDriveSpaceInfo(driveId) {
    let driveSpaceInfo = await this.state.m365wrapper.drive.getDriveSpaceInfo(driveId);

    if (driveSpaceInfo.isSuccess) {
      let total = document.getElementById('drive-space-total');
      let used = document.getElementById('drive-space-used');
      let remaining = document.getElementById('drive-space-remaining');

      total.innerHTML = 'Total space: ' + driveSpaceInfo.data.total;
      used.innerHTML = 'Used space: ' + driveSpaceInfo.data.used;
      remaining.innerHTML = 'Remaining space: ' + driveSpaceInfo.data.remaining;
    }
  }

  render() {

    let userDetailsResult = this.state.m365wrapper.user.getMyAccountInfo()
    let title;

    if (this.state.userInfo != undefined) {
      title = <h1>Welcome {userDetailsResult.data.name}!</h1>;
    }
    else {
      title = <h1>Welcome! Please perform the login</h1>;
    }

    return (
      <div className="App">
        <div className='container'>
          <div className='row'>
            <div className='col-12 text-center'>
              {title}
            </div>
          </div>
          <div className='row'>
            <div className='col-6 offset-3'>
              <div className='row'>
                <div className='col-4 d-grid gap-2'>
                  <button type="button" class="btn btn-primary" onClick={() => this.loginPopup()}>Login</button>
                </div>
                <div className='col-4 d-grid gap-2'>
                  <button type="button" class="btn btn-primary" onClick={() => this.logoutPopup()}>Logout</button>
                </div>
                <div className='col-4 d-grid gap-2'>
                  <button type="button" class="btn btn-primary" onClick={() => this.getMyApps()}>Get Apps</button>
                </div>
              </div>
              <div className='row mt-2'>
                <div className='col-4 d-grid gap-2'>
                  <button type="button" class="btn btn-primary" onClick={() => this.getDriveSpaceInfo('{DRIVE ID}')}>Get drive space info</button>
                </div>
              </div>
            </div>
          </div>
          <div className='row'>
            <div className='col-12 text-center'>
              <div id="office-apps"></div>
            </div>
          </div>
          <div className='row'>
            <div className='col-12 text-center'>
              <div id="drive-space-total"></div>
              <div id="drive-space-used"></div>
              <div id="drive-space-remaining"></div>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

export default App;
