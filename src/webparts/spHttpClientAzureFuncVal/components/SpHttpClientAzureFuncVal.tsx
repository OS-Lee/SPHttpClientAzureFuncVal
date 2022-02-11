import * as React from 'react';
import styles from './SpHttpClientAzureFuncVal.module.scss';
import { ISpHttpClientAzureFuncValProps } from './ISpHttpClientAzureFuncValProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { HttpClient, HttpClientResponse, SPHttpClient, ISPHttpClientOptions, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IUserProfile {
  // FirstName: string;  
  // LastName: string;      
  // Email: string;  
  // Title: string;  
  // WorkPhone: string;  
  // DisplayName: string;  
  // Department: string;  
  // PictureURL: string;      
  // UserProfileProperties: Array<any>;
  DisplayName: string;
  PersonalUrl: string;
  PictureUrl: string;
  UserProfileProperties: Array<any>;
}

export interface IRestData {
  profileResponse: IUserProfile;
  azureapiResponse: any;
}

export default class SpHttpClientAzureFuncVal extends React.Component<ISpHttpClientAzureFuncValProps, IRestData> {

  public constructor(props: ISpHttpClientAzureFuncValProps, state: IRestData) {
    super(props);

    this.state = {
      profileResponse: {} as IUserProfile,
      azureapiResponse: ""
    };

  }

  private async GetUserProfile(): Promise<IUserProfile> {
    return this.props.context.spHttpClient.get(`${this.props.context.pageContext.web.absoluteUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  protected async callAzureFunction(): Promise<string> {
    const myOptions: ISPHttpClientOptions = {
      headers: new Headers(),
      method: "GET",
      mode: "cors"
    };

    var AzureWebApi = "https://spfxapp2.azurewebsites.net/api/HttpTrigger1?code=Iyru3gyJOG5o9jpuU0F4PtOTkP0XsUP8Px4mTYKeYTv8NTnmcFcSPg==&name=test";
    return this.props.context.httpClient.get(AzureWebApi, HttpClient.configurations.v1, myOptions).then((azureapiresponse: HttpClientResponse) => {
      return azureapiresponse.text();
    });
  }

  public componentDidMount() {
    var reactHandler = this;
    console.log("componentDidMount");
    this.GetUserProfile().then((myUserProfile: IUserProfile) => {
      console.log(myUserProfile);
      reactHandler.setState({ profileResponse: myUserProfile });
    });

    this.callAzureFunction().then((azureWebapiresponse: string) => {
      console.log(azureWebapiresponse);

      this.setState({
        azureapiResponse: azureWebapiresponse
      });
    });
  }

  public render(): React.ReactElement<ISpHttpClientAzureFuncValProps> {
    return (
      <div className={styles.spHttpClientAzureFuncVal}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
              <p className={styles.description}>{escape(this.props.description)}</p>
              <div>User Profile-1:</div>
              {this.state.profileResponse.DisplayName}:{this.state.profileResponse.PictureUrl}
              <img src={this.state.profileResponse.PictureUrl} ></img>

              <div>Azure web api call response:</div>
              {this.state.azureapiResponse}
            </div>
          </div>
        </div>
      </div>
    );
  }
}
