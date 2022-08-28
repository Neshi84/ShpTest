import { Version } from "@microsoft/sp-core-library";
import {
  PropertyPaneCheckbox,
  PropertyPaneDropdown,
  PropertyPaneToggle,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./HelloWorldWebPart.module.scss";
import * as strings from "HelloWorldWebPartStrings";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";

export interface IHelloWorldWebPartProps {
  description: string;
  test: string;
  test1: boolean;
  test2: string;
  test3: boolean;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
  Name:string;
  LastName:string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = "";

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${
      !!this.context.sdks.microsoftTeams ? styles.teams : ""
    }">
      <div class="${styles.welcome}">
        <img alt="" src="${
          this._isDarkTheme
            ? require("./assets/welcome-dark.png")
            : require("./assets/welcome-light.png")
        }" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(
          this.context.pageContext.user.displayName
        )}!</h2>
        <div>${this._environmentMessage}</div>
      </div>
        <div>
         <h3>Welcome to SharePoint Framework!</h3>
          <div>Web part description: <strong>${escape(
            this.properties.description
          )}</strong></div>
          <div>Web part test: <strong>${escape(
            this.properties.test
          )}</strong></div>
          <div>Loading from: <strong>${escape(
            this.context.pageContext.web.title
          )}</strong></div>
        </div>
            <input type="text" id="name"></input>
            <input type="text" id="lastName"></input>
            <input type=submit id="save" value="Save"></input>
        <div>
        </div>
      <div id="spListContainer" />
</section>`;

    this._setButtonEventHandlers();
    this._renderListAsync();
  }

  private _setButtonEventHandlers():void{

    this.domElement.querySelector("#save").addEventListener('click',()=>{
      const spOpts: ISPHttpClientOptions = {
      body: `{Title: "${escape(this.context.pageContext.user.displayName)}",Name: "${escape(this.domElement.querySelector("#name")["value"])}",LastName: "${escape(this.domElement.querySelector("#lastName")["value"])}"}`
    };
    this._makePOSTRequest(spOpts);
      alert("Hello "+this.domElement.querySelector("#name")["value"]+" "+this.domElement.querySelector("#lastName")["value"]+"!");
    })
  }

  private _makePOSTRequest(spOpts: ISPHttpClientOptions): void {

    console.log(spOpts)
    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('TestList')/items`, SPHttpClient.configurations.v1, spOpts)
      .then((response: SPHttpClientResponse) => {
        // Access properties of the response object. 
        console.log(`Status code: ${response.status}`);
        console.log(`Status text: ${response.statusText}`);

        //response.json() returns a promise so you get access to the json in the resolve callback.
        response.json().then((responseJSON: JSON) => {
          console.log(responseJSON);
           this._renderListAsync();
        });
      });
}

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('TestList')/items`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = "";
    console.log(items)
    items.forEach((item: ISPList) => {
      html += `
        <ul class="${styles.list}">
          <li class="${styles.listItem}">
            <span class="ms-font-l">${item.Title}<input type="text" value=${item.Name}></input><input type="text" value=${item.LastName}></input></span>

          </li>
        </ul>`;
    });

    const listContainer: Element =
      this.domElement.querySelector("#spListContainer");

    listContainer.innerHTML = html;
  }

  private _renderListAsync(): void {
    this._getListData().then((response) => {
      this._renderList(response.value);
    });
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) {
      // running in Teams
      return this.context.isServedFromLocalhost
        ? strings.AppLocalEnvironmentTeams
        : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost
      ? strings.AppLocalEnvironmentSharePoint
      : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    this.domElement.style.setProperty("--bodyText", semanticColors.bodyText);
    this.domElement.style.setProperty("--link", semanticColors.link);
    this.domElement.style.setProperty(
      "--linkHovered",
      semanticColors.linkHovered
    );
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: "Description",
                }),
                PropertyPaneTextField("test", {
                  label: "Multi-line Text Field",
                  multiline: true,
                }),
                PropertyPaneCheckbox("test1", {
                  text: "Checkbox",
                }),
                PropertyPaneDropdown("test2", {
                  label: "Dropdown",
                  options: [
                    { key: "1", text: "One" },
                    { key: "2", text: "Two" },
                    { key: "3", text: "Three" },
                    { key: "4", text: "Four" },
                  ],
                }),
                PropertyPaneToggle("test3", {
                  label: "Toggle",
                  onText: "On",
                  offText: "Off",
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
