import * as React from 'react';
import styles from './SpListDispReactDemo.module.scss';
import { ISpListDispReactDemoProps } from './ISpListDispReactDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as jquery from 'jquery'

export interface IListItemWPState {
  listItems: [
    {
      "Title": "",
      "Client": { "Title": "" },
      "Status": "",
      "Id": "",
    }
  ]
}

export default class SpListDispReactDemo extends React.Component<ISpListDispReactDemoProps, IListItemWPState> {

  public static siteURL: string = "";

  public constructor(props: ISpListDispReactDemoProps, state: IListItemWPState) {
    super(props);
    this.state = {
      listItems: [
        {
          "Title": "",
          "Client": { "Title": "" },
          "Status": "",
          "Id": "",
        }
      ]
    }
    SpListDispReactDemo.siteURL = this.props.siteURL;
  }

  public componentDidMount() {
    let reactContext = this;
    jquery.ajax({
      url: `${SpListDispReactDemo.siteURL}/_api/web/lists/getbytitle('Order_Details')/items?$select=Client/Title,Id,Status,Title&$expand=Client/Title`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (result) {
        console.log(JSON.stringify(result));
        reactContext.setState({
          listItems: result.d.results
        });
        console.log(result.d.results);
      },
      error: function (jqXHR, textstatus, error) {
      }
    });
  }

  public render(): React.ReactElement<ISpListDispReactDemoProps> {
    return (
      <div className={styles.panelStyle}>
        <table className={styles.tableStyle} >
          <tr className={styles.headerStyle}>
            <td className={styles.CellStyle}>Title</td>
            <td className={styles.CellStyle}>Id</td>
            <td className={styles.CellStyle}>Client</td>
            <td className={styles.CellStyle}>Status</td>
          </tr>
          {
            this.state.listItems.map(function (listitem, llistitemkey) {
              let itemURL: string = `${SpListDispReactDemo.siteURL}//Lists/Order_Details/dispform.aspx?ID=${listitem.Id}`;

              return (
                <tr className={styles.rowStyle}>
                  <td>
                    <a className={styles.CellStyle} href={itemURL} >{listitem.Title}</a>
                  </td>
                  <td className={styles.CellStyle} >{listitem.Id}</td>
                  <td className={styles.CellStyle} >{listitem.Client.Title}</td>
                  <td className={styles.CellStyle} >{listitem.Status}</td>
                </tr>
              );
            })
          }
        </table>
        <br></br>
        <h3>-- Orderd List -- </h3>
        <ol>
          {
            this.state.listItems.map(function (listitem, llistitemkey) {
              let itemURL: string = `${SpListDispReactDemo.siteURL}//Lists/Order_Details/dispform.aspx?ID=${listitem.Id}`;
              return (
                <li>
                  <a className={styles.label} href={itemURL} >
                    <span>{listitem.Title}</span> - <span>{listitem.Id} </span> - <span>{listitem.Client.Title}</span> - <span>{listitem.Status}</span>
                  </a>
                </li>
              );
            })
          }
        </ol>
      </div>
    );
  }
}
