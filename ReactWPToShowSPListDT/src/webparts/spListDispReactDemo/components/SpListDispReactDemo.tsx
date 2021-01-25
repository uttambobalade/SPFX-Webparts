import * as React from 'react';
import styles from './SpListDispReactDemo.module.scss';
import { ISpListDispReactDemoProps } from './ISpListDispReactDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import * as jquery from 'jquery'

export interface IListItemWPState {
  listItems: [
    {
      "Title": "",
      "Client": "",
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
          "Client": "",
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
      url: `${SpListDispReactDemo.siteURL}/_api/web/lists/getbytitle('Order_Details')/items`,
      type: "GET",
      headers: { 'Accept': 'application/json; odata=verbose;' },
      success: function (result) {
        reactContext.setState({
          listItems: result.d.results
        });
      },
      error: function (jqXHR, textstatus, error) {
      }
    });

  }

  public render(): React.ReactElement<ISpListDispReactDemoProps> {
    return (
      <div className={styles.spListDispReactDemo}>
        <table className={styles.row} >
          {
            this.state.listItems.map(function (listitem, llistitemkey) {
              let itemURL: string = `${SpListDispReactDemo.siteURL}//Lists/Order_Details/dispform.aspx?ID=${listitem.Id}`;

              return (
                <tr>
                  <td>
                    <a className={styles.label} href={itemURL} >{listitem.Title}</a>
                  </td>
                  <td className={styles.label} >{listitem.Id}</td>
                  <td className={styles.label} >{listitem.Client}</td>
                  <td className={styles.label} >{listitem.Status}</td>
                </tr>
              );
            })
          }
        </table>
        <ol>
          {
            this.state.listItems.map(function (listitem, llistitemkey) {
              let itemURL: string = `${SpListDispReactDemo.siteURL}//Lists/Order_Details/dispform.aspx?ID=${listitem.Id}`;
              return (
                <li>
                  <a className={styles.label} href={itemURL} >
                    <span>{listitem.Title}</span> - <span>{listitem.Id} </span> - <span>{listitem.Client}</span> - <span>{listitem.Status}</span>
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
