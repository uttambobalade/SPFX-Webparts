import * as React from 'react';
import styles from './AnonymousApInExtLib.module.scss';
import { IAnonymousApInExtLibProps } from './IAnonymousApInExtLibProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IAnonymousAPIState } from './IAnonymousAPIState';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';

export default class AnonymousApInExtLib extends React.Component<IAnonymousApInExtLibProps, IAnonymousAPIState> {

  public constructor(props: IAnonymousApInExtLibProps, state: IAnonymousAPIState) {
    super(props);
    this.state = {
      id: null,
      name: null,
      username: null,
      email: null,
      address: { street: null, suite: null, city: null, zipcode: null },
      phone: null,
      website: null,
      company: { name: null, catchPhrase: null, bs: null }
    };
  }

  public getUserDetails(): Promise<any> {
    let url = this.props.siteURL + "/" + this.props.userId;

    return this.props.context.httpClient.get(url, HttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        return response.json();
      })
      .then(jsonres => {
        return jsonres;
      }) as Promise<any>;
  }

  public invokeAPIandSetState() {
    this.getUserDetails().then(response => {
      this.setState({
        id: response.id,
        name: response.name,
        username: response.username,
        email: response.email,
        address: { street: response.address.street, suite: response.address.suite, city: response.address.city, zipcode: response.address.zipcode },
        phone: response.phone,
        website: response.website,
        company: { name: response.company.name, catchPhrase: response.company.catchPhrase, bs: response.company.bs }
      });
    });
  }

  public componentDidMount() {
    this.invokeAPIandSetState();
  }

  public componentDidUpdate(prevProps: IAnonymousApInExtLibProps, prevState: IAnonymousAPIState, prevContext: any) {
    this.invokeAPIandSetState();
  }

  public render(): React.ReactElement<IAnonymousApInExtLibProps> {
    return (
      <div className={styles.anonymousApInExtLib}>
        <span className={styles.container}><h2>User Details</h2></span><br />
        <div><strong>ID: </strong>{this.state.id}</div><br />
        <div><strong>Name: </strong>{this.state.name}</div><br />
        <div><strong>UserName: </strong>{this.state.username}</div><br />
        <div><strong>Email: </strong>{this.state.email}</div><br />
        <div><strong>Address: </strong>{this.state.address.street}, {this.state.address.suite},{this.state.address.city}, {this.state.address.zipcode}</div><br />
        <div><strong>Phone: </strong>{this.state.phone}</div><br />
        <div><strong>Website: </strong>{this.state.website}</div><br />
        <div><strong>Company: </strong>{this.state.company.name}, {this.state.company.catchPhrase}, {this.state.company.bs}</div><br />
      </div>
    );
  }
}
