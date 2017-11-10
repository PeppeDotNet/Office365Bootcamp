import * as React from 'react';
import styles from './Sample.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IContact } from '../model/IContact';
import { css, Persona, PersonaSize, PersonaPresence, Spinner } from 'office-ui-fabric-react';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import pnp from 'sp-pnp-js';

export interface ISampleProps {
  list: string;
  context: IWebPartContext
}

export interface ISampleState {
  contacts: IContact[];
  isLoading: boolean;
}

export default class Sample extends React.Component<ISampleProps, ISampleState> {
  constructor(props: ISampleProps, state: ISampleState) {
    super(props);

    this.state = {
      contacts: [], isLoading: false
    };
  }

  public render(): React.ReactElement<ISampleProps> {
    const loading: JSX.Element = this.state.isLoading ? <div style={{ margin: '0 auto' }}><Spinner label={'Loading...'} /></div> : <div/>;
    const listOfContacts: JSX.Element[] = this.state.contacts.map((contact) => { 
      return (
        <Persona
        primaryText={contact.DisplayName}
        secondaryText={contact.Mail}
        imageUrl={contact.ImageUrl}
        size={PersonaSize.large}
        presence={PersonaPresence.none}
        key={contact.Id} />
      )});
    return (      
      <div>
        <h1>Contacts</h1>
        <div>
          {loading}
          {listOfContacts}
        </div>
      </div>
    );
  }

  public componentDidMount() {
    this.loadContacts(this.props.list);
  }

  public componentWillReceiveProps(props) {
    this.loadContacts(props.list);
  }

  private loadContacts(list: string): void {
    var results: IContact[] = [];
    if(this.props.list) {
      this.setState({ contacts: results, isLoading: true });   
      pnp.setup({ sp: { baseUrl: this.props.context.pageContext.web.absoluteUrl } });
      pnp.sp.web.lists.getById(this.props.list).items.select("Id", "WebPage","FullName","Email").get().then((items) => {
        items.forEach(item => {
          results.push({ Id: item.Id, DisplayName: item.FullName, ImageUrl: item.WebPage.Url, Mail: item.Email });
        });
        this.setState({ contacts: results, isLoading: false });
      });
    }
  }
}
