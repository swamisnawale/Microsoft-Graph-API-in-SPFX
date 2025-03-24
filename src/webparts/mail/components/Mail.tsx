import * as React from "react";
import styles from "./Mail.module.scss";
import type { IMailProps } from "./IMailProps";
import "@pnp/graph/mail";
import "@pnp/graph/users";
import { SPFx, graphfi } from "@pnp/graph";
interface IMail {
  subject: string;
  webLink: string;
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
}
export default class Mail extends React.Component<
  IMailProps,
  { myMails: IMail[] | any; loading: boolean; isMailboxAssigned: boolean }
> {
  constructor(props: IMailProps) {
    super(props);
    this.state = {
      myMails: [],
      loading: true,
      isMailboxAssigned: false,
    };
  }

  async componentDidMount(): Promise<void> {
    await this.getMails();
  }

  async getMails() {
    const graph = graphfi().using(SPFx(this.props.context));
    const currentUser = graph.me;
    const mailboxSettings = await currentUser
      .mailboxSettings()
      .then(async () => {
        const emailsFromInbox = await currentUser.messages();

        this.setState({
          myMails: emailsFromInbox,
          loading: false,
        });
        console.log("mailboxSettings", mailboxSettings);
      })
      .catch((err) => {
        alert("Mailbox is not assigned to you");
      });
  }

  async sendEmail() {
    let emailID = prompt("Enter email ID");
    const graph = graphfi().using(SPFx(this.props.context));
    const currentUser = graph.me;
    const draftMessage: any = {
      subject: "PnPjs Test Message",
      importance: "low",
      body: {
        contentType: "html",
        content: "This is a test message!",
      },
      toRecipients: [
        {
          emailAddress: {
            address: emailID,
          },
        },
      ],
    };
    await currentUser.sendMail(draftMessage);
  }
  public render(): React.ReactElement<IMailProps> {
    return (
      <section className={styles.mail}>
        <h2>My Emails</h2>

        <br />
        <button onClick={() => this.sendEmail()}> Send email</button>

        {this.state.loading && <p>Loading...</p>}

        {this.state.myMails.map((mail: IMail) => {
          return (
            <>
              <h4>{mail.from.emailAddress.name}</h4>
              <h2>{mail.subject}</h2>
              <a href={mail.webLink} target="_blank">
                Open email
              </a>
              <hr />
            </>
          );
        })}
      </section>
    );
  }
}
