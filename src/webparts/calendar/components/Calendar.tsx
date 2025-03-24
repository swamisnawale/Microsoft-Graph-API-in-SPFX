import * as React from "react";
import styles from "./Calendar.module.scss";
import type { ICalendarProps } from "./ICalendarProps";
import "@pnp/graph/users";
import "@pnp/graph/calendars";
import { SPFx, graphfi } from "@pnp/graph";

interface IEvent {
  subject: string;
  webLink: string;
  start: {
    dateTime: string;
  };
  end: {
    dateTime: string;
  };
}

export default class Calendar extends React.Component<
  ICalendarProps,
  { myEvents: IEvent[] | any; loading: boolean }
> {
  constructor(props: ICalendarProps) {
    super(props);
    this.state = {
      myEvents: [],
      loading: true,
    };
  }

  async componentDidMount(): Promise<void> {
    await this.getMyEvents();
  }

  async getMyEvents() {
    const graph = graphfi().using(SPFx(this.props.context));
    const events = await graph.users
      .getById(this.props.context.pageContext.user.email)
      .events();

    this.setState({
      myEvents: events,
      loading: false,
    });

    console.log(events);
  }

  async createEvent() {
    const graph = graphfi().using(SPFx(this.props.context));
    let eventName = prompt("Enter event name");
    const eventData: any = {
      subject: eventName,
      body: {
        contentType: "HTML",
        content: "Does late morning work for you?",
      },
      start: {
        dateTime: "2025-03-25T12:00:00",
        timeZone: "Pacific Standard Time",
      },
      end: {
        dateTime: "2025-03-26T14:00:00",
        timeZone: "Pacific Standard Time",
      },
      location: {
        displayName: "Harry's Bar",
      },
    };
    await graph.users
      .getById(this.props.context.pageContext.user.email)
      .calendar.events.add(eventData);
  }

  public render(): React.ReactElement<ICalendarProps> {
    return (
      <section className={styles.calendar}>
        <h2>My Calendar</h2>

        <br />
        <button onClick={() => this.createEvent()}>Create event</button>
        {this.state.myEvents.map((event: IEvent) => {
          return (
            <>
              <h1>{event.subject}</h1>
              <a href={event.webLink} target="_blank">
                See event
              </a>
              <p>{new Date(event.start.dateTime).toDateString()}</p>
              <p>{new Date(event.end.dateTime).toDateString()}</p>
              <hr />
            </>
          );
        })}
      </section>
    );
  }
}
