import * as React from 'react';
import styles from './PersonalCalendar.module.scss';
import * as strings from 'GloboSkeletonWebPartStrings';
import { IPersonalCalendarProps, IPersonalCalendarState, IMeeting, IMeetings } from '.';
import { Spinner, SpinnerSize} from 'office-ui-fabric-react/lib/components/Spinner';
import { List } from 'office-ui-fabric-react/lib/components/List';
import { Link } from 'office-ui-fabric-react/lib/components/Link';

export default class PersonalCalendar extends React.Component<IPersonalCalendarProps, IPersonalCalendarState> {
    constructor(props: IPersonalCalendarProps) {
        super(props);

        this.state = {
            meetings: [],
            loading: true,
            error: undefined
        };
        this._loadMeetings();
    }

    private _getTimeZone(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            this.props.graphClient
                //get the mailbox settings
                .api(`me/mailboxSettings`)
                .version('v1.0')
                .get((err: any, res: microsoftgraph.MailboxSettings): void => {
                    resolve((res.timeZone || 'GMT Standard Time')); // deals with cases where the user has not opened outlook and set a timezone
                })
        })
    }

    private _loadMeetings(): void {
        if (!this.props.graphClient) {
          return;
        }
    
        // update state to indicate loading and remove any previously loaded
        // meetings
        this.setState({
          error: null,
          loading: true,
          meetings: []
        });
    
        const date: Date = new Date();
        const now: string = date.toISOString();
        // set the date to midnight today to load all upcoming meetings for today
        date.setUTCHours(23);
        date.setUTCMinutes(59);
        date.setUTCSeconds(0);
        date.setDate(date.getDate());
        const midnight: string = date.toISOString();
    
        this._getTimeZone().then(timeZone => {
          this.props.graphClient
            // get all upcoming meetings for the rest of the day today
            .api(`me/calendar/calendarView?startDateTime=${now}&endDateTime=${midnight}`)
            .version("v1.0")
            .select('subject,start,end,showAs,webLink,location,isAllDay')
            .top(20)
            .header("Prefer", "outlook.timezone=" + '"' + timeZone + '"')
            // sort ascending by start time
            .orderby("start/dateTime")
            .get((err: any, res: IMeetings): void => {
              if (err) {
                // Something failed calling the MS Graph
                this.setState({
                  error: err.message ? err.message : strings.Error,
                  loading: false
                });
                return;
              }
    
              // Check if a response was retrieved
              if (res && res.value && res.value.length > 0) {
                this.setState({
                  meetings: res.value,
                  loading: false
                });
              }
              else {
                // No meetings found
                this.setState({
                  loading: false
                });
              }
            });
        });
      }
    
      private _getDuration = (meeting: IMeeting): string => {
        if (meeting.isAllDay) {
          return strings.AllDay;
        }
    
        const startDateTime: Date = new Date(meeting.start.dateTime);
        const endDateTime: Date = new Date(meeting.end.dateTime);
        // get duration in minutes
        const duration: number = Math.round((endDateTime as any) - (startDateTime as any)) / (1000 * 60);
        if (duration <= 0) {
          return '';
        }
    
        if (duration < 60) {
          return `${duration} ${strings.Minutes}`;
        }
    
        const hours: number = Math.floor(duration / 60);
        const minutes: number = Math.round(duration % 60);
        let durationString: string = `${hours} ${hours > 1 ? strings.Hours : strings.Hour}`;
        if (minutes > 0) {
          durationString += ` ${minutes} ${strings.Minutes}`;
        }
    
        return durationString;
      }
    
      private _onRenderCell = (item: IMeeting, index: number | undefined): JSX.Element => {
        const startTime: Date = new Date(item.start.dateTime);
        const minutes: number = startTime.getMinutes();
    
        return <div className={`${styles.meetingWrapper} ${item.showAs}`}>
          <Link href={item.webLink} className={styles.meeting} target='_blank'>
            <div className={styles.start}>{`${startTime.getHours()}:${minutes < 10 ? '0' + minutes : minutes}`}</div>
            <div className={styles.subject}>{item.subject}</div>
            <div className={styles.duration}>{this._getDuration(item)}</div>
            <div className={styles.location}>{item.location.displayName}</div>
          </Link>
        </div>;
      }
    
      public render(): React.ReactElement<IPersonalCalendarProps> {
        return (
          <div className={styles.personalCalendar}>
          <div id='root'></div>
            {
              this.state.loading &&
              <Spinner label={strings.Loading} size={SpinnerSize.large} />
            }
    
            {
              this.state.meetings &&
                this.state.meetings.length > 0 ? (
                  <div>
                    <h2>{strings.Heading}</h2>
                    <List items={this.state.meetings}
                      onRenderCell={this._onRenderCell} className={styles.list} />
                    <Link href='https://outlook.office.com/owa/?path=/calendar/view/Day' target='_blank'>{strings.ViewAll}</Link>
                  </div>
                ) : (
                  !this.state.loading && (
                    this.state.error ?
                      <span className={styles.error}>{this.state.error}</span> :
                      <span className={styles.noMeetings}>{strings.NoMeetings}</span>
                  )
                )
            }
          </div>
        );
      }
}
