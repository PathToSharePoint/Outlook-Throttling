import * as React from 'react';
import styles from './OutlookEvents.module.scss';
import { IOutlookEventsProps } from './IOutlookEventsProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { MSGraphClient } from '@microsoft/sp-http';

const OutlookEvents = (props) => {

  async function addOutlookEvents() {
    // Run sequential calls to prevent Outlook throttling
    const datesISO = [
      "2021-04-01",
      "2021-04-02",
      "2021-04-03",
      "2021-04-04",
      "2021-04-05",
      "2021-04-06",
      "2021-04-07",
      "2021-04-08",
      "2021-04-09",
      "2021-04-10",
      "2021-04-11",
      "2021-04-12",
      "2021-04-13",
      "2021-04-14",
      "2021-04-15",
      "2021-04-16",
      "2021-04-17",
      "2021-04-18",
      "2021-04-19",
      "2021-04-20",
      "2021-04-21"
    ];

    for (let i = 0; i < 20; i++) {
      let startDateISO = datesISO[i];
      let endDateISO = datesISO[i + 1];

      await props.context.msGraphClientFactory.getClient()
        .then((client: MSGraphClient): void => {
          client
            .api("/me/events")
            .post({
              "subject": "Reservation",
              "isAllDay": true,
              "showAs": "free",
              "start": {
                "dateTime": startDateISO + "T00:00:00",
                "timeZone": "UTC"
              },
              "end": {
                "dateTime": endDateISO + "T00:00:00",
                "timeZone": "UTC"
              }
            });
        });
    }
  }

  return (
    <div className={styles.outlookEvents}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            <span className={styles.title}>Welcome to SharePoint!</span>
            <p className={styles.subTitle}>Customize SharePoint experiences using Web Parts.</p>
            <button onClick={addOutlookEvents} type="button">Add Outlook Events</button>
          </div>
        </div>
      </div>
    </div>
  );
}

export default OutlookEvents;
