//参考 https://developers.google.com/apps-script/advanced/calendar?authuser=0#listing_events
function sendMessageToChatWork() {
  const users = listAllUsers();
  const events = getEventListFromUsers(users);
  let sendText = '';

  let allDayScheduleText = '[info][title]終日予定[/title]';

  for (var i = 0; i < events.length; i++) {
    const event = events[i];

    if (!event.eventStartDate) {
      allDayScheduleText = allDayScheduleText + 'タイトル：' + event.title + '　作成者：' + event.creator + '\n';
      continue;
    }

    sendText = sendText + '[info][title]';
    sendText = sendText + '時間：' + event.eventStartDate + '〜' + event.eventEndDate;

    sendText = sendText + '　タイトル：' + event.title;

    if (event.resourceName) {
      sendText = sendText + '　会議室：' + event.resourceName;
    }

    sendText = sendText + '[/title]';
    if (event.attendees) {
      sendText = sendText + '参加者：' + event.attendees;
    } else {
      sendText = sendText + '参加者：' + event.creator;
    }
    sendText = sendText + '[/info]';
  }

  if (allDayScheduleText === '[info][title]終日予定[/title]') {
    allDayScheduleText = '';
  } else {
    allDayScheduleText = allDayScheduleText + '[/info]';
  }

  sendText = sendText + allDayScheduleText;

  const client = ChatWorkClient.factory({token: 'c4c571858b7c40896b67487c1612152a'});
  client.sendMessage({room_id: '96421512', body: sendText});
}

function getEventListFromUsers(users) {
  const timezone = Session.getScriptTimeZone();
  const RFC3339format = 'yyyy-MM-dd\'T\'HH:mm:ss.SXXX';

  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const minTime = now.toLocaleTimeString('en-US', {hour12: false});
  let date = Utilities.formatDate(now, timezone, RFC3339format);
  const timeMin = Utilities.formatString('%s', date);

  now.setHours(23, 59, 59, 59);
  const maxTime = now.toLocaleTimeString('en-US', {hour12: false});

  date = Utilities.formatDate(now, timezone, RFC3339format);
  const timeMax = Utilities.formatString('%s', date);

  const eventsFromUsers = [];

  for (let i = 0; i < users.length; i++) {
    const user = users[i];

    // https://developers.google.com/calendar/api/v3/reference/events/list
    const eventsFromUser = Calendar.Events.list(user.email, {
      timeMin: now.toISOString(),
      singleEvents: true,
      orderBy: 'startTime',
      timeMin: timeMin,
      timeMax: timeMax,
      // maxResults: 10
    });

    for (let ii = 0; ii < eventsFromUser.items.length; ii++) {
      let event = eventsFromUser.items[ii];
      let eventInfo = {};
      eventInfo.title = event.summary;
      eventInfo.creator = findUserFullName(users, event.creator.email, '');
      eventInfo.user = user.email;

      // 時間指定のイベントの場合、開始時間を格納
      if (event.start.dateTime) {
        const eventStartDate = new Date(event.start.dateTime);
        eventInfo.eventStartDate = eventStartDate.toLocaleTimeString('en-US', {hour12: false});
        const eventEndDate = new Date(event.end.dateTime);
        eventInfo.eventEndDate = eventEndDate.toLocaleTimeString('en-US', {hour12: false});
      }

      // イベントに他の参加者がいる場合
      if (event.attendees) {
        let attendees = [];
        let resourceName = null;
        for (let iii = 0; iii < event.attendees.length; iii++) {
          const attende = event.attendees[iii];

          // リソース(会議室とか)の場合
          if (attende.resource) {
            resourceName = attende.displayName;
            continue;
          }
          attendees.push(findUserFullName(users, attende.email, attende.displayName));
        }

        eventInfo.attendees = attendees;
        if (resourceName !== null) {
          eventInfo.resourceName = resourceName;
        }
      }
      eventsFromUsers.push(eventInfo);

    }
  }

  const filteredUniqueEvents = filterUnique(eventsFromUsers);
  return filteredUniqueEvents.sort(alphabetically(true));
}

function filterUnique(events) {
  return events.filter(function (element, index, self) {
    const sample = self.findIndex(function (e) {
      return e.title === element.title;
    });
    return sample === index;
  });
}

// 参考 https://developers.google.com/apps-script/advanced/admin-sdk-directory
function listAllUsers() {
  let page;
  const result = [];
  // https://developers.google.com/admin-sdk/directory/reference/rest/v1/users/list
  page = AdminDirectory.Users.list({
    domain: 'eyemovic.com',
    viewType: 'domain_public',
  });
  const users = page.users;
  for (let i = 0; i < users.length; i++) {
    const user = users[i];
    result.push({name: user.name.fullName, email: user.primaryEmail});
  }
  return result;
}

// null含むソート
function alphabetically(ascending) {
  return function (a, b) {
    if (a.eventStartDate === b.eventStartDate) {
      return 0;
    } else if (a.eventStartDate == null) {
      return 1;
    } else if (b.eventStartDate == null) {
      return -1;
    } else if (ascending) {
      return a.eventStartDate < b.eventStartDate ? -1 : 1;
    } else {
      return a.eventStartDate < b.eventStartDate ? 1 : -1;
    }
  };
}

// メールアドレスから名前を取得。組織外のユーザはdisplayNameから。無ければメールアドレス
function findUserFullName(users, email, displayName) {
  for (let i = 0; i < users.length; i++) {
    const user = users[i];
    if (user.email === email) {
      return user.name;
    }
  }
  if (displayName) {
    return displayName;
  }
  return email;
}
