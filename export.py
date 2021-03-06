import json, os, pytz
from datetime import datetime, timedelta

from exchangelib import Account, Credentials, DELEGATE 
from exchangelib import CalendarItem, EWSDateTime
from exchangelib.items import MeetingRequest, MeetingCancellation

from icalendar import Calendar, Event, vCalAddress, vText, vDate

def readCfg():
    with open(os.path.dirname(os.path.realpath(__file__))+'/config.json', 'r') as file:
        config = json.load(file)
    return config

def connectToMsx():
    cfg = readCfg()
    today = datetime.today().replace(hour=0, minute=0, second=0)
    
    credentials = Credentials(username=cfg['auth']['username'], password=cfg['auth']['password'])
    account = Account(primary_smtp_address=cfg['auth']['email'], credentials=credentials, autodiscover=True, access_type=DELEGATE)

    calendarItems = account.calendar.view (
        start = account.default_timezone.localize(EWSDateTime.from_datetime(today)) + timedelta(days=cfg['export']['last']*-1), 
        end   = account.default_timezone.localize(EWSDateTime.from_datetime(today)) + timedelta(days=cfg['export']['next']),
    )

    return calendarItems

def generateIcs(calendarItems):
    cfg = readCfg()
    cal = Calendar()

    for key, value in cfg['calAttr'].items():
        cal.add(key, value)

    for item in calendarItems:
        try:
            tzOverwriteStart = pytz.timezone(item._start_timezone.zone)
            tzOverwriteEnd   = pytz.timezone(item._end_timezone.zone)
            item.start       = item.start.astimezone(tz=tzOverwriteStart)
            item.end         = item.end.astimezone(tz=tzOverwriteEnd)
        except:
            print("Meh. timezones...")

        event = Event()
        if item.sensitivity == 'Private':
            event.add('summary', "{} {}".format('🔒', item.subject))
        else:
            event.add('summary',    item.subject)
        if (item.is_cancelled):
            event['summary'] = "{} {}".format('❌', event['summary'])
        elif (item.my_response_type == 'NoResponseReceived'):
            event['summary'] = "{} {}".format('❓', event['summary'])

        event.add('uid',            item.id)
        event.add('dtstart',        item.start)
        event.add('dtend',          item.end)
        if item.text_body is not None:
            event.add('description',    item.text_body)


        if item.is_all_day:
            item.start= item.start + timedelta(days=1)
            event['dtstart'] =        vDate(item.start.replace(hour=0, minute=0, second=0))
            item.end= item.end + timedelta(days=1)
            event['dtend'] =          vDate(item.end.replace(hour=0, minute=0, second=0))

        if item.subject.startswith(tuple(cfg['manipulation']['convertToAllday']['startsWith'])):
            event['dtstart'] =        vDate(item.start.replace(hour=0, minute=0, second=0))
            event['dtend'] =          vDate(item.end.replace(hour=0, minute=0, second=0))
    
        if item.location is not None:
            event.add('location',   item.location)

        if item.categories is not None:
            event.add('CATEGORIES', item.categories)
            if cfg['manipulation']['prependCategoriesToBody']:
                out = '\n'.join(item.categories)
                if event.has_key('description'):
                    event['description'] = out + '\n\n' + str(event['description'])
                else:
                    event.add('description', out)

        organizer = vCalAddress("MAILTO:{}".format(item.organizer.email_address))
        organizer.params['cn']      = item.organizer.name
        organizer.params['ROLE']    = vText('CHAIR')
        event.add('organizer', organizer)

        if item.required_attendees is not None:
            for x in item.required_attendees:
                attendee = vCalAddress("MAILTO:{}".format(x.mailbox.email_address))
                if ((x.mailbox.email_address.lower() == cfg['auth']['email'].lower()) and (item.my_response_type is not None)):
                    x.response_type = item.my_response_type
                attendee.params['cn']       = x.mailbox.name 
                attendee.params['ROLE']     = vText('REQ-PARTICIPANT')
                if ParticipationState(x.response_type) != False:
                    attendee.params['PARTSTAT'] = ParticipationState(x.response_type)
                event.add('attendee', attendee, encode=0)

        if item.optional_attendees is not None:
            for x in item.optional_attendees:
                attendee = vCalAddress("MAILTO:{}".format(x.mailbox.email_address))
                if ((x.mailbox.email_address.lower() == cfg['auth']['email'].lower()) and (item.my_response_type is not None)):
                    x.response_type = item.my_response_type
                attendee.params['cn']       = x.mailbox.name 
                attendee.params['ROLE']     = vText('OPT-PARTICIPANT')
                if ParticipationState(x.response_type) != False:
                    attendee.params['PARTSTAT'] = ParticipationState(x.response_type)
                event.add('attendee', attendee, encode=0)

        cal.add_component(event)

    write(cfg['export']['filename'], cal.to_ical().decode('utf-8'))

    # For better Debugging:
    # def readable(cal):
    #     return cal.to_ical().decode('utf-8').replace('\r\n', '\n').strip()
    # write(cfg['export']['filename']+".debug", readable(cal))

def write(filename, content):
    f = open(filename, "w")
    f.write(content)
    f.close()

def ParticipationState(response_type):
    # https://icalendar.org/iCalendar-RFC-5545/3-2-12-participation-status.html
    if response_type == 'Accept':
        return vText("ACCEPTED")
    elif response_type == 'Decline':
        return vText("DECLINED")
    elif response_type == 'Tentative':
        return vText("TENTATIVE")
    else:
        return False

if __name__ == '__main__':
    generateIcs(connectToMsx())