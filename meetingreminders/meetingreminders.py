import asyncio
import os
import re
from datetime import datetime, timedelta
from discord.ext import commands
from cogs.utils.dataIO import dataIO
from __main__ import send_cmd_help

try:
    from apiclient import discovery
    from oauth2client import client
    from oauth2client.service_account import ServiceAccountCredentials
    oauth2Available = True
except:
    oauth2Available = False

try:
    import httplib2
    httplib2Available = True
except:
    httplib2Available = False

try:
    import dateutil.parser
    dateutilAvailable = True
except:
    dateutilAvailable = False

try:
    from pytz import timezone
    pytzAvailable = True
except:
    pytzAvailable = False

class MeetingReminders:
    """Watches a Google calendar and reminds attendees when meetings approach"""

    KEEP_HISTORY = timedelta(hours=1)
    AUTO_REFRESH_INTERVAL = 60 # One minute

    def __init__(self, bot):
        self.bot = bot
        self.settings = dataIO.load_json("data/meetingreminders/settings.json")
        self.calendars = {}
        self.soon_notified = []
        self.now_notified = []

    ####
    # Commands
    ####
    @commands.group(pass_context=True, no_pm=True)
    async def meetings(self, ctx):
        """Commands for viewing and managing meeting reminders."""
        if ctx.invoked_subcommand is None:
            await send_cmd_help(ctx)

    @meetings.command(pass_context=True, no_pm=True)
    async def sharewith(self, ctx, email):
        """Shares the google calendar with a specified email. All viewers will be able to create and modify events."""
        if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
            await self.bot.say("That doesn't look like a valid email to me.")
            return

        rule = {'role': 'writer', 'scope': {'type': 'user', 'value': email}}
        client = self._get_google_client(ctx.message.server.id)
        client.acl().insert(calendarId='primary', body=rule).execute()
        await self.bot.say("Successfully added user. Check your Google calendar list.")

    @meetings.command(pass_context=True, no_pm=True)
    async def list(self, ctx, tz=None):
        """Lists the next five meetings. Optionally provide a display timezone like 'US/Pacific'."""
        server_id = ctx.message.server.id
        settings = self._get_settings(server_id)
        zone = timezone(settings['timezone'])
        if tz:
            try:
                zone = timezone(tz)
            except:
                await self.bot.say("Time zone provided isn't valid. Using %s instead."%zone)

        if not server_id in self.calendars or self.settings[server_id]['creds_file'] == 'None':
            await self.bot.say("No calendar for this server. Have you run [p]meetings creds <filename> yet?")
        else:
            events = self.calendars[server_id]
            if len(events) > 0:
                await self.bot.say("------------------------------------\n"+"\n\n------------------------------------\n".join([self._meeting_str(m, zone) for m in events]))
            else:
                await self.bot.say("No upcoming meetings found.")

    @meetings.command(pass_context=True, no_pm=True)
    async def refresh(self, ctx):
        """Immediately loads changes from the source calendar. Usually unecessary due to automatic updates."""
        self._load_calendars()
        await self.bot.say("Refreshed calendar.")

    @meetings.command(pass_context=True, no_pm=True)
    async def creds(self, ctx, filename):
        """Sets the filename holding credentials to use for managing meetings. The file should be stored in the data/meetingreminders directory. To obtain one, follow this documentation: https://developers.google.com/identity/protocols/OAuth2ServiceAccount"""
        settings = self._get_settings(ctx.message.server.id)
        old_file = settings['creds_file']
        try:
            settings['creds_file'] = filename
            self._load_calendars()
            await self.bot.say("Done setting credentials for this server.")
        except Exception as err:
            settings['creds_file'] = old_file
            await self.bot.say("Uh oh, looks like that file didn't work: %s"%err)

        self._save_settings()

    @meetings.command(pass_context=True, no_pm=True)
    async def timezone(self, ctx, zone):
        """Sets the default timezone in which to display meeting times. Use values like US/Pacific, Europe/London, or Asia/Hong_Kong"""
        settings = self._get_settings(ctx.message.server.id)
        try:
            timezone(zone)
            settings['timezone'] = zone
            self._save_settings()
            await self.bot.say("Set timezone for this server to %s."%zone)
        except Exception as err:
            await self.bot.say("Couldn't update timezone: %s"%err)

    @meetings.command(pass_context=True, no_pm=True)
    async def remindertime(self, ctx, minutes: int):
        """Sets the number of minutes before meetings to remind attendees. Set to 0 for no advance warning, only a reminder when the meeting is starting."""
        settings = self._get_settings(ctx.message.server.id)
        settings['soon'] = minutes
        self._save_settings()
        await self.bot.say("Set reminder time (in minutes) for this server to %s."%minutes)

    ####
    # Helper functions
    ####
    def _get_google_client(self, server_id):
        settings = self._get_settings(server_id)
        scopes = ['https://www.googleapis.com/auth/calendar.readonly', 'https://www.googleapis.com/auth/calendar']
        credentials = ServiceAccountCredentials.from_json_keyfile_name(os.path.join('data/meetingreminders', settings['creds_file']), scopes=scopes)
        http_auth = credentials.authorize(httplib2.Http())
        return discovery.build('calendar', 'v3', http=http_auth)

    def _load_calendars(self):
        for server_id in self.settings.keys():
            client = self._get_google_client(server_id)
            time_min = datetime.now(timezone('UTC')) - self.KEEP_HISTORY
            self.calendars[server_id] = client.events().list(calendarId='primary', timeMin=time_min.isoformat(), maxResults=5, singleEvents=True, orderBy='startTime').execute()['items']

    def _meeting_str(self, meeting, zone):
        start_time = dateutil.parser.parse(meeting['start']['dateTime']).astimezone(zone)
        end_time = dateutil.parser.parse(meeting['end']['dateTime']).astimezone(zone)
        start_time_format = '%a %b %-d, %l:%M%P'
        end_time_format = '%l:%M%P %Z'
        time_string = "%s - %s" % (start_time.strftime(start_time_format), end_time.strftime(end_time_format))

        return "**Title:** %s\n**Time:** %s\n**Description:** %s" % (meeting['summary'], time_string, meeting.get('description', "_no description provided_"))

    def _get_settings(self, server_id):
        if not server_id in self.settings:
            self.settings[server_id] = {'timezone':'US/Eastern', 'creds_file':'None', 'soon':60}
            self._save_settings()

        return self.settings[server_id]

    def _save_settings(self):
        dataIO.save_json("data/meetingreminders/settings.json", self.settings)

    async def _pm_attendees(self, meeting, server_id, msg):
        try:
            for username in [u.strip('@ ') for u in meeting.get('description', '').splitlines()[0].split(',')]:
                server = self.bot.get_server(server_id)
                if server:
                    member = server.get_member_named(username)
                    if member:
                        await self.bot.send_message(member, msg)
                    else:
                        print("Meeting %s has non member attendee: %s"%(meeting['summary'], username))
        except Exception as err:
            print("Error PMing members for %s: %s"%(meeting['summary'], err))

    ####
    # Event loop
    ####
    async def check_meetings(self):
        while self is self.bot.get_cog("MeetingReminders"):
            try:
                self._load_calendars()
            except Exception as err:
                print("Error loading calendars: %s"%err)

            for server_id, calendar in self.calendars.items():
                settings = self._get_settings(server_id)
                zone = timezone(settings['timezone'])

                for meeting in calendar:
                    now = datetime.now(timezone('UTC'))
                    meeting_start = dateutil.parser.parse(meeting['start']['dateTime'])
                    if settings['soon'] > 0 and meeting_start < (now + timedelta(seconds=settings['soon'])) and not meeting in self.soon_notified:
                        self.soon_notified.append(meeting)
                        await self._pm_attendees(meeting, server_id, "Your meeting starts soon!\n%s"%self._meeting_str(meeting, zone))
                    elif meeting_start < now and not meeting in self.now_notified:
                        self.now_notified.append(meeting)
                        await self._pm_attendees(meeting, server_id, "Your meeting is starting now!\n%s"%self._meeting_str(meeting, zone))

            # Prune notified sets so they don't get too big
            time_min = datetime.now(timezone('UTC')) - self.KEEP_HISTORY
            self.soon_notified = [m for m in self.soon_notified if dateutil.parser.parse(meeting['start']['dateTime']) > time_min]
            self.now_notified = [m for m in self.now_notified if dateutil.parser.parse(meeting['start']['dateTime']) > time_min]

            await asyncio.sleep(self.AUTO_REFRESH_INTERVAL)

####
# Bot setup
####
def check_folders():
    folders = ("data", "data/meetingreminders/")
    for folder in folders:
        if not os.path.exists(folder):
            print("Creating " + folder + " folder...")
            os.makedirs(folder)

def check_files():
    if not os.path.isfile("data/meetingreminders/settings.json"):
        print("Creating default settings.json...")
        dataIO.save_json("data/meetingreminders/settings.json", {})

def setup(bot):
    if not oauth2Available:
        raise RuntimeError("You need to run 'pip3 install google-api-python-client'")
    if not httplib2Available:
        raise RuntimeError("You need to run 'pip3 install httplib2'")
    if not dateutilAvailable:
        raise RuntimeError("You need to run 'pip3 install python-dateutil'")
    if not pytzAvailable:
        raise RuntimeError("You need to run 'pip3 install pytz'")

    check_folders()
    check_files()
    cog = MeetingReminders(bot)
    loop = asyncio.get_event_loop()
    loop.create_task(cog.check_meetings())

    bot.add_cog(cog)

