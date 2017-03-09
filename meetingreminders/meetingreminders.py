import urllib
import asyncio
import time
import os
from datetime import datetime
from discord.ext import commands
from cogs.utils.dataIO import dataIO
from __main__ import send_cmd_help

try:
    import vobject
    vobjectAvailable = True
except:
    vobjectAvailable = False

try:
    from pytz import timezone
    pytzAvailable = True
except:
    pytzAvailable = False

class MeetingReminders:
    """Watches a Google calendar and reminds attendees when meetings approach"""

    OLD_MEETING_CUTOFF = 86400 # One day
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
    async def list(self, ctx, tz=None):
        """Lists all of the future meetings. Optionally provide a display timezone like 'US/Pacific'."""
        server_id = ctx.message.server.id
        settings = self._get_settings(server_id)
        zone = timezone(settings['timezone'])
        if tz:
            try:
                zone = timezone(tz)
            except:
                await self.bot.say("Time zone provided isn't valid. Using %s instead."%settings['timezone'])

        if not server_id in self.calendars or self.settings[server_id]['ics_url'] == None:
            await self.bot.say("No calendar for this server yet. Have you run [p]meetings url <url> yet?")
        else:
            meetings = self.calendars[server_id]
            await self.bot.say("------------------------------------\n"+"\n\n------------------------------------\n".join([self._meeting_str(m, zone) for m in meetings]))


    @meetings.command(pass_context=True, no_pm=True)
    async def refresh(self, ctx):
        """Immediately loads changes from the source calendar."""
        self._load_calendars()
        await self.bot.say("Refreshed calendar.")

    @meetings.command(pass_context=True, no_pm=True)
    async def url(self, ctx, url):
        """Sets the .ics URL to use for managing meetings."""
        settings = self._get_settings(ctx.message.server.id)
        old_url = settings['ics_url']
        try:
            settings['ics_url'] = url
            self._save_settings()
            self._load_calendars()
            await self.bot.say("Done setting URL for this server.")
        except Exception as err:
            settings['ics_url'] = old_url
            await self.bot.say("Couldn't update url: %s"%err)

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
    def _load_calendars(self):
        for server_id, settings in self.settings.items():
            ics = urllib.request.urlopen(settings["ics_url"])
            cal_str = ics.read().decode('utf-8')
            ics.close()

            self.calendars[server_id] = self._process_ical(cal_str)

    def _process_ical(self, calendar_str):
        calendar = vobject.readOne(calendar_str)
        meetings = []
        for vevent in calendar.vevent_list:
            if type(vevent.dtstart.value) is datetime and vevent.dtstart.value.timestamp() > (time.time() - self.OLD_MEETING_CUTOFF):
                # Currently, we only pay attention to meetings with a start time within a day ago
                # TODO: Decide how to handle all-day events
                meeting = {'start': vevent.dtstart.value, 'end': vevent.dtend.value, 'summary':vevent.summary.value, 'description':vevent.description.value}
                meetings.append(meeting)

        return sorted(meetings, key=lambda x: x['start'])

    def _meeting_str(self, meeting, zone):
        start_time_format = '%a %b %-d, %l:%M%P'
        end_time_format = '%l:%M%P %Z'
        time_string = "%s - %s" % (meeting['start'].astimezone(zone).strftime(start_time_format), meeting['end'].astimezone(zone).strftime(end_time_format))
        return "**Title:** %s\n**Time:** %s\n**Description:** %s" % (meeting['summary'], time_string, meeting['description'])

    def _get_settings(self, server_id):
        if not server_id in self.settings:
            self.settings[server_id] = {'timezone':'US/Eastern', 'ics_url':None, 'soon':60}
            self._save_settings()

        return self.settings[server_id]

    def _save_settings(self):
        dataIO.save_json("data/meetingreminders/settings.json", self.settings)

    async def _pm_attendees(self, meeting, server_id, msg):
        try:
            for username in [u.strip('@ ') for u in meeting['description'].splitlines()[0].split(',')]:
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
                    start = meeting['start'].timestamp()
                    if settings['soon'] > 0 and start < (time.time() + int(settings['soon'])*60) and not meeting in self.soon_notified:
                        self.soon_notified.append(meeting)
                        await self._pm_attendees(meeting, server_id, "Your meeting starts soon!\n%s"%self._meeting_str(meeting, zone))
                    elif start < time.time() and not meeting in self.now_notified:
                        self.now_notified.append(meeting)
                        await self._pm_attendees(meeting, server_id, "Your meeting is starting now!\n%s"%self._meeting_str(meeting, zone))

            # Prune notified sets so they don't get too big
            self.soon_notified = [m for m in self.soon_notified if m['start'].timestamp() > (time.time() - self.OLD_MEETING_CUTOFF)]
            self.now_notified = [m for m in self.now_notified if m['start'].timestamp() > (time.time() - self.OLD_MEETING_CUTOFF)]

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
    if not vobjectAvailable:
        raise RuntimeError("You need to run 'pip3 install vobject'")
    if not pytzAvailable:
        raise RuntimeError("You need to run 'pip3 install pytz'")

    check_folders()
    check_files()
    cog = MeetingReminders(bot)
    loop = asyncio.get_event_loop()
    loop.create_task(cog.check_meetings())

    bot.add_cog(cog)
