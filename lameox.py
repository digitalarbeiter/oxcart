#!/usr/bin/env python3
import click
import colorful as cf
import datetime
import json
import os
import re
import pytz

import oxcart


def date_time(s):
    if isinstance(s, datetime.datetime):
        return s
    if isinstance(s, datetime.date):
        return datetime.datetime(s.year, s.month, s.day)
    try:
        return datetime.datetime.strptime(s, "%Y-%m-%dT%H:%M")
    except:
        pass
    try:
        return datetime.datetime.strptime(s, "%Y-%m-%d %H:%M")
    except:
        pass
    return datetime.datetime.strptime(s, "%Y-%m-%d")


@click.group()
@click.option("--debug", is_flag=True)
@click.option("--server", help="OX server Ajax API URL", default="https://office.mailbox.org/ajax")
@click.pass_context
def cli(ctx, debug, server):
    """ OpenExchange (OX) command line client.
    """
    ctx.ensure_object(dict)
    ctx.obj["debug"] = debug
    ctx.obj["ox"] = oxcart.OX(
        server,
        os.environ.get("OX_USERNAME"),
        os.environ.get("OX_PASSWORD"),
        debug=debug,
    )


@cli.group()
@click.pass_context
def calendar(ctx):
    """ OX Calendar related commands.
    """
    pass


@calendar.command()
@click.pass_context
def folders(ctx):
    """ List all calendar folders.
    """
    ox = ctx.obj["ox"]
    all_folders = ox.folders.all_folders("calendar")
    click.echo(cf.cyan(f"All calendars:"))
    for folder in all_folders:
        click.echo(
            f"Calendar ID {folder['id']}: {cf.green(folder['title'])} "
            + json.dumps(folder, indent=4, sort_keys=True)
        )


@calendar.command()
@click.option("--start", help="start date (default: today)", type=date_time, default=datetime.date.today())
@click.option("--end", help="end date (default: tomorrow)", type=date_time, default=datetime.date.today()+datetime.timedelta(days=1))
@click.pass_context
def list(ctx, start, end):
    """ List all calendar appointments between start and end date.
    """
    ox = ctx.obj["ox"]
    click.echo(f"Appointments from {start} to {end}")
    for appt in ox.calendar.all_(start, end):
        click.echo(appt)


@calendar.command()
@click.option("--pattern", help="appointment title pattern")
@click.option("--startletter", help="appointment title starting letter")
@click.pass_context
def search(ctx, pattern, startletter):
    """ Search for calendar appointments.
    """
    ox = ctx.obj["ox"]
    if pattern:
        click.echo(f"Appointments matching {pattern}")
        for appt in ox.calendar.search(pattern=pattern):
            click.echo(appt)
    else:
        click.echo(f"Appointments starting with {startletter}")
        for appt in ox.calendar.search(startletter=startletter):
            click.echo(appt)


@calendar.command()
@click.option("--title", help="appointment title", required=True)
@click.option("--location", help="location of appointment (default: none)")
@click.option("--notes", help="notes for the appointment (default: none)")
@click.option("--folder", help="calendar folder for the appointment (default: the one calendar, if it exists)", type=int)
@click.option("--start", help="start date of the appointment", type=date_time, required=True)
@click.option("--end", help="end date of the appointment", type=date_time, required=True)
@click.pass_context
def create(ctx, title, start, end, location, notes, folder):
    """ Create an appointment.
    """
    ox = ctx.obj["ox"]
    timezone = open("/etc/timezone").read().strip()
    click.echo(f"add appointment {title}")
    ox.calendar.create(
        oxcart.OxAppointment(
            id=None,
            folder=folder,
            title=title,
            start_date=start,
            end_date=end,
            timezone=timezone,
            full_time=False,
            location=location,
            note=notes,
            recurrence=None,
            raw=None,
        )
    )


def parse_owa_headers_de_DE(header):
    """ German (de_DE) OWA header line structure:
        (title) (weekday) (date) (time) - (time)(location)
        AuSu-Daily Do 28.07.2022 10:30 - 11:00Gather AuSu-Tisch
    """
    title, _weekday, start_day, start_time, end_time, location = re.match(
        r"^(.*?) (\S+) (\d\d.\d\d.\d\d\d\d) (\d\d:\d\d) - (\d\d:\d\d)(.*)$",
        header,
    ).groups()
    start_date = datetime.datetime.strptime(start_day+" "+start_time, "%d.%m.%Y %H:%M")
    end_date = datetime.datetime.strptime(start_day+" "+end_time, "%d.%m.%Y %H:%M")
    return title, start_date, end_date, location


PARSE_OWA_HEADERS = {
    "de_DE": parse_owa_headers_de_DE,
}


@calendar.command()
@click.option("--folder", help="calendar folder for the appointment (default: the one calendar, if it exists)", type=int)
@click.option("--locale", help="OWA locale (for parsing date/time)", default="de_DE")
@click.option("--list-locales", help="List supported OWA locales", is_flag=True)
@click.option("--yes/-y", help="auto-create appointment without confirmation prompt", is_flag=True)
@click.option("--timezone", help="time zone for appointment date/time (default: system tz)", default=None)
@click.argument("file", type=click.File())
@click.pass_context
def paste_owa(ctx, folder, file, locale, list_locales, yes, timezone):
    """ Create an appointment from OWA copy/paste.

        To parse the OWA header, the locale is needed to recognize the dates.
        Currently there's only a parser for German OWA settings (leading weekday,
        mm.dd.yyyy date format, 24h clock). To extend, write a header parser for
        the desired locale, and put it into PARSE_OWA_HEADERS.
    """
    if list_locales:
        click.echo("Avaliable OWA locales:")
        for locale in sorted(PARSE_OWA_HEADERS.keys()):
            click.echo(f"    * {locale}")
        return

    ox = ctx.obj["ox"]
    contents = [line for line in file]
    header = contents[0].strip()
    title, start_date, end_date, location = PARSE_OWA_HEADERS[locale](header)
    if len(contents) > 1 and contents[1].strip() != "":
        click.echo(f"{cf.yellow('second line should be empty:')} {contents[1]}")
    if len(contents) > 2:
        notes = "\n".join(line.strip() for line in contents[2:])
    else:
        notes = None
    create_appointment(ox, yes, folder, title, start_date, end_date, timezone, location, notes)


@calendar.command()
@click.option("--folder", help="calendar folder for the appointment (default: the one calendar, if it exists)", type=int)
@click.option("--locale", help="OWA locale (for parsing date/time)", default="de_DE")
@click.option("--list-locales", help="List supported OWA locales", is_flag=True)
@click.option("--yes/-y", help="auto-create appointment without confirmation prompt", is_flag=True)
@click.option("--timezone", help="time zone for appointment date/time (default: system tz)", default=None)
@click.argument("file", type=click.File())
@click.pass_context
def paste_office365_de(ctx, folder, file, locale, list_locales, yes, timezone):
    """ Create an appointment from Office 365 copy/paste.

        To parse the Office 365 header, the locale is needed to recognize the dates.

        German examples:

        Follow-Up: APIv2-Abschaltung
        Do, 2022-08-25 14:00 – 15:00
        https://solute.whereby.com/ausu

        Team-Connection
        Fr, 2022-08-26 13:30 – 14:00
        Serie
        https://solute.whereby.com/ausu
    """
    SERIES = {
        "de_DE": "Serie",
        "en_US": "Series",
    }
    ox = ctx.obj["ox"]
    contents = [line.strip() for line in file]
    title = contents[0]
    if locale == "de_DE":
        _weekday, start_day, start_time, end_time = re.match(
            r"^(\S+), (\d\d\d\d-\d\d-\d\d) (\d\d:\d\d) – (\d\d:\d\d)$",
            contents[1],
        ).groups()
        start_date = datetime.datetime.strptime(start_day+" "+start_time, "%Y-%m-%d %H:%M")
        end_date = datetime.datetime.strptime(start_day+" "+end_time, "%Y-%m-%d %H:%M")
    else:
        # TODO: more date locales; even the German one is weird...
        raise ValueError(f"cannot handle dates for locale {locale}")
    if contents[2] == SERIES[locale]:
        location = contents[3]
    else:
        location = contents[2]
    notes = None
    create_appointment(ox, yes, folder, title, start_date, end_date, timezone, location, notes)


def create_appointment(ox, yes, folder, title, start_date, end_date, timezone, location, notes):
    click.echo(
        f"{cf.cyan('appointment from copy/paste:')} {title} {cf.cyan('from')} "
        f"{start_date} {cf.cyan('to')} {end_date} {cf.cyan('at')} {location}"
    )
    if notes:
        click.echo(f"{cf.cyan('notes')}:\n{notes}")
    # Time zones with OX are weird. Even when given the time zone for start and
    # end time, we have to adjust start and end time by the offset of the time
    # zone against UTC. The OX docs are a bit thin here; I just tinkered with
    # the OX API until it seemed to work :-/
    # See section 3. in the docs:
    # https://documentation.open-xchange.com/components/middleware/http/7.10.1/index.html
    if not timezone:
        timezone = open("/etc/timezone").read().strip()
    tz = pytz.timezone(timezone)
    time_displacement = pytz.utc.localize(start_date) - tz.localize(start_date)
    if yes or click.prompt(f"{cf.cyan('create appointment? [y/n]')}").lower() == "y":
        appointment = oxcart.OxAppointment(
            id=None,
            folder=folder,
            title=title,
            start_date=start_date + time_displacement,
            end_date=end_date + time_displacement,
            timezone=timezone,
            full_time=False,
            location=location,
            note=notes,
            recurrence=None,
            raw=None,
        )
        ox.calendar.create(appointment)
        click.echo(f"{cf.green('appointment created.')}")
    else:
        click.echo(f"{cf.yellow('appointment not created.')}")


if __name__ == "__main__":
    cli(obj={})
