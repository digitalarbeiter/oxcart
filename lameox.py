#!/usr/bin/env python3
import click
import colorful as cf
import datetime
import json
import os

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
    click.echo(f"add appointment {title}")
    ox.calendar.create(
        oxcart.OxAppointment(
            id=None,
            folder=folder,
            title=title,
            start_date=start,
            end_date=end,
            full_time=False,
            location=location,
            note=notes,
            recurrence=None,
            raw=None,
        )
    )


if __name__ == "__main__":
    cli(obj={})
