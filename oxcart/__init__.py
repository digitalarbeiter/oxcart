import colorful as cf
from dataclasses import dataclass
from datetime import datetime, timezone
from enum import IntEnum
import requests
import time
import urllib


class OXError(Exception):
    def __init__(self, user, url, data, resp):
        self.user = user
        self.url = url
        self.data = data
        # resp is something like this: {
        #    'error': 'Your session expired. Please login again.',
        #    'error_params': ['b023e74cd29b40a0bc8348334a84d2ea'],
        #    'categories': 'TRY_AGAIN',  # may also be list of str
        #    'category': 4,
        #    'code': 'SES-0203',
        #    'error_id': '82558447-5082714',
        #    'error_desc': 'Your session b023e74cd29b40a0bc8348334a84d2ea expired. Please start a new browser session.',
        # }
        self.resp = resp

    def __str__(self):
        return f"OXError({self.user}@{self.url}({self.data}) -> {self.resp}"


class OX:
    def __init__(self, base_url, username, password, *, debug=False):
        self.debug = debug
        self.base_url = base_url
        self.user = None
        self.session = None
        self._auth_record = None
        self.login(username, password)
        self.calendar = OxCalendar(self)
        self.folders = OxFolders(self)

    def login(self, username, password):
        # Documentation on OX login process:
        # https://documentation.open-xchange.com/7.10.3/middleware/login_and_sessions/session_lifecycle.html
        self.session = requests.Session()
        resp = self.POST(
            "/login",
            params={
                "action": "login",
                "client": "open-xchange-oxcart",
                "staySignedIn": True,
            },
            data={
                "name": username,
                "password": password,
            }
        )
        # response looks like this (plus some cookies): {
        #     "session":"f0fb8c2565cf46edbb4c728072a98187",
        #     "user":"schemitz@mailbox.org",
        #     "user_id":3,
        #     "context_id":7667555,
        #     "locale":"de_DE",
        # }
        self._auth_record = resp
        self.user = resp["user"]
        self.session.headers.update({"session": resp["session"]})
        print(cf.green(f"logged in as {self.user}"))

    def logout(self):
        self.GET(
            "/login",
            params={
                "action": "logout",
            }
        )
        print(cf.green(f"logged out {self.user}"))
        self._auth_record = None
        self.user = None
        self.session = None

    def POST(self, url, *, params, data):
        return self._request("POST", url, params=params, data=data)

    def PUT(self, url, *, params, data):
        return self._request("PUT", url, params=params, json=data)

    def GET(self, url, *, params):
        return self._request("GET", url, params=params, json=None)

    def _request(self, method, url, *, params=None, data=None, json=None):
        if not params:
            params = {}
        if self._auth_record:
            params["session"] = self._auth_record["session"]
        url = f"{self.base_url}{url}"
        if params:
            url += "?" + urllib.parse.urlencode(params)
        if self.debug:
            print(f"{cf.cyan(method)} {cf.blue(url)} {data or json}")  # COOKIES: {self.session.cookies}
        resp = None
        while not resp:
            if data:
                resp = self.session.request(method, url, data=data)
            elif json:
                resp = self.session.request(method, url, json=json)
            else:
                resp = self.session.request(method, url)
            if resp.status_code != 200:
                print(f"HTTP {method} {url} data={data}", cf.red(resp.content))
                resp.raise_for_status()
            if self.debug:
                print(cf.cyan("    response:"), resp.text)  # resp.cookies
            if not resp.text:
                return None
            resp = resp.json()
            if "error" in resp:
                if resp["categories"] != "TRY_AGAIN":
                    print(f"HTTP {method} {url} data={data}", cf.red(resp["error_desc"]))
                    raise OXError(self.user, url, data, resp)
                else:
                    print(cf.yellow("TRY_AGAIN in 1 second"))
                    time.sleep(1)
                    resp = None
        if "data" in resp:
            return resp["data"]
        return resp


class OxFolders:
    def __init__(self, ox):
        self.ox = ox

    def all_folders(self, type_="contacts"):
        all_ = self.all_(type_)
        for access in ["private", "public"]:
            for folder in all_.get(access, []):
                yield self.get_(folder[0])

    def get_(self, id_):
        resp = self.ox.GET("/folders", params={"action": "get", "id": id_})
        return resp

    def all_(self, type_):
        resp = self.ox.GET("/folders", params={"action": "allVisible", "columns": "1,2,3,4,100,102", "content_type": type_})
        return resp


class OxRecurrenceInterval(IntEnum):
    NONE = 0
    DAILY = 1
    WEEKLY = 2
    MONTHLY = 3
    YEARLY = 4


@dataclass
class OxRecurrence:
    id: str
    interval: OxRecurrenceInterval
    start: datetime
    days: list

    def to_ox(self):
        ox = {
            "recurrence_type": int(self.interval),
        }
        if self.interval >= OxRecurrenceInterval.WEEKLY:
            ox.days = sum(2**((n+1)%7) for n in self.days)
        # TODO all the rest of it
        return ox

    @staticmethod
    def Once():
        return OxRecurrence(id=None, interval=OxRecurrenceInterval.NONE, start=None, end=None, days=[])

    @staticmethod
    def Daily(start, end):
        return OxRecurrence(id=None, interval=OxRecurrenceInterval.DAILY, start=start, end=end, days=[])

    @staticmethod
    def Weekly(start, end, days):
        return OxRecurrence(id=None, interval=OxRecurrenceInterval.WEEKLY, start=start, end=end, days=days)

    @staticmethod
    def none_to_ox():
        return {
            "recurrence_type": 0,
        }

    @staticmethod
    def from_ox(resp):
        # In Python datetime, Monday is 0 and Sunday is 6.
        # (Source: https://docs.python.org/3/library/datetime.html#datetime.datetime.weekday)
        # In OpenXchange, bit 0 [is] indicating sunday.
        # (Source: https://documentation.open-xchange.com/components/middleware/http/7.10.1/index.html#!/Calendar/createAppointment_0)
        if resp["recurrence_type"] == OxRecurrenceInterval.NONE:
            return None
        # Daily appointment:
        {
            "created_by":3,
            "uid":"49bf5b57-0c39-4aa6-ab9b-fe2a6600429b",
            "creation_date":1656536280566,
            "organizer":"schemitz@mailbox.org",
            "organizerId":3,
            "modified_by":3,
            "last_modified_utc":1656529080566,
            "last_modified":1656536280566,
            "id":10,
            "folder_id":26,

            "participants":[{"id":3,"type":1}],
            "users":[{"id":3,"confirmation":1}],
            "confirmations":[],

            "number_of_attachments":0,

            "title":"Täglich",
            "start_date":1656709200000,
            "end_date":1656712800000,
            "timezone":"Europe/Berlin",
            "color_label":0,
            "private_flag":False,
            "full_time":False,
            "shown_as":1,
            "alarm":15,

            "recurrence_type":1,
            "interval":1,
            "occurrences":10,
            "sequence":0,
            "recurrence_id":10,
            "recurrence_start":"1656633600000",
        }
        # "recurrence_type":1,"interval":1,"occurrences":10,"recurrence_id":10,"recurrence_start":"1656633600000"
        # 'recurrence_type': 2, 'days': 4, 'interval': 1, 'recurrence_id': 5, 'recurrence_start': '1655769600000'
        days = []
        if resp["recurrence_type"] > int(OxRecurrenceInterval.DAILY):
            for n in range(0, 7):
                if resp["days"] & (1 << n):
                    days.append((n+6)%7)  # correct to Python style weekday (Monday=0)
        return OxRecurrence(
            id=resp["id"],
            interval=OxRecurrenceInterval(resp["recurrence_type"]),
            start=datetime.fromtimestamp(int(resp["recurrence_start"])/1000, tz=timezone.utc),
            days=days,
        )


@dataclass
class OxAppointment:
    id: str
    folder: str
    title: str
    start_date: datetime
    end_date: datetime
    full_time: bool  # FIXME?
    location: str
    note: str
    recurrence: OxRecurrence
    raw: dict

    @staticmethod
    def Once(folder, start_date, end_date, title, note=None, location=None, full_time=False):
        return OxAppointment(
            id=None,
            folder=folder,
            title=title,
            start_date=start_date,
            end_date=end_date,
            full_time=full_time,
            location=location,
            note=note,
            recurrance=OxRecurrence.Once(),
        )

    @staticmethod
    def Daily(folder, start_date, end_date, title, note=None, location=None, full_time=False):
        return OxAppointment(
            id=None,
            folder=folder,
            title=title,
            start_date=start_date,
            end_date=end_date,
            full_time=full_time,
            location=location,
            note=note,
            recurrance=OxRecurrence.Daily(),
        )

    @staticmethod
    def Weekly(folder, start_date, end_date, title, note=None, location=None, full_time=False):
        ...

    def to_ox(self):
        ox = {
            "folder_id": self.folder,
            "title": self.title,
            "start_date": int(self.start_date.strftime("%s")) * 1000,
            "end_date": int(self.end_date.strftime("%s")) * 1000,
        }
        if self.location:
            ox["location"] = self.location
        if self.note:
            ox["note"] = self.note
        if self.recurrence:
            ox.update(self.recurrence.to_ox())
        else:
            ox.update(OxRecurrence.none_to_ox())
        return ox

    @staticmethod
    def from_ox(resp):
        return OxAppointment(
            id=resp["id"],
            folder=resp["folder_id"],
            title=resp["title"],
            start_date=datetime.fromtimestamp(resp["start_date"]/1000, tz=timezone.utc),
            end_date=datetime.fromtimestamp(resp["end_date"]/1000, tz=timezone.utc),
            location=resp.get("location", None),
            note=resp.get("note", None),
            full_time=resp.get("full_time"),
            raw=resp,
            recurrence=OxRecurrence.from_ox(resp),
        )


class OxCalendar:
    def __init__(self, ox):
        self.ox = ox
        self._folders = set()

    def all_(self, start: datetime, end: datetime):
        start = int(start.strftime("%s")) * 1000
        end = int(end.strftime("%s")) * 1000
        appointments = self.ox.GET("/calendar", params={"action": "all", "columns": "1,20", "start": start, "end": end})
        for appt_id, folder_id in appointments:
            self._folders.add(folder_id)
            yield self.get_(appt_id, folder_id)

    def list_(self, appts):
        # Like a multi get, but with selective numerical columns
        resp = self.ox.PUT("/calendar", params={"action": "list", "columns": "1,20,200"}, data=appts)
        print(resp)
        return resp

    def search(self, *, pattern: str=None, startletter: str=None):
        query = {}
        if pattern:
            query["pattern"] = pattern
        if startletter:
            # FIXME: really weird, startletter="B" matches lots of stuff without "B"
            query["startletter"] = startletter
        appointments = self.ox.PUT("/calendar", params={"action": "search", "columns": "1,20"}, data=query)
        for appt_id, folder_id in appointments:
            self._folders.add(folder_id)
            yield self.get_(appt_id, folder_id)

    def get_(self, id_, folder):
        """ Get appointment with given (id, folder).

            Returns an OxAppointment.
        """
        # PUT action=list is a bit like a multi get, but one needs to specify which
        # columns one wants (and returns them as a list, in column order), whereas GET
        # action=get always returns the entire appointment as a dict.
        self._folders.add(folder)
        resp = self.ox.GET("/calendar", params={"action": "get", "id": id_, "folder": folder})
        return OxAppointment.from_ox(resp)

    def create(self, appointment):
        try:
            appointment = appointment.to_ox()
        except AttributeError:
            pass  # no to_ox(), assume it's already "oxy"
        if not "folder_id" in appointment and len(self._folders) == 1:
            # if we know of only one folder, use it as default
            appointment["folder_id"] = list(self._folders)[0]
        assert appointment["folder_id"] in self._folders
        resp = self.ox.PUT("/calendar", params={"action": "new"}, data=appointment)
        if "conflicts" in resp:
            raise OXError()  # FIXME don't be so harsh maybe
        return resp["id"]
