import colorful as cf
from dataclasses import dataclass
from datetime import datetime, timezone
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
            print(f"{cf.cyan(method)} {cf.blue(url)} {data}")  # COOKIES: {self.session.cookies}
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


@dataclass
class OxAppointment:
    id: str
    title: str
    start_date: datetime
    end_date: datetime
    location: str
    note: str
    raw: dict


class OxCalendar:
    def __init__(self, ox):
        self.ox = ox

    def search(self, *, pattern=None, startletter=None):
        query = {}
        if pattern:
            query["pattern"] = pattern
        if startletter:
            query["startletter"] = startletter
        appointments = self.ox.PUT("/calendar", params={"action": "search", "columns": "1,20"}, data=query)
        for appt in appointments:
            yield self.get_(appt[0], appt[1])

    def get_(self, id_, folder):
        resp = self.ox.GET("/calendar", params={"action": "get", "id": id_, "folder": folder})
        appt = OxAppointment(
            id=resp["id"],
            title=resp["title"],
            start_date=datetime.fromtimestamp(resp["start_date"]/1000, tz=timezone.utc),
            end_date=datetime.fromtimestamp(resp["end_date"]/1000, tz=timezone.utc),
            location=resp.get("location", None),
            note=resp.get("note", None),
            raw=resp,
        )
        return appt

    def all_(self, start, end):
        start = start.strftime("%s")
        end = end.strftime("%s")
        resp = self.ox.GET("/calendar", params={"action": "all", "columns": "1,2,3,4,100,102", "start": start, "end": end})
        print(resp)
        return resp

    def list_(self):
        resp = self.ox.PUT("/calendar", params={"action": "list", "columns": "1,2,3,4,100,102"}, data={})
        print(resp)
        return resp
