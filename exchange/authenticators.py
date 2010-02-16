# vim: set filencoding=utf8
"""
Exchange Server Authenticators

@author: Mike Crute (mcrute@gmail.com)
@organization: SoftGroup Interactive, Inc.
@date: April 26, 2009
"""

import urllib
from Cookie import SimpleCookie
from httplib import HTTPSConnection
from datetime import datetime, timedelta
from exchange import AuthenticationException


class CookieSession(object):
    """
    CookieSession implmenents the authentication protocol for the
    Exchange web interface.
    """

    AUTH_DLL = "/exchweb/bin/auth/owaauth.dll"

    def __init__(self, server, username=None, password=None):
        self.server = server
        self.username = username
        self.password = password
        self.cache_timeout = timedelta(minutes=15)
        self._token = None
        self._last_modified = None

    @property
    def has_expired(self):
        if not self._last_modified:
            return False

        update_delta = datetime.now() - self._last_modified
        return (update_delta >= self.cache_timeout)

    @property
    def is_authenticated(self):
        return bool(self._token) and not self.has_expired

    @property
    def token(self):
        if not self.is_authenticated:
            self._authenticate()
            self._check_auth()

        return self._token

    @token.setter
    def token(self, token):
        self._last_modified = datetime.now()
        self._token = token.strip()

    def _authenticate(self):
        # Another idiotic Exchange issue, you MUST pass the redirect
        # destination, and what's more if you don't pass it the name
        # of the current server + /echange it will fail to auth

        params = urllib.urlencode({
            "destination": "https://{0}/exchange".format(self.server),
            "username": self.username,
            "password": self.password,
            })

        conn = HTTPSConnection(self.server)
        conn.request("POST", self.AUTH_DLL, params)
        response = conn.getresponse()

        cookie = SimpleCookie(response.getheader("set-cookie"))
        self.token = cookie.output(attrs=[], header='', sep=';')

    def _check_auth(self):
        # Grrr... Exchange is idiotic, instead of returning the correct error
        # code with the auth they force you to follow a redirect. If you
        # don't get a 200 (you'll get a 302) then you can rest assured that
        # your authentication failed.

        conn = HTTPSConnection(self.server)
        conn.request("GET", '/exchange/',
                        headers=dict(Cookie=self.token))
        resp = conn.getresponse()

        if not resp.status == 200:
            raise AuthenticationException
