# -*- coding: utf-8 -*-
"""
Exchange Server Authenticators

@author: Mike Crute (mcrute@gmail.com)
@organization: SoftGroup Interactive, Inc.
@date: April 26, 2009
@version: $Rev$

$Id$
"""
import urllib

from copy import copy
from httplib import HTTPSConnection
from Cookie import SimpleCookie


class ExchangeAuthenticator(object):

    _auth_cookie = None
    authenticated = False
    username = None

    def __init__(self, web_server):
        self.web_server = web_server

    def authenticate(self, username, password):
        """
        Authenticate the user and cache the authentication so we aren't
        hammering the auth server. Should hanlde expiration eventually.
        """
        self.username = username

        if self._auth_cookie:
            return self._auth_cookie

        self._auth_cookie = self._do_authentication(username, password)
        return self._auth_cookie

    def _do_authentication(self, username, password):
        raise NotImplemented

    def patch_headers(self, headers):
        raise NotImplemented


class CookieAuthenticator(ExchangeAuthenticator):

    AUTH_DLL = "/exchweb/bin/auth/owaauth.dll"

    def _do_authentication(self, username, password):
        """
        Does a post to the authentication DLL to fetch a cookie for the session
        this can then be passed back to the exchange API for servers that don't
        support basicc HTTP auth.
        """
        params = urllib.urlencode({ "destination": "https://%s/exchange" % (self.web_server),
                                    "flags": "0",
                                    "username": username,
                                    "password": password,
                                    "SubmitCreds": "Log On",
                                    "trusted": "4"
                                    })

        conn = HTTPSConnection(self.web_server)
        conn.request("POST", self.AUTH_DLL, params)
        response = conn.getresponse()

        cookie = SimpleCookie(response.getheader("set-cookie"))
        cookie = ("sessionid=%s" % cookie["sessionid"].value, "cadata=%s" % cookie["cadata"].value)

        self.authenticated = True
        return "; ".join(cookie)

    def patch_headers(self, headers):
        out_headers = copy(headers)
        out_headers["Cookie"] = self._auth_cookie
        return out_headers
