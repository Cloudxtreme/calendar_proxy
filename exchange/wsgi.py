# vim: set filencoding=utf8
"""
Calendar WSGI App

@author: Mike Crute (mcrute@ag.com)
@organization: SoftGroup Interactive, Inc.
@date: February 15, 2010
"""

from exchange.commands import FetchCalendar
from exchange.authenticators import CookieSession


class CalendarApp(object):

    def __init__(self, exchange_server, user, password):
        self.session = CookieSession(exchange_server,
                                        username=user, password=password)

    def __call__(self, environ, start_response):
        start_response('200 OK', [])
        command = FetchCalendar(self.session)
        calendar = command.execute()
        return calendar.as_string()
