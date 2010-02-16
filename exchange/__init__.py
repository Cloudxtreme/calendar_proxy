# vim: set filencoding=utf8
"""
Exchange Server Handling Code

@author: Mike Crute (mcrute@gmail.com)
@organization: SoftGroup Interactive, Inc.
@date: April 26, 2009
"""

from datetime import tzinfo, timedelta


class ExchangeException(Exception):
    "Exception that is thrown by all Exchange handling code."


class AuthenticationException(ExchangeException):
    "Exception that is raised when authentication fails."


class EST(tzinfo):

    def tzname(self, dt):
        return "EST"

    def utcoffset(self, dt):
        return timedelta(0)

    dst = utcoffset
