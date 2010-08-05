# vim: set filencoding=utf8
"""
Exchange Commands

@author: Mike Crute (mcrute@gmail.com)
@organization: SoftGroup Interactive, Inc.
@date: November 10, 2008

This is a set of classes that starts to define a set of classes for
fetching data using Exchange's WebDAV API. This is still pretty
development code but it does the trick. Watch out, it doesn't consider
many corner cases.
"""

import re

import dateutil.parser as date_parser
import xml.etree.cElementTree as ElementTree

from string import Template
from httplib import HTTPSConnection
from datetime import datetime, timedelta, date, tzinfo
from exchange import ExchangeException, EST
from icalendar import Calendar, Event as _Event, Alarm
from icalendar import vCalAddress, vText


class InvalidEventError(Exception):
    pass


class Organizer(object):

    def __init__(self, cn, email):
        self.cn = cn
        self.email = email

    def __repr__(self):
        return u"vCalAddress(%s)" % str.__repr__(self)

    def ical(self):
        return "CN=%s:%s" % (self.cn, self.email)

    def __str__(self):
        return self.ical()


class EST(tzinfo):

    def tzname(self, dt):
        return "EST"

    def utcoffset(self, dt):
        return timedelta(0)

    dst = utcoffset


class EmailParser(object):

    EXCHANGE_FORMAT = re.compile("([^<]*)<([^>]*)>")

    def parse(self, email):
        if not email:
            return

        match = self.EXCHANGE_FORMAT.match(email)
        if match:
            return match.groups()


class Event(_Event):

    def __init__(self):
        super(Event, self).__init__()
        self.start_date = None
        self.end_date = None

    def _get_element_text(self, element, key):
        value = element.find(key)

        if hasattr(value, 'text'):
            return value.text

    def add_text(self, element, key, add_as=None):
        value = self._get_element_text(element, key)

        add_as = key if not add_as else add_as
        self.add(add_as, value)

    def add_organizer(self, element):
        value = self._get_element_text(element, 'organizer')

        parsed = EmailParser().parse(value)
        if not parsed:
            return

        name = parsed[0].strip().strip('"')

        organizer = vCalAddress('MAILTO:%s' % parsed[1])
        organizer.params['cn'] = vText(name)
        organizer.params['ROLE'] = vText('CHAIR')

        self.add('organizer', organizer, encode=0)

    def add_date(self, element, key, add_as=None):
        value = date_parser.parse(self._get_element_text(element, key))

        if key == 'start_date':
            self.start_date = value
        elif key == 'end_date':
            self.end_date = value
        else:
            add_as = key if not add_as else add_as
            self.add(add_as, value)

    def finalize(self):
        if not self.start_date and self.end_date:
            raise InvalidEventError()

        delta = self.end_date - self.start_date
        if delta.days >= 1:
            start_date = date(self.start_date.year, self.start_date.month, self.start_date.day)
            end_date = date(self.end_date.year, self.end_date.month, self.end_date.day)
        else:
            start_date = self.start_date
            end_date = self.end_date

        if self.start_date > datetime.now(tz=EST()):
            alarm = Alarm()
            alarm.add("action", "DISPLAY")
            alarm.add("description", "REMINDER")
            alarm.add("trigger", timedelta(minutes=-15))
            self.add_component(alarm)

        self.add('dtstart', start_date)
        self.add('dtend', end_date)


class ExchangeRequest(object):

    def __init__(self, session, service, method='GET'):
        self.session = session
        self.service = service
        self.method = method
        self.server = session.server
        self.username = session.username
        self._headers = { 'Content-Type': 'text/xml',
                         'Depth': '0', 'Translate': 'f' }

    @property
    def headers(self):
        self._headers['Cookie'] = self.session.token
        return self._headers

    @property
    def request_url(self):
        path = '/'.join(['exchange', self.username, self.service])
        return '/{0}'.format(path)

    def get_response(self, query=None):
        connection = HTTPSConnection(self.server)
        connection.request(self.method, self.request_url, query,
                            headers=self.headers)
        resp = connection.getresponse()

        if int(resp.status) > 299 or int(resp.status) < 200:
            raise ExchangeException("%s %s" % (resp.status, resp.reason))

        return resp.read()


class ExchangeCommand(object):
    """
    Base class for Exchange commands. This really shouldn't be constructed
    directly but should be subclassed to do useful things.
    """

    def __init__(self, session):
        self.session = session

    def _get_xml(self, **kwargs):
        """
        Try to get an XML response from the server.
        @return: ElementTree response
        """
        kwargs["username"] = self.session.username
        xml = self._get_query(**kwargs)

        req = ExchangeRequest(self.session, self.exchange_method,
                                    self.dav_method)
        resp = req.get_response(Template(xml).substitute(kwargs))

        return ElementTree.fromstring(resp)

    def _get_query(self, **kwargs):
        """
        Build up the XML query for the server. Mostly just does a lot
        of template substitutions, also does a little bit of elementtree
        magic to to build the XML query.
        """
        declaration = ElementTree.ProcessingInstruction("xml", 'version="1.0"')

        request = ElementTree.Element("g:searchrequest", { "xmlns:g": "DAV:" })
        query = ElementTree.SubElement(request, "g:sql")
        query.text = Template(self.sql).substitute(kwargs)

        output = ElementTree.tostring(declaration)
        output += ElementTree.tostring(request)

        return output


class FetchCalendar(ExchangeCommand):

    exchange_method = "calendar"
    dav_method = "SEARCH"

    sql = """
        SELECT
            PidLidAllAttendeesString                 AS attendees,
            "urn:schemas:calendar:location"          AS location,
            "urn:schemas:calendar:organizer"         AS organizer,
            "urn:schemas:calendar:meetingstatus"     AS status,
            "urn:schemas:httpmail:normalizedsubject" AS subject,
            "urn:schemas:calendar:dtstart"           AS start_date,
            "urn:schemas:calendar:dtend"             AS end_date,
            "urn:schemas:calendar:timezone"          AS timezone_info,
            "urn:schemas:httpmail:textdescription"   AS description
        FROM
            Scope('SHALLOW TRAVERSAL OF "/exchange/${username}/calendar/"')
        WHERE
            NOT "urn:schemas:calendar:instancetype" = 1
            AND "DAV:contentclass" = 'urn:content-classes:appointment'
        ORDER BY
            "urn:schemas:calendar:dtstart" ASC
        """

    def execute(self, alarms=True, alarm_offset=15, **kwargs):
        exchange_xml = self._get_xml(**kwargs)
        calendar = Calendar()

        for item in exchange_xml.getchildren():
            item = item.find("{DAV:}propstat").find("{DAV:}prop")

            event = Event()
            event.add_text(item, 'subject', add_as='summary')
            event.add_text(item, 'location')
            event.add_text(item, 'status')
            event.add_text(item, 'description')
            event.add_date(item, 'start_date')
            event.add_date(item, 'end_date')
            event.add_organizer(item)

            try:
                event.finalize()
                calendar.add_component(event)
            except InvalidEventError:
                print "Rejected event"

        return calendar
