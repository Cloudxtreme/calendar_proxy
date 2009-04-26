# -*- coding: utf-8 -*-
"""
Exchange Commands

@author: Mike Crute (mcrute@gmail.com)
@date: November 10, 2008
@version: $Revision$

This is a set of classes that starts to define a set of classes for
fetching data using Exchange's WebDAV API. This is still pretty
development code but it does the trick. Watch out, it doesn't consider
many corner cases.

$Id$
"""
import xml.etree.cElementTree as etree

from httplib import HTTPSConnection
from string import Template
from datetime import datetime, timedelta

import dateutil.parser
from icalendar import Calendar, Event, Alarm

from exchange import ExchangeException


class ExchangeCommand(object):
    """
    Base class for Exchange commands. This really shouldn't be constructed
    directly but should be subclassed to do useful things.
    """

    #: Base URL for Exchange commands.
    BASE_URL = Template("/exchange/${username}/${method}")

    #: Basic headers that are required for all requests
    BASE_HEADERS = {
       "Content-Type": 'text/xml; charset="UTF-8"',
       "Depth": "0",
       "Translate": "f",
       }

    def __init__(self, server, authenticator=None):
        self.server = server
        self.authenticator = authenticator

    def _get_xml(self, **kwargs):
        """
        Try to get an XML response from the server.
        @return: ElementTree response
        """
        if not self.authenticator.authenticated:
            raise ExchangeException("Not authenticated. Call authenticate() first.")

        # Lets forcibly override the username with the user we're querying as
        kwargs["username"] = self.authenticator.username

        xml = self._get_query(**kwargs)
        url = self.BASE_URL.substitute({ "username": self.authenticator.username,
                                         "method": self.exchange_method })
        query = Template(xml).substitute(kwargs)
        send_headers = self.authenticator.patch_headers(self.BASE_HEADERS)

        conn = HTTPSConnection(self.server)
        conn.request(self.dav_method.upper(), url, query, headers=send_headers)
        resp = conn.getresponse()

        # TODO: Lets determine authentication errors here and fix them.
        if int(resp.status) > 299 or int(resp.status) < 200:
            raise ExchangeException("%s %s" % (resp.status, resp.reason))

        return etree.fromstring(resp.read())

    def _get_query(self, **kwargs):
        """
        Build up the XML query for the server. Mostly just does a lot
        of template substitutions, also does a little bit of elementtree
        magic to to build the XML query.
        """
        declaration = etree.ProcessingInstruction("xml", 'version="1.0"')

        request = etree.Element("g:searchrequest", { "xmlns:g": "DAV:" })
        query = etree.SubElement(request, "g:sql")
        query.text = Template(self.sql).substitute(kwargs)

        output = etree.tostring(declaration)
        output += etree.tostring(request)

        return output


class FetchCalendar(ExchangeCommand):
    exchange_method = "calendar"
    dav_method = "search"

    sql = """
        SELECT
            "urn:schemas:calendar:location"          AS location,
            "urn:schemas:httpmail:normalizedsubject" AS subject,
            "urn:schemas:calendar:dtstart"           AS start_date,
            "urn:schemas:calendar:dtend"             AS end_date,
            "urn:schemas:calendar:busystatus"        AS busy_status,
            "urn:schemas:calendar:instancetype"      AS instance_type,
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

            # These tests may look funny but the result of item.find
            # does NOT evaluate to true even though it is not None
            # so, we have to check the interface of the returned item
            # to make sure its usable.

            subject = item.find("subject")
            if hasattr(subject, "text"):
                event.add("summary", subject.text)

            location = item.find("location")
            if hasattr(location, "text"):
                event.add("location", location.text)

            description = item.find("description")
            if hasattr(description, "text"):
                event.add("description", description.text)

            # Dates should always exist
            start_date = dateutil.parser.parse(item.find("start_date").text)
            event.add("dtstart", start_date)

            end_date = dateutil.parser.parse(item.find("end_date").text)
            event.add("dtend", end_date)

            if item.get("timezone_info"):
                """This comes back from Exchange as already formatted
                ical data. We probably need to parse and re-construct
                it unless the icalendar api lets us just dump it out.
                """
                pass

            if alarms and start_date > datetime.now(tz=EST()):
                alarm = Alarm()
                alarm.add("action", "DISPLAY")
                alarm.add("description", "REMINDER")
                alarm.add("trigger", timedelta(minutes=alarm_offset))
                event.add_component(alarm)

            calendar.add_component(event)

        return calendar
