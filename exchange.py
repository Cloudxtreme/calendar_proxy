"""
Exchange API -> iCal Fie Proxy

@author: Mike Crute (mcrute@gmail.com)
@date: November 10, 2008
@version: $Revision$

This is a set of classes that starts to define a set of classes for
fetching data using Exchange's WebDAV API. This is still pretty 
development code but it does the trick. Watch out, it doesn't consider
many corner cases.

== Config File Format ==
[exchange]
server = <your_mail_server>

[authentication]
user = <your_username>
password = <your_password>

[server]
bind_address = 0.0.0.0
bind_port = 8000
"""

######################################################################
# Standard Library Imports
######################################################################
import urllib
import xml.etree.cElementTree as etree

from getpass import getpass
from copy import copy
from httplib import HTTPSConnection
from string import Template
from Cookie import SimpleCookie
from datetime import datetime, timedelta, tzinfo

######################################################################
# Third-Party Library Imports
######################################################################
#: These can all be found on pypi
import dateutil.parser
from icalendar import Calendar, Event, Alarm


######################################################################
# Exceptions
######################################################################
class ExchangeException(Exception):
    """
    Exception that is thrown by all Exchange handling code.
    """
    pass


######################################################################
# Exchange Enumerations
######################################################################
class CalendarInstanceType(object):
    """
    Enum for Calendar Instance Types
    @see: http://msdn.microsoft.com/en-us/library/ms870457(EXCHG.65).aspx
    
    This ended up not being used but its probably good to keep around for 
    future use.
    """
    #: Single appointment
    SINGLE = 0
    
    #: Master recurring appointment
    MASTER = 1
    
    #: Single instance of a recurring appointment
    INSTANCE = 2
    
    #: Exception to a recurring appointment
    EXCEPTION = 3


######################################################################
# EST Timezone
######################################################################
class EST(tzinfo):

    def tzname(self, dt):
        return "EST"

    def utcoffset(self, dt):
        return timedelta(0)

    dst = utcoffset

######################################################################
# Exchange Authenticators
######################################################################
class ExchangeAuthenticator(object):
    """
    Exchange Authenticator Interface
    """
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
        raise NotImplemented("Implement in a subclass.")
        
    def patch_headers(self, headers):
        raise NotImplemented("Implement in a subclass.")

class CookieAuthenticator(ExchangeAuthenticator):
    #: Authentication DLL on the exchange server
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
        """
        Patch the headers dictionary with authentication information and 
        return the patched dictionary. I'm not a big fan of patching 
        dictionaries in-place so just make a copy first.
        """
        out_headers = copy(headers)
        out_headers["Cookie"] = self._auth_cookie
        return out_headers

######################################################################
# Exchange Command Base Class
######################################################################
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


######################################################################
# Exchange Commands
######################################################################
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


######################################################################
# Testing Server
######################################################################
if __name__ == "__main__":
    import sys
    from BaseHTTPServer import HTTPServer, BaseHTTPRequestHandler
    from ConfigParser import ConfigParser
    
    config = ConfigParser()
    config.read("exchange.cfg")
    
    username = config.get("exchange", "user")
    password = getpass("Exchange Password: ")

    class CalendarHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            print "> GET CALENDARS"

            server = config.get("exchange", "server")
            fetcher = FetchCalendar(server)
            
            authenticator = CookieAuthenticator(server)
            authenticator.authenticate(username, password)
            fetcher.authenticator = authenticator
            
            calendar = fetcher.execute()
            
            self.wfile.write(calendar.as_string())
            self.wfile.close()
            
    try:
        bind_address = config.get("local_server", "address")
        bind_port = int(config.get("local_server", "port"))
        
        print "Exchange iCal Proxy Running on port %d" % bind_port
        
        server = HTTPServer((bind_address, bind_port), CalendarHandler)
        server.serve_forever()
    except KeyboardInterrupt:
        print "\n Fine, be like that."
        sys.exit()
