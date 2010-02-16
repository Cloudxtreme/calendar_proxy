# vim: set filencoding=utf8
"""
Exchange Calendar Proxy Server

@author: Mike Crute (mcrute@gmail.com)
@organization: SoftGroup Interactive, Inc.
@date: April 26, 2009
"""

from getpass import getpass
from ConfigParser import ConfigParser
from BaseHTTPServer import HTTPServer, BaseHTTPRequestHandler

from exchange.commands import FetchCalendar
from exchange.authenticators import CookieAuthenticator


class CalendarHandler(BaseHTTPRequestHandler):

    def do_GET(self):
        print('* Fetching Calendars')

        fetcher = FetchCalendar(self.server.exchange_server)
        authenticator = CookieAuthenticator(self.server.exchange_server)

        authenticator.authenticate(self.server.user, self.server.password)
        fetcher.authenticator = authenticator

        calendar = fetcher.execute()
        self.wfile.write(calendar.as_string())

        # This seems to work on Linux but not Mac OS. ~mcrute
        if hasattr(self.wfile, 'close'):
            self.wfile.close()


def main(config_file='exchange.cfg'):
    config = ConfigParser()
    config.read(config_file)
    bind_address = config.get('local_server', 'address')
    bind_port = int(config.get('local_server', 'port'))

    print('Exchange iCal Proxy Running on port {0:d}'.format(bind_port))

    server = HTTPServer((bind_address, bind_port), CalendarHandler)
    server.exchange_server = config.get('exchange', 'server')
    server.user = config.get('exchange', 'user')

    if config.has_option('exchange', 'password'):
        server.password = config.get('exchange', 'password')
    else:
        server.password = getpass('Exchange Password: ')

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print '\n All done, shutting down.'


if __name__ == '__main__':
    main()
