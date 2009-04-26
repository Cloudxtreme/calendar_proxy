# -*- coding: utf-8 -*-
"""
Exchange Calendar Proxy Server

@author: Mike Crute (mcrute@gmail.com)
@organization: SoftGroup Interactive, Inc.
@date: April 26, 2009
@version: $Rev$

$Id$
"""
from getpass import getpass
from ConfigParser import ConfigParser
from BaseHTTPServer import HTTPServer, BaseHTTPRequestHandler

from exchange.commands import FetchCalendar
from exchange.authenticators import CookieAuthenticator


class CalendarHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        print('> GET CALENDARS')

        fetcher = FetchCalendar(self.server.exchange_server)
        authenticator = CookieAuthenticator(self.server.exchange_server)

        authenticator.authenticate(self.server.user, self.server.password)
        fetcher.authenticator = authenticator

        calendar = fetcher.execute()
        self.wfile.write(calendar.as_string()).close()


def get_un_pass(config):
    username = config.get('exchange', 'user')

    if config.has_option('exchange', 'password'):
        password = config.get('exchange', 'password')
    else:
        password = getpass('Exchange Password: ')

    return username, password


def get_host_port(config):
    bind_address = config.get('local_server', 'address')
    bind_port = int(config.get('local_server', 'port'))

    return bind_address, bind_port


def main(config_file='exchange.cfg'):
    config = ConfigParser().read(config_file)
    server_cfg = get_host_port(config)

    print('Exchange iCal Proxy Running on port {0:d}'.format(server_cfg[1]))

    server = HTTPServer(server_cfg, CalendarHandler)
    server.exchange_server = config.get('exchange', 'server')
    server.user, server.password = get_un_pass(config)

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print '\n All done, shutting down.'


if __name__ == '__main__':
    main()
