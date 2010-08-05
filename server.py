#!/usr/bin/env python
# vim: set filencoding=utf8
"""
Exchange Calendar Proxy Server

@author: Mike Crute (mcrute@gmail.com)
@organization: SoftGroup Interactive, Inc.
@date: April 26, 2009
"""

from os import path
from util import config_dict
from exchange.wsgi import CalendarApp
from wsgiref.simple_server import make_server


def main():
    config = config_dict(path.expanduser('~/.exchange.cfg'))
    password = open(path.expanduser('~/.exchange.pass'), 'r').read()

    try:
        app = CalendarApp(config['exchange']['server'],
                          config['exchange']['user'],
                          password)

        make_server(config['local_server']['address'],
                    config['local_server']['port'], app).serve_forever()
    except KeyboardInterrupt:
        print '\n All done, shutting down.'


if __name__ == '__main__':
    main()
