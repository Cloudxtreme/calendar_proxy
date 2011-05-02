#!/usr/bin/env python
# vim: set filencoding=utf8
"""
Exchange Calendar CLI App

@author: Mike Crute (mcrute@ag.com)
@organization: American Greetings Interactive
@date: May 02, 2011
"""

from os import path
from util import config_dict
from exchange.wsgi import CalendarApp
from wsgiref.simple_server import make_server


def main():
    config = config_dict(path.expanduser('~/.exchange.cfg'))
    password = open(path.expanduser('~/.exchange.pass'), 'r').read()

    app = CalendarApp(config['exchange']['server'],
                      config['exchange']['user'],
                      password)

    print app(None, lambda x, y: None)


if __name__ == '__main__':
    main()
