# vim: set filencoding=utf8
"""
Exchange Proxy Utility Functions

@author: Mike Crute (mcrute@ag.com)
@organization: American Greetings Interactive
@date: February 15, 2010
"""

from collections import defaultdict
from ConfigParser import ConfigParser


def config_dict(config_file):
    """
    Load a ConfigParser config as a dictionary. Will also
    attempt to do simple stuff like convert ints.
    """
    config = ConfigParser()
    config.read(config_file)
    output = defaultdict(dict)

    for section in config.sections():
        for key, value in config.items(section):
            if value.isdigit():
                value = int(value)

            output[section][key] = value

    return dict(output)

