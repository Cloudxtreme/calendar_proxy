# vim: set filencoding=utf8
"""
Timezone Definitions

@author: Mike Crute (mcrute@gmail.com)
@organization: SoftGroup Interactive, Inc.
@date: April 26, 2009
"""

from datetime import tzinfo, timedelta


class EST(tzinfo):

    def tzname(self, dt):
        return "EST"

    def utcoffset(self, dt):
        return timedelta(0)

    dst = utcoffset
