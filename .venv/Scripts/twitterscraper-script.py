#!c:\users\mandiaye\documents\code\.venv\scripts\python.exe
# EASY-INSTALL-ENTRY-SCRIPT: 'twitterscraper==0.2.7','console_scripts','twitterscraper'
__requires__ = 'twitterscraper==0.2.7'
import re
import sys
from pkg_resources import load_entry_point

if __name__ == '__main__':
    sys.argv[0] = re.sub(r'(-script\.pyw?|\.exe)?$', '', sys.argv[0])
    sys.exit(
        load_entry_point('twitterscraper==0.2.7', 'console_scripts', 'twitterscraper')()
    )
