#!/bin/bash
HOST=ftp.zachkirsch.com #This is the FTP servers host or IP address.
USER=debatehelper@zachkirsch.com #This is the FTP user that has access.

# Read Password
echo -n "Password: "
read -s PASS

# Send to site at godaddy server

## ENTERING FTP SHELL ##
ftp -in << EOF
open $HOST
user $USER $PASS

mkdir css
mkdir img
mkdir php
mkdir js
mkdir static
mkdir updates

mput *.html
cd css
lcd css
mput *
cd ../static
lcd ../static
mput *
cd ../img
lcd ../img
mput *.png *.ico
cd ../php
lcd ../php
mput *.php
cd ../js
lcd ../js
mput *.js
cd ../updates
lcd ../updates
mput *

close
bye
