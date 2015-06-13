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

mput *.html
cd css
lcd css
mput master.css

close
bye
