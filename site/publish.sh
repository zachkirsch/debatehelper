#!/bin/bash
HOST=ftp.zachkirsch.com #This is the FTP servers host or IP address.
USER=zachkirsch         #This is the FTP user that has access to the server.

# Read Password
echo -n "Password: "
read -s PASS

# Send to site at godaddy server

## ENTERING FTP SHELL ##
ftp -in << EOF
open $HOST
user $USER $PASS
cd public_html

mkdir css

mput *.html
cd css
lcd css
mput master.css

close
bye
