#!/bin/bash    
HOST="ftp.zachkirsch.com"
USER="debatehelper@zachkirsch.com"
LCD="~/Documents/Developer/debatehelper/site/"
RCD="/"
#DELETE="--delete"
FAIL_EXIT="set cmd:fail-exit yes;"




# ---------------------------------- #
# ---------------------------------- #
# ---------------------------------- #

# Read Password
echo -n "Password: "
read -s PASS
echo
FTPURL="ftp://$USER:$PASS@$HOST"

# check if mkdir required
echo 'Checking if directory needs to be created on server...'
checkfolder=$(lftp -c "open -u $USER,$PASS $FTPURL; ls $RCD")
if [ "$checkfolder" == "" ];
then
	MKDIR="mkdir $RCD;"
	echo "  Directory created"
else
	echo "  Directory already exists"
fi

echo 'Uploading modified files...'
lftp -c "
$FAIL_EXIT
open '$FTPURL';
lcd $LCD;
$MKDIR
cd $RCD;
mirror --reverse \
	   $DELETE \
	   --verbose \
	   --exclude-glob .git/ \
	   --exclude-glob .DS_Store \
	   --exclude-glob *.sh \
	   --exclude-glob .* \
	   --exclude-glob unused/"
