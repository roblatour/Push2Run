This directory contains a Push2Run settings file and up to seven backups of the Push2Run settings file.

The Push2Run settings file is named 'user.config'.

The Push2Run settings backup files are named 'user-1.config' to 'user-7.config'.

The most recent backup file is named 'user-1.config', the least recent backup file is named 'user-7.config'.

If you don't see the '.config' part of the filename then from the menu bar of Microsoft File Explorer click on 'View', and then check the option 'File name extensions'.

If you don't see the modified date associated with each file then from the menu bar of Microsoft File Explorer click on 'View', and then click on 'Details'.

Backups of your settings file are made when you first start Push2Run and at the most once a day only if your settings file has changed.

Accordingly, even if you are running Push2Run every day you may not have a backup for each of the last seven days, this is normal.

If you leave your computer on, Push2Run will determine if a new set of backups is required every morning just after midnight.

To restore your settings from a backup (also please see the notes below as additional steps may be required depending on your situation):
   1. exit Push2Run
   2. rename the file 'user.config' to 'user-old.config' 
   3. rename the backup file of your choice to 'user.config'
   
Notes:

1. After you have taken the steps above and start Push2Run again it will automatically update its backups.
   This means the least recent backup (usually 'user-7.config') will be lost.
   Accordingly, you may want to copy all the files in this directory elsewhere as an additional backup before proceeding with the restore process noted above.

2. If you use Push2Run's password protection feature your password is stored in the settings file in an encrypted fashion.
   This encrypted password is used to match the password you enter to gain access to your Push2Run database.
   Accordingly, if you have changed your password, or if you have changed the password protection on your database from on to off or from off to on, then you may not have access to your database information unless you also restore your database to the same time-period that the settings file is being restored from.
   
3. Later, once you have restarted Push2Run and you have confirmed the backup has been restored to your satisfaction, you can delete the 'user-old.config' file.  
