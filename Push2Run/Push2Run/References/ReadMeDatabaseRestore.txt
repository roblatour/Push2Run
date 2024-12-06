This directory contains a Push2Run database file and up to seven backups of the Push2Run database file.

The Push2Run database file is named 'Push2Run.db3'.

The Push2Run database backup files are named 'Push2Run-1.db3' to 'Push2Run-7.db3'.

The most recent backup file is named 'Push2Run-1.db3', the least recent backup file is named 'Push2Run-7.db3'.

If you don't see the '.db3' part of the filename then from the menu bar of Microsoft File Explorer click on 'View', and then check the option 'File name extensions'.

If you don't see the modified date associated with each file then from the menu bar of Microsoft File Explorer click on 'View', and then click on 'Details'.

Backups of your database file are made when you first start Push2Run and at the most once a day only if your database file has changed.

Accordingly, even if you are running Push2Run every day you may not have a backup for each of the last seven days, this is normal.

If you leave your computer on, Push2Run will determine if a new set of backups is required every morning just after midnight.

To restore your database from a backup (also please see the notes below as additional steps may be required depending on your situation):
   1. exit Push2Run
   2. rename the file 'Push2Run.db3' to 'Push2Run-old.db3' 
   3. rename the backup file of your choice to 'Push2Run.db3'
   
Notes:

1. After you have taken the steps above and start Push2Run again it will automatically update its backups.
   This means the least recent backup (usually 'Push2Run-7.db3') will be lost.
   Accordingly, you may want to copy all the files in this directory elsewhere as an additional backup before proceeding with the restore process noted above.

2. If you use Push2Run's password protection feature your password is stored in the Push2Run settings file in an encrypted fashion.
   This encrypted password is used to match the password you enter to gain access to your Push2Run database.
   Accordingly, if you have changed your password, or if you have changed the password protection on your database from on to off or from off to on, then you may not have access to your database information unless you also restore your settings file to the same time-period that the database file is being restored from.
   
3. Later, once you have restarted Push2Run and you have confirmed the backup has been restored to your satisfaction, you can delete the 'Push2Run-old.db3' file.  
