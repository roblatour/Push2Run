
Push2Run Version History
==============================================================

Welcome to the Push2Run Version History page.  
  
This page details the changes made with various versions of Push2Run.  
  
* * *
<div class="comparison">

 **Version** | **Release date** | **Update**
------ | ------- | ----- |
5.0 | 2024-12-xx | Released as open source on Github <br>Improvements to allow for faster processing when main window is not visible<br>Updated several underlying components
4.8.3 | 2024-06-14 | Added support for Pushover authorization when 2FA is required
4.8.2 | 2024-03-10 | Corrected a bug introduced in v4.8.1 when using up or down key arrow on the main window when the viewing of all cards is off.  
4.8.1 | 2024-03-09 | An option has been added to the main window to filter out cards that are off.<br>Other minor updates and corrections.
4.8 | 2023-08-20 | Corrected variable text not being correctly processed when using MQTT.
4.7 | 2023-06-23 | Added support for multiple Pushbullet titles.<br>For example, the Pushbullet title filter can now be set to something like “Push2Run W11-Server , From Alexa” which will allow Push2Run to react to Pushbullet pushes with either “Push2Run W11-Server” or “From Alexa” in their Pushbullet title.<br>This change has been made in support of the PC Pusher Alexa skill in development by Push2Run user ‘LazySpaniard’.<br>With this change Push2Run can support the use of variable phrases when using Alexa, the PC Pusher skill and Pushbullet (without the need for IFTTT).<br>For example, with the ‘PC Pusher’ Amazon skill enabled you can say:<br>“Alexa, tell my computer to play 99 red balloons“<br>and have Push2Run do that for you.<br>For more information, please see: [http://pcpusher.s3-website.us-east-2.amazonaws.com/howto/install/](http://pcpusher.s3-website.us-east-2.amazonaws.com/howto/install/)<br>Added support for {PRTSC} and {ALT}{PRTSC} in the ‘Keys to send’ Push2Run card field.<br>Other minor updates and corrections.
4.6 | 2023-02-07 | Change to support running in the Windows Sandbox<br>Another shot at fixing the pesky first time password prompt issue (but still not fully resolved).<br>Corrected could not tab from mqtt port field in options
4.5 | 2022-10-04 | Improved MQTT Publish and Subscribe logic:<br>- Great flexibility is now provided for defining MQTT topics and payloads.<br>- Empty payloads can now be published.  
4.4 | 2022-05-25 | Added support for ‘Everything’ searches. For more information please see: [https://www.push2run.com/help/everything.html](https://www.push2run.com/help/everything.html)<br>Further updated/improved processing when MQTT topics are changed in the Options window.<br>Unsubscribe to all MQTT topics when Push2Run is shutdown, or the user disables MQTT processing.<br>Corrected a bug with the auto update check feature continuing to work when disabled in the settings.
4.3 | 2022-02-06 | Corrected variable text (‘$’) not being handled correctly in some cases<br>Push2Run cards which are dragged and dropped to the Desktop or a file folder will now have underscores used in place of spaces in their file names<br>Corrected a bug where the Systray icon was not going red when Push2Run is paused / the master switch is turned off<br>Streamlined processing when MQTT topics are changed in the Options windows<br>Strongly named Push2Run and Push2RunReloader to make them more tamper resistant
4.2 | 2022-01-30 | Added support for Regex groups in the 'Open', 'Parameters' and 'Keys to send' fields<br>Corrected an issue with importing some .p2r cards<br>Corrected a problem with the Push2Run systray icon showing up with a red background when only Dropbox is used as a trigger<br>Other minor updates and corrections<br>
4.1 | 2022-01-28 | Added an automatic re-connect to the MQTT server should it become temporarily unavailable.<br>Added the ability to publish a MQTT message.<br>Added support for Regex expression matching on the ‘Listen for’ field of the Push2Run card.<br>Corrected an issue where the status bar showed MQTT as not connected when it was connected.<br>
4 | 2022-02-26 | Added MQTT triggering
3.7.1 | 2022-01-10 | Updated to suppress duplicate network notifications as the program attempts to reconnect to a service that has gone down but may now be up again.
3.7.0 | 2022-01-03 | Windows notifications have been added. While these are turned off by default, they can be turned on from Push2Run – Options – Notification window.<br>Changed the way the program processes command line requests, allowing this feature to work on some systems where before it would not.<br>For longer term supportability, the program has been updated to use Microsoft’s dot net framework version 4.8
3.6.2 | 2021-10-28 | Should Push2Run's setting file be found to be corrupt at start-up and a suitable backup is available, Push2Run will now prompt you to do an automated recovery.  Of note, Push2Run already periodically backs up your settings and database files automatically.  
3.6.1 | 2021-06-06 | Now supports ‘Listen for’ phrases of: \* something $<br>For example, instead of:<br>set my computer’s volume to $<br>and/or<br>set my pc’s volume to $<br>you can now use:<br>\* volume to $<br>Other minor additional checks when saving a Push2Run card
3.6 | 2021-04-02 | Added a feature to run more than one Push2Run card based on separating words.<br>For example, instead of saying:<br>“OK Google, tell my computer to open the calculator”<br>followed by<br>“OK Google, tell my computer open word”<br>You can now just say"<br>“OK Google, tell my computer to open the calculator and open word”<br>For a deeper explanation of this new feature you are welcome to watch the following Youtube video: [https://youtu.be/mDwJ2fT8rBE](https://youtu.be/mDwJ2fT8rBE)<br>Other minor updates and corrections  
3.5.3 | 2021-03-14 | Added new Import options:<br> - Automatically turn on imported cards<br>(prior to this release cards were turned off when imported)<br>- Confirm when importing is done<br>- Update the description of imported cards to end with ‘(Imported)’<br>Other minor updates and corrections  
3.5.2 | 2021-03-04 | Corrected program unloading when filtering on text that does not exist  
3.5.1 | 2021-01-31 | Corrected a serious bug with undoing an action when a filter is on
3.5 | 2021-01-30 | Improved automatic notifications of new Push2Run releases<br>Added semi-automatic download and install of new releases<br>Added option to skip notifications of the current release if you would rather not upgrade to it, while still allowing you to be automatically notified of the next release<br>The program’s changelog (this file) is now viewable from within the program itself, no need to visit the website just to see what’s changed<br>Extended undo capabilities to more actions<br>Other minor updates and corrections
3.4.3 | 2021-01-16 | Added automatic daily backups (seven generations) of the Push2Run database and settings files<br>Corrects initial install problem on some machines which caused the program to prompt for a database password when it should not have<br>Updated installer to support installing on a Raspberry Pi running Windows 10<br>(also Push2Run now tested and working on this platform)<br>Other minor updates and corrections  
3.4.2 | 2021-01-01 | Adds support for wild card expressions in the 'Listen for' field, for example:<br>'\* calculator' can now replace, 'open the calculator', 'start the calculator', 'run the calculator', etc..<br>Added support for substituting Window’s environmental variables in the open, start directory, and parameters field<br>Other minor updates and corrections  
3.4.1 | 2020-11-13 | Corrected bug introduced in version 3.4 whereby a card with a 'Listen for' phrase ending with a $ was not being processed correctly  
3.4 | 2020-11-11 | Added functionality to optionally perform specific actions when there are no matching or enabled cards for the command you have given<br>Other minor updates and corrections  
3.3  | 2020-10-23 | Added a safeguard against runaway queries to Pushover  
3.2  | 2020-10-11 | Added functionality to import and export all cards to and from the database  
3.1  | 2020-09-17 | Added full support for multiple triggers<br>Added 'Clear' button to Session log window  
3.0.3 | 2020-04-24 | Updated to handle Dropbox synchronization issues
3.0.2 | 2020-03-31 | Update to consider Dropbox CreatedAt time when processing requests triggered by Dropbox - this to allow Push2Run to ignore old commands issued, for example, when the PC was in sleep mode.<br>Of note, to take advantage of this feature, your IFTTT applet needs to be updated as described in step 14 here: [https://push2run.com/setup\_dropbox.html](https://push2run.com/setup_dropbox.html)<br>Corrected auto lookup and edit of program name when only the program name, for example, when "chrome.exe" is entered in a Push2Run card's Open field.
3.0.1 | 2020-03-20 | Updated to handle new Pushover server responses  
3 | 2020-02-16 | Added support for Dropbox based triggered processing  
2.5.4 | 2020-02-09 | Minor updates related to upgrading third party libraries  
2.5.3 | 2019-12-11 | Corrected program crash when using Pushbullet and the Title field is left blank on the IFTTT applet<br>Added a fix to ensure the status of all Push2Run cards is correctly interpreted each time a command is issued  
2.5.2 | 2019-12-01 | Corrected manual check for update  
2.5.1 | 2019-11-24 | Corrected issue with authenticating Pushover credentials  
2.5 | 2019-11-23 | Added support for Pushover
2.4 | 2019-09-26 | Added functionality to prevent Pushbullet accounts from automatically expiring. With this change, Pushbullet accounts will only expire 31 days after Push2Run's last use.<br>Changed the pop-up notifying the user of a Push2Run version upgrade so that it doesn't prevent Push2Run background processing.  
2.3 | 2019-08-22 | Updated approach for managing memory, may help some with the program using high cpu after 48 hours of continuous use  
2.2 | 2019-04-20 | Added option to save session log to disk<br>Improved accessibility (switch on/off wording)<br>Fixed looping startup for Windows 7 with admin rights / uac on/off<br>Updated log at startup to show more info<br>Corrected typo in the word Admin on the add/change window
2.1.2 | 2019-02-27 | Use TLS 1.2 to connect to Pushbullet<br>Suppress command line entries like StartAdmin on startup<br>Handle switching of primary network connection (when docking / undocking)
2.1.1 | 2019-02-08 | Can now have Push2Run send keystrokes to the currently active window by setting the Push2Run card's open field to "Active Window" (without the quotes)<br>Corrected allow the program to start with administrative privileges at system startup time in accordance with the startup options chosen
2.1 | 2019-01-23 | Added sending keystrokes<br>Added ability to set window state of window opening<br>Auto heal startup code – ensure master record is at the top of the main window<br>Ensure message box windows appear on top<br>Corrected problem if clicked 'x' on the new password window
2.0.5 | 2018-09-03 | corrected issue with an action potentially running a second time after PC awakes from sleep<br>retain setting for session log auto-scroll feature
2.0.4 | 2018-07-01 | corrected application checking for update when the option to check for update is unchecked
2.0.3 | 2018-05-20 | added ability to edit a filtered Push2Run card back in 
2.0.2 | 2018-05-20 | a variety of menu actions, such as: add, delete, edit, undo, ... are now disabled when the filter is on - this prevents an issue with undo clearing the database.<br>when saving a Push2Run card to disk, the characters '/' and '\\' are now substituted with the '-' character in the filename.<br>This resolves an issue where the Push2Run card could not be saved to disk if its filename contained one or more '/' or '\\' characters.
2.0.1 | 2018-05-11 | fixed run from menu command when there is a variable ($) in the Push2Run card's parameter field
2 | 2018-05-10 | added a filter feature on the Main window, this to help you find the Push2Run card you want much faster - especially if you have a large collection of Push2Run cards<br>added 'copy to clipboard' button to Session Log Window<br>improved confirming program locations for special programs
1.9 | 2018-04-10 | corrected lag with program response when Pushbullet Access Token is not yet entered
1.8 | 2018-03-10 | save main window columns widths<br>fixed loading of a card when it is opened from the desktop<br>save cards with a file extension of .p2r rather than .P2R
1.7 | 2018-03-03 | added a link to the Community Support Forum on the About/Help window<br>fixed some drag and drop issues  
1.6 | 2018-02-18 | added drag and drop support from the Main window for Push2Run cards, allowing Push2Run entries to be saved and shared<br>when manually running a Push2Run card which includes a '$' (variable text placeholder)  the program will now prompt for text you would like to use as your variable text<br>when removing the password requirement, the program will first require that the current password is entered<br>restore Pushbullet connection if the network or Pushbullet connection is lost and later restored
1.5 | 2018-02-10 | the words entered in the 'Listen for' field are now automatically converted to lower case<br>in the 'Open' and 'Parameters' fields the blanks between spoken words can now be automatically removed or replaced by something else (for example a comma)  
1.4 | 2018-02-04 | corrected install and uninstall behaviors<br>stopped the session log from automatically closing when dragging and dropping information between windows<br>no longer log issues when trying to make the newly launched program the top most window 
1.3 | 2018-01-30 | re-establish Pushbullet connection following a PC hibernation event  
1.2 | 2018-01-26 | corrected communications with Pushbullet where a decimal is needed in place of a comma for some languages (for example Swedish and Serbian)  
1.1 | 2018-01-24 | added support for variable spoken command endings, such as OK Google tell my computer to search for $ (with the $ being the variable command ending)<br>updates to reduce start-up prompt for password issue  
1 | 2018-01-20 | initial release
</div>

* * *  

Copyright © 2018 - 2024 Rob Latour.  
All Rights Reserved.
