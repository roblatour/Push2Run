# Push2Run - Setup instructions      
  
Welcome to the Push2Run set up page.  
  
In short, Push2Run needs to be triggered to do things and this webpage explains how to do that.  
  
There are various options for triggering Push2Run, these include using:  

- **an Amazon Alexa device, the [PC Commander](https://PCCommander.net/) skill and Pushbullet.**  This set-up provides maximum versatility as it supports the use variable expressions - meaning you can say a phrase that doesn't need to be fully pre-determined.  For example, you could say "Alexa, tell my computer to play 99 Red Balloons" (or any other song you like);

- **an Apple iPhone/iPad and either** [**Dropbox**](https://www.dropbox.com/) **or** [**Pushbullet**](https://www.pushbullet.com)**.**  This set-ups also provides maximum versatility as it too supports the use variable expressions.   However, you need to include a slight pause after the trigger phrase, for example you could say "Siri, tell my computer to" (pause for a few seconds and then say) "play 99 Red Balloons";

- **a Google Device, or other service supported by [IFTTT](https://ifttt.com),**<br>**and one or more of the following: Dropbox, Pushbullet and/or** [**Pushover**](https://pushover.net/).  With these set-ups, you can say something like "Ok Google, (trigger) shutdown my computer".   However, the things you say will need to be specific, and configured as unique IFTTT applets ahead of time;

- **MQTT.** With this set up you can have Push2Run control your computer based on a MQTT messages;  
  
- **the Windows command line**; and

- **the Push2Run application itself.**

  
Detailed instructions for the various options above are either linked to or written up further below.  
  

Now, a few words about the use of PC Commander, IFTTT, Dropbox, Pushbullet and/or Pushover in conjunction Push2Run ...  
  
All services are quite reliable and have many features going beyond those needed in support of Push2Run.  
  
IFTTT offers both free and paid subscription plans; here is where you can find more information about that:  [https://ifttt.com/plans](https://ifttt.com/plans) .  Also, at one point IFTTT was able to work with Google in such a way that variable text was supported.  Sadly, however this was discontinued by Google in September 2022.  The result of this is where it once was that only one IFTTT applet was required for most, if not all the needs of Push2Run, now one specific applet is need for each specific action you would like Push2Run to take.   
  
PC Commander is an Alexa skill specifically designed to work hand in hand with Pushbullet to trigger Push2Run.  Using Alexa, PC Commander, Pushbullet and Push2Run you can execute a wide range of tasks on your Windows computer using simple voice commands. Whether you’re looking to open applications, control media playback or manage system settings, using PC Commander provides an excellent, user-friendly and efficient way to enhance your home automation setup.  PC commanders offers 150 free uses a month, but also a paid subscription if you need more (\*).  
  
Here too is a recap of the Dropbox, Pushbullet and Pushover services, with some key differences as they relate to their use with Push2Run (as of the last time I looked):  

<div class="comparison">

 **Category** | **Dropbox** | **Pushbullet** | **Pushover**
------ | ----- | ----- | ----- |
**Free account** | unlimited time; 500 pushes per month | unlimited time; 500 pushes per month | 30 days; 25,000 pushes
**Paid account (\*)** | Not required for use with Push2Run but if you would like additional Dropbox space please see [here](https://www.dropbox.com/individual/plans-comparison)| $4.99 US per month or $39.99 US per year | one time license of $4.99 US per platform; Push2Run only needs the Desktop platform
**Use with IFTTT and Push2Run** | each IFTTT applet allows you to identify on exactly which computer(s) you want a particular IFTTT triggered command to run| each IFTTT applet allows you to identify on exactly which computer(s) you want a particular IFTTT triggered command to run|each IFTTT applet triggers one or all Pushover devices, so if you are running Push2Run on three or more computers you may not be able to tailor exactly which computers you want a particular IFTTT triggered command to run
**Status reporting** | There is no status at bottom of the main window or systray icon | Status at bottom of the main window and systray icon confirms when the connection is active | Status at bottom of the main window and systray icon confirms when the connection is active
**Misc. notes** | Requires at least 7 seconds between commands. Push2Run only monitors the one Dropbox folder (which you specify) and only interacts with one file (which you specify) in that folder. |  If you are using the free Pushbullet service and if you go beyond a month not using either Push2Run with Pushbullet, or Pushbullet separately, then Pushbullet may deactivate your free account in which case you may need to set up a new one. | Service may require re-authentication if you switch network services when Push2Run is running. For example, if you switch between using your regular network and a VPN the service may require re-authentication.

</div>
  
  
**(\*) Please note:**<br>Push2Run is funded solely by the [donations](https://www.buymeacoffee.com/roblatour) of kind people such as yourself. IFTTT, Dropbox, Pushbullet, Pushover and PC Commander do not provide any form of commissions or kick-backs in support of Push2Run.  


For further set up instructions please click on a link below for triggering Push2Run with:  

<div class="comparison">

**Triggered by** | Please also see the Misc. Notes <br>(directly below) where indicated
------ | ----- |
[an Apple iPhone/iPad with Dropbox](setup_Apple_Dropbox.md) | A
[an Apple iPhone/iPad with Pushbullet](setup_Apple_Pushbullet.md) | A
[an Apple iPhone/iPad with Pushover](setup_Apple_Pushover.md) | A
[an Amazon Alexa device, the PC Commander skill and Pushbullet](https://PCCommander.net/howto/push2run/) | 
[a Google Assistant device with IFTTT and Dropbox](setup_Google_IFTTT_Dropbox.md) | B, C
[a Google Assistant device with IFTTT and Pushbullet](setup_Google_IFTTT_Pushbullet.md) | B, C
[a Google Assistant device with IFTTT and Pushover](setup_Google_IFTTT_Pushover.md) | B, C
[MQTT](setup_mqtt.md) | 
[the Windows command line](#part3) | 
[Push2Run itself.](#part4) | 
</div>


**Misc. Notes:**
<br> 

**A.**<br>
For all Apple iPhone/iPad set ups:  
  
Please also see '[Using Everything](everything.html)' for optional integration with the Everything freeware file search utility.  
  
**B.** <br>
For all set ups involving IFTTT:  
  
It is possible to set up Push2Run to use one, two or three of the Dropbox, Pushbullet and Pushover services in conjunction with IFTTT. However, for this you will require a subscription to [IFTTT's Pro service](https://ifttt.com/plans).  
  
IFTTT's Pro service provides for several advantages, but in particular for use with Push2Run it allows the triggering of more than one action.
  
The reason you might want to do this is for redundancy. For example, if one of Dropbox, Pushbullet or Pushover was down or responding slowly, triggered actions would still flow through the other two.
  
Push2Run has built in logic to ignore multiple triggered actions for the same thing. So, in short, Push2Run works the same with or without redundant services, however it is more reliable with redundant services enabled.  
  
To set up an IFTTT Pro applet just follow the instructions above for setting up a Dropbox, Pushbullet, or Pushover applet. Once done, add either one or two of the other services as additional actions. Please refer to the set up instructions above for the details on how to set up each service.  
   
**C.**<br>
For all set ups involving Google Assistant device:
  
In working with my Google Assistant device and IFTTT, I found a way to say "OK Google, do something" instead of "OK Google, activate do something" and have it work.  
  
If you would like to do this too, then to get this to work go into the Google Home app on your phone, tap the 'Routines' icon, tap the colourful '+' sign in the bottom right of the screen, tap 'Add Starter', tap 'When I say to Google Assistant', enter phase you are working with but do not include the word 'activate' - for example "open the calculator" (without the quotes), tap 'Add starter', tap 'Add Action', tap 'Try adding your own', enter the same phase as above but this time be sure to include the 'activate' at the beginning of it - for example "activate open the calculator" (without the quotes), tap 'Done'.  
  
Of note, with the above Google will not respond with a confirmation. If you would like to also hear a confirmation, then tap on 'Add action', tap on 'Communicate and announce', tap on the check box beside 'Make an announcement, tap on the option 'Announce on the device that started the routine', enter whatever you want in 'Your Message' - for example "OK, opening the calculator", tap on the check symbol at the top right of the screen, tap 'Save', tap 'Done'.  
  
Tap 'Save', and leave the Google Home app.  
  
After that, I can say "OK Google, open the calculator" and it will work.  


* * *

<a name="part3" id="part3"></a>
**Triggering with the Windows command line**
  
Push2Run does not need any configuration to allow it to trigger actions from the command line.  
  
To open the Windows command line, just press the \[Windows\] and \[R\] keys together at the same time, and enter:  
cmd  
  
So for example, out of the box, from the Windows Command line you can simply type:  
  
"C:\\Program Files\\Push2Run\\Push2Run" open the calculator  
  
and it will work.  
  

* * *
<a name="part4" id="part4"></a>
**Triggering by using Push2Run itself**
  
From within the Push2Run main window, you can run any card  by left clicking on it and pressing F12 or by selecting 'Actions' - 'Run' from the main menu.   
  
You can also run it by right clicking on it, and selecting 'Run'.  
  
* * *
  
**For additional help**  
  
Please see the [Push2Run Help documentation](help_v4.9.0.0.md).

* * *
 ## Support Push2Run

 To help support Push2Run, or to just say thanks, you're welcome to 'buy me a coffee'<br><br>
[<img alt="buy me  a coffee" width="200px" src="https://cdn.buymeacoffee.com/buttons/v2/default-blue.png" />](https://www.buymeacoffee.com/roblatour)
* * *
Copyright © 2018 - 2024 Rob Latour
* * *