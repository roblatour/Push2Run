
# Push2Run 
## How to send your own text, to be spoken aloud, to your Google device
   
Welcome to the Push2Run 'How to send your own text, to be spoken aloud, to your Google device' page.   
  
This page provides 'how to' instructions for sending a custom message to your Google device from your Windows PC or laptop.  
  
By following the instruction on this page you should be able to type a command  from the Windows Command Prompt where the words you type will be spoken through your Google Home, Google Mini, and/or Google Max.  
  
It also explains on how to integrate this functionality within Push2Run so that you can, for example, say to your Google Device:  
  
"**OK Google ask my computer if it is on"**, and then hear the response:  
  
"**Yes, your computer is on**" (if it is).  
  

* * *
  
To do this you will need download, install and configure some additional software called [Cast](https://github.com/roblatour/Cast)
  
The following details all of that:    

**Step 1 \- Download and install Cast for Windows**

  
To download the Cast installer for Windows please : [click here](https://6ec1f0a2f74d4d0c2019-591364a760543a57f40bab2c37672676.ssl.cf5.rackcdn.com/CastSetup.exe)  

  
Once you have downloaded the file, double click on it to install it.  
  
Please take note of the directory in which Cast was installed (Cast will tell you where that is as part of the install process).  Normally, the directory is 'C:\\Program Files\\Cast'  


**Step 2 - Get the Google device name(s) that you want to send your text to**

  
If you want to send your text all your Google devices, you can skip this step.  
  
Otherwise, you will need to know the names of the devices to which you want to send your text.  
  
Press the Windows Key and the 'R' key at the same time.  
  
In the pop-up box that appears, type in:

        CMD

and press enter.  
  
At the C:> prompt that appears, enter the following:

        c:

        cd \\

        cd "Program Files\\Cast"

        cast -inventory

and press enter.  
  
(the above assume Cast was installed in the directory  'C:\\Program Files\\Cast')  
  
You should now see an inventory of your devices, showing their device names.  
    

**Step 3 \- Run a quick test**

  
In this example, lets assume you want to send your text to two devices named 'Office Home' and 'Basement Mini'.  
  
Continuing from Step 2 above, enter the following:

         cast -device "Office Home" "Basement Mini" -text This is a test

 you should hear your Google device say "this is a test".  
  
Notes:  
  
If you don't hear anything from one of your devices, it may be because it is muted or has its volume set really low.  Check the inventory results for that.  If you need to fix that you can issue the following command as needed:

         cast -device "Office Home" -unmute -volume 30

  
For more information about what Cast can do, just type

         cast -help

  
**Step 4 - Integrate with Push2Run**
  
Double click on the following Push2Run card to download it, and then open it to install it within Push2Run  
  
[![Cat Notify](/images/CastNotify.jpg)](/misc/cast_confirm_computer_is_running.p2r)
  
Notes:  
   
i. you will need to change the location of the cast.exe program if it was not installed in C:\\Program Files\\Cast  
  
ii. in the above example, cast will broadcast the text to all your Google Devices.  If you want to broadcast your text to specific devices only, in the parameters field you would enter the same type of command as is shown in Step 3 above (but exclude the word Cast).  For example you would enter in the parameters field:  
\-device "Office Home" "Basement Mini" -text yes, your computer is on  
  
iii. the [Push2Run Setup Instructions](setup.md) show you how to create an IFTTT applet that is triggered by the expression, "OK Google tell my computer to $".    However, in this example you find it better to create another IFTTT applet that is triggered by the expression, "OK Google ask my computer if it is on" and has a fixed PushBullet body text of "if it is on".   

That's it, hope this write up will be of use to you.  


* * *
 ## Support Push2Run

 To help support Push2Run, or to just say thanks, you're welcome to 'buy me a coffee'<br><br>
[<img alt="buy me  a coffee" width="200px" src="https://cdn.buymeacoffee.com/buttons/v2/default-blue.png" />](https://www.buymeacoffee.com/roblatour)
* * *
Copyright © 2018 - 2025 Rob Latour
* * *
