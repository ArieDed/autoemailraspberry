# autoemailraspberry

Simple Python 2 script I made to be run on a raspberry pi zero connected to a sound output and a screen. Also a momentary push button is connected to the raspberry on pin GPIO 23. The script checks a outlook mail map for unread messages with a specific subject. The subject is either for just text emails or for emails with photos or videos. For all the unread emails the script will either read them out loud and/or show the photo or video on the screen. For every text email, the script will also send a stock email back to the messenger. The variables to change to make it work are:

* YOUR_OUTLOOK_EMAIL: Your own outlook email address
* YOUR_OUTLOOK_PASSWORD: The password for said email address
* OUTLOOK_MAIL_FOLDER: The folder to look into for unread emails
* SUBJECT_FOR_TEXT: Subjects for emails with just text
* SUBJECT_FOR_PHOTO_OR_VIDEO: Subject for emails with photos or videos
* REPLY_EMAIL_SUBJECT: Subject of reply email
* MESSAGEFOREMAIL: Message of the reply email (2x, both in the html message and the plain text message)

It uses a python library to read the outlook emails that can be found [here](https://github.com/awangga/outlook). Furthermore, the other python package that is needed is "rpi.gpio".

I have used this script and the raspberry for my grandpa with Parkinson. Due to him not being able to move very well and the corona lockdown, people were able to send messages and photos and videos using my website, which could then be read out loud or shown using this script. The messages were sent using the WPForms plugin for wordpress.
