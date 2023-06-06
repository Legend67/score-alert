# score-alert

Python-Based mail notification system for students scoring below average

Create two mails 
1. Sender and receiver
2. Enable 2FA on sender mail to generate custom app password 

## Steps to create an app password, you need 2-Step Verification on your sender Google Account.

If you use 2-Step-Verification and get a "password incorrect" error when you sign in, you can try to use an app password.

1. Go to your Google Account.
2. Select Security
3. Under "Signing in to Google," select 2-Step Verification.
4. At the bottom of the page, select App passwords.
5. Enter a name that helps you remember where you’ll use the app password.
6. Select Generate.
7. To enter the app password, follow the instructions on your screen. The app password is the 16-character code that generates on your device.
Select Done.

Copy your app password and add it in the place of password in the code

### Update the sender_mail and receiver_mail fields with youe newly created mails
After successfully running the code select the excel file you want to use to send alerts.

You can go through the sample excel file.
