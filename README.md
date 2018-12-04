# Reply All Warning for Outlook

This is a VBA script to present a confirmation dialog to the user if they click on Reply All in Outlook.  The idea is to discourage accidentally replying to all.

### Requirements

I have tested this on Outlook 2016.  Other supported versions of Outlook that are allowed to run VBA should also work.

### Limitations

If Outlook is configured to have a preview pane visible, clicking on Reply All from the preview may skip the warning dialog. A work-around would be to disable the reading pane, or remove the Reply All button from the preview pane.

### Warning

Because this code functions as a macro, you will need to allow Outlook to run macros, which is not ideal.  You can help make this a little more secure by self-signing the code and restricting Outlook to only running signed code.

## Usage
Enable macros (or self-sign it and enable signed macros)

Enable developer tools in Outlook.  Navigate to Project1->Microsoft Outlook Objects->ThisOutlookSession.  Copy and paste the text from replyall.txt into ThisOutlookSession.




