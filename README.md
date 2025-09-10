# SecretSanta-Excel-Helper
Excel (Mac &amp; Windows) helper for the Secret Santa generator by Arcanis

🎄 Secret Santa Email Generator (macOS & Windows Excel)
This tool lets you generate personalized Secret Santa email drafts using Excel — no programming needed. It will create an output for using at the Secret Santa Generator https://arcanis.github.io/secretsanta/. On the Send Email page there is a macro that will open draft emails to all of the recipients that says (Name's) Seret Santa for (This year) and will include the link to their secret Santa as well as a custom message you can add in the text box. You can use the spreadsheet without macros, it just means the extra text and the autogeneration of emails won't work.

📋 What to Do:
Fill in the Category column if you want to filter by different groups
Fill in the Buyer column
Fill in the Permanent Exclusions columns on the right of the table.
Set the year in the Secret Santa sheet (named cell This_Year)
Define how many years to look back (to avoid having repeats each year)Write "Yes" or "No" in the column for this year to denote if they are participating this year and that will populate the Send Email sheet and exclude people's names from the output if they are not participating this year
Use the text box on the Send Email Sheet to include any group-wide notes, rules, or gift ideas.Copy the Output column from the Send Email sheet and paste it into the secrete Santa generator which has an option for a text input instead of writing each name.Generate the links and then copy and paste the links into the Link column in the Send Email sheet.  ⚠️WARNING: THE OUTPUT FROM THE WEBSITE IS IN ALPHABETICAL ORDER
Run the macro:
Press Alt + F8, select EmailLinks, and click Run
Or use the “Send Emails” button
✅ Your default email app (e.g., Outlook, eM Client, Apple Mail) will open a pre-filled draft for each recipient.
Subject line and body are automatically personalized.
In my email client, I need to press return after the link to turn it into a hyperlinkNext Year, follow the links to find out who each person had and replace the "Yes" with the name of who they had. delete all of the "No" text so that it doesn't show up in the output. just leave that cell blank.

⚠️ Notes:
Requires Excel 365 on macOS or Windows
Works with any email client that supports mailto: links
Long messages may be truncated — keep it short for best results
Your default email account will be used

Conditional Formatting:
row will be highlighted green if "yes" in current year
row will be highlighted grey if "no" in current year
cell will be highlighted yellow if the name doesn't appear in the list of names
cell will be highlighted red if the name appears twice for that year
🔒 No email is sent automatically — you can review/edit drafts before sending.


