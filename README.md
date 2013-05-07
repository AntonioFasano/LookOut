LookOut
=======

Get Outlook Freedom


Motivation
----------

I am a seasoned user of Microsoft Outlook, but I don’t like it. Unfortunately I have been stuffing contacts in it for about 15 years. Migrating contacts to a different platform is not simple since they are stored in a very deep folder tree. Contact folders, and particularly nested contact folders, are a feature rather specific of Outlook. Categorizing contacts in some way is crucial to me, but every contact system has a specific classification approach (folders, categories, group, tags etc.) making it difficult to share contact data and preserving the original structure. Not to mention that, while classifying mails is always possible, many applications still have no facilities alike for contacts. Notably Thunderbird has no labelling for contacts and your choices are limited to the address book.

Recently I discovered a Thunderbird extension adding categories to contacts and including the possibility to search contacts by category. MoreFunctionsForAddressBook, http://nic-nac-project.de/~kaosmos/morecols-en.html

The time had come for the big switch.

Outlook limitations
-------------------


Outlook is probably one of the most powerful PIM system and these macros are partly a consequence and an evidence of that. Outlook is also very mouse oriented. In a typical session you move back and forth from the navigation pane to the content area. If the navigation tree is very simple, few nodes and a low nesting level, then you can stand the mouse journeys, when not a keyboard addict; with a very deep structure, moving from a contact folder to a mail folder and then to the calendar etc. is very time consuming.

Keyboard shortcuts could help, but their design is rather weird. With Ctrl-6 you visualize the contact folders in the navigation pane (yes, CTRL-6), then you’d like to scroll the tree with the arrows but, surprise, the focus is not on the folder. Even if you help yourself with the mouse and select an upper node, the focus is immediately moved away (very unnaturally). To keep the selection you have to double click, but not too fast or you will activate the rename feature.

Real time search could really help to get to the item you need. Here again Outlook has plenty of powerful options to search in every possible way, but you have to click way t0o icons and dialogs, so it’s not really real time. No comparison with Gmail, where you write /Chess and you immediately get all your contacts playing your favourite game.

My humble opinion is that Outlook is a very mature product. When it was introduced, writing an email and searching for the recipient’s details was something quite occasional in the day of a computer user, therefore the responsiveness of the interface hardly affected the overall user experience. Nowadays, since we process electronically an uncountable number of social tasks per day, the total valuable time spent in searching the proper folder with the proper item can entail a huge impact on productivity and on the level of stress at the end of the day.

Summing up everything you are thinking of doing will be done in Outlook, and much more, but everything will be done very slowly, and too slow for the present competition pace.

What do I need to know for going further?
-----------------------------------------

Of course you are supposed to have a basic understanding of Outlook, but you don’t need to have a specific knowledge of Outlook macros or the language with whom they are written, Visual Basic for Applications (VBA). Anyway this is not a tutorial on Outlook macros. If you ignore completely this subject, you might find difficult to follow these lines and prefer to read some introductory material first.

Bear in mind that the code, the menus and keyboard shortcuts mentioned below refer to Outlook 2010. Slight, but subtle, differences may apply to your version. If someone will give it, I will publish his/her feedback here on this regard.

Usual disclaimers apply, that meaning you had better to back up your valuable Outlook data before running these macros as errors may hide in the code causing unwanted side effects.

Given this, if you are already used to VBA, just add a VBA module, probably named “catman”, and do skip this section; if you are not, try to follow the subsequent steps.

First and foremost you have to enable macros. By default, for security reasons, macros are disabled. To enable macros, assuming they are not such, go to:

File → Options → Trust Centre → Trust Centre Settings… → Macro Settings

Choose “Notification for all macros”. Why?

With this option you are notified, just once per session, of the presence of potentially risky macros. Enabling all macros without notifications adds a useless tier of risk and, of course, enabling only digitally signed macros would require a special certification procedure for macros you are going to write.

Now type ALT-F11 and you will enter the VBA IDE. If not visible already, show the “Immediate Window” typing CTRL-G and show the “Project Explorer” typing CTRL-R (or use the “View” menu). In the “Project Explorer” expand the node “Modules”. Go to menu Insert → Module (there is also a button for this, but it may not insert a module item, if not set so). You will see a new module in “Modules” node, possibly named like “Module 1”. Name it something more meaningful, such as “catman”. To do this, select the module name and type F4. A properties bar will appear, where you have the possibility to change the name field of the module as you like.

At this point make your IDE yours. You can undock, resize and close every bar of the IDE, much like every Windows window. Note that, in the following you will need only the central code window (resizable, but not closable) and the “Immediate Window”. The latter is used to insert commands on behalf of the user (in our case to run the macros) or to print information on behalf of the macros themselves.
Please note that after typing subroutines/functions in the code window, they will appear in the left drop down menu as an item in alphabetical order, so you will be able to easily browse them.
