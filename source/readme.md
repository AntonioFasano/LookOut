LookOut   {#mainpage}
=========

LookOut is a Microsoft Outlook class macro intended to manage contacts.
The main objective is to move contacts  organisation from a folder structure to a
 category structure. The former is peculiar of Outlook, so you are locked-in.
 Hopefully, by using these macros, you will be able to move your contacts data to in other applications,  not based on a folder classification (e.g. Google contacts or Thunderbird) without losing the structure, hence the name. 


Among the duties performed by LookOut there are: the conversion of contact folders in categories; the export of entire folder trees, the detection and management of duplicate contacts; the inclusion of a category field when exporting as VCF. 

LookOut is  intended for Outlook 2010. Chances there are it will work in other versions.
You don't need to be knowledgeable with VBA to use LookOut.


Disclaimer
----------

Despite this software  helped me dramatically in my contact system migration, it  is still at a beta stage: I had no time to test it thoroughly and writing sound documentation. 

As the  macros  described here affect  Outlook data files, I strongly encourage to backup them  before playing with LookOut. To this end I  suggest that you identify the Outlook data file(s) containing critical data, close Outlook and make a copy of the files. If something goes wrong, close Outlook again and replace the  data files with the back-upped copy.

If you want to learn more about Outlook data file read ahead the section "Outlook folder structure".




Motivation
-----------

I am a seasoned user of Microsoft Outlook, but I don't like it. Unfortunately I have been stuffing contacts in it for about 15 years. Migrating contacts to a different platform is not simple since they are stored in a very deep folder tree. Contact folders, and particularly nested contact folders, are a feature rather specific of Outlook. Categorizing contacts in some way is crucial to me, but every contact system has a specific classification approach (folders, categories, group, tags etc.) making it difficult to share contact data and preserving the original structure. Not to mention that, while classifying mails is always possible, many applications still have no facilities alike for contacts. Notably Thunderbird has no labelling for contacts and your choices are limited to the address book.

Recently I discovered a Thunderbird extension adding categories to contacts and including the possibility to search contacts by category. MoreFunctionsForAddressBook, [MoreFunctionsForAddressBook](http://nic-nac-project.de/~kaosmos/morecols-en.html)

The time had come for the big switch.




Outlook limitations
-------------------

Outlook is probably one of the most powerful PIM system and these macros are partly a consequence and an evidence of that. Outlook is also very mouse oriented. In a typical session you move back and forth from the navigation pane to the content area. If the navigation tree is very simple, few nodes and a low nesting level, then you can stand the mouse journeys, when not a keyboard addict; with a very deep structure, moving from a contact folder to a mail folder and then to the calendar etc. is very time consuming.

Keyboard shortcuts could help, but their design is rather weird. With Ctrl-6 you visualize the contact folders in the navigation pane (yes, CTRL-6), then you'd like to scroll the tree with the arrows but, surprise, the focus is not on the folder. Even if you help yourself with the mouse and select an upper node, the focus is immediately moved away (very unnaturally). To keep the selection you have to double click, but not too fast or you will activate the rename feature.

Real time search could really help to get to the item you need. Here again Outlook has plenty of powerful options to search in every possible way, but you have to click way t0o icons and dialogs, so it's not really real time. No comparison with Gmail, where you write /Chess and you immediately get all your contacts playing your favourite game.

My humble opinion is that Outlook is a very mature product. When it was introduced, writing an email and searching for the recipient's details was something quite occasional in the day of a computer user, therefore the responsiveness of the interface hardly affected the overall user experience. Nowadays, since we process electronically an uncountable number of social tasks per day, the total valuable time spent in searching the proper folder with the proper item can entail a huge impact on productivity and on the level of stress at the end of the day.

Summing up everything you are thinking of doing will be done in Outlook, and much more, but everything will be done very slowly, and too slow for the present competition pace.


Installing LookOut
------------------

### If you are  knowledgeable with VBA

Make sure you that security level are properly set in `Trust Centre Settings` then in  Outlook VBA IDE import the file `LookOut.cls` and set a reference to  `Microsoft Scripting Runtime` and `Windows Script Host Object Model` (the sub `ExportVcf` need them, other procedures don't).

If you like to make it more manually,  you can create yourself a new class module, named `LookOut` and copy the "VBA content" of  `LookOut.cls`. This means excluding the header lines containing the class module proprieties, i.e. version attributes etc.


### If you are not  knowledgeable with VBA


Of course you are supposed to have a basic understanding of Outlook, but you don't need to have a specific knowledge of Outlook macros or the language with whom they are written, Visual Basic for Applications (VBA). Anyway this is not a tutorial on Outlook macros. If you ignore completely this subject, you might find difficult to follow these lines and prefer to read some introductory material first.

Bear in mind that the code, the menus and keyboard shortcuts mentioned below refer to Outlook 2010. Slight, but subtle, differences may apply to your version. If someone will give it, I will publish his/her feedback here on this regard.

Usual disclaimers apply, that meaning you had better to back up your valuable Outlook data before running these macros as errors may hide in the code causing unwanted side effects. 

Let's start the setup.


First and foremost you have to enable macros. By default, for security reasons, macros are disabled. To enable macros, assuming they are not such, go to:

    File ->  Options  ->  Trust Centre  ->   Trust Centre Settings...  ->   Macro Settings

Choose `Notification for all macros`. Why?

With this option you are notified, just once per session, of the presence of potentially risky macros. Enabling all macros without notifications adds a useless tier of risk and, of course, enabling only digitally signed macros would require a special certification procedure for macros you are going to write.


Identify the  file `LookOut.cls`. It is on the same page/site/package from which you are reading this document.

Now type ALT-F11 and you will enter the VBA IDE. If not visible already, show the `Project Explorer` typing CTRL-R and the  `Immediate Window` typing  CTRL-G (or selecting them in the `View` menu).  Therefore go to menu `File -> Import File...`; a dialog windows opens, use it to import the file `LookOut.cls`. In the `Project Explorer`, you will see a new element in the `Class Modules`, named `LookOut`. 

In order to export the contacts (in vCard format) you need to set the so called VBA _references_. In the VBA IDE select `Tools->References`. When the dialog opens, find and check the entries `Microsoft Scripting Runtime` and `Windows Script Host Object Model` and click `OK`. This is necessary to make the procedure `ExportVcf` work; indeed all other procedures don't need these references. 

You are done with setup.

Anyway, note that you can   undock, resize and close every component of the VBA IDE, much like every Windows window. As you will run macros through `Immediate Window` and read  here information printed by the macros themselves, you might want adjust it a little. 


<!--  
Go to menu `Insert -> Class Module` (there is also a button for this, but it may not insert a module item, if not set so). In the `Project Explorer`, you will see a new element in the `Class Modules' node, possibly named like ``Module 1''. Name it something more meaningful, such as ``catman''. To do this, select the module name and type F4. A properties bar will appear, where you have the possibility to change the name field of the module as you like.

At this point make your IDE yours. First, if not visible, show the `Immediate Window` via CTRL-G and show the `Project Explorer` typing CTRL-R (or use the `View` menu).
You can undock, resize and close every bar of the IDE, much like every Windows window. Note that, in the following you will need only the central code window (resizable, but not closable) and the `Immediate Window`. The latter is used to insert commands on behalf of the user (in our case to run the macros) or to print information on behalf of the macros themselves.
Please note that after typing subroutines/functions in the code window, they will appear in the left drop down menu as an item in alphabetical order, so you will be able to easily browse them.

 -->

Outlook folder structure
------------------------

LookOut deals with Outlook folder structure. To take advantage of it you need to grasp how contacts  are organised into contact folders. If you already know this, jump to the next section, else read ahead.


Every Outlook item, emails, notes, contacts, etc. is stored in _Outlook folders_, much like Windows files are. Outlook folders can be nested and they are stored in an **Outlook Store**. Outlook has a _navigation pane_ with a _folder tree_, where you can browse the stores, their folders/subfolders and the included items. Every store corresponds to a _top level nodes_ in the navigation pane; of course many users may have just one store to fit their needs. On the file system point of view, each store corresponds to a file with the extension `.PST`, defined **Outlook Data File** or sometimes _Personal Folder file_.

Out of the box Outlook 2010 comes with the (single) store named `Outlook Data File` associated with the file `%USERPROFILE%\Documents\Outlook Files\Outlook.pst`. For many users this is all they need other may want to create new stores segregate objects with different scopes, e.g., to play with the following code, you might want to create a new store for safeguarding your data. Selecting  `Home->New Items->More Items->Outlook Data File...` you can create a new `.PST` files with the associated store, where you can create, copy or move objects.


For macro purposes,  folders in an  Outlook store are denoted by a path reflecting their position in the navigation pane, similar  to an ordinary file system path, that is:

    "\\My Store\folder\subfolder\..."

Bear in mind that a store name may be different from the name of the related Outlook Data File. To learn about the file system path of a store, select it in the navigation tree, right click on `Data File Properties` and click `Advanced`. In the following dialog, the first textbox, labelled `Name`, shows the name of the data store and it is editable, that is you can change the name; the second textbox, labelled `Filename` (not editable), shows the full path of the Data File related to this store.

Outlook has also a notion of **default folder** for given types of items, e.g. **default folders for contacts**. This is nothing special: for most tasks, if no specific folder is given, the default one for the given items is used. There are *global default folders* and _per store default folders_. If you don't make changes, the original folders that come with the initial store after the product setup are also the default folders. Later this article will show how to use the `GetFolder` function to print the path of default contact folders.

When the macros presented here will ask you to input a contact folder path, often you will be allowed to give as an alternative this parameter:

    DefaultFolder:=True

As you have guessed, in this way the global contact folder will be used.

Another option is to give the folder path like this:

    "\\My Store"

In this case the default contact folder of the store named `My Store` will be used.

Note that paths are not case sensitive.


Using LookOut macros
--------------------

If you want you might associate VBA macros with fancy buttons and menus, but this is beyond the scope of this documentation. Here we will issue LookOut commands by typing them  in the `Immediate Window` of the VBA IDE.

A typical session consists of typing lines like these:

    Set ol = new LookOut

    ol.Commnad1 ...
    ol.Commnad2 ...
     
    ...
     
    Set ol = Nothing


`Set OL = new LookOut` initializes the program (that is the class `LookOut`). We then issue the commands to accomplish specific tasks and eventually we do some memory cleanup with  `Set OL = Nothing`.

Normally you need the initialisation only once per session, that is when you open Outlook and start using LookOut. Some special events, such when you click on the reset button in VBA IDE, delete the object `ol`. Therefore, should  you get an `object required` error, simply issue `Set ol = new LookOut` again.



### Get Default Folders 

As many commands accept as input the default contact folder, you may want to learn  its path with respect to Outlook tree nodes.

In order to print the path of the global default contact folder, type the command:

	  ? ol.GlobalDefaultPath

You may substitute `?` with `Debug.Print`.
To print the path of the default contact folder of the store named `My Store`, type the command:

	 ? ol.StoreDefaultPath("\\My Store")

You get an error if no store with this name exists or you miss `\\`.



### Convert folder names in category names

The main motivation to for LookOut is to transform Outlook folder structure for contacts in an equivalent category structure. Categories (or labels or groups in other contact management systems) can be exported and easily understood by other applications. To accomplish this `Folder2Cat` attaches category names to contacts on the basis of the names of folders containing them.

If a folder is nested, the category names include also the chain of parent folders, but the name of the root contact folder will not be added as a category. Indeed every contact is in the root, so this category won't be very meaningful. So if a contact belongs to the Outlook folder `\\My Store\Contacts\foo\bar\etc\`, the three categories `foo, bar, etc` will be added to every contact in this folder and in this order (not the alphabetical one), but not `Contacts`.

Previous category names are preserved unless they duplicate new ones. So, if some contact in  `\\\\My Store\Contacts\foo\bar\etc\`  has already the categories `bar, baz`, final categorisation will be `baz, foo, bar, etc`. Similarly, if the folder path contains duplicate names, only the first occurrence will be added as a category name. Therefore, if a contact belongs to the folder `\\My Store\Contacts\foo\bar\foo\`, the two categories `foo, bar` will be added to its contacts (without name repetitions). Category comparison is not case sensitive.
 
You need to provide to `Folder2Cat` a folder path and the  categorisation will be applied to all contacts belonging to the passed folder and to those in the possible nested folders. Therefore issuing: 


	ol.Folder2Cat "\\My Store\Contacts\foo\bar\"

categorisation will be applied to the folder `\\My Store\Contacts\foo\bar\` and all subfolders, if any.


If you want to use the global default contact folder (which supposedly contains all the contacts in your address book), issue:

	ol.Folder2Cat DefaultFolder:=True
	
And, if you want to use the default contact folder of `\\My Store`:

	ol.Folder2Cat "\\My Store"

Instead of setting a single category for each nested folder, it is possible to concatenate them via a dot,  e.g. `cat.subcat`. Setting the `GoogleStyle:=True` will achieve this:

	ol.Folder2Cat "\\My Store\Contacts\foo\bar\", GoogleStyle:=True


### Add and Remove categories



If you want to add the category `NewCat` to the contacts in  the  folder `\\My Store\folder\to\categorize`, issue the command:

    ol.AddCat `\\My Store\folder\to\process",  "NewCat"

`AddCat` will not affect contacts in possible subfolders of `NewCat`. Besides existing category names are preserved unless they duplicate they the new one. Since the conflicting old category name will be deleted and the new one will be appended, when there are two or more category names you will just see a permutation of names, that is the old conflicting name will appear as the last category name. Given this, you will see no effect when there is just one category name like the new added one.


If you want to remove the  category `Cat2Del`  to all contacts in a the folder `\\My Store\folder\to\process`, issue the command:

    ol.DelCat "\\My Store\folder\to\process", "Cat2Del"

The category name `Cat2Del` will be removed from all contacts in folder `\\My Store\folder\to\process`. `DelCat`  will not affect contacts in possible subfolders of `\\My Store\folder\to\process`. 

If the category asked for removing is not attached to some contacts, no warning or error will be generated and, of course, they will be unaffected by `Cat2Del`.


After deleting a single category from a single folder, we now turn to delete all categories from a whole contact folder tree, that is from a given folder and all its (possible)  subfolders of the tree root folder. 


    ol.DelAllCats "\\My Store\folder\to\process"


### Contacts with empty names 

Outlook allows to store contacts without filling name fields. This can cause a problems
when exporting contact data.

If you want to learn about possible empty name contacts in `\\My Store\folder\to\check`, use:

    ol.ShowEmpty "\\My Store\folder\to\check"


You can fix empty names, by filling Outlook `LastName` field with `FileAs` field (always present). Issue:

    ol.FillEmpty "\\My Store\folder\to\check"

As usual, to check/fix the global default folder, issue: 

    ol.ShowEmpty DefaultFolder:=True

and

    ol.FillEmpty DefaultFolder:=True


### Flatten folder

Once we map folders to categories, with `Folder2Cat`, there is no need to to keep contacts in separate folders. We can can put flatten them into a single folder; so the process of  transforming the folder strutcture into  the category structure is completed.

To flatten contacts into the tree whose parent folder is `\\My Store\Tree\Parent` into the single folder `\\My Store\path\to\FlatFolder` issue:


	ol.FlattenFolderCopy "\\My Store\path\to\TreeParent", "\\My Store\path\to\FlatFolder"
	


In case you want to copy and delete contacts in `\\My Store\Tree\Parent`, issue:

	ol.FlattenFolderMove  "\My Store\path\to\TreeParent", "\\My Store\path\to\FlatFolder"

### Find Duplicates

When using  Outlook folder structure you can have the same contacts stored in different folders. This is not very advisable  when all contacts are in the same folder under different categories.

`FindWDup` helps you identify duplicate contacts in a given folder tree.

	ol.FindWDup  "\\My Store\folder\to\process" 

will find duplicates in the same folder or in different subfolders of "\\My Store\folder\to\process".
The function will print the  `FileAs` field and the path of duplicates in the immediate window.

`FindWDup`  considers "weak duplicates" as contacts with the same having the same `LastNameAndFirstName` or  `FileAs` field. Obviously other contact fields might not be  the same, so  these are simply  referring to the same person.



`CopyUnique` copies only unique contacts from a given folder to another folder. 

	ol.CopyUnique  "\\My Store\folder\to\source" "\\My Store\folder\to\dest"

will copy non-duplicated (unique) contacts from `\\My Store\folder\to\source` folder `\\My Store\folder\to\dest` folder. Contacts in  subfolders of `\\My Store\folder\to\source`, if any,  are not included. 


`CopyUnique` considers duplicates as contacts with the same having the same `LastNameAndFirstName` or  `FileAs` field and modification time. 

It is reasonable to consider these kind of contacts total duplicate. The opposite is not true, as there can be duplicated contacts with all fields equals except modification time.



### The final cut

After categorising folders we are able to export contacts without losing their grouping structure.
We will export in **vCard** format, each exported contact will be a file with a  `.vcf`  extension. This format is compatible with Thunderbird and Gmail.

As was observed in the installation section,  the actual export procedure, `ExportVcf`,  needs references to `Microsoft Scripting Runtime` and `Windows Script Host Object Model`. If you missed this step, it is now a necessary  one. 


When exporting,  `ExportVcf` will add  Outlook contact categories (created by `Folder2Cat` and/or manually by the user) to the vCards. So, if you import the vCards, for example in Thunderbird equipped with [MoreFunctionsForAddressBook](http://nic-nac-project.de/~kaosmos/morecols-en.html) add-in, you will find the same Outlook categories.

The command: 

    ol.ExportVcf "export\path", "\\store\folder\to\export"

will export as `.vcf`-files all contact contained in Outlook folder `\\store\folder\to\export` and its subfolders to the file system path `export\path`.

As usual: 

    ol.ExportVcf  "export\path", DefaultFolder:=True

will export from the global default contact folder.


The export folder path can have environment  variables, e.g. `%%USERPROFILE%\Desktop\ExportVcf`.

The name of  the `.vcf`-files is given by the  `FileAs` field of the contacts. 
Note that, when copying contacts to `export\path`, the folder structure is not kept. Therefore, if the same contact is stored in different folders, you endup with duplicate files.

If a duplicate contact is found, the procedure stops exporting further contacts and skips to the next folder in the queue, if any.


`ExportVcf` produces a  vCard for each contact. Some applications, such as Gmail, require a single vCards to be combined in a single multi-contact vCard.

To combine the vCards in `export\path` into a single multi-contact vCard, issue:

    ol.MultiVcf  "export\path", "path\to\MultivCard.vcf"

where `path\to\MultivCard.vcf` is the pathname of the multi-contact vCard to be created.


`MultiVcf` does not start  combining  vCards, is any of them  results already to be a multi-contact vCards. Possible  non-vCard files the source folder are ignored.



<!--  LocalWords:  Google Thunderbird VBA IDE LookOut cls subfolders vCard vcf vCards
 -->
