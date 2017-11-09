# MsWordVersioning
MS Word script for (basic) version control of documents.

## What does it do?

* Allows creating new main or sub versions (ie 1.5 -> 2.0 or 1.5 -> 1.6) of a document. This includes:
  * Exporting a pdf and a comment-only Word document, for example for sharing or archiving.
  * Adding a line to the change log for the document.
* Automatically keeping track of version number, and allows inserting this as a field in your document.

### Other features

* Set the folder name, to where exported documents are saved. (Defaults to document name.)
* Set the bookmark used for inserting change log for the document.
* Tries to use the same language as the Ms Word interface. Currently US English (fallback) and Swedish supported.

## Why document versioning?

* You write a document in several steps, and want to keep track of each version of the document.
* You mistaningly believe that you write a document once-and-for-all, but later on realize that you need to update it.
* You want to circulate a document for comments.
* You are tired of multiple versions of the same document and comparing last save date to see which is the most recent.
* You are tired of trying to merge together different versions of a document, that have been edited independently. Possibly in several steps.
* You are tired of realizing too late that there was *another* version of the same document, with suggested changes that you have not taken into account.
* You are tired of document names with prefix *FINAL final*.

## Why not document versioning?

* You have really limited disk space.

## How to install

1. Copy the script into Word. This is probably easiest done by (a) open the script file here on GitHub, press the *raw* button and copy all the content, (b) pressing alt + F8 in Word to open the macro dialogue and then pressing edit or create, followed by (c) pasting all the code, saving and closing the code environment.
2. Allow running macros in Word. This is done by going to file -> options -> security center -> setting for security center and allowing macros. **Note that this also means that other macros can run in macro-enabled documents. Don't open macro-enabled documents you don't trust.**
3. To be able to use the script you also need to save your document(s) as macro-enabled documents, or macro-enabled document templates.

To use the script more smoothly, it is also recommended that you add some buttons to the toolbar.

4. Right-click on the toolbar and select customize.
5. Add the macros to the desired place in the toolbar. It's easier to find the macros if you limit the list of actions to only macros.
6. Rename or give new icons to the new actions, if you like.

## How to use

* newSubVersion: This prompts you for a message for the change log, whereafter the version numer is increased a minor step (ie 1.5 -> 1.6). Timestamp and change log message is added to the change log, then a pdf and a locked Word document is exported.
* newMainVersion: Like the previous macro, but increases version number a major step (ie 1.5 -> 2.0).
* insertVersionNumber: Inserts a field at the cursor showing the version number of the document.
* setVersionLog: Marks the selected section as the change log (by inserting a bookmark), which means that new log messages will be added there. Works better if a section is selected, than if nothing is selected.
* changeExportFolder: Prompts you for a name for the folder to use for exports. No folder name (default) means that the folder will have the same name as the document.
* about: Displays a link leading to the project page on GitHub (here).

## Extra information for tech people

* Version number is saved as a document property, which can be edited through the relevant settings in Word.
* The change log is found using a bookmark, which can be changed manually through the relevant settings in Word.
* The interface language is detected through the application setting. Additional translations can be easily added -- look at the end of the script file.
