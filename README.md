README - BadAss Library project documentation.<br>
1/11/24.    wmk.
###Modification History.
<pre><code>8/23/22.    wmk.   original document.
10/23/22.	wmk.    Project Overview section documentation.
11/5/22.    wmk.    Project Overview section updated.
1/11/24.    wmk.    Updated for Accounting/BadAss *git project.
</code></pre>
<h3 id="IX">Document Sections.</h3>
<pre><code><a href="#1.0">link</a> 1.0 Project Overview - overall project description.
<a href="#2.0">link</a> 2.0 Project Files - files comprising the project.
<a href="#3.0">link</a> 3.0 Build Process - step-by-step library build instructions.
</code></pre>
<h3 id="1.0">1.0 Project Overview.</h3>
The BadAss project contains all files that Calc or Excel need in order
to build the BadAss library for either app. The primary files required
for building the library are in the project folder. The individual macros'
source code is in the subfolders /Module1 and /Module2, and /Dialogs.

The /BadAss folder within the project is the current BadAss Library exported from
Calc. For now, the "build" process does not touch that folder. Any build of the
BadAss Library from the Geany project builds the resultant Module1 or Module2
.xba file in the src/Releae folder folder. This provides a level of separation
where any new build is separated from the existing library code.

In other words, the /BadAss folder is used for "export" operations from Calc
back into the \*git project. The folder /BadAss/Release is used for "import"
operations to Calc for testing/integrating new code. Once the code has been
tested, it should then be exported to the /BadAss parent folder completing the
cycle.

The EditBas geany project is used for extracting, modifying, and re-integrating
macros within the BadAss Calc library. It uses its own shells in addition to
those defined in the Accounting/BadAss/Procs-Dev folder. The EditBas project is
managed from the Accounting/BadAss/Projects-Geany folder.

The library .xba files that contain all of the formalized macro definition
code may be built from the source code files by using the *make* utility.
The makefile for the build is *MakeBadAssLibrary.tmp* which is modified
by a DoSed.sh shell prior to executing the build.<br>
**Caution:** The library .xba or .dlg files may be overwritten by exporting
the running library from the Calc Macros/Organizer. This will make the "build"
source .bas files out-of-date with the running library. Whenever this is
done it will be necessary to re-sync the .bas source files to the code that
was exported into the .xba files for the build process to correctly regenerate
the library code.<br><a href="#IX">Index</a>
<h3 id="2.0">2.0 Project Files.</h3>
 The files fall into 4 categories:
1) macro source code, 2) dialog source files, 3) library definition files, 
4) library Build files.

<b>Macro Source Code.</b>
The macro source code for the library is spread across multiple Basic source code
files having the *.bas* file extension. This allows for easy maintenance of each
macro without having to edit in either the Calc or Excel app macro interface. The
*.bas* files are Basic code files without any XML markup (as needed by a "loaded"
library in the spreadsheet apps). Each module <module-name> has its source code
stored in its own BadAss/< module-name > folder.

The actual library source code that gets installed into the apps with the >macros>Organizer
extension is in "module" files that are collections of *.bas* files organized into
single files that are in XML Basic markup language. These files follow the naming
convention *< module-name >.xba*. They have an XML markup header preceding the Basic
code, and an XML markup footer following the Basic code. The Basic code within the
*.xba* file is recognizable, but has special characters "marked up" with XML coding
(e.g. the apostrophe (') character is marked up as (&)apos;). A shell utility is
provided that converts between the *.xba* file format and the *.bas* file format
for ease of editing.

The files XBAHeader.txt and XBAFooter.txt are stored with the project to provide
the Build process with the appropriate XML code wrapper for the *.xba* file that
the *make* utility creates for the BadAss library. *< module-name >* is edited into
the XBAHeader.txt file XML so the resultant .xba is properly tagged.

<b>Dialog Source Files.</b>
The dialog source files are stored in folder BadAss/Dialogs. All dialog source files
have the *.xdl* filename extension. There is no *frame* interface for interactively
editing dialogs outside of Calc. The dialog source files stored in BadAss/Dialogs
are images of the \*WINGIT_PATH/Libraries-Project/BadAss/\*.xdl files. All editing
of the *.xdl* files is done interactively within the Calc/Edit Macros facility.

The *.xdl* files are XML DOCTYPE dlg:Window markup files. It is possible to make
manual changes to the .xdl files within the BadAssLibrary/Dialogs folder using
an editor. The CopyToGit.sh shell file may be used to copy the ./Dialogs folder
to the \*WINGIT_PATH/Libraries-Project/BadAss \*.xdl files. All bets are off
if this is done, as the editing is much easier performed in the graphic environment
of the Calc/Edit Macros tools.

<b>Library Definition Files.</b>
The library definition files are stored in the BadAss project folder. These files have
the filename extensions *.xlb* and *.xlc*.<br><a href="#IX">Index</a>

<b>Library Build Files.</b>
The library build files are stored in the BadAss project folder. These files are
of multiple types with filenames <i>Make\*</i>, \*.sh, \*.txt and are used by the utility
apps *make*, *sed*, and *awk*.<br><a href="#IX">Index</a>

<h3 id="3.0">3.0 Build Process.</h3>
<br><a href="#IX">Index</a>

