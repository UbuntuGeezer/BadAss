README.md - EditBas FLsara86777lib project documentation.<br>
11/12/23.	wmk.
###Modification History.
<pre><code>6/24/23.    wmk.   original document for FLsara86777.
11/12/23.   wmk.    documentation change for FLsara86777lib
Legacy mods.
3/7/22.     wmk.   original document.
3/8/22.     wmk.   Supporting Files section added.
</code></pre>
<h3 id="IX">Document Sections.</h3>
<pre><code><a href="#1.0">link</a> 1.0 Project Description - overall project description.
<a href="#2.0">link</a> 2.0 Dependencies - dependencies for EditBas operations.
<a href="#3.0">link</a> 3.0 Setup - step-by-step instructions for various builds done by project.
<a href="#4.0">link</a> 4.0 Supporting Files - files integrated with project processes.
</code></pre>
<h3 id="1.0">1.0 Project Description.</h3>
The EditBas project provides tools for editing an integrating .bas code
modules into the FLsara86777 Calc library. The library source code is maintained
in the folder set TerrCode86777/FLsara86777lib. A utility uses *awk* to extract
modules from the current library. It uses *sed* to replace modules within
the Territories library. It prompts the user to use *Calc* to update the
Territories library in the *GitHub/Libraries Project* folder. It uses
*make* to  perform the above operations.

Prior to working with the EditBas project, the \*LibBuild project was
been used to maintain the library folders. See the [LibBuild](file:///media/ubuntu/Windows/GitHub/Libraries-Project/FLsara86777/src/Projects-Geany/LibBuild/README.html) documentation
for the Version.2.0.x and prior versions.

Under Version.3.0.0 and newer versions, all library maintenance for the
FLsara86777 library macros is performed under the TerrCode86777/FLsara86777lib
folders. See the
[FLsara86777lib](file:///media/ubuntu/Windows/GitHub/TerrCode86777/FLsara86777lib/README.html) documentation for details of all the
folders used in maintaining Version.3.0.0 and newer versions. The configuration
file for the Territories system contains two library base pointers:
<pre><code>
libbase=*WINGIT_PATH/TerrCode86777/FLsara86777lib
libbase2=*WINGIT_PATH/TerrCode86777/Territorieslib/FLsara86777
</code></pre>
\*libbase is the base path of this (FLsara86777lib) repository. These
environment variables are present in all versions Version.3.0.9 and newer.

EditBas depends upon minimal special formatting of each subroutine/function
within the Calc library. Each module (subroutine/function) should be preceded
by a line containing *'// modulename.bas*. Each module should
be followed by a line containing *'/\*\*/*. These lines delimit the subroutine/function
to allow extraction, replacement or deletion of the module.

The project flow follows these logical steps:
<pre><code>    Use Calc to Export the current Territories library to GitHub/FLsara86777

    1. run *GetXBAModule.sh* shell from project to make a working copy of the
     current code module from GitHub/Territories in the /Import folder.
    
    2. run *CopyXBAoverBAS.sh* shell from project to copy the current code module
     from the /Import folder to the Basic/< xbamodule > folder. (LibBuild has
     already initialized < xbamodule >Bas.txt to be the list of .bas blocks
     within < xbamodule >.bas)
     
    3. *make* MakeExtractBas to extract the subroutine/function from the local
     copy of the .xba file; the extracted code will be in < procname >.bas
     in the Basic/< xbamodule > folder (< procname > is the sub/function name).
    
    4. edit the source code in < procname >.bas to make the desired changes
      OR
     *make* MakeDeleteXBAbas to remove the < procname > from the local
      .xba file. **Note:** If you delete a < procname > from the local
      .xba and want to note this as a permanent change, it would be wise
      to make a notation in the /Basic/<procname>.bas that this code
      has been removed from the running library.
      OR
     create a new < procname >.bas with a new macro definition
      
    when your editing is complete:
    5. move on to the next .bas to modify
      OR (with caution)
     *make* MakeReplaceBas to replace the old < procname > code with the
      edited < procname > code.
	  OR
	   DoSedAdd.sh to set up to add the new < procname >.bas into the library
	   *make* MakeAddBas to add the new < procname > to the < xbamodule >Bas.txt
	   list of .bas blocks to build into the library.
	   
     repeat steps 2 - 5 for all <procname>,s you desire to modify.
</code>/<pre>    
To test all the new code added above, the \*Lib library in the running
Calc system will need to be replaced. This involves several steps:
<pre><code>    1. Replace the existing .xba for the module in the GitHub/Libraries-Project/Territories
  folder. Use the *PutXBAModule.sh* project shell to accomplish this. This shell
  automatically cycles the previous .xba module to a file named *old<modulename>.xba*
  where <modulename> is the original module filename (e.g. Module1).

    2.Use Calc to reload the Territories library from GitHub into the running
     system. The Organizer has an Import function to do this; be sure to
     check the "Replace existing libraries" checkbox to ensure that the
     GitHub version replaces the current running Territories version.
    
    3. Test your changed <procnames> with the appropriate Territory activities
     that will exercise the new code.
    
    If you are satisfied with all of your changes, update the /Basic folder
    with all of the changed .bas files:
    
    4. run UpdateBasicFiles.sh to update all of the .bas files in /Basic that
     have been changed in the project folder. This will also remove the .bas
     files from the project folder, allowing fresh extracts to obtain more
     <procfile> source without the clutter of old code.
     
     If you are not satisfied with all of your changes... you're on your own.. 
</code></pre>
<br><a href="#IX">Index</a>
<h3 id="2.0">2.0 Dependencies.</h3>
When working on a Calc library, use the \*chglib aliased command to set up the
environment variables for library management. For FLsara86777 the 
\*Lib environment var should be set to FLsara86777.

Each .xba module has an XML header identifying the file as an XML BASIC
file. The EditBas project uses the GetModuleXMAHdr shell to extract this XML
header into the file < module >XbaHdr.bas within the module's folder. The .xba
modules all have a < /script:module > termination line at the end of the file
which is removed when the .xba module is disassembled, and restored when the
.xba is re-assembled into the /Release folder.

.bas blocks are defined within an .xba module by starting and ending with block
delimiter lines. Each block within the macro source code begins with the line:
<pre><code>
	'// < block-name >.bas
	 where the < block-name > is the macro name that appears within the
	 Calc macro organizer
</code></pre>
Each block within the macro source code ends with the line:
<pre><code>
	'/**/
</code></pre>
These lines look line comment lines to the XML Basic interpreter, so are ignored.

**Standardized Block Names.**<br>
Within each module to facilitate consistent processing by EditBas some
standardized block names are used. These names are reserved for the initial
blocks that will be present in every .xba module. The constants defined within
these blocks precede the macro definitions, so are module-wide. They typically
define paths, system parameters, and constants that are specific to the
Territory system for which the module is defined.
<pre><code>
< module-name >Hdr - module header containing standardized Module header information.
publicsMM - public constants used module-wide in macros.
Main - Main sub standard within every Module, usually an empty sub.
</code></pre>
These module names MUST be the first names in any < xbamodule >Bas.txt listing
so that they are included at the very beginning of the module, prior to any maro
definitions.
<br><a href="#IX">Index</a>
<h3 id="3.0">3.0 Setup.</h3>
There are several setups for the EditBas project.

**Setup for ExtractBas.**
<pre><code>Build Menu:
	ensure the .xba in GitHub/Libraries-Project/Territories is current by
	 using Calc>Tools>Macros>Organize Macros>Basic>[Organizer][Libraries]>Territories[Export](*)Export as Basic Library
	 to export the Territories .xba files to GitHub/Libraries-Project (this 
	 will automatically overwrite the Territories folder under /Libraries-Project)
	 
    copy the appropriate .xba from GitHub/Libraries-Project/Territories to
     the project folder (e.g. Module1.xba) by using the GetXBAModule.sh
     shell file
     
	edit 'sed' command parameters with basmodule and *bafile* parameters
	run *sed* from the Build Menu
	run *Make from the Build Menu to extract the source to a .bas file
</code></pre>
**Setup for AddBas.**
To add a new .bas block into  the library do the following:
<pre><code>
	run DoSedAdd.sh with the < basmodule > name, the new .bas block name,
	 and the position the .bas block should be inserted within the new library.
	 if no position is specified, the new block will be added at the end.
	
	run *make to add the new < basmodule > name into the < basmodule >Bas.txt
	 file at the desired position
</code></pre>
At this point the .bas block will be inserted into the rebuild of the
< xbamodule >.xba file in the /Release folder.
br><a href="#IX">Index</a>
<h3 id="4.0">4.0 Supporting Files.</h3>
<pre><code>CopyXBAoverBAS.sh - shell to copy extracted XBA module.bas file over
 existing /Basic/module.bas file.

GetXBAModule.sh - shell to copy XBA module from GitHub/Territories folder
 to project folder; typically done to extract current source from tracked
 .xba library file.
 
 DoSed.sh - shell to edit *basmodule* and *bafile* into .tmp makefiles
  for project.

DoSedAdd.sh - shell for setting up MakeAddBas operations.

DoSedBuild.sh - shell for setting up MakeBuidLib operation.

DoSedDel.sh - shell fo setting up MakeDelBas operations.

MakeAddBas.tmp - makefile template for inserting new block name into < xbamodule >Bas.txt

MakeBuildLib.tmp - makefile tempate for rebuilding library module.
 (MakeBuildLib) will be generated in Basic/< xmamodule > folder.)

MakeDelBas.tmp - makefile template for removing block name from < xbamodule >Bas.txt

MakeExtractBas.tmp - makefile template for extraction builds.

ModuleHdr.xba - (read-only) XML module header lines for use when replacing
  module into <xmafile>.

scratch.xba - intermediate .xba file used when replacing module into
  <xmafile>.

</code></pre>
<br><a href="#IX">Index</a>
