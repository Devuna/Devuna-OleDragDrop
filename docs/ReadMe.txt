Using the OleDragDrop template for Clarion 10:

There are two templates in the KCR_OleDragDrop template set.

KCR_GlobalOleDragDrop is the Devuna Global OLE Drag and Drop Extension template.
Add it to the Global Extensions of your application.  It is required by the other
template(s) in the set.  It has a check box to allow you to turn off code generation.
If you check this option, there will be no Ole Drag and Drop code generated in the
entire application.  You can disable it at the procedure level with the next template.

KCR_OleDragDrop is the Devuna OLE Drag and Drop Extension to add Ole Drag and Drop
to your procedure.  You need to have a local CSTRING variable to contain the drag and
drop text.  Define it before populating the template.  


This template has several prompts:

Drag and Drop Control - select the control that is to be the source/target for the drag and drop
If the control you selected is also an ABC BrowseBox control the the template will automatically
generate the code to set the OleData to the BrowseBox contents.

Drag and Drop Data - the local CSTRING variable you created to hold the drag/drop text

Include Drag Support - uncheck to disable generating code to support the drag operation

Include Column Heading - if this is an ABC BrowseBox add the column headers to the OleData.

Field Delimiter - choose either CRLF (Carrige Return Line Feed) or TAB as a field separator.
If you are pasting into Excel, use CRLF to put the data into different rows; use TAB to
put the data into different columns.

Effects Allowed - these control the effects that are allowed by the drop source.  You must
have at least one of these selected.  Keystate modifiers have the following standard OLE
effects:
ctrl = copy
shift = move
ctrl+shift = link
It is up to you to respond appropriately in your code to the returned effect after a successful
drop.  For copy you don't need to do anytihing; for move, you would need to delete the queue
entry (for a listbox) or the selected text (for text or entry controls).  Link is provided
for completeness; you'll have to decide how to handle this yourself.

Include Drop Support - uncheck to disable generating code to support the drop operation

OleDrop Event Value - when a successful drop has been completed, the generated code will
post an event back to your main window.  You can set the value for that event here.
the default is 09000h

***WARNING***
If you are using a List control as your Drag and Drop control, make sure it does not have
a DRAGID or a DROPID as Clarion list box drag and drop support will cause problems.
***WARNING***

An example application is provided in the 
CSIDL_COMMON_DOCUMENTS\SoftVelocity\Clarion10\Accessory\Devuna\Examples\Ole Drag and Drop directory.
This is the standard People.App with Ole Drag and Drop support added to the people browse.

There are four interface classes included in the package.  These classes are wrappers for the similarly
name interfaces.  Please see the interface documentation on MSDN for details of the interface
methods.  Besides the interface methods, each class has several common methods: contruct,
destruct, init, kill, getinterfaceobject, getliblocation, and _ShowErrorMessage:

construct - standard constructor

destruct - standard destructor

init - standard initialization for class ... must be called before other methods
        IDropSourceClass.Init - no parameters

        IDropTargetClass.Init(hwnd) - pass the handle of the window that is to be the drop target

        IDataObjectClass.Init(fmtetc, stgmed, count) -  fmtetc pointer to a tagFORMATETC group array
                                                        stgmed pointer to a stgmed group array
                                                        count long number of array elements

        IEnumFORMATETCClass.Init(fmtetc, count) -       fmtetc pointer to a tagFORMATETC group array
                                                        count long number of array elements
                                                        
kill - for clean up on exit

getinterfaceobject - returns a pointer to the raw interface

getliblocation - returns nothing for these classes

_ShowErrorMessage - used by the class to display internal errors


If you want to know more details about how OLE Drag and Drop works, and also what I used as a guide,
visit http://www.catch22.net/tuts/dragdrop/1

There is also an OleData class that has just one method GetOleData
This class method is called from a subclassed winproc.  As such it runs on a different
stack than your main code.  This means that you should only access global and module data
from this procedure.

This procedure is called just prior to initiating the drag operation to allow you to prepare
the data for transfer.  Please look at the comments and default generated code.

To use the Classes and Templates with older clarion versions, you will need to copy the following
Clarion 10 files to the appropriate libsrc folder:
svapi.inc
svapifnc.inc
svcom.inc
svcom.clw
svcomdef.inc
winerr.inc


Please email your questions, comments, or sugestions to rrogers@devuna.com

Thanks for choosing our Ole Drag and Drop Template set!

Here is a list of the files installed in the CSIDL_PROGRAM_FILES\SoftVelocity\Clarion10\accessory folder:
libsrc\win\cDataObject.clw - Class Implementation
libsrc\win\cDataObject.inc - Class Declarations
template\win\OleDragDrop.tpl - the templates
images\DragDrop.ico - template icon


Here is a list of the files installed in the CSIDL_COMMON_DOCUMENTS\SoftVelocity\Clarion10\Accessory\Devuna
Documents\Ole Drag and Drop\licence.txt - your licence
Documents\Ole Drag and Drop\readme.txt - this file
Examples\Ole Drag and Drop\people.app
Examples\Ole Drag and Drop\people.dct
Examples\Ole Drag and Drop\people.tps
