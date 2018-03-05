# Transformation Script Development Menu
Dynamic menu to allow easy execution of design-time Transformation script for testing, workarounds, etc.

This menu allows for easy execution of design time scripts that work on individual documents, whole folders, or neither.  It will dynamically list all of the project script functions that require no parameters or require only an XDoc/XFolder (with any amount of optional parameters allowed). 

It should be used from Project Builder in either of two ways:
1. Providing an XFolder as a parameter, from an event like Batch_Open. Execute the event from the Runtime Script Events button (lightning bolt).
2. Providing an XDocument as a parameter, from a document level extraction event like Document_BeforeExtract, in a separate class not otherwise used in the project. Execute the event by selecting the class, selecting the document, then extracting the document.

When an XFolder is provided, Document functions will run on each doc in the folder.  When an XDoc is provided, Folder functions will run on the doc's parent folder.  Not all operations will work when using the parent folder from a document event.
