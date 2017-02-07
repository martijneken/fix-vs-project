# fix-vs-project
This is a Visual Studio Extension to edit VB/C# project files for use as Rhino/Grasshopper plugins.

Specifically, project files are edited to:
- Set copylocal = true on project and local DLLs, but false on Rhino/Grasshopper DLLs
- Add post build events to rename DLLs to GHA, and copy output to ..\Output
- Find the Rhino install directory and set it to the RHINO environment variable
- Set a debugging action to launch $(RHINO)

The plugin shows up as a series of menu items in the Project context menu (on right click).
