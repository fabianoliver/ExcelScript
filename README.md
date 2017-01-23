# ExcelScript
Roslyn-based Excel AddIn that allows to write, compile, and run C# code right in Excel using simple user-defined functions

Aside from doing simple immediate evaluations, some of the core features are:

- You can create re-usable scripts on the fly in your worksheets, which will be stored in an in-memory cache with a used-defined string key.
- You can then invoke that script using its key anywhere as often as you like
- You can define and pass parameters to your scripts, so they can be invoked with arguments
- Range-objects are fully supported as input parameters (ExcelScript takes care of converting & marshalling these internally)
- You can create the parameter definitions & script manually, or just parse a function and have the AddIn extract and generate all necessary handles automatically
- You can even register your script, meaning the AddIn will register a User Defined function with the name and all parameters you defined to invoke your script
- A helper function lets you format & syntax highlight your C#-code in Excel
- You can serialize/deserialize your scripts to strings or files

Some bonus features:

- You can debug your scripts. To do so, try the following:
  1. Open the addin and the Example.xlsm
  2. Attach the debugger of your choice to the Excel.exe process (e.g. open Visual Studio -> Debug -> Attach to Process -> Excel)
  3. In any script, e.g. lets say on Sheet "Parse" in cell B18, enter "System.Diagnostics.Debugger.Break();"
  4. When this script is run, your debugger should behave as expected; break execution, show the code, show local variables, etc.
- The library supports execution in different AppDomains. By default, a script will run in a shared appdomain between all other scripts; you could also pass in an option to run in a shared appdomain with the main addin (ExcelDna's default domain), or to create a seperate appdomain for that script.
- By default, all functions are marked non-volatile. Unfortunately, that means that after opening a workbook, you need to refresh each formula that creates a script by hand or macro. ExcelScript offers an experimental way around this, though:
  1. Open The ExcelScript-AddIn.xll.config
  2. set TagDirtyMethodCalls to true
  3. This will mark each cell creating a handle (script, parameter and so on) to dirty once. Could be particularly useful in manual calculation mode (in auto calculation mode, all scripts would run on the next recalculation, so not very useful).

Project includes an exemplary Excel-file to demonstrate usage. Please see the Wiki on Github for further guidelines.
