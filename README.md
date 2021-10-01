# opensolver-libre-port
Port of OpenSolver for Excel to LibreOffice

## Installation Notes
Clone the repository and open openSolverPortTestDoc.ods.

## Usage Notes
1. Click the "Run Solver" button in the first sheet of the openSolverPortTestDoc.ods
(Note: May not work the first time with WrappedTargetException, but should work afterwards).
2. Enter the proper objective cell addresses, variable cell addresses, constraints cell addresses (e.g. B2, C2:D4, etc).
3. Choose objective sense(maximize or minimize. "Exact value of" functionality is in UI but may not work properly) and operators properly.
4. Run solver by clicking Solve! button.

## Developer's Notes
Porting a macro based extension from Excel to LibreOffice is not as simple as it seems.
The two macro coding languages, Visual Basic for Applications (VBA) and LibreOffice Basic (BASIC) are similar
in many ways, but are just different enough to be largely incompatible. The following notes are meant to
explain various work-arounds and adaptations we have made to allow the system to function. Due to the severe
lack of documentation for LibreOffice, there are a couple resources that would be helpful to continue work
on the system. [This book](http://www.pitonyak.org/book/) is extremely helpful, as it is the most complete
documentation on BASIC we were able to find. 
[VBA Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/excel) is helpful for
understanding VBA. [The API](https://api.libreoffice.org/) for LibreOffice is also a helpful reference.

### Introduction to the system
There are three major sections within the OpenSolver codebase. `Standard`, `ClassModules`, and `Dialogs`.  

#### Standard
This is where the main Solver functionality is stored, and where the API that users can interact with is
stored. This code is used for all Solvers within OpenSolver. The following are explanations of the modules
currently present within the port:

- Debug

    This module is used as an isolated module to test functions and classes. Outside of development testing
    purposes this module serves no purpose.
    
- OpenSolverAPI

    The API for OpenSolver. Most functions that Users will have direct contact with and influence over are
    present here. Almost all other functions present within OpenSolver are called from functions within this
    module. Here is where the variables used by the solvers are stored and retrieved, using the Named Ranges
    functionality of LibreOffice Calc. This is the most complete module.
    
    Functions present but incomplete or untested:
    1. RunOpenSolver
    
- OpenSolverConstants

    Constant values, enumerations, and functions that access them are largely stored here.
    
    Functions present but incomplete or untested:
    1. ReverseRelation
    
- OpenSolverErrorHandler

    Error Handling for all of OpenSolver eventually routes through here. Most functions call at least one of
    these functions on an error.
    
    Functions present but incomplete or untested:
    1. ClearError
    2. ReportError
    3. RaiseUserCancelledError
    
- OpenSolverIO

    IO functionality is implemented here. This includes accessing the file system, and accessing the
    workbook.
    
    Functions present but incomplete or untested:
    1. GetExistingFilePathName
    2. JoinPaths
    3. FileOrDirExists
    4. DeleteFileAndVerify
    5. SolverDirIsPresent
    
- OpenSolverLibre

    This is a good place to store functions that are necessary for OpenSolver to work on LibreOffice.
    Currently this only contains the function `LibreSolve`, which effectively short-circuits the current
    lack of full implementation on `RunOpenSolver` to allow the default system solver of LibreOffice to run
    using the inputs generated through the API. This function will eventually become obsolete when
    `RunOpenSolver` has a full implementation.
    
- OpenSolverModelValidation

    Validation for many of the inputs passed into the API happens here. If the input cannot be validated
    naturally in the original function, a validation function is created here.
    
- OpenSolverRangeUtils

    This module contains the Utility functions that deal primarily with ranges of cells. Utilities like
    merging multiple ranges, or ensuring that merged cell selection is correctly referencing the merged cell.
    
    Functions present but incomplete or untested:
    1. CheckRangeContainsNoAmbiguousMergedCells
    2. TestCellsForWriting
    3. SetDifference
    4. ProperUnion
    
- OpenSolverStoredNames

    Named range functionality is handled here. Saving named ranges or values, and retrieving those values are
    both handled within this module. Most of these functions are simply wrappers that do data type
    manipulation on the inputs so convert them all to doubles, and then store those doubles on the sheet.
    Due to how data types work in LibreOffice, anything that sets a name on the sheet must have access to
    both the LibreOffice workbook and the VBA worksheet. This is normally handled automatically by
    `GetActiveSheetIfMissing` from `OpenSolverIO`. However, if a VBA sheet is passed in, the LibreOffice
    workbook must also be passed in for setting to work.
    
    Another note: `GetSheetNameAsValueOrRange` is currently functional, but not as it should be in VBA.
    This is because we were not able to get sheet names to be correctly included in the saved names in both
    saving through VBA and saving through BASIC. If this functionality becomes required later on (i.e. it
    matters if there are different sheets with different variables) then this will need to be modified, as
    will `SetNameOnSheet` and `SetNamedRangeOnSheet` to correctly handle sheet names.
    
- OpenSolverUtils

    Utility functions that are not specifically for ranges. The top of this file contains operating system
    specific code. This is largely non-functional and definitely requires further investigation and
    implementation.
    
    Functions present but incomplete or untested:
    1. ConvertLocale
    2. MakeSpacesNonBreaking
    3. StripNonBreakingSpaces
    4. SolverSummary
    5. UpdateStatusBar
    6. ForceCalculate
    7. RemoveRangeOverlap
    
- SolverCommon

    Functions common to all solvers, that either set up or solve the problem. __This is entirely untested__.

#### ClassModules
This library contains all class modules used by OpenSolver. This is where the specific solvers are
implemented, and where the major solver model code is. __This library is entirely untested__. Note that all
Class Modules must begin with `Option Compatible` and `Option ClassModule` to function properly.

- COpenSolver

    This builds the solver models used by the various solvers. This includes getting the various variables
    and settings that most solvers require.
    
- CSolverCbc

    This is the specific implementation of the solver for the CBC Linear Solver. This solver is the primary
    concern for this project, and should be first priority when implementing the actual solvers.
    
- ISolver

    The definition of the solver interface, to be implemented by the actual solvers.
    
- ISolverFile

    Interface for using a model file when solving - _not entirely sure what this does_.
    
- ISolverLinear

    Interface for sensitivity analysis. Contains an implementation of writing the Constraint Sensitivity
    Table.
    
- ISolverLocal

    Interface for local functionality - _not entirely sure what this does_.
    
- ISolverLocalExec

    Interface for file system execution for solvers that are an executable file.

#### Dialogs
Macros in Module 1 interacts with 'DemoSolver' Dialog. Some of the main macros related to Dialogs are:

- StartDialog

    Generate a working pop-up screen of DemoSolver GUI.

- setAndDrawTargetCells

    Sets target cell(objective cell) using function from OpenSolverAPI, and draws objects that highlight the cells.

- setAndDrawVarCells

    Sets variable cells(decision variable cells) using function from OpenSolverAPI, and draws objects that highlight the cells.    

- setAndDrawConstCells1 (and 2-8 are essentially same, except that those are optional where 1 is not)

    Sets constraint cells using function from OpenSolverAPI, and draws objects that highlight the cells.

- onclick_finalSolve

    Sets all variable cells and draw highlighters using macros above, and runs solver. Deletes constraint named ranges after running solver.


### Helpful notes on translation
- Due to the relative lack of documentation for BASIC, it is a good idea to try the VBA commands first.
    * This can be done by including `Option VBAsupport 1` at the top of all new modules. This allows much
    of VBA to function within LibreOffice
    * The downside to this is that many VBA functions are not fully supported. Be sure to watch inputs and
    outputs of functions to ensure they are handling the variables as expected.
    * Also, when using data types exclusive to VBA, variables must be explicitly declared as that data type,
    or as an object. The variant data type does not correctly handle VBA types.
- Function returns are handled by assigning the function name a value, as if it were a variable. In VBA this
allows the function name to be handled like a variable, and can be passed in as a parameter to a function so
that the return value can be set in other functions. This is not supported in BASIC.
    * A workaround for this is to create a temporary variable that is passed into the function instead of the
    function name. This variable can then be modified in that function. Set the function name to the value of
    that variable to have the same effect as VBA.
- Optional function parameters cannot be set in the function declaration, like they can in VBA. To have the
same functionality, use the function `IsMissing()` to determine whether the optional variable is present,
then set the variable to the default if it returns true.
    + A side effect of this is that optional variables that do not have a default value still must be
    handled. Perform the same operation as stated above, but assign objects the value `Nothing` or strings
    as an empty string to avoid errors.
- Make full use of the breakpoint and watch features when testing code. Variables can change in unexpected
ways when being passed between functions, especially if a VBA type is not explicitly defined (as above).
- Though unfortunately discovered too late into production to be useful to us, it is possible that VBA
ranges contain within them a reference to the BASIC cellRange equivalent. This could be useful for
transitioning between the two going forward.
- Note that VBA array indexes start at 1, while BASIC array indexes start at 0.

### Helpful notes on User Interface
- Note that Dialog in LibreOffice Basic is quite volatile - it may crash just when you are changing name of a component in your Dialog. So be sure to save your work frequently!
- When you see an error and if you are using global variables, try using local variable that would grab the same data and run it again. In many of the debugging cases, for some reason we do not fully understand, this resolved the error.
- When assigning colors, using RGB ( 0 , 0 , 255 ) did not yield the color Blue - instead, it returned Red. Hence, try using decimal values (for instance, 255 for Blue, 16711680 for Red) instead.
