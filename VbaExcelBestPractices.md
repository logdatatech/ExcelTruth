# Excel VBA Best Practice

<!-- http://codematic.net/excel-development/excel-vba/best-practice/overview.htm -->

<!-- TOC -->autoauto- [Excel VBA Best Practice](#excel-vba-best-practice)auto  - [Fundamental Imperative](#fundamental-imperative)auto  - [Formal Development Process](#formal-development-process)auto  - [Spreadsheet Development Lifecycle](#spreadsheet-development-lifecycle)auto  - [Coupling](#coupling)auto  - [Fan in / fan out](#fan-in--fan-out)auto  - [Development Environment Enhancements](#development-environment-enhancements)auto  - [Design - 2 Questions, 2 Approaches](#design---2-questions-2-approaches)auto    - [2 Questions](#2-questions)auto    - [2 Approaches](#2-approaches)auto  - [Design Considerations](#design-considerations)auto  - [High Quality Code - General](#high-quality-code---general)auto    - [Naming conventions](#naming-conventions)auto  - [High Quality Code - Modules](#high-quality-code---modules)auto  - [High Quality Code - Procedures](#high-quality-code---procedures)auto  - [High Quality Code - Blocks](#high-quality-code---blocks)auto  - [High Quality Code - Variables](#high-quality-code---variables)auto  - [High Quality Code - Excel](#high-quality-code---excel)auto  - [High Quality Code - Classes](#high-quality-code---classes)auto  - [Avoiding Common Errors](#avoiding-common-errors)auto  - [Debugging and Testing](#debugging-and-testing)auto  - [Documentation Standards](#documentation-standards)auto    - [Third Party Options](#third-party-options)auto  - [VBA Code Convention for Excel VBA](#vba-code-convention-for-excel-vba)auto  - [Extending Excel](#extending-excel)auto  - [VBA Best Practice Summary](#vba-best-practice-summary)autoauto<!-- /TOC -->

> (Note this advice is specific to Excel VBA in business applications

---

## Fundamental Imperative

The fundamental imperative is quite simply to manage complexity, a theme repeated in many walks of life. Without this focus complexity rapidly gets out of control.

* Manage complexity (McConnell)
* Solution complexity grows at 4 x the rate of problem complexity. (Glass)
* Things should be as simple as they can be, but no simpler (Einstein)
* K.I.S.S. keep it simple stupid.

This principle should drive all other work.

Easier to build, easier to test, easier to document, easier to use, etc. There should be no conflict

---

## Formal Development Process

Systems Development lifecycle (these don't all flow neatly like a waterfall, but are the main elements of the process of delivering working, fit for purpose software)

* **Requirements** \- understanding what the system should do, this will change as the project progresses - be ready for that
* **Analysis** \- sorting through the requirements, understanding the implications, dealing with conflicts
* **Logical Design** \- a theoretical design, perhaps on paper of the sort of things the system will need to satisfy the requirements. This element drives the next one
* **Technology Choice** \- based on the logical design, what is the most appropriate technology to implement the system? This step is so often missed.
* **Physical Design** \- converting the logical design into the actual technology - eg if the system needs a relational database or an OLAP database, this stage defines the table structure.
* **Construction** \- actually building the system - involves going back through all the stuff as new requirements are identified etc.
* **Test** \- each component should be tested (unit testing) as soon as possible and throughout development, then they should be tested together (integration testing). System testing aims to ensure the model does what the developers think it should. User Acceptance testing aims to check it does what the users need, (at this stage that should be clear and well known to the developers).
* **Release** \- making the system available, live to the users, may include deploying the system to servers and/or desktops.
* **Maintain** \- keep the system working, adapt to changes in requirements
* In some shape or form.
* Understanding the requirements is critical and difficult
* Some sort of structured approach is useful
* Build a library of useful code and references

---

## Spreadsheet Development Lifecycle

Oh! I need a tool - clickety-click, reasonableness check, release, (Test & Maintain in live environment).

---

## Coupling

How strongly two separate elements depend on each other

Low coupling is better, especially through clearly defined interfaces

High coupling often means hidden dependencies which generally leads to incorrect modifications (side effects)

Example: hard coded cell addresses in VBA code
Example: Gobal/Public variables

---

## Fan in / fan out

Fan in / fan out refers to the shape of the dependencies of a particular routine.

Low fan in - a VBA routine depends on only a few other units

Good because it minimises dependencies and reduces complexity

High fan out - a VBA routine is used by many others

Good because it minimises duplication
Good because routine is more thoroughly exercised

---

## Development Environment Enhancements

In order to develop decent code, you need decent tools. The VBA IDE is certainly pretty tired looking these days there are still some things we can do to improve on the default set up.

<!-- TODO:http://codematic.net/excel-development/excel-vba/excel-vba-ide.htm -->

Option explicit
Don't warn about syntax errors
Actual set up - which windows are visible etc
Use a code library
Use MZ tools
Other tools - call tree, indenter, reference checker, BYG
Source control, even just exporting as text

---

## Design - 2 Questions, 2 Approaches

In terms of design there are two key questions and two key approaches.

### 2 Questions

What will happen when things go right ?
What will happen when things go wrong ?

It is important to consider these two cases, many problems are created later in the life of the system if the second one is ignored early on.

### 2 Approaches

What will the system do ?
What real world objects am I modelling ?

The first question leads to a procedural design, the second leads to a more object oriented one. OO has lots of benefits in many large scale developments. Often for smaller systems, which many spreadsheet systems are, a procedural approach is the simplest to implement and maintain. If all developers have only OO experience though then that must be factored in, and may result in an OO system because that is easier for the development team.

There is no best, or right and wrong, people who advocate one or the other very strongly, often just have no experience of the other. Certain technologies can also influence choice - Excel/VBA is not very rich from an OO point of view so a functional or procedural approach often makes sense. C# on the other hand has very powerful object oriented features and work best in OO style projects.

---

## Design Considerations

* Simplicity
* Testability
* Cohesion and coupling
* Fan in fan out
* Encapsulation and separation
* Reusability
* What does it do, or what are you modelling?
* N tier - especially msgbox and error handling
* Consider future changes and/or enhancements
* Design patterns eg: singleton, factory, facade
* Design 'normal' case and 'error' case

---

## High Quality Code - General

### Naming conventions

Option 1 fit in with current standards (eg Reddick 3 letter)
Option 2 give everything a simple, distinctive, non abbreviated descriptive name
Procedures - verb noun combination
Variables - add scope prefix, data type maybe, usage more useful - eg m_rwCustomer for row in customer table or colMonth for month column.
Only use a type prefix if you know what it is doing for you and you value that. (eg probably don't need 'obj')

> Avoid most comments - make meaning crystal clear in the code
> Move non standard code and workarounds to named routines, and comment if necessary
> refactor early and often
> build test routines as you go, and at least step through in the debugger

* Recognise when you move from procedure based design to module/object based (eg procs need to return >1 value, lots of parameters get passed all around)
* Object Oriented designs can be useful, but can cost performance
* The biggest factor on performance is design
* Be aware of Excel/VBA limitations, and clear on its strengths
* Use 'TODO' and maybe 'Enhancement' comments (don't release until TODOs cleared - (use edit find in project)
* Use shift+F2 for definitions, F8, Shift+F8 and Ctrl+Shift+F8 (step out) and stop
* Note compiler is weak compared to C++ (eg:won't catch assigning a string to a variable of type double)
* Use code cleaner regularly (indenter,line joiner, blank line management)
* Differentiate style from quality
* Use the most restrictive scope that works for variables and routines

---

## High Quality Code - Modules

* As solution complexity increases think in terms of modules rather than procedures
* Use enums to add understanding
* Could use types but probably better to go straight to classes.
* Module name should explain what it does, a comment at the top may be useful.
* May need 1 or 2 global variables, but may need lots of module level variables
* A module should represent a bunch of closely related things eg: use common data
* Use Option private module otherwise public routines will be listed in the user defined functions list. (Excel Front End)
* Forms should usually call straight out to a class or module to do the real work.

---

## High Quality Code - Procedures

* Consider the anatomy of a procedure/object - initialisation, process, clear up, error handling
* Short routines 40-50 lines max (one screens worth) (easier to understand - but not proven to reduce errors)
* A procedure should do one thing well
* Prefer functions (Boolean success/failure)
* Use parameters rather than global variables (max 7+/- 2)
* Avoid application.run - breaks VBA error handling
* Use separate routines for separate error handling
* Consider error handling early, use break on all errors when testing
* Use environ rather than api where possible (simple)
* All call tree parents should have error handling
* Test performance don't guess (see perfmon)
* Protect procedures from bad input (including malicious input)

---

## High Quality Code - Blocks

Code block start and end together

If using If.then code an Else if, add a comment, only remove it if you are sure there is no 'else' case (missing the else is a common error)

Code the expected behaviour first, usually.

Use do while and do until loops, be sure they will end, I generally mainly use for next and for each next, finding out the end first helps with status bar updates.

In a Select Case always have a default case, possibly with an error warning.

---

## High Quality Code - Variables

* Don't reuse variables for different purposes, avoid 'temp'
* Use clear data types, and control them (watch out for implicit coercion)
* Variants are often easier to work with than arrays
* Pass parameters byVal if they are not to be changed (watch for coercion)
* Explicitly use byRef in an input parameter is to be changed, but watch for signs to redesign
* Don't implicitly use the default property, be explicit (eg range.value)
* Magic numbers and strings should be made into constants.
* Use explicit data casting Cstr() Cbool() etc
* Code with early binding if at all possible (new), change to late binding to support multiple versions or if component may not be installed. Late can be slower and may need better error handling.
* Can define all variables together at top or just before they are needed both have pros and cons.
* Boolean variable should be named so True or False make sense - eg use done rather than status
* Booleans should be positive eg if not found then. rather than if not notFound then.
* Never mess with a loop counter within the loop.

---

## High Quality Code - Excel

* It is usually dangerous to hardcode a range reference in VBA.
* Use range names to connect code to worksheet ranges
* Choose a single cell range name as a start point
* Choose a multi-cell range to allow the user to insert/delete rows/columns
* Use in cell text where possible, closely ties the mechanics to what the user uses.
* Always check cell text thoroughly before depending on it
* Note data validation is easily bypassed
* Give sheets meaningful code names and use them (Not Index Numbers)
* Take the first few rows and columns for system use, hide them if necessary
* Check the selection type with typeOf if you are going to use it.
* Generally avoid selecting and copying - set values instead where possible.
* Be clear and explicit which workbook and worksheet code should operate on, especially ThisWorkbook v ActiveWorkbook for add-ins.
* Pulling Range.Value or Range.Formula into a variant is much faster than looping the cells.
* Use worksheet.cells(r,c) structure rather than range(""A"" & someNumber)
* Be consistent how you use ranges (between Range() and Cells() for example)
* Remember the power of R1C1 notation
* Its often easier to set the formulaR1C1 than the formula of a range.
* If distributing code with references, consider making it late bound.
* Use standard environment management routines (calc, events etc there is one here)
* Use Excel functionality wherever possible
* Use doEvents where needed
* Work upwards when deleting rows, otherwise counters get messed up
* Watch for inadvertently firing events
* don't hog the onCalculate event

---

## High Quality Code - Classes

* Classes result from asking 'what am I modeling?' rather than 'what is this system to do?'
* Full class based design may be overkill for most VBA applications
* Hard/impossible to combine the best of spreadsheets with proper OO - data hiding is not possible.
* Classes may be useful when a set of routines share lots of data and it becomes cumbersome to pass around as parameters
* Or when you need to return a complex type from a routine

---

## Avoiding Common Errors

Consider lifetime - initialise, set, use, destroy (CRUD - Create, read, update, delete)

Off by one errors, especially ranges into variants, arrays, loops and ubound()

Watch for index cross talk

Dim x, y as integer => x is a variant

StrComp, InStr - check the documentation

---

## Debugging and Testing

* Debugging
* Immediate window
* Locals
* Debug.print debug.assert
* Testing
* Pre-conditions and post conditions
* Use a 'test' module, and application.run for private routines
* Test early and test often
* Keep all test code - it can act as a specification, and can give comfort that later changes do not break anything
* Get good test data - realistic, test full range including boundary conditions, but weight testing towards real world use.
* Execution testing v static testing and code inspection/review.
* Unit testing, system testing, user acceptance testing.

---

## Documentation Standards

* If you design for simplicity only very minor additional documentation is needed
* Design and build the user Interface with the Users needs primary, the documentation will automatically be there
* Design and build the business logic parts with the maintainers needs primary, the main documentation will automatically be there, expand as required.
* Integral tests should explicitly clarify intent.
* Reports should contain enough description to be meaningful.
* External documentation is almost always so out of date its worse than useless.
* Excessive documentation is too hard to plod through
* Poorly targeted documentation is pointless
* Working software is more useful than documentation
* Sometimes documentation is important
* Call tree print out should be enough for most uses
* Be clear who the target audience is
* Should be auto generated from source code if it is to be up to date.
* Stepping through code is often the easiest way to understand it.

### Third Party Options

PUP provides a nice summary

MZ Tools creates a weighty document

TODO: there are more options to add.

---

## VBA Code Convention for Excel VBA

* Always use Option Explicit **usually** use Option Private Module.
* Give Module level variables a scope qualifier prefix g*, p* or m\_ (Global (all open workbooks), Public (this workbook), or Module). Scope everything as tightly as possible. Use procedure level scope where possible and pass values by parameters.
* Use meaningful names for all procedures and variables. Do not bother with convoluted data type prefixes, that is inappropriate for business level applications, makes the code harder to read and adds very little of value. The compiler will pick up any obvious data type errors. Used mixed case descriptive names - if the name is too long then probably the procedure is too.
* Procedures should fit on one screen - ie be 40-50 lines long maximum.
* Avoid most comments - make the executable code meaningful and simple instead.
* Avoid magic numbers and strings - use constants.
* Never comment what the code does - that should be crystal clear from the code, comment WHY something is done, especially if it is unusual. Add a couple of sentences to provide an overview of a module or class.
* Pass parameters ByVal (ByRef is the default) - only use ByRef where you intend to modify the parameter and pass the change back to the caller.
* Avoid Application.Run where possible as it breaks the error handling stack.
* Use additional tools. See links page for some suggestions. TODO: Link to ToolsPage ?
* Be aware of other options if VBA appears inappropriate for certain aspects of the project.
* Vary any rules you the developer feel do not promote clarity, simplicity and safety.
* If you really want to write high quality code read Code Complete 2 by Steve McConnell - and then apply it.

---

## Extending Excel

Excel is powerful not perfect

Leverage benefits whilst managing weaknesses by using complimentary technologies

Data: VBA, ADO, ODBC, OLAP, .net, COM, Info Bridge, XML, DDE, Web Queries, SOX Solution Accelerator

Logic: VBA, COM, .net, xll, Pivots, Filters

Search the web for vast array of samples, questions, tips and tools. Don't Reinvent things unless needed.

---

## VBA Best Practice Summary

Manage complexity
Consider security carefully
Excel/VBA not the best tool for everything
Be defensive, especially with inputs
Understand your data
Aim for cohesive models with low coupling
Use names to connect VBA to worksheets
Use complimentary technologies where appropriate
Manage documentation
