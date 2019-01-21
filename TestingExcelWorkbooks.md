<!-- http://codematic.net/excel-development/excel-dev-general/excel-testing.htm -->

# Testing Excel workbooks

Any serious Excel development should undergo a variety of testing throughout the various stages of its life.

## Lifecycle overview

Most spreadsheets will pass through the following stages (in some shape or form) one or more times during their lifetime :

* Requirements development
* Analysis
* Logical Design
* Physical Design
* Construction
* **Testing**
* Release
* Maintenance

---

## Types of Testing

There are 3 commonly discussed types of testing

### Unit testing

Performed by the developer before, during or hopefully shortly after building a 'unit'. In spreadsheets a unit is not clearly defined, but could be considered to be a formula or a block of similar formulas. In VBA it would be a procedure and/or maybe a module or class. The developer tests that the part they have built does what they think it should, in all the uses they envisage.

### System or integration Testing

Performed by the developer usually, but could be a dedicated tester (preferred) once the whole 'system' is nearing completion. This tests that the whole thing works together as one in the expected way. Ideally this should be done partially as early as possible so no nasty surprises surface just before the due date.

### User Acceptance testing

Performed by the customers of the system, this is often done towards the end of the project, the users test that the system does what they expect the way they want it to. Leaving this until late in the project represents a very real risk that the development team will build the wrong thing. The more agile approaches bring this forward and get user feedback as early as possible as components are developed.

---

## Testing methods

There are two commonly discussed testing methods

### Execution testing

Here data is passed through a process and the results evaluated to see if they match expectations. Also called black box testing - there is no consideration of the workings, simply if the system produces the expected results for all test input data then the testing would be considered a success.

### Inspection

Here the individual units of logic are reviewed to see if they make sense and can be expected to produce the correct results. This is also known as white box testing as the inner workings of the system are investigated. This type of testing is extremely difficult for the original developer to perform.

---

## Testing Spreadsheets

So that's the theory, but how do we ensure the spreadsheets we have built or depend on are fit for purpose.

The first things to consider are

* how correct does the spreadsheet have to be ?
* what confidence level is required ?
* How important is it ?
* How reliable has it been so far ?

Testing properly takes time and is therefore expensive, its essential to focus that effort on the most critical areas.

Also its important to understand the threat, to be clear what we are testing for. Is it just that certain numbers are within an acceptable materiality level of the 'correct' ones. Are we testing for potential fraud? Are we testing that confidentiality is not breached?

Testing is likely to be much more useful if you believe the spreadsheet is wrong and you are looking to see where, rather than believing it is correct and looking for some 'proof'. Sadly research clearly and consistently demonstrates that a significant proportion of live commercial spreadsheets give the wrong answers.

This section concentrates on testing of spreadsheets during development

TODO: Link , testing of live spreadsheets is discussed here and in more detail in Simons presentation to Eusprig 2006 here (ppt) and here (pdf).

This section assumes a fairly general spreadsheet design, and that the spreadsheet contains the following elements:

* Input data
* Reference data
* Calculations
* Results

---

## Practical testing

Appropriate testing is governed by the type of spreadsheet to a certain extent, some or all of the following may or may not be useful.

Clear all input data - results should be in an expected state (often zero), do the same with the reference data.

Populate every possible input and reference data cell with a 1 (one) - check they all flow through as expected

Populate the model with a known valid data set and confirm results agree

Populate the model with invalid data (input and reference) and check the impact, test data should be weighted towards realistic values, but should incorporate boundary values and invalid values. Input validation logic should be thoroughly exercised to ensure the validations fire whenever they should and never when they should not.

Consider carefully any possible missing inputs reference data or logic.

Vary the input data by a known proportion (say 20%) check the results move by the appropriate proportion (if the relationship is linear).

Use monte carlo simulation and tornado charts to check input/result relationships act as expected.

Use visualisation software to watch the relative values of cells as various inputs are changed.

Map the types of cells (formula, value or text) to look for unexpected types.

Review a list of unique formulas for possible issues like embedded constants

All the stuff from this powerpoint TODO: http://codematic.net/files/Papers/SimonMurphyEusprig2006.pps

Rip the model into a flat data table and point pivot tables at it, confirm results to formula based version

Trace formulas through all dependents/precedents for logical consistency

Ensure formula consistency where appropriate by re-copying and pasting formulas over whole blocks.

Add cross check totals wherever possible

Search for error results and deal with them (use select special and tab)

If you have used data validation (and you probably should) then be sure to click circle invalid and confirm there is no invalid data that is not part of the testing.

Use alternative approaches to confirm values, either different formulas or data filtering and/or pivot tables. For example break down a SUMIF to highlight those values that meet the criteria, then SUM them manually to confirm.

It is useful to keep as much of the test structure as possible with the spreadsheet as this means it can be rechecked in production easily. The down side is the test elements can complicate a model, this can be minimised by using a dedicated test sheet and hiding it when not required.

> If you are not using commercial quality spreadsheet testing and quality assurance products then you are probably wasting an awful lot of time and effort, and you will certainly be missing errors the tools find easily.
