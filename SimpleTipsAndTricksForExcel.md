
<!-- 
http://exceluser.com/formulas/eight-tips-to-make-excel-easier-to-use.htm -->

# Tips and Tricks for Excel

simple tips that will make Excel easier to use, and also make your reports more accurate.

## Tip 1. When you sum columns of numbers, begin and end with border rows

Before Excel was invented, a building contractor was using Lotus 1-2-3 to prepare a complex bid that included a list of costs, somewhat like the first, shorter list here. To total the costs in cell A5, he used a formula like this:
A5:   =SUM(A2:A4)
At the last minute, he realized that he had omitted a major cost. So he inserted a row at the top of his list of costs, entered the value, recalculated, and printed his bid. But the row in which he had inserted the cost was outside the range of cells in his SUM formula.
The second figure above illustrates the result. The new value of 500,000 looks like it should be included in the SUM, but it obviously isn't. He won the contract, but lost his shirt.
The contractor sued Lotus for his losses. But the court ruled that Lotus wasn't responsible for the contractor's bad spreadsheet techniques.
These days, if the contractor had inserted a row at the bottom of the column and entered the value immediately above the formula, Excel might have updated the formula to include that new value. These days, Excel is supposed to make changes like that automatically, but it often fails to do so...and I never rely on this feature.
And certainly, if the SUM formula had been in another part of the worksheet, Excel would not have updated the formula's reference.
One way to avoid the problem when the new value is placed at the top is to begin the sum range in the cell with the label. That is, the formula in cell A5 in the left figure above could have been:
A5:   =SUM(A1:A4)
But a more general way to avoid this kind of costly mistake is to add border rows above and below your data, rows that you include in your SUM formulas. To illustrate, in the example at the right, the formula for the cell shown is:
A8:   =SUM(A2:A7)
For these "graycell lists", I always color the border rows gray, as shown. This makes them obvious and reminds me what purpose they serve.
Of course, when you insert rows for your new data, always insert them between the pair of gray border rows. This insures that your SUM formulas continue to yield the correct results.
These gray border rows are kind of ugly, of course. So I always reserve this technique for worksheets that I never intend to distribute.

## Tip 2: Specify the default number of worksheets you really need

By default, Excel includes three sheets in a new workbook. But for most Excel work, we rarely use more than one sheet. Therefore, when the other two empty sheets are left in your workbook, you—or whoever receives your workbook—must check the two empty worksheets to make sure that something important isn't included in them.
It's much easier to include just one worksheet in a new workbook, and then add more sheets as you need them.
To change the default setting of three sheets...
In Excel 2010 and after, choose File, Options, to launch the Excel Options dialog. Select the General tab, and then set Include this many sheets to 1. Then choose OK.
It’s easy to add a new worksheet, of course. Just click on the Insert Worksheet tab, which is the small tab to the right of the right-most worksheet tab at the bottom of your workbook.

## Tip 3: Get control of your Enter key

Excel’s default setting moves your selection down one cell each time you press Enter. This behavior is useful for entering numbers in a column, but it's an irritating "feature" at any other time.
Here’s how to prevent Excel from moving your selection down a row each time you press Enter:

In Excel 2010 and after, choose File, Options, to launch the Excel Options dialog. Select the Advanced tab. Remove the checkmark from After pressing enter, move selection. Then choose OK.

## Tip 4: Change Excel's default file folder when you open a new file

Typically, Excel’s default file path is: C:\Documents and Settings\Owner\My Documents
But if you're like most Excel users, you probably save your Excel files somewhere else entirely. To change this default setting so that Excel uses a more convenient default location for your files, follow these steps:
Choose File, Options, to launch the Excel Options dialog. Select the Save tab. In the box labeled Default file location enter the full path to your desired default folder. Then choose OK.

## Tip 5: Turn off those gridlines

Gridlines are distracting because they clutter your worksheet. Also, because gridlines don’t print, you get a better view of what your printed version will look like if they're not displayed at all.
To turn off your worksheet’s gridlines:
In Excel 2010 and after. Choose View, Show, then uncheck Gridlines.

## Tip 6: Get help about any spreadsheet function

To get help for any function, type an equal sign in an empty cell, the function you need help with, and then a left parenthesis.
For example, if you want help with the PMT function, type...
=PMT(
...and then press Ctrl+a (that is, hold down Ctrl and press a) to launch Excel’s Function Arguments dialog. This dialog provides short but helpful information about each argument for the function you typed.
If you need more help, you can launch the full Help topic for the function. To do so, click the blue Help on this function link, which is at the bottom-left corner of the Function Arguments dialog.

## Tip 7: Get help entering arguments in a worksheet function

Suppose you know how to use the PMT function, but want to add arguments easily.
First enter =PMT( as you did before. But this time, press Ctrl+Shift+a. Excel now displays this formula in the Formula Bar:
=pmt(rate,nper,pv,fv,type)
Press Enter.
Your active cell will show the #NAME? error because Excel is trying to make sense of the argument names. No problem. Double-click on the rate argument and select the cell with your interest rate. Double-click on the nper argument to set up its value. Then continue with the other arguments.
Follow these steps with any Excel function.
At times, however, Excel will not accept the formula when you press Enter. For example, if you type =SUMIFS( in an empty cell, and then press Ctrl+Shift+a, Excel will show this text in your Formula Bar:
=SUMIFS(sum_range,criteria_range,criteria,...)
But when you try to enter this formula, Excel returns an error dialog because it can make no sense of the ellipsis (the three dots), which indicate that you can have more pairs of criteria_range and criteria arguments. To fix this problem, remove the ellipsis and the last comma from the formula. You now can enter it with no problem.

## Tip 8: Launch format dialogs easily

You can modify the settings for virtually every object in Excel—a cell, a chart, a drawing object, and so on—by right-clicking on the object and choosing a menu item that begins with "Format". For example, to modify a cell, you would choose Format Cells; to modify a line in a chart, you would choose Format Data Series, and so on.
But here’s an easier way: To format any object in Excel, first select the object and then press Ctrl+1. If you select an image, Excel launches the Format Picture dialog. If you select a cell, Excel launches the Format Cells dialog. And so on.
Using Ctrl+1 is particularly useful in charts because it’s often easier to select a chart object than it is to right-click it.