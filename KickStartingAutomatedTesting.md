<!-- https://sysmod.wordpress.com/2013/09/30/notes-from-softtest-kick-starting-automated-testing/ -->

# Notes from SoftTest “Kick Starting Automated Testing”

## Paul Gerrard’s talk was a call to review requirements.

He distinguished definitions of Check vs Test.
In Safety tests, they speak of a check for the part being “in good working order”
Look up Behaviour Driven Devt – a requirements driven structured language eg Cucumber, and DefoSpam.
You don’t want to end up with a long, brittle chain of unmaintainable test scripts.
Alternative: robot framework, write your own interface
See BusinessStoryMethod.com, SP.QA free story platform

---

## John McHale, Storm tech, presented Microsoft ALM, VS Ultimate 2012

MS Test Center for manual testing with recording and playback.
Test Impact: What LOC were tested, to see what tests might be impacted by a change in code.

---

## Dorothy Graham DorothyGraham.co.uk and on blogspot.com

Efficiency of finding bugs:
Regression testing 9%
Manual testing 28%
Exploratory testing 58%

---

## See BCS Grove report “Experience-driven Test automation”

High maintenance cost of testware leads to abandonment. What happens to all your scripts if you change tools?
Solution: Separate tests from scripts. (I am reminded of the saying that there is no problem in computer science that cannot be solved by another level of indirection / interface)
Tools don’t replace testers, they support them.
Tester don’t need to learn coding: lose a good tester and gain a poor programmer.
Test automation is a developer’s job.
Deadline pressure usually pushes you back into manual testing.

---

## David Evans, Neuri:

Attitude of learning from mistakes
Testing is never finished, merely stopped; defend your stopping heuristics.
Choosing which features to test = What do we have to let out that we assume is not working?
A bug report does not necessarily cause a bug fix to be added to a work queue.
Every bug (discovered in exploratory testing) is evidence of a missing (unit) test.
Value of a test is speed – timeliness to take action.
He is not fond of end-to-end testing – but I think a counter-example might be the Wobbly Bridge.
Transport analogy: the speed of the bus is never the point, but the service. So, neither is the speed of testing (goes against value argument?)
“Testing is slowing down development” –> Slowing the production of non-functioning code.
Coverage: lines of code are not paths.

---
