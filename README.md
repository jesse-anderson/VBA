Table of Contents
## [**VBA**](#VBA) 
1. [MSLC Tracker](#MSLC-Tracker-Code)
2. [MonteCarlo Pi Approximation](#MonteCarlo)
3. [False Position](#False)
4. [Modified False Position](#ModFalse)
5. [Van der Waals False Position Examples](#VDW)
6. [Simple Stats in VBA](#Stats)
## [Non-Coded Excel Examples](#NoCode)
1. [Finding Roots](#FindRoots)
2. [**Integration**](#Integrate)
## **VBA** <a name = "VBA"></a>
All VBA code that I have written.

#### [MSLC Tracker Code](https://github.com/jesse-anderson/VBA/tree/main/MSLC-Attendance-Tracking) <a name = "MSLC-Tracker-Code"></a>
###### This VBA program was written to automate the task of taking attendance in tutoring/discussion sections through blackboard collaborate. BB collab exports a .csv file that is then downloaded by me and can contain multiple rejoins and is often times not alphabetized. So inputting attendance for a student is annoying as is counting the amount of students, the average time they were in the session, and what percentage of the time I was tutoring in relation to the average time a student was present. This VBA code automates all of that based on the input of how long the person was tutoring. It alphabetizes the names and deletes doubles and adds the times together. After that it gives attendance data such as average time in session, count of students, and how long students were in the session in relation to how long the session went on. Note that the time for the particular session referenced was 150 minutes. 

#### [MonteCarlo Approximation Code](https://github.com/jesse-anderson/VBA/tree/main/MonteCarlo-Method)<a name = "MonteCarlo"></a>
###### [MonteCarlo Pi Approximation Info](https://blogs.sas.com/content/iml/2016/03/14/monte-carlo-estimates-of-pi.html): Wrote this code in my VBA class. I wasn't assigned this problem as homework or anything but rather did it in preparation for the exam so I could quickly use snippets if I lucked out and got it during the exam. I did not. :(

#### [Modified False Position Code](https://github.com/jesse-anderson/VBA/tree/main/Modified-False-Position)<a name = "ModFalse"></a> / [False Position Code](https://github.com/jesse-anderson/VBA/tree/main/Original-False-Position)<a name = "False"></a>
###### [False Position Info](https://mathworld.wolfram.com/MethodofFalsePosition.html): Solves for a root within a given bracket, note that a good 'guess' for the bracketed root's Left/Right end points is required.
###### [Modified False Position Info](https://www.charlesrcook.com/archive/2012/11/14/modified-false-position-method-in-c-accepting-a-function-pointer.aspx): Same as false position with some modifications to make more efficient, see relevant literature.

#### [Original Van der Waals Code](https://github.com/jesse-anderson/VBA/tree/main/Original-vanderwaals) / [Modified Van der Waals Code](https://github.com/jesse-anderson/VBA/tree/main/Modified-vanderwaals) <a name = "VDW"></a>
###### Application of False Position / Modified False Position to the Van der Waals EOS to find volume.

#### [Statistics Using Random Numbers Code](https://github.com/jesse-anderson/VBA/tree/main/Stats-with-Random-Numbers) <a name = "Stats"></a>
###### Simple statistics using random numbers in VBA, nothing special..

## **No Code Excel** <a name = "NoCode"></a>
"Code" written using excel spreadsheet functions
#### [Finding Roots Folder](https://github.com/jesse-anderson/VBA/tree/main/Excel-No-Code/Finding-Roots)<a name = "FindRoots"></a>
###### Root finding using Bisection Method, False Position, Goal Seek, Newton Method, Secant Method, and the root of the Bessel Function between 2 and 5.
#### Integration<a name = "Integrate"></a>
