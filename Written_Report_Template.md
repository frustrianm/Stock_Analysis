# VBA of Wall Street - FRM's Tec Data Bootcamp Challenge #2

## Overview of Project

### Purpose
Throughout the module, the main objective was to help Steve analyse in faster fashion the stocks he has chosen to offer his parents better investing opportunities rather than their initial thought of DAQO. In the end, our VBA code would allow him to analyze any a small stock porfolio, and it did. But Steve, eager to use our code for bigger and broader analysis, asks us to refactor the code, which brought us here.
## Results

### Code Comparison
- Original VBA Code

![VBA_Module_Code](https://user-images.githubusercontent.com/96660344/149711193-70f6effc-39db-4302-9814-4cd823749793.png)


- Refactored VBA Code

![VBA_Challenge_Code](https://user-images.githubusercontent.com/96660344/149711200-b3555f6e-8616-40d7-9ab1-cda97636a047.png)


- Main differences

As stated in a previous commit, I was not able to make the "yearValue" variable to correctly run in the module code, and that for me was the biggest surprise in the refactored code, as it did run perfectly without having to type the hard data, as I did to get the 2017 and 2018 following timer screenshots.

Also, in the module code no index variable was available, being reasonable, as said in the challenge intrusctions, that it ran well for small data sample only. In the refactored code, tickerIndex index was created from the array of previous tickers and we also defined 3 new variables linked to tickerIndex, which are the ones that would allow us to run the VBA code with larger data samples.

Away from that, code reusing was useful and the structure of the code, with the nested loop and the conditionals remains almost the same. For me, that was the enlightment moment with the challenge, as I felt I was doing the same but knowing it was in a different way, and that it ran better than my original module code.


### 2017
- Code Results
![VBA_Challenge_2017](https://user-images.githubusercontent.com/96660344/149708802-14e64bc7-3002-4834-a7f4-15daaeefb8ee.png)


- Execution times

Original VBA code:

![VBA_Module_2017 timer](https://user-images.githubusercontent.com/96660344/149708707-d61c5eba-2843-4d87-9c84-a3cbb257e5dd.png)

Refactored VBA code:

![VBA_Challenge_2017 timer](https://user-images.githubusercontent.com/96660344/149708738-ae24124e-0bd8-4a2b-a6f6-07724234f977.png)


### 2018
- Code Results
![VBA_Challenge_2018](https://user-images.githubusercontent.com/96660344/149708935-1bd0e3d6-9326-4020-af6d-2b899027cf55.png)


- Execution times

Original VBA code:

![VBA_Module_2018 timer](https://user-images.githubusercontent.com/96660344/149708716-d360cce5-227f-406b-832d-3ffc0cfc21fc.png)

Refactored VBA code:

![VBA_Challenge_2018 timer](https://user-images.githubusercontent.com/96660344/149708755-a7a96115-921f-498c-90a9-d50fa91f6fed.png)

## Summary

### What are the advantages or disadvantages of refactoring code?
Advantages:
- The code gets more functions than the original one.
- We can reuse past code to rethink and shape new solutions.
- A broader variable definition in the refactoring can lead to more usage of the VBA code, extending its limits.

Disadvantages:
- Not a disadvantage as so in my opinion, but when I first though of refactoring in my mind the code was going to be ran in less code lines, and for instance, I even got the sensation the refactorde code was larger and more complex, but following the logic of being more powerful. Not a disadvantage as I said, but more like kind of a belief non-coders have (not me because I did the Hello World! progam already :) )
- Refactoring can take more time than the one at our disposal if the initial code instructions are not well established and new functions are asked on the go.
- Despite being asked to run the VBA code faster, only in 2018 the refactored code did so.

### How do these pros and cons apply to refactoring the original VBA script?
The refactoring help us to improve a code wihout taking away its original functions, and in this case, what Steve asked us was for a more powerful VBA code which could run more years than 2017 and 2018, with the refactoring part of the previous code was cleaned and in the same lines new variables and intructions were added to reach Steve's goal. Even though, as Alex said in one of the module's classes, not being broken, the refactoring process allow us to upgrade our code, doing the same in new fashion and even with a more sourceful and potent outcoume.

