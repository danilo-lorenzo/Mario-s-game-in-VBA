# Mario's Game in VBA

My life changed when I discovered that Excel was not used just for painting tables and adding data. We have a fantastic tool in our hand, capable of making extremely elaborate databases and sending emails automatically.



>“Everyone in this country should learn how to program because it teaches you how to think” – Steve Jobs



❓ **Definition:**

But what is Visual Basic for Applications (VBA) ? Well, some time ago when you wanted to run a series of commands in sequence in a programming language, it was necessary to tell the computer, line by line, what the command to be executed should be. This, in the early eighties, was called Macro.

![img1](https://media-exp1.licdn.com/dms/image/C4E12AQH0RFI6tTV7Xg/article-inline_image-shrink_1000_1488/0/1581811071665?e=1614816000&v=beta&t=f9WmR8BXw3gNHtL80BXHpDIncBJYd_4vLE_B5TeXwIU)

The term was after the Lotus 1-2-3 spreadsheet and it is used whenever a program is able to implement a method that performs successive actions from a command menu. Everything changed significantly in 1995 when the Microsoft Visual Basic 4 was released, making it possible from now on to program all Microsoft Office applications with a single language; enable direct interactions between brand applications. Although all the code was made with a single language (VBA), the term macro code was generalized for the proposals with the objective of making the spreadsheet processes more automated.



💻**Introduction to the Developer tab:**


First of all, it is recommended that the user has a basic notion of Excel's functionalities to make the most of the tool's potential (I will leave a series of useful links at the end of this ReadMe).

Let's get started! The first step is to make sure that the developer tab is enabled. If it is not, rest assured!

![img2](https://media-exp1.licdn.com/dms/image/C4E12AQHDpQsDryxorw/article-inline_image-shrink_1000_1488/0/1581813772147?e=1614816000&v=beta&t=-vezkQ0r-QBH8r1Qxz-rzXfTRFq5locpWA4ApUrUaG8)

To enable it, follow this path: File >> Options >> Customize Ribbon >> developer tab. Select the checkbox and click OK.

![img3](https://media-exp1.licdn.com/dms/image/C4E12AQHsaobYxnZqkg/article-inline_image-shrink_1000_1488/0/1581814204786?e=1614816000&v=beta&t=1PWbhtmru_j4Sy__OSUE_D_rpHjZ15jvE7iWjjssQaw)

With the developer tab enabled, you will be able to:

![img4](https://media-exp1.licdn.com/dms/image/C4E12AQHzW6K9aZT70g/article-inline_image-shrink_1000_1488/0/1581819956987?e=1614816000&v=beta&t=2KdFb0qVNBkhrC1uMf0afLuYT2F0nGcQCL7Y60d0HOY)

1. Record Macros: When clicking on "record macro" a window will appear requesting the name and description of the macro, from the moment you press "OK" Excel will start recording all the actions done, combining the actions in a kind of "function" that can be called by the user at any time.

2. Use Relative References: if this option is enabled before recording the macro, references will be created around the cell in which the macro was activated. For example: suppose that a code responsible for deleting a line is written in line 6. Using Relative References if the same code is executed on line 18, the deleted line will be 18 and not 6.

3. Visual Basic: this tab is where the programming actually takes place, it will be where we will focus on this Program. When clicking on the button the Visual Basic editor is activated.

4. Macros: displays the macro manager, in which the user can rename, execute or delete an already created macro.

5. Macro Security: a window will appear when this option is selected asking for user preferences regarding macro settings. Among them the situations in which the macros should work or not.

![img5](https://media-exp1.licdn.com/dms/image/C4E12AQFgkKLWruYvCw/article-inline_image-shrink_1000_1488/0/1581820429776?e=1614816000&v=beta&t=1GM2zRPvgr2ssujcgt4z5BTy9R0FAefj7t8AMH5S_pk)


6. Insert: in the insert option, several controls are offered for the document, the most used being the form control - command buttons, combo boxes, checkboxes, etc. (I will address each of the types of control in the forms topic).

7. Design mode: allows a much more complete view of the form structure. You can see the header, detail and footer sections. You cannot see additional data while you are making design changes; However, there are some tasks that you can perform more easily in Design view than in Layout view

8. Properties: Displays the Worksheet properties (I will address some of the Worksheet elements individually in the forms topic).

9. Show Code: enables the Visual Basic editor in the Worksheet module and spreadsheet that are activated. This module differs from the others in that it acts directly on the spreadsheet, not needing to be activated (called by the user); will be running from the moment the spreadsheet is activated.

🌟 **The Game:**

![img6](https://media-exp1.licdn.com/dms/image/C4E12AQFy-ImSSwRpjg/article-inline_image-shrink_1000_1488/0/1587161941089?e=1614816000&v=beta&t=kYttA_Fue-37PtD8yV_9obq1rtFDbZYE8KdRZ4mbtAM)

Objective: make a program that makes Mario walk around the scene and collect coins. Furthermore, when the game starts the player's name, the start time, the end time and the duration of the game must be recorded in a table automatically. Ex.:

![img7](https://media-exp1.licdn.com/dms/image/C4E12AQFqbfrPGGQdtw/article-inline_image-shrink_1000_1488/0/1587162257192?e=1614816000&v=beta&t=amHj8dJ6psuhQvaN6KyMrqNW946DEAZ_OnyOubtQ8q4)

How we will do it: We will create two independent Userforms and 3 programming modules that should interact with each other




