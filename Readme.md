# Dagpengekort app
This is a small console application meant to simulate the dagpengekort system in Denmark.
The program is written in C#.

It shall have 3 functions as stated by the assignment:

- Read the dagpengekort (xcel file)
- Calculate the payment based on the dagpengekort
- Produce a payment file in JSON-format

The generated JSON files will be placed in a Folder created by the program.
The new folder will be placed in the app folder using the absolute path.

## Dagpenge Situations
A list of situations used to test the applications

- A dagpengekort with no registrations

- A dagpengekort with vacation

- One with sickenss

- Two cards, only the second one is correct

- A card with some work hour, where "Teknisk belægning" should be added to "Fradrag per dag" becomes 7,4

- A card with both vacation and work hours

- A card where too much was paid last month, so additional hours has been added here, to make up for what the person owes.

- A card where the person starts working, card has work hours and techincal addition

## Dagpengekort Rules

- Every month has 160,33 hours. Always.

- Reduction (HR) = days with technical addition * 7,4

- If a work day has more than 7,4 hours, then the full work day hours are counted instead. (example 8 hour work day counts as 8 hour reduction)

- If a work day has less than 7,4 hours, then technical addition will be added, so that it totals 7,4 hours with the work hours for that day.

- Dagpenge is only paid for the month if a person has minimum 14,8 hours left in that month


