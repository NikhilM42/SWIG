# Welcome to the SWIG wiki!
Standard Work Instruction Generator

This script was started because I got tired of making standard work instructions by hand and was just dipping my toe into python. No one else really used it so here it lies gathering dust.
The script reads a text file of instructions and generates an excel workbook. Plenty of room for improvement if I ever come back to this script. **WARNING** There is no error handling baked into this currently.
The textfile has to follow some special rules as follows:

## Rules:

In general each line must start with a lower case letter to specify whether the line is a title or an instruction block and if it is an instruction block how much width the block takes up. The worksheet is divided into a 3 column grid format.

"t" - *title*
- To make the line a title also consumes the entire worksheet width

"r" - *regular*
- To make the instruction block consume one-third of the worksheet width

"m" - *medium*
 - To make the instruction block consume two-thirds of the worksheet width

"h" - *half*
- To make the instruction block consume half the worksheet width

"w" - *wide*
- To make the instruction block consume the entire worksheet width

## Special Rules:

"NOTE"
- can be added at the end of a regular instruction if you wish to add any special attention instructions to the bottom of the instruction block. Will not be parsed for any lines that have been declared as titles.

## Example:

```
tInstructions
rStep 1
rStep 2
rStep 3
wStep 4 NOTE: THE STEPS ARE PLACEHOLDERS FOR INSTRUCTIONS!
hStep 5
hStep 6
mStep 7
rStep 8
```

## Limitations:
You have to combine instruction block sizes so that they consume all 3 columns cleanly. Note the example section above:

Titles and the "w" will consume 3 columns
> tInstructions\
wStep 4 NOTE: THE STEPS ARE PLACEHOLDERS FOR INSTRUCTIONS!

Regular blocks will consume only 1 column so combining them groups of 3 is valid
>rStep 1\
rStep 2\
rStep 3

Or combine a regular block with a medium block

>mStep 7\
rStep 8

Half blocks will only be used with other half blocks

>hStep 5\
hStep 6

## To Use
Simply call the script from command line within its folder and provide it the name of the text file that has the instructionse when it asks. The script will automatically look inside the *instructions_src* folder located locally within the folder the script is in. The script will also generate the excel workbook in a folder called *output_swi*.