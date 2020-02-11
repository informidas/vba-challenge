# vba-challenge
VBA scripts and corresponding reports for challenge assignments

----

## Constructs used in solving the Challenge

* Ranges 

* Loop Structures

* If Then Else

* Comments

* Variables

* Msgbox and Stop statements to test variable assignment and subsection of my code


----

## Pseudo Code

There are many approahes to solving the challenge. Here is the approach I found most effective is breaking the assignment is small tasks or functional activities

#### Step 1 - Create a Subroutine

Create a Subroutine to act as a container for my code


#### Step 2 - Declare Variables needed

Declare the variables needed for ranges and counters


#### Step 3 - Initialize Variables

Although not stictly needed, I declare them to ensure I am confident about their starting values. <br>
As part of the initialization process I like to set the values for my last row and column since this is something I will use when manipulating datasets <br>

* set the ticker variable equal (=) to the contents of cell A2
* set the row counter equal (=) to 2

#### Step 4 - Create a Nested Loop structure to move through dataset on each sheet

The two dimensions of looping that will be performed are: <br>

* Looping through the sheets (using a For Each structure)
* looping through the datasheet on each sheet (using For / Next or While / Wend structure)

#### Step 5 - Begin Processing each sheet

For each sheet in the workbook: <br>
  * Sort the dataset in ascending order by date to ensure the data is organized
  * using a for loop start moving down row by row until you get to a cell where the value in the A column **does not match** the value currently assigned to the ticker variable. i.e. This means we have gone past the last row matching the current ticker
  * While we are looping, kep adding the value in the Total Stock Volume column to current *Total Stock Volume* variable

#### Step 6 - Set the values for First Open Day and Last Open Day for the current ticker

Assign the value  the Last Close Day column to the variable created for *Last Close Day* <br>
Assign the value from the First Open Day column to the variable created for *First Open Day*


#### Step 7 - Populate Summary Table

Use the ticker, Last Close Day and First Day Open, Total Stock Volumes variables to populate the values in a summary table in columns *I, J and K*.


#### Step 8 - Define a range for the Summary Table 

* Find the last row in the summary data section
* Use the last row to define a range for the data summary table
* Sort the *Percentage Changed* column in **Descending** order to get the *Largest Percentage Increase* metric 
* Sort the *Percentage Changed* column in **Ascending** order to get the *Largest Percentage Decrease* metric 
* Sort the *Total Stock Volume* column in **Descending** order to get the *Highest Total Stock Volume* metric 

#### Step 9  - Manipulate Summary table to find the Largest Percentage Increase / Decrease

* Set the values for the ticker showing the greatest percentage increase

* Set the values for the ticker showing the greatest percentage decrease

* Set the values for the ticker showing the greatest total volume

#### Step 10 - Audit Log

Keep track of the sheets you processed and append to a File Processing log.  This is not strictly required but good practice. Print a message to the user when the processing has completed, showing the files that hav been processed.

 




