# CSVParseClass
###  Clarion class for parsing delimiter separated files
##### by Jeff Slarve

#### Features
* Autodetects common column separators and line endings (but has provisions to support uncommon separators too).
* Loads entire CSV into memory so it can be accessed like a grid.
* For those with a StringTheory license, there's a derived version of class that loads a file via StringTheory, which works a lot faster than SystemString (and uses less memory).
* Optional dynamic user interface for viewing data in a listbox.
* Generate a Clarion FILE, GROUP, or QUEUE structure that can be compiled into a Clarion app.
* Simple to use.
* Free.

#### Unusual concepts put to work

* Instead of creating copies of the data (via NEW()), this class points to the original buffer with &STRING references. <em>Depending on the number of rows/columns, this can save a lot of memory.</em>
* This is a simple example of dynamic data structures (although in this case, all of the data is strings, but the concept works for pretty much any data type with my Caster class).
* "Really" PRIVATE methods. Take a look at the MAP structure in JSCSVParseClass.clw, then look at the corresponding procedure code for those procedures. These procedures behave as if they're members of the class, but they're not included in the .INC file and don't clutter up the .EXP with extra stuff that nobody can use anyway.
* Take a look at the GetColumnLabel() method and how the "LegalChars" string is put to use. It's faster than using INSTRING().
* An &STRING references variable is 8 bytes in size, with 2 LONGs inside. The first LONG represents the ADDRESS() of the data. The 2nd LONG represents the SIZE() of the data. The GetCellLen() method looks at that 2nd LONG to quickly get at that size. 

#### Example projects

###### NOTE: These examples are not intended to be the ultimate in UI design. One of these days, I'll make them really perty. But they're functional. :-)

* CSVReader.sln demonstrates the UI.
* CSVReaderST.sln also demonstrates the UI, but does the loading of the file with StringTheory.
* SimpleCSVProcess.sln is a bare bones simple example of how to fill your own queue with data from the CSV.

#### TODO

There are quite a few features that I've meant to add, such as: 
* Import/export of the grid map (so it doesn't have to re-parse every time you load the same CSV)
* Export to XML/JSON/SQL.
* Support of escaped double quotes ("").

#### Known drawbacks/limitations
* Need to add support for escaped double quotes ("").
* There's a finite size of data that can be loaded into a 32 bit process. Putting LARGE_ADDRESS into the .EXP will allow more bytes to be loaded, but if the file exceeds the maximum size, then this class will not be suitable.
* Although no extra memory is used for COPIES of the data, memory does need to be allocated for all of the &STRING references used for each "cell". Sometimes this can be as much as the data itself, sometimes not. So oftentimes, the allocated memory <em> could </em> be as much as classes that DO create copies of the data. The number of rows and columns are what determine the amount of memory used.


## NOTE: This has not been extensively tested. It works on what I've thrown at it. If you have a CSV that this class fails with, I'd appreciate an example.
