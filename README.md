# CSVParseClass
A Clarion class for parsing delimiter separated files

This class uses a few different concepts, that I wrote a while back. 
It creates a run-time grid of &STRING references that point to the original data in the CSV.
This is a simple example of dynamic data structures (although in this case, all of the data is strings, but the concept works for pretty much any data type with my Caster class).

I would greatly appreciate if you could give it a try on various CSVs, with different separators and line endings and let me know if there are any problems.

Regarding all of the PRIVATE stuff in the class, you are welcome to do whatever you wish with it, but you really have to tread carefully.

There are quite a few features that I've meant to add, such as import/export of the grid map (so it doesn't have to re-parse every time you load the same CSV), and export of XML/JSON.

Perhaps I could do a presentation on it some day. 

## NOTE: This has not been extensively tested. It works on what I've thrown at it. Usually the problem is with the wrong separator or line endings.
