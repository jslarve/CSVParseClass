!ABCIncludeFile

!MIT License
!
!Copyright (c) 2021 Jeff Slarve
!
!Permission is hereby granted, free of charge, to any person obtaining a copy
!of this software and associated documentation files (the "Software"), to deal
!in the Software without restriction, including without limitation the rights
!to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
!copies of the Software, and to permit persons to whom the Software is
!furnished to do so, subject to the following conditions:
!
!The above copyright notice and this permission notice shall be included in all
!copies or substantial portions of the Software.
!
!THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
!IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
!FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
!AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
!LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
!OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
!SOFTWARE.
!

OMIT('_EndOfInclude_',_JSCSVParsePresent_)

  INCLUDE('ABPOPUP.INC'),ONCE 
  INCLUDE('SystemString.inc'),ONCE

_JSCSVParsePresent_      EQUATE(1)
                                       
JSCSV:FirstRowIsLabels   EQUATE(01h)   !A flag to tell class that first row of CSV is for labels only.
JSCSV:ColumnLen          ITEMIZE(1),PRE()
Max                        EQUATE               !A switch for GetColumnLen - Get Max length
Total                      EQUATE               !A switch for GetColumnLen - Get Total length
                         END
JSCSV:BufferStatus       ITEMIZE(0),PRE()
NotLoaded                  EQUATE
Loaded                     EQUATE
                         END                
JSCSVColumnDataGroupType GROUP,TYPE    !A structure for accessing the individual cells of data
Data                       &STRING     !A reference to the actual data for a cell
                         END           !              
JSCSVColumnInfoGroupType GROUP,TYPE    !A structure for accessing the info of individual cells of data
Addr                       LONG        !Data Address
Len                        LONG        !Length
                         END           !              

JSCSVParseClass          CLASS,TYPE,MODULE('JSCSVParseClass.clw'),LINK('JSCSVParseClass.clw',_ABCLinkMode_),DLL(_ABCDLLMode_)
!Public Properties  ---------------------------------------------------------------------------------------------------------
Popup                      &PopupClass                                !Reference to an ABC Popup
!Public Methods     ---------------------------------------------------------------------------------------------------------
AdjustColumnWidth          PROCEDURE(LONG pColumnp,LONG pMaxWidth=300)!Attempt to adjust columns to data width
Construct                  PROCEDURE                                  !Automatic Constructor
Destruct                   PROCEDURE,VIRTUAL                          !Automatic Destructor
GenerateFileDef            PROCEDURE(<STRING pLabel>,<STRING pFileName>),STRING !Generate a FILE declaration for the loaded CSV
GenerateClarionStructure   PROCEDURE(STRING pLabel,STRING pType,<STRING pPre>),STRING !Generate a Clarion structure, such as RECORD, GROUP, or QUEUE.
GetDataName                PROCEDURE(LONG pData),STRING               !Get Text name of Separators/LineEndings
GetCellLen                 PROCEDURE(LONG pRow,LONG pColumn),LONG     !Retrieves the length of the data contained in a specific row/column
GetCellValue               PROCEDURE(LONG pRow,LONG pColumn),*STRING  !Retrieves the data contained in a specific row/column
GetColumnCount             PROCEDURE(),LONG                           !Retrieves the number of columns
GetColumnLabel             PROCEDURE(LONG pColumn,BYTE pForClarion=FALSE),STRING!Retrieves the name of the specified column.
GetColumnNumber            PROCEDURE(STRING pName),LONG
GetBufferSize              PROCEDURE(),LONG                           !Gets the size of the buffer
GetBufferStatus            PROCEDURE(),LONG                           !Gets status of the buffer (JSCSV:BufferStatus:Loaded or NotLoaded)
GetColumnLen               PROCEDURE(LONG pColumn,LONG pLimit=0,LONG pType=JSCSV:ColumnLen:Max),LONG!Gets the maximum or total width of a column
GetMaximumColumnTextWidth  PROCEDURE(LONG pColumn,LONG pLimit=0),LONG !Gets the maximum width of a listbox column
GetMemoryUsed              PROCEDURE(LONG pWhich),LONG
GetRowCount                PROCEDURE(),LONG                           !RECORDS(SELF.Q)
GetRowLen                  PROCEDURE(LONG pRow,LONG pLimit=0,LONG pType=JSCSV:ColumnLen:Max),LONG   !Gets the maximum or total width of a the cells of a row
LoadBuffer                 PROCEDURE(*STRING pBuffer)                 !Assigns the buffer of CSV data
LoadFile                   PROCEDURE(STRING pFileName),LONG,PROC      !Load a file
SetColumnLabel             PROCEDURE(LONG pColumn,STRING pLabel),LONG,VIRTUAL,PROC!Virtual method gets called when setting the labels of the columns, in case you want to customize label
SetFileSpecs               PROCEDURE(<STRING pSeparator>,<STRING pLineEnding>,<LONG pFlags>)        !Set the CSV separator, line endings, or flags
SetKnownDelimiters         PROCEDURE(STRING pDelimiters)              !Set which delimiters to search for in auto-detect
SetListBox                 PROCEDURE(LONG pFEQ,<WINDOW pOwnerWindow>) !Sets the listbox used to view the data.
SetLoading                 PROCEDURE                                  !Lets the VLB proc know not to do its thing while buffer is in flux
TakeEvent                  PROCEDURE(LONG pEvent),LONG,PROC,VIRTUAL
TakeProgress               PROCEDURE(LONG pProgressPct,LONG pProgress,LONG pRows),VIRTUAL
!Private Properties ---------------------------------------------------------------------------------------------------------
ColumnCount                LONG,PRIVATE                               !Number of columns - Calculated in InitColumns 
DataChanged                BYTE,PRIVATE                               !Data has changed
EventsRegistered           LONG,PRIVATE                               !Are the events currently registered?
FEQ                        LONG,PRIVATE                               !Listbox FEQ
Win                        &WINDOW,PRIVATE                            !Owner Window
Flags                      LONG,PRIVATE                               !Bitwise flags - such as JSCSV:FirstRowIsLabels
KnownDelimiters            &STRING,PRIVATE                            !Known column separators for detection
RowCount                   LONG,PRIVATE                               !Number of rows    - Calculated in ParseRows
Separator                  CSTRING(5),PRIVATE                         !Defaults to comma separated - Use SetFileSpecs() to change.
LineEnding                 CSTRING(11),PRIVATE                        !Defaults to CRLF (<13,10>) - Use SetFileSpecs() to change.
Buffer                     &STRING,PRIVATE                            !A reference to the entire contents of CSV file
RefBuffer                  &STRING,PRIVATE                            !A reference to a string that contains individual references to all of the cells.
ColumnDefQ                 &JSCSVColumnDefQueueType,PRIVATE           !A queue of column names
Q                          &JSCSVQ,PRIVATE                            !To manage the rows of data
QPointer                   LONG,PRIVATE                               !So we don't keep re-getting the queue for every column
Len                        LONG,PRIVATE                               !Number of bytes in the buffer (the contents of the CSV file)
rDummy                     &STRING,PRIVATE                            !Null Dummy String
SS                         &SystemStringClass,PRIVATE                 !Mostly used just to load a file into the buffer. 
BufferStatus               LONG,PRIVATE                             !
                         END

_EndOfInclude_
