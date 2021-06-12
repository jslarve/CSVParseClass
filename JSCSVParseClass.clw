  MEMBER

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


JSCSVQ  QUEUE,TYPE
Len       LONG    !Size of the &STRING in the Line column, so we don't have to keep calling LEN().
Line     &STRING  !A direct reference to a particular row of the original CSV buffer (not new'd)
Columns  &STRING  !A pseudo record of &STRING references that represent the data inside the Line column above. This reference is part of SELF.Refbuffer.
        END
   
   INCLUDE('JSCSVParseClass.inc'),ONCE 
   INCLUDE('KEYCODES.CLW'),ONCE

  MAP
    JSCSVGetTempFileAndPathName(),STRING
    JSCSVDetectLineEnding(*STRING pBuffer,LONG pMaxBytes=0FFFFh,<STRING pDefault>),STRING
    MODULE('')
      JSCSVGetTempPath(ULONG,*CSTRING),RAW,ULONG,PASCAL,NAME('GetTempPathA'),DLL(1) 
      JSCSVGetTempFilename(*CSTRING,*CSTRING,ULONG,*CSTRING),RAW,ULONG,PASCAL,NAME('GetTempFileNameA'),DLL(1)
    END
     
     !Private Methods
    InitColumns     (JSCSVParseClass pSELF),PRIVATE                         !Count columns and prepare data area to receive the references it will house
    ParseColumns    (JSCSVParseClass pSELF),PRIVATE                         !Parses a row of CSV columns into a string of string references
    ParseRows       (JSCSVParseClass pSELF),LONG,PRIVATE,PROC               !Parses the buffer into rows of data 
    RegisterEvents  (JSCSVParseClass pSELF),PRIVATE                         !For use with the listbox
    RegisterHandler (JSCSVParseClass pSELF),BYTE,PRIVATE                    !Used for REGISTER()
    SetElementRef   (JSCSVParseClass pSELF,LONG pColumn,*STRING pDataToAssign),PRIVATE !Sets a reference for the cell of data at the current row, specified column.
    SetFormatString (JSCSVParseClass pSELF),PRIVATE                         !Sets the format string, depending on number of columns
    UnRegisterEvents(JSCSVParseClass pSELF),PRIVATE                         !Unregister previously registered events     
    VLBproc         (JSCSVParseClass pSELF,LONG xRow, SHORT xCol),STRING,PRIVATE !Virtual Listbox Procedure, used to display data
  END

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Adjusts the width of a column</summary>
!!! <param name="pColumn">Column where data is located</param>
!!! <param name="pCharacterWidth">Width per character</param>
!======================================================================================================================================================
JSCSVParseClass.AdjustColumnWidth  PROCEDURE(LONG pColumn,LONG pMaxWidth=300)
CurColumn LONG
Pixels    BYTE
  CODE

!This is a "quick and dirty" column adjustment because we already have the length of each cell stored as part of the &STRING reference.
!If you wanted to "really" adjust the column widths correctly, you'd want to get the actual width of the text with the current (proportional) font info for every cell.
!Instead, we're just getting the width of the cell with the most text, which may or may not really be the widest because of proportional font.

  Pixels = SELF.Win{PROP:Pixels}
  SELF.Win{PROP:Pixels} = TRUE
  IF pColumn > 0  !Specified a specific column
    CurColumn = pColumn
    DO AdjustColumn
  ELSE
    LOOP CurColumn = 1 TO SELF.GetColumnCount() 
      DO AdjustColumn
    END
  END  
  SELF.Win{PROP:Pixels} = Pixels
  
AdjustColumn ROUTINE
  
  IF SELF.Win $ SELF.FEQ {PROPLIST:Exists,CurColumn}
    SELF.Win  $ SELF.FEQ {PROPLIST:width,CurColumn} = SELF.GetMaximumColumnTextWidth(CurColumn,pMaxWidth) 
  END
   
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Automatic constructor</summary>
!======================================================================================================================================================
JSCSVParseClass.Construct PROCEDURE

  CODE
  
  SELF.Buffer     &= NULL
  SELF.RefBuffer  &= NULL
  SELF.Q          &= NEW JSCSVQ
  SELF.SS         &= NEW SystemStringClass
  SELF.rDummy     &= NULL
  SELF.SetFileSpecs(',','<13,10>',JSCSV:FirstRowIsLabels) !Default to Comma separated, CRLF
  SELF.SetKnownDelimiters(',<9>:;| ')
  SELF.ColumnCount = 1 
  SELF.RowCount    = 0 
  SELF.Popup      &= NEW PopupClass
  SELF.Popup.Init()
  SELF.Popup.AddItem('Copy Cell','COPYCELL')
  SELF.Popup.AddItem('Copy Column (Excel Friendly)','COPYCOLUMN')
  SELF.Popup.AddItem('Copy Row (Excel Friendly)','COPYROW')
  SELF.Popup.AddItem('Copy Row (Original)','COPYROWCSV')
  SELF.Popup.AddItem('Export to XML, JSON, etc.','EXPORT') 
  
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Automatic destructor</summary>
!======================================================================================================================================================
JSCSVParseClass.Destruct      PROCEDURE

  CODE
  
  DISPOSE(SELF.RefBuffer)
  DISPOSE(SELF.Q)
  SELF.Popup.Kill
  DISPOSE(SELF.Popup)
  DISPOSE(SELF.SS)
  DISPOSE(SELF.KnownDelimiters)
  !SELF.Buffer is either owned by SELF.SS OR it could be externally passed in, so we don't dispose that.

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the name of the passed data</summary>
!!! <param name="pData">Data to be named. 1=Separator, 2=LineEnding</param>
!!! <returns>STRING Name</returns>
!======================================================================================================================================================
JSCSVParseClass.GetDataName PROCEDURE(LONG pData)!,STRING
ReturnVal CSTRING(21)

  CODE
  
  CASE pData
  OF 1
    CASE SELF.Separator
    OF ','
      ReturnVal = 'Comma'
    OF '<9>'
      ReturnVal = 'Tab'
    OF ':'
      ReturnVal = 'Colon'
    OF ';'
      ReturnVal = 'SemiColon'
    OF '|'
      ReturnVal = 'Pipe'
    OF ' '
      ReturnVal = 'Space'
    ELSE
      ReturnVal = pData
    END
  OF 2  
    CASE SELF.LineEnding
    OF '<13,10>'
      ReturnVal = 'Windows'
    OF '<10>'
      ReturnVal = 'UNIX'
    OF '<13>'
      ReturnVal = 'Mac'
    ELSE
      ReturnVal = pData
    END
  END  
  RETURN ReturnVal

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the length of the value contained in a cell</summary>
!!! <param name="pRow">Row where data is located</param>
!!! <param name="pColumn">Column where data is located</param>
!!! <returns>LONG value</returns>
!======================================================================================================================================================
JSCSVParseClass.GetCellLen   PROCEDURE(LONG pRow,LONG pColumn)!,LONG    !Retrieves the length of the data contained in a specific row/column
rLen        &LONG

  CODE

  IF pRow <> SELF.QPointer
    GET(SELF.Q,pRow + CHOOSE(BAND(SELF.Flags,JSCSV:FirstRowIsLabels)))
    IF ERRORCODE()
      RETURN 0
    END
    SELF.QPointer = pRow
  END  
  CASE pColumn
  OF 1 TO SELF.ColumnCount 
    IF NOT SELF.Q.Columns &= NULL
      rLen &= ((ADDRESS(SELF.Q.Columns) + ((pColumn-1) * 8)) + 4)
      IF NOT rLen &= NULL
        RETURN rLen
      END
    END
  END
  RETURN 0


!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the value contained in a cell</summary>
!!! <param name="pRow">Row where data is located</param>
!!! <param name="pColumn">Column where data is located</param>
!!! <returns>A STRING reference</returns>
!======================================================================================================================================================
JSCSVParseClass.GetCellValue PROCEDURE(LONG pRow,LONG pColumn)!,*STRING
rColumnData &JSCSVColumnDataGroupType

  CODE

  IF pRow <> SELF.QPointer
    GET(SELF.Q,pRow + CHOOSE(BAND(SELF.Flags,JSCSV:FirstRowIsLabels)))
    IF ERRORCODE()
      RETURN SELF.rDummy
    END
    SELF.QPointer = pRow
  END  
  CASE pColumn
  OF 1 TO SELF.ColumnCount 
    IF NOT SELF.Q.Columns &= NULL
      rColumnData &= (ADDRESS(SELF.Q.Columns) + ((pColumn-1) * 8))
      IF NOT rColumnData &= NULL
        RETURN rColumnData.Data 
      END  
    END
  END
  RETURN SELF.rDummy

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the number of columns</summary>
!!! <returns>Number of columns</returns>
!======================================================================================================================================================
JSCSVParseClass.GetColumnCount PROCEDURE()!,LONG
  
  CODE
  
  RETURN SELF.ColumnCount

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the label of a column</summary>
!!! <param name="pColumn">Column where label is located</param>
!!! <returns>A STRING</returns>
!======================================================================================================================================================
JSCSVParseClass.GetColumnLabel        PROCEDURE(LONG pColumn)!,STRING
ColumnDef     &JSCSVColumnDefGroupType
  
  CODE
  
  ColumnDef &= ADDRESS(SELF.ColumnDefBuffer) + (pColumn * SIZE(JSCSVColumnDefGroupType))
  RETURN CLIP(ColumnDef.Name)
 
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the number of bytes contained in the buffer (this is the size of the CSV file)</summary>
!!! <returns>A LONG containing the value</returns>
!======================================================================================================================================================
JSCSVParseClass.GetBufferSize         PROCEDURE()!,LONG

  CODE
  
  RETURN SELF.Len

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the current status of the buffer (JSCSV:BufferStatus:Loaded or JSCSV:BufferStatus:NotLoaded) </summary>
!!! <returns>A LONG containing the status</returns>
!======================================================================================================================================================
JSCSVParseClass.GetBufferStatus       PROCEDURE()!,LONG   

  CODE
  
  RETURN SELF.BufferStatus

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the maximum Len of a column</summary>
!!! <param name="pColumn">Column which column to test</param>
!!! <param name="pLimit">Maximum Len to test</param>
!!! <param name="pType">Maximum or Total</param>
!!! <returns>A LONG</returns>
!======================================================================================================================================================
JSCSVParseClass.GetColumnLen   PROCEDURE(LONG pColumn,LONG pLimit=0,LONG pType=JSCSV:ColumnLen:Max)!,LONG!Gets the maximum or total width of a column
Ndx1       LONG
CurLen     LONG
MaxLen     LONG

  CODE

  MaxLen = 0
  LOOP Ndx1 = 1 TO SELF.GetRowCount()
    CASE pType
    OF JSCSV:ColumnLen:Max
      CurLen = SELF.GetCellLen(Ndx1,pColumn)
      IF CurLen > MaxLen
        MaxLen = CurLen
      END
      IF pLimit
        IF MaxLen => pLimit
          BREAK
        END
      END
    OF JSCSV:ColumnLen:Total
      MaxLen += SELF.GetCellLen(Ndx1,pColumn)
    END
  END
  RETURN MaxLen


!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the maximum width of a column</summary>
!!! <param name="pColumn">Column which column to test</param>
!!! <param name="pLimit">Maximum Len to test</param>
!!! <returns>A LONG</returns>
!======================================================================================================================================================
JSCSVParseClass.GetMaximumColumnTextWidth PROCEDURE(LONG pColumn,LONG pLimit=0)!,LONG!Gets the maximum width of a listbox column
Ndx1        LONG
CurLen      LONG
MaxLen      LONG
MaxRowNdx   LONG
StringFEQ   LONG
CurTarget   &WINDOW
HeaderText  CSTRING(256)
ColumnWidth LONG

  CODE

  IF SELF.Win &= NULL OR SELF.FEQ = 0
    RETURN 0
  END
  CurTarget &= SYSTEM{PROP:Target}
  SetTarget(SELF.Win)
  HeaderText =  SELF.FEQ{PROPList:Header,pColumn} & ' (' & pColumn & ')'
  MaxLen = LEN(HeaderText)
  LOOP Ndx1 = 1 TO SELF.GetRowCount()    
    CurLen = SELF.GetCellLen(Ndx1,pColumn)
    IF CurLen > MaxLen
      MaxRowNdx = Ndx1
      MaxLen    = CurLen
      HeaderText = SELF.GetCellValue(MaxRowNdx,pColumn) 
    END
  END
  StringFEQ                 =  CREATE(0,CREATE:String,0)
  StringFEQ{PROP:FontName } =  SELF.FEQ{PROP:FontName}
  StringFEQ{PROP:FontSize } =  SELF.FEQ{PROP:FontSize}
  StringFEQ{PROP:FontStyle} =  SELF.FEQ{PROP:FontStyle}
  StringFEQ{PROP:Text     } =  HeaderText
  DISPLAY(StringFEQ)
  ColumnWidth               =  StringFEQ{PROP:Width} + 10
  IF pLimit
    IF MaxLen => pLimit
       MaxLen = pLimit
    END
  END
  DESTROY(StringFEQ)
  SETTARGET
  RETURN ColumnWidth
  
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the number of rows contained in the CSV file</summary>
!!! <returns>A LONG containing the value</returns>
!======================================================================================================================================================
JSCSVParseClass.GetRowCount  PROCEDURE()!,LONG

  CODE
  
  RETURN RECORDS(SELF.Q) - CHOOSE(BAND(SELF.Flags,JSCSV:FirstRowIsLabels))

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Retrieve the maximum Len of the cells of a row</summary>
!!! <param name="pRow">Row which column to test</param>
!!! <param name="pLimit">Maximum Len to test</param>
!!! <param name="pType">Maximum or Total</param>
!!! <returns>A LONG</returns>
!======================================================================================================================================================
JSCSVParseClass.GetRowLen   PROCEDURE(LONG pRow,LONG pLimit=0,LONG pType=JSCSV:ColumnLen:Max)!,LONG!Gets the maximum or total width of the cells of a row
Ndx1       LONG
CurLen     LONG
MaxLen     LONG

  CODE

  MaxLen = 0
  LOOP Ndx1 = 1 TO SELF.GetColumnCount()
    CASE pType
    OF JSCSV:ColumnLen:Max
      CurLen = SELF.GetCellLen(pRow,Ndx1)
      IF CurLen > MaxLen
        MaxLen = CurLen
      END
      IF pLimit
        IF MaxLen => pLimit
          BREAK
        END
      END
    OF JSCSV:ColumnLen:Total
      MaxLen += SELF.GetCellLen(pRow,Ndx1)
    END
  END
  RETURN MaxLen

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Load a Delimiter Separated File</summary>
!!! <param name="pFileName">Name of the Delimiter Separated File</param>
!!! <returns>Bytes loaded</returns>
!======================================================================================================================================================
JSCSVParseClass.LoadFile     PROCEDURE(STRING pFileName)!,LONG          !Load a file
DummySS   SystemStringClass !To trick SystemString class into giving up its FILEBUFFERS=
DummyPath CSTRING(FILE:MaxFilePath+1)
  CODE

  IF NOT EXISTS(CLIP(pFileName))
    RETURN 0
  END
  SELF.SetLoading
  SELF.SS.FromFile(pFileName)
  DummyPath = JSCSVGetTempFileAndPathName() !Create a zero byte dummy file
  IF EXISTS(DummyPath)                      !Make sure it exists
    DummySS.FromFile(DummyPath)             !Load it into a SystemString object
    REMOVE(DummyPath)                       !Delete it
  END                                       !This frees up the memory created by the previous "FILEBUFFERS=xxx" in SystemString.FromFile()
  SELF.LoadBuffer(SELF.SS.GetStringRef())
  RETURN SELF.SS.Length()

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Initialize the class with the CSV data and parse it for use</summary>
!!! <returns>No return value</returns>
!======================================================================================================================================================
JSCSVParseClass.LoadBuffer   PROCEDURE(*STRING pBuffer)

  CODE 
  
  FREE(SELF.Q)
  SELF.ColumnCount = 0
  SELF.DataChanged = TRUE
  SELF.BufferStatus = JSCSV:BufferStatus:NotLoaded
  SELF.FEQ{PROP:Format} = '31L(2)|M~Loading......~'
  IF SELF.FEQ
    DISPLAY(SELF.FEQ)
  END
  DISPOSE(SELF.RefBuffer)
  SELF.Buffer &= pBuffer
  SELF.Len     = SIZE(pBuffer)
  IF SELF.ParseRows()
    SELF.ParseColumns()
  ELSE
    !TODO - Need to let user know that the line endings are probably wrong.
  END
  SELF.SetFormatString  
  SELF.BufferStatus = JSCSV:BufferStatus:Loaded
  SELF.DataChanged = TRUE
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Private method for preparing data</summary>
!!! <returns>No Return Value</returns>
!======================================================================================================================================================
JSCSVParseClass.InitColumns  PROCEDURE !,PRIVATE
Ndx1          LONG,AUTO
SaveNdx       LONG,AUTO
SeparatorLen  LONG,AUTO
ColumnCount   LONG,AUTO
InsideQuote   BYTE
ColumnDef     &JSCSVColumnDefGroupType
ColumnNameQ   QUEUE
Label           STRING(60)
              END

  CODE
  
  SeparatorLen = LEN(SELF.Separator)
  DISPOSE(SELF.RefBuffer)
  DISPOSE(SELF.ColumnDefBuffer)

  DO CountColumns

  IF SELF.ColumnCount > 0
     SELF.ColumnDefBuffer &= NEW STRING(SELF.ColumnCount * SIZE(JSCSVColumnDefGroupType) )
     CLEAR(SELF.ColumnDefBuffer,-1)
     SELF.RefBuffer &= NEW STRING(SELF.ColumnCount * 8 * RECORDS(SELF.Q))
     CLEAR(SELF.RefBuffer,-1)
     DO LabelColumns
  END

LabelColumns ROUTINE

  LOOP Ndx1 = 0 TO SELF.ColumnCount - 1
    ColumnDef &= ADDRESS(SELF.ColumnDefBuffer) + (Ndx1 * SIZE(JSCSVColumnDefGroupType))
    IF NOT BAND(SELF.Flags,JSCSV:FirstRowIsLabels)
      ColumnDef.Name = 'Column ' & Ndx1 + 1
    ELSE
      GET(ColumnNameQ,Ndx1+1)
      ColumnDef.Name = ColumnNameQ.Label
    END
    SELF.SetColumnLabel(Ndx1+1,ColumnDef)
  END

CountColumns ROUTINE

  InsideQuote = FALSE
  ColumnCount = 1
  GET(SELF.Q,1)   
  SELF.QPointer = POINTER(SELF.Q) 
  IF SELF.Q.Line &= NULL
    EXIT
  END
  SaveNdx   = 1
  LOOP Ndx1 = SaveNdx TO SELF.Q.Len 
    IF Ndx1 > SELF.Q.Len 
      BREAK
    END
    IF SELF.Q.Line[Ndx1] = '"'
      IF InsideQuote
        ColumnNameQ.Label = SELF.Q.Line[SaveNdx : Ndx1-1]
        ADD(ColumnNameQ)
        SaveNdx      = Ndx1 + SeparatorLen + 1
        IF Ndx1 < SELF.Q.Len
          Ndx1        += 1
          ColumnCount += 1
        END
        InsideQuote  = FALSE
        CYCLE
      ELSE
        InsideQuote  = TRUE
        SaveNdx      = Ndx1 + 1
        CYCLE
      END
    END
    IF NOT InsideQuote
      IF (SELF.Q.Line[Ndx1 : Ndx1 + SeparatorLen - 1] = SELF.Separator) 
        ColumnNameQ.Label = SELF.Q.Line[SaveNdx : Ndx1 - 1]
        ADD(ColumnNameQ)
        SaveNdx      = Ndx1 + 1
        ColumnCount += 1
      ELSIF (Ndx1 => SELF.Q.Len)  
        ColumnNameQ.Label = SELF.Q.Line[SaveNdx : Ndx1]
        ADD(ColumnNameQ)
        BREAK
      END
    END  
  END
  SELF.ColumnCount = ColumnCount

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Private method for parsing columns</summary>
!!! <returns>No Return Value</returns>
!======================================================================================================================================================
JSCSVParseClass.ParseColumns  PROCEDURE !,PRIVATE
Ndx1          LONG,AUTO !Generic counter
LineNdx1      LONG,AUTO !Line counter
SaveNdx       LONG,AUTO !Counter 
SeparatorLen  LONG,AUTO !Length of a the separator property
ColumnNdx     LONG,AUTO !Column counter
InsideQuote   BYTE,AUTO !Flag determined by whether we are currently inside a quoted column as we parse
ColumnAddress LONG,AUTO !The ADDRESS() of the current SELF.Q.Columns buffer
NullString    &STRING   !Just a NULL string to return
Recs          LONG,AUTO !Number of records in queue
ProgressTic   LONG      !Frequency of progress display. This area needs work.

  CODE

  NullString   &= NULL  
  Recs          = RECORDS(SELF.Q)
  IF NOT Recs
    RETURN
  END
  IF SELF.Separator[1] = '<1>'
    DO DetectSeparator
  END
  SeparatorLen = LEN(SELF.Separator)
  IF SeparatorLen < 1
    RETURN
  END
  IF SELF.RefBuffer &= NULL
    SELF.InitColumns
  END
  ProgressTic   = Recs / 50
  IF NOT ProgressTic
    ProgressTic = Recs
  END
  LOOP LineNdx1 = 1 TO Recs
    GET(SELF.Q,LineNdx1)
    SELF.QPointer   = LineNdx1
    SELF.Q.Columns &= ADDRESS(SELF.RefBuffer) + ((LineNdx1-1) * 8 * SELF.ColumnCount) & ':' & 8 * SELF.ColumnCount
    ColumnAddress   = ADDRESS(SELF.Q.Columns)
    PUT(SELF.Q)
    SaveNdx     = 1
    ColumnNdx   = 0
    InsideQuote = FALSE
    LOOP Ndx1 = SaveNdx TO SELF.Q.Len
      IF ColumnNdx > SELF.ColumnCount 
        BREAK
      END
      IF Ndx1 <= SELF.Q.Len
        IF SELF.Q.Line[Ndx1] = '"'
          IF InsideQuote
            SELF.SetElementRef(ColumnNdx,SELF.Q.Line[SaveNdx : Ndx1-1])
            SaveNdx = Ndx1+SeparatorLen+1
            InsideQuote = FALSE
            ColumnNdx += 1
            Ndx1 += 1
            CYCLE
          ELSE
            InsideQuote = TRUE
            SaveNdx = Ndx1 + 1
            CYCLE
          END
        END
      ELSE 
        InsideQuote = FALSE
      END
      IF NOT InsideQuote
        IF SELF.Q.Line[Ndx1 : Ndx1 + SeparatorLen-1] = SELF.Separator 
          SELF.SetElementRef(ColumnNdx,SELF.Q.Line[SaveNdx : Ndx1-1])
          SaveNdx = Ndx1 + 1
          ColumnNdx += 1
        ELSIF Ndx1 = SELF.Q.Len 
          SELF.SetElementRef(ColumnNdx,SELF.Q.Line[SaveNdx : SELF.Q.Len])
          ColumnNdx += 1
          BREAK
        END
      END  
    END
    IF NOT LineNdx1 % ProgressTic
       SELF.TakeProgress(LineNdx1 / SELF.GetRowCount() * 100, LineNdx1, SELF.GetRowCount() )
       YIELD()
    END
  END

DetectSeparator ROUTINE

  DATA
EligibleChars BYTE,DIM(255)
CharCount     USHORT,DIM(2,255) 
RowCount      USHORT
  CODE
  !This still needs work, but it appears to be operational :)
  !Basically counts each occurrence of eligible characters within 2 rows and checks to see if the numbers match.
  !It has worked on everything I tested it with, but needs some re-vamping
  IF NOT Recs
    EXIT
  END
  CLEAR(EligibleChars)
  Clear(CharCount)
  LOOP Ndx1 = 1 TO LEN(SELF.KnownDelimiters)                !Looping through known delimiters 
    EligibleChars[ VAL(SELF.KnownDelimiters[Ndx1]) ] = TRUE !and adding them to the lookup table of EligibleChars
  END
  IF Recs => 2
    RowCount = 2
  ELSE
    RowCount = Recs
  END 
  LOOP LineNdx1 = 1 TO RowCount
    GET(SELF.Q,LineNdx1)
    LOOP Ndx1 = 1 TO SELF.Q.Len
      IF SELF.Q.Line[Ndx1] = '"'
        IF InsideQuote
          InsideQuote = FALSE
          CYCLE
        ELSE
          InsideQuote = TRUE
          CYCLE
        END
      END
      IF NOT InsideQuote
        IF EligibleChars[VAL(SELF.Q.Line[Ndx1])]
          CharCount[LineNdx1,VAL(SELF.Q.Line[Ndx1])] += 1
        END  
      END
    END
  END
 
  LOOP Ndx1 = 1 TO LEN(SELF.KnownDelimiters)
    IF CharCount[1, VAL(SELF.KnownDelimiters[Ndx1]) ] 
      IF RowCount = 2
        IF CharCount[1,VAL(SELF.KnownDelimiters[Ndx1])] = CharCount[2,VAL(SELF.KnownDelimiters[Ndx1])]
          SELF.Separator   = SELF.KnownDelimiters[Ndx1]
          SELF.ColumnCount = CharCount[1,VAL(SELF.KnownDelimiters[Ndx1])] + 1
          BREAK
        END
      ELSE !Do less validation because we only have one row
        SELF.Separator   = SELF.KnownDelimiters[Ndx1]
        SELF.ColumnCount = CharCount[1,VAL(SELF.KnownDelimiters[Ndx1])] + 1
        BREAK
      END
    END
  END
  InsideQuote = FALSE
  LineNdx1 = 0
  Ndx1 = 0
  
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Private method for parsing rows</summary>
!!! <returns>Number of rows</returns>
!======================================================================================================================================================
JSCSVParseClass.ParseRows     PROCEDURE !,LONG
Ndx1          LONG,AUTO
SaveNdx       LONG,AUTO
LineEndingLen LONG,AUTO

  CODE   

  IF SELF.Buffer &= NULL
    RETURN 0
  END
  IF SELF.LineEnding = ''
    SELF.LineEnding = JSCSVDetectLineEnding(SELF.Buffer)
    IF SELF.LineEnding = ''
      RETURN 0
    END
  END
  LineEndingLen = LEN(SELF.LineEnding)
  IF SELF.GetBufferSize() < LineEndingLen
    RETURN 0
  END
  SaveNdx = 1
  LOOP 
    Ndx1 = INSTRING(SELF.LineEnding,SELF.Buffer,1,SaveNdx)
    IF NOT Ndx1
      BREAK
    END
    SELF.Q.Columns &=  NULL
    SELF.Q.Line    &=  SELF.Buffer[SaveNdx : Ndx1-1]
    SELF.Q.Len      =  LEN(SELF.Q.Line)
    ADD(SELF.Q)
    SaveNdx         =  Ndx1 + LineEndingLen
  END
  RETURN RECORDS(SELF.Q)
  
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Private method for setting reference values</summary>
!!! <returns>no return value</returns>
!======================================================================================================================================================
JSCSVParseClass.SetElementRef   PROCEDURE(LONG pColumn,*STRING pDataToAssign)
rColumnData &JSCSVColumnDataGroupType

  CODE

!This method casts a JSCSVColumnDataGroupType structure over a specific offset of the column data buffer. 
!It then assigns a reference to the ACTUAL data to the rColumnData.Data &STRING reference.
!It's functionally the same as declaring one thing OVER another, except it allows you to create dynamic 
!pseudo records and do some groovy stuff with it. It doesn't have to just be &STRING, either. 
!You could use this technique for pretty much any data type or reference and generate dynamic "records" on the fly.
!But since CSV is just text, &STRING works just fine for this purpose.

  CASE pColumn
  OF 0 TO SELF.ColumnCount - 1
    IF NOT SELF.Q.Columns &= NULL
      rColumnData &= (ADDRESS(SELF.Q.Columns) + (pColumn * 8))
      IF NOT rColumnData &= NULL
        rColumnData.Data &= pDataToAssign
      END  
    END
  END

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Private method for setting listbox format</summary>
!!! <returns>no return value</returns>
!======================================================================================================================================================
JSCSVParseClass.SetFormatString PROCEDURE
ColumnText EQUATE('60L(2)|M~ ~C')
SS         SystemStringClass
Ndx1       LONG

   CODE
   
  IF SELF.Win &= NULL
    RETURN
  END
  IF SELF.FEQ = 0
    RETURN
  END   
  IF NOT SELF.ColumnCount
    SELF.Win $ SELF.FEQ{PROP:Format} =  '60L(2)|M~NO COLUMNS TO LOAD - Check separator & line endings.~L'
    RETURN
  END
  LOOP Ndx1 = 1 TO SELF.ColumnCount + 1 !Add 1 extra dummy column so the last column sizes OK
    SS.Append(ColumnText)
  END
  SELF.Win $ SELF.FEQ{PROP:Format} = SS.ToString()
  LOOP Ndx1 = 1 TO SELF.ColumnCount 
    SELF.Win $ SELF.FEQ{PROPLIST:Header,Ndx1}  = CLIP(LEFT(SELF.GetColumnLabel(Ndx1-1))) & ' (' & Ndx1 & ')'
  END
  !To do - support multiple column properties, detect whether numeric and thus s/b right-justified, and allow user choice for various properties.
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Virtual Method</summary>
!======================================================================================================================================================
JSCSVParseClass.SetColumnLabel   PROCEDURE(LONG pColumn,*JSCSVColumnDefGroupType pColumnDef)!,VIRTUAL

   CODE
   
   !Virtual Method - If you derive it, you can override pColumnDef.Name and other properties in the future.
  
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Sets the CSV file specs</summary>
!!! <param name="pSeparator">Separator character (defaults to comma)</param>
!!! <param name="pLineEnding">Line Ending (defaults to CRLF)</param>
!!! <param name="pFlags">Flags</param>
!!! <returns>Nothing</returns>
!======================================================================================================================================================
JSCSVParseClass.SetFileSpecs     PROCEDURE(<STRING pSeparator>,<STRING pLineEnding>,<LONG pFlags>)
  
  CODE

  IF NOT OMITTED(pSeparator)    
    SELF.Separator  = pSeparator   
  END   

  IF NOT OMITTED(pLineEnding)    
    SELF.LineEnding  = pLineEnding   
  END   

  IF NOT OMITTED(pFlags)
    SELF.Flags = pFlags
  END

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Sets the eligible separators for auto-detecting column separators</summary>
!!! <param name="pDelimiters">A string of unique characters that could be used as a separator</param>
!!! <returns>Nothing</returns>
!======================================================================================================================================================
JSCSVParseClass.SetKnownDelimiters PROCEDURE(STRING pDelimiters)

   CODE
   
   IF NOT LEN(pDelimiters)
     RETURN
   END
   DISPOSE(SELF.KnownDelimiters)
   SELF.KnownDelimiters &= NEW STRING(LEN(pDelimiters))
   SELF.KnownDelimiters = pDelimiters
   
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Sets the listbox for use by class</summary>
!!! <param name="pFEQ">FEQ of the listbox</param>
!!! <param name="pOwnerWindow">OPTIONAL - Window that owns the listbox</param>
!!! <returns>Nothing</returns>
!======================================================================================================================================================
JSCSVParseClass.SetListBox            PROCEDURE(LONG pFEQ,<WINDOW pOwnerWindow>)
rW  &WINDOW
Ndx  LONG

  CODE

  SELF.FEQ = pFEQ

  IF NOT OMITTED(pOwnerWindow)
    rW &= pOwnerWindow
  ELSE
    rW &= SYSTEM{PROP:Target}
  END
  SELF.Win &= rW
  
  IF NOT rW $ pFEQ{PROP:Type} = CREATE:list
    MESSAGE('Must past a Clarion listbox FEQ')
    RETURN
  END
  SELF.Win $ SELF.FEQ{PROP:FROM}     = SELF.Q
  SELF.Win $ SELF.FEQ{PROP:VLBVal}   = ADDRESS(SELF)
  SELF.Win $ SELF.FEQ{PROP:VLBProc}  = ADDRESS(SELF.VLBproc)
  SELF.Win $ SELF.FEQ{PROP:Alrt,255} = CtrlC
  SELF.Win $ SELF.FEQ{PROP:Alrt,255} = MouseRight
  SELF.Win $ SELF.FEQ{PROP:Alrt,255} = MouseLeft2

  SELF.SetFormatString 
  SELF.RegisterEvents
  
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Set loading status to NotLoaded so VLB doesn't try to access buffer</summary>
!======================================================================================================================================================
JSCSVParseClass.SetLoading   PROCEDURE
  
    CODE

  SELF.BufferStatus     =  JSCSV:BufferStatus:NotLoaded
  SELF.ColumnCount      =  0
  SELF.FEQ{PROP:Format} =  '31L(2)|M~Loading......~'
  
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Private method for virtual listbox</summary>
!!! <returns>String containing cell data</returns>
!======================================================================================================================================================
JSCSVParseClass.VLBproc     PROCEDURE(LONG xRow, SHORT xCol)!,STRING,PRIVATE   
ROW:GetRowCount  EQUATE(-1)
ROW:GetColCount  EQUATE(-2)
ROW:IsQChanged   EQUATE(-3)
  
  CODE
    
  CASE xRow
  OF ROW:GetRowCount
    IF SELF.BufferStatus = JSCSV:BufferStatus:NotLoaded
      RETURN 0
    END
    IF NOT RECORDS(SELF.Q)
      RETURN 0
    END
    RETURN RECORDS(SELF.Q) - CHOOSE(BAND(SELF.Flags,JSCSV:FirstRowIsLabels))
  OF ROW:GetColCount
    RETURN SELF.ColumnCount
  OF ROW:IsQChanged 
    IF SELF.DataChanged
      SELF.DataChanged = FALSE
      RETURN TRUE
    END
    RETURN FALSE
  END
  IF SELF.BufferStatus = JSCSV:BufferStatus:Loaded
    RETURN SELF.GetCellValue(xRow,xCol)
  END
  RETURN 'Not Loaded'
    
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Virtual method that gets called when it's time to display progress</summary>
!!! <param name="pProgressPct">Percent completed</param>
!!! <param name="pProgress">Current row</param>
!!! <param name="pRows">Total number of rows to complete</param>
!!! <returns>Nothing</returns>
!======================================================================================================================================================
JSCSVParseClass.TakeProgress  PROCEDURE(LONG pProgressPct,LONG pProgress,LONG pRows)!,VIRTUAL

  CODE
   
   !Virtual Method

!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Registers events for the listbox</summary>
!!! <returns>Nothing</returns>
!======================================================================================================================================================
JSCSVParseClass.RegisterEvents  PROCEDURE

  CODE

  IF NOT SELF.FEQ
    RETURN
  END
  IF SELF.EventsRegistered
    RETURN
  END
  
  REGISTER(EVENT:PreAlertKey, ADDRESS(SELF.RegisterHandler),ADDRESS(SELF),,SELF.FEQ)
  REGISTER(EVENT:AlertKey,    ADDRESS(SELF.RegisterHandler),ADDRESS(SELF),,SELF.FEQ)
  REGISTER(EVENT:CloseWindow, ADDRESS(SELF.RegisterHandler),ADDRESS(SELF))
  SELF.EventsRegistered = TRUE
        
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Handles the registered events</summary>
!!! <returns>Nothing</returns>
!======================================================================================================================================================
JSCSVParseClass.RegisterHandler    PROCEDURE !,BYTE !Used for REGISTER()

  CODE
     
  RETURN CHOOSE(SELF.TakeEvent(EVENT())=1,Level:Notify,Level:Benign)
        
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Handles events</summary>
!!! <param name="pEvent">The EVENT() to handle</param>
!!! <returns>Nothing</returns>
!======================================================================================================================================================
JSCSVParseClass.TakeEvent  PROCEDURE(SIGNED pEvent)
CurColumn   LONG
CurRow      LONG
RowCount    LONG
CurPos      LONG
CurPosEnd   LONG
CurCellLen  LONG
MouseRow    LONG
MouseCol    LONG
WorkString  &STRING
WorkSize    LONG

  CODE

  IF pEvent = EVENT:CloseWindow
    SELF.UnRegisterEvents
    RETURN 0
  END
  
  IF FIELD() <> SELF.FEQ
    RETURN 0
  END

  CASE pEvent
  OF EVENT:AlertKey
    CASE KEYCODE()
    OF CtrlC
      SETCLIPBOARD(SELF.GetCellValue(SELF.Win $ SELF.FEQ{PROP:Selected},SELF.Win $ SELF.FEQ{PROP:Column}))
    OF MouseRight
      MouseCol = SELF.Win $ SELF.FEQ{PROPLIST:MouseDownField} 
      MouseRow = SELF.FEQ{PROPLIST:MouseDownRow}
      CASE MouseCol
      OF 1 TO SELF.ColumnCount
      ELSE 
        RETURN 0
      END
      SELF.Win $ SELF.FEQ{PROP:Selected} = MouseRow
      SELF.Win $ SELF.FEQ{PROP:Column}   = MouseCol
      CASE SELF.Popup.Ask()
      OF 'COPYCELL'
        SETCLIPBOARD(SELF.GetCellValue(SELF.Win $ SELF.FEQ{PROP:Selected},SELF.Win $ SELF.FEQ{PROP:Column}))         
      OF 'COPYCOLUMN'
        DO CopyColumn
      OF 'COPYROW' 
        DO CopyRow
      OF 'COPYROWCSV'
        GET(SELF.Q,SELF.Win $ SELF.FEQ{PROP:Selected}+CHOOSE(BAND(SELF.Flags,JSCSV:FirstRowIsLabels))) 
        SELF.QPointer = POINTER(SELF.Q)         
        SETCLIPBOARD(SELF.Q.Line)
      OF 'EXPORT'
        MESSAGE('Will add this functionality, but not yet.|Had to stop somewhere, but there''s lots of cool things to be done.','Pardon our dust',ICON:Clarion)
      END
    OF MouseLeft2
      IF SELF.Win $ SELF.FEQ{PROPLIST:MouseDownRow} = 0
        CurColumn = SELF.Win $ SELF.FEQ{PROPLIST:MouseDownField}
        SELF.AdjustColumnWidth(CurColumn)   
      END
    END
  END
  
  RETURN 0
  
CopyColumn ROUTINE
  RowCount = SELF.GetRowCount()
  IF NOT RowCount
    EXIT
  END
  CurColumn   = SELF.Win $ SELF.FEQ{PROPList:MouseDownField}
  WorkSize    = SELF.GetColumnLen(CurColumn,,JSCSV:ColumnLen:Total) + (RowCount * 2) - 2
  WorkString &= NEW STRING(WorkSize)
  IF CurColumn
    CurPos = 1
    LOOP curRow = 1 TO RowCount
      CurCellLen = SELF.GetCellLen(CurRow,CurColumn)
      CurPosEnd  = CurPos + CurCellLen + 1
      IF CurPosEnd > WorkSize   
        IF CurPos <= WorkSize     
          CurPosEnd = WorkSize   
        ELSE
          BREAK !If we somehow messed up somewhere :)
        END
      END
      WorkString[CurPos : CurPosEnd] = SELF.GetCellValue(curRow,CurColumn) & '<13,10>' !The last CRLF will be truncated on the last row
      CurPos += (CurCellLen + 2)
    END
  END 
  SETCLIPBOARD(WorkString)
  DISPOSE(WorkString)
  
CopyRow    ROUTINE

  CurRow = SELF.Win $ SELF.FEQ{PROP:Selected}
  IF CurRow
    CurPos = 1
    WorkString &= NEW STRING(SELF.GetRowLen(CurRow,,JSCSV:ColumnLen:Total) + (SELF.GetRowCount() ) - 1)
    LOOP CurColumn = 1 TO SELF.GetColumnCount()
      CurCellLen = SELF.GetCellLen(CurRow,CurColumn)
      WorkString[CurPos : CurPos + CurCellLen + 1] = SELF.GetCellValue(curRow,CurColumn) & CHOOSE(CurColumn < SELF.GetColumnCount(),'<9>','')
      CurPos += (CurCellLen + 1)
    END  
  END
  SETCLIPBOARD(WorkString)
  DISPOSE(WorkString)
    
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Unregisters the registered events</summary>
!!! <returns>Nothing</returns>
!======================================================================================================================================================
JSCSVParseClass.UnRegisterEvents  PROCEDURE

  CODE
       
  IF NOT SELF.EventsRegistered
      RETURN
  END
  
  UNREGISTER(EVENT:PreAlertKey, ADDRESS(SELF.RegisterHandler),ADDRESS(SELF),,SELF.FEQ)
  OMIT('Prevent C11.1 Crash',_C111_)
    UNREGISTER(EVENT:AlertKey,    ADDRESS(SELF.RegisterHandler),ADDRESS(SELF),,SELF.FEQ)
  !Prevent C11.1 Crash
  UNREGISTER(EVENT:CloseWindow, ADDRESS(SELF.RegisterHandler),ADDRESS(SELF))
  SELF.EventsRegistered = FALSE        
              
!-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
JSCSVGetTempFileAndPathName  PROCEDURE
!-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
TempFilename        STRING(259)                                      
lpPathName          CSTRING(244)                                     
lpPrefixString      CSTRING(4)                                       
wUnique             ULONG                                            
lpTempFilename      CSTRING(259)                                     
nBufferLength       ULONG                                            
TotalLength         ULONG                                            
                                                                     
  CODE                                                               
  TempFilename   = ''                                                
  lpPathName     = 0                                                 
  lpPrefixString = 'JSCSV'                                              
  wUnique        = 0                                                 
  nBufferLength  = 243                                               
  TotalLength = JSCSVGetTempPath(nBufferLength,lpPathName)                  
  IF NOT TotalLength or TotalLength > 243                            
    lpPathName = '.'                                                 
  END                                                                
  IF JSCSVGetTempFileName(lpPathName,lpPrefixString,wUnique,lpTempFilename)     
    TempFilename = lpTempFilename                                    
  END                                                                
  RETURN(LONGPATH(CLIP(TempFilename)))
            
!-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
JSCSVDetectLineEnding        PROCEDURE(*STRING pBuffer,LONG pMaxBytes=0FFFFh,<STRING pDefault>)!,STRING
!-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Ndx1       USHORT,AUTO
Terminator CSTRING(11)
BufferMax  LONG

  CODE

  IF NOT OMITTED(pDefault)    
    Terminator = pDefault
  END
  BufferMax = pMaxBytes
  IF SIZE(pBuffer) < pMaxBytes !Max record size in basic driver is 65520 bytes, so we'll test up to 0FFFFh by default - Not sure if this is a Clarion RECORD limit, or limit within the record text between line endings.
    BufferMax = SIZE(pBuffer)    
  END
  
  LOOP Ndx1 = 1 TO BufferMax-1
    CASE pBuffer[Ndx1] 
    OF '<10>'
      Terminator = '<10>'
      BREAK
    OF '<13>'
      IF pBuffer[Ndx1+1] = '<10>'
        Terminator = '<13,10>'
      ELSE
        Terminator = '<13>'
      END
      BREAK
    END
  END
  RETURN Terminator
            