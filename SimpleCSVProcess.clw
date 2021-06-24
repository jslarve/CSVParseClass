
  PROGRAM

  INCLUDE('JSCSVParseClass.inc'),ONCE 

OMIT('***')
 * Created with Clarion 11.1
 * User: Jeff Slarve
 * Date: 6/18/2021
 * Time: 10:15 PM
 * 
 * Simple Example to fill your own queue
 ***

MyQueue       QUEUE  
id              STRING(4)
first_name      STRING(13)
last_name       STRING(18)
email           STRING(39)
meds            STRING(114) !This column is in a completely different position in the example CSV than the others, so we can test the GetColumnNumber() method.
              END

Columns       GROUP ! Column numbers as detected within CSV file
id              SHORT
first_name      SHORT
last_name       SHORT
email           SHORT
meds            SHORT
              END


Window WINDOW('Loading your own queue from CSV'),AT(,,557,224),CENTER,GRAY,SYSTEM, |
      FONT('Segoe UI',9)
    LIST,AT(2,2,553,220),USE(?LIST1),VSCROLL,FROM(MyQueue),FORMAT('26L(2)|M~ID~C' & |
        '(2)52L(2)|M~First Name~C(2)53L(2)|M~Last Name~C(2)115L(2)|M~Email~C(2)2' & |
        '0L(2)|M~Meds~C(2)')
  END


CSV  JSCSVParseClass
Ndx  LONG


  MAP
  END

  CODE

  IF NOT CSV.LoadFile('.\SampleData\CSVdemo2.Comma.CRLF.csv')
    MESSAGE('Did not load file.')
    RETURN
  END

  !This code attempts to locate a column by name and set up the column numbers. 
  !That way (if you get the column name correct), you don't have to know the column number ahead of time.
  !If you already know the column number, or there's no header record, you don't need this part.
  Columns.id         = CSV.GetColumnNumber('id')
  Columns.first_name = CSV.GetColumnNumber('first_name')
  Columns.last_name  = CSV.GetColumnNumber('last*') !using just a part of the label here to test the wildcard search
  Columns.email      = CSV.GetColumnNumber('email') 
  Columns.meds       = CSV.GetColumnNumber('meds')  
   
  !Loop through the "virtual grid" and fill the queue.  
  LOOP Ndx = 1 TO CSV.GetRowCount()
    MyQueue.ID         = CSV.GetCellValue(Ndx,Columns.id) !Fetching a value, like a grid
    MyQueue.first_name = CSV.GetCellValue(Ndx,Columns.first_name)
    MyQueue.last_name  = CSV.GetCellValue(Ndx,Columns.last_name)
    MyQueue.email      = CSV.GetCellValue(Ndx,Columns.email)
    MyQueue.meds       = CSV.GetCellValue(Ndx,Columns.meds)
    ADD(MyQueue)
  END
  OPEN(Window)
  ACCEPT
  END
  
  