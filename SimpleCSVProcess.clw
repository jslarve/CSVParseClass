
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

MyQueue                              QUEUE  
id                                     STRING(4)
first_name                             STRING(13)
last_name                              STRING(18)
email                                  STRING(39)
                                     END

Window WINDOW('Loading a queue from CSV'),AT(,,315,224),CENTER,GRAY,FONT('Segoe UI',9),SYSTEM
    LIST,AT(2,2,311,220),USE(?LIST1),FROM(MyQueue),FORMAT('26L(2)|M~ID~C(2)52L(2' & |
        ')|M~First Name~C(2)53L(2)|M~Last Name~C(2)20L(2)|M~Email~C(2)'),VSCROLL
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
  LOOP Ndx = 1 TO CSV.GetRowCount()
    MyQueue.ID         = CSV.GetCellValue(Ndx,1) 
    MyQueue.first_name = CSV.GetCellValue(Ndx,2)
    MyQueue.last_name  = CSV.GetCellValue(Ndx,3)
    MyQueue.email      = CSV.GetCellValue(Ndx,4)
    ADD(MyQueue)
  END
  OPEN(Window)
  ACCEPT
  END
  
  