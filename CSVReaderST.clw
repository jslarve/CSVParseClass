  PROGRAM

!Example program using the StringTheory version of the loader.

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

  
  INCLUDE('JSCSVParseClassST.inc'),ONCE 

CSV        CLASS(JSCSVParseClassST)
TakeProgress PROCEDURE(LONG pProgressPct,LONG pProgress,LONG pRows),DERIVED
           END

OMIT('***')
 * Created with Clarion 11.0
 * User: Jeff Slarve
 * Date: 3/13/2021
 * Time: 7:27 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 ***

  MAP
  END

Progress1        BYTE
LineEnding       CSTRING(11)
Separator        CSTRING(11)
FirstRowIsLabels BYTE(TRUE)
SeparatorGroup   GROUP,PRE()
Comma              STRING(',')
Tab                STRING('<9>')
Colon              STRING(':')
SemiColon          STRING(';')
Pipe               STRING('|')
Space              STRING(' ')
AutoDetect         STRING('<1>')
                 END
Window WINDOW('CSV Parser Demo'),AT(,,707,268),CENTER,GRAY,IMM,MAX,FONT('Segoe UI',9), |
      RESIZE
    PROGRESS,AT(5,3,117,7),USE(PROGRESS1),RANGE(0,100)
    STRING('Progress Text'),AT(7,13,117),USE(?ProgressText)
    PROMPT('Separator:'),AT(210,11),USE(?PROMPT2)
    COMBO(@s20),AT(245,9,58,12),USE(Separator),TIP('If your desired separator is' & |
        ' not listed, enter CHR(YourASCIICode) (no quotes) OR ''YourCharacter'' ' & |
        '(in quotes).'),DROP(10),FROM('Comma|Tab|Colon|SemiColon|Pipe|Space|Auto' & |
        'Detect'),FORMAT('20L(2)|M')
    OPTION('LineEnding'),AT(307,2,160,23),USE(LineEnding),BOXED
      RADIO('Windows'),AT(310,12,39),USE(?LineEndingRADIO1),TIP('Windows Style -' & |
          ' AKA 0d0ah'),VALUE('<13,10>')
      RADIO('Mac'),AT(354,12,24),USE(?LineEndingRADIO2),TIP('Mac Style - AKA 0dh'), |
          VALUE('<13>')
      RADIO('UNIX'),AT(382,12,29),USE(?LineEndingRADIO3),TIP('UNIX Style - AKA 0ah'), |
          VALUE('<10>')
      RADIO('Auto Detect'),AT(415,12,46,10),USE(?LineEndingRADIO4),TIP('Auto det' & |
          'ect line endings'),VALUE('<0>')
    END
    CHECK('First Row is Labels'),AT(474,10,65),USE(FirstRowIsLabels),TIP('First ' & |
        'row of the CSV contains the label of each column.')
    BUTTON('Gen Clarion Structure'),AT(545,4,73,20),USE(?GenerateClarionStructureButton) |
        ,TIP('Generate a Clarion structure that can be compiled in another program.')
    BUTTON('&Open'),AT(621,4,26,20),USE(?OpenButton),TIP('Open a CSV file')
    BUTTON('&Close'),AT(679,4,26,20),USE(?CloseButton),STD(STD:Close), |
        TIP('Exit the application.')
    LIST,AT(4,29,701,237),USE(?CSVList),HVSCROLL,COLUMN,FROM('')
    BUTTON('&Reload'),AT(650,4,26,20),USE(?ReloadButton),TIP('Reload the current' & |
        ' CSV file from disk.')
  END
StartTime LONG
EndTime   LONG
CSVFile   STRING(FILE:MaxFilePath)

  CODE

  Separator = 'AutoDetect'
  BIND(SeparatorGroup)
  OPEN(Window)  
  DO SizeList
  StartTime = CLOCK()
  CSVFile = LONGPATH('.\SampleData\CSVdemo2.Comma.CRLF.csv')
  DO SetSpecs  
   ACCEPT
     CASE EVENT()
     OF EVENT:Sized
       DO SizeList
     OF EVENT:OpenWindow        
        SETCURSOR(CURSOR:Wait)
        CSV.SetListBox(?CSVList)
         IF CSV.LoadFile(CSVFile)
           HIDE(?PROGRESS1)
           HIDE(?ProgressText)
           EndTime = CLOCK()
           DO SetCaption
         ELSE
           MESSAGE('"' & CLIP(CSVFile) & '" was not loaded.')
         END
        SETCURSOR()
     OF EVENT:ACCEPTED
       CASE FIELD()
       OF ?FirstRowIsLabels OROF ?LineEnding OROF ?Separator
         DO SetSpecs
       OF ?OpenButton
         IF FILEDIALOG('Select CSV',CSVFile,'CSV Files|*.CSV|Text Files|*.TXT',FILE:KeepDir+FILE:LongName)
           DO LoadFile
         END
       OF ?ReloadButton
         DO LoadFile
       OF ?GenerateClarionStructureButton  
         CASE MESSAGE('What do you want to generate to your Windows clipboard for pasting into your app?','Generate Clarion Structure',ICON:Question,'File|Queue|Group|Cancel')
         OF 1
           SETCLIPBOARD(CSV.GenerateFileDef('MyFile',CSVFile))       
         OF 2
           SETCLIPBOARD(CSV.GenerateClarionStructure('MyQueue','QUEUE','My'))       
         OF 3
           SETCLIPBOARD(CSV.GenerateClarionStructure('MyGroup','GROUP','My'))       
         END  
       END
     END  
   END

LoadFile   ROUTINE

  IF NOT EXISTS(CSVFile)
    MESSAGE('File does not exist: "' & CLIP(CSVFile) & '"')
    EXIT
  END

  Progress1 = 0
  DISPLAY(?PROGRESS1)
  UNHIDE(?PROGRESS1)
  UNHIDE(?ProgressText)
  StartTime = CLOCK()
  DO SetSpecs
  CSV.LoadFile(CSVFile)
  EndTime = CLOCK()
  DO SetCaption
  HIDE(?PROGRESS1)
  HIDE(?ProgressText)
  SELECT(?CSVList)

SetCaption ROUTINE

  0{PROP:Text} =  '[ST] ' & CLIP(LEFT(FORMAT(CSV.GetBufferSize() / 1024,@n20))) & ' KBytes ' & CLIP(LEFT(FORMAT(CSV.GetRowCount(),@n20))) & ' rows  ' &|
                  CLIP(LEFT(FORMAT(CSV.GetColumnCount(),@n20))) & ' columns in ' & (EndTime - StartTime) * .01 & ' seconds. [' & CLIP(CSVFile) & '][' &|
                  CSV.GetDataName(1) & '][' & CSV.GetDataName(2) & ']'

SetSpecs   ROUTINE

  CSV.SetFileSpecs(EVALUATE(Separator),LineEnding,FirstRowIsLabels)
     
SizeList   ROUTINE     

  ?CSVList{PROP:Height} = 0{PROP:Height} - ?CSVList{PROP:Ypos} - 4
  ?CSVList{PROP:Width}  = 0{PROP:Width}  - ?CSVList{PROP:Xpos} - 4

CSV.TakeProgress      PROCEDURE(LONG pProgressPct,LONG pProgress,LONG pRows)!,DERIVED

  CODE
   
  Progress1 = pProgressPct
  ?ProgressText{PROP:Text} = pProgress & ' of ' & pRows
  DISPLAY(?Progress1)