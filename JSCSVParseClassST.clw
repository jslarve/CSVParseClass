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
   
   INCLUDE('JSCSVParseClassST.inc'),ONCE    
   
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Automatic constructor</summary>
!======================================================================================================================================================
JSCSVParseClassST.Construct PROCEDURE

  CODE
  
  SELF.ST         &= NEW StringTheory
  
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Automatic destructor</summary>
!======================================================================================================================================================
JSCSVParseClassST.Destruct      PROCEDURE

  CODE
  
  DISPOSE(SELF.ST)
   
!------------------------------------------------------------------------------------------------------------------------------------------------------
!!! <summary>Load a Delimiter Separated File</summary>
!!! <param name="pFileName">Name of the Delimiter Separated File</param>
!!! <returns>Bytes loaded</returns>
!======================================================================================================================================================
JSCSVParseClassST.LoadFile     PROCEDURE(STRING pFileName)!,LONG          !Load a file
  CODE

  IF NOT EXISTS(CLIP(pFileName))
    RETURN 0
  END
  SELF.SetLoading  
  IF NOT SELF.ST.LoadFile(pFileName)
    RETURN 0
  END
  SELF.LoadBuffer(SELF.ST.GetValuePtr())
  RETURN SELF.ST.Length()
   