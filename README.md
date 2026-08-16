Copyright 2022 xiangwuxw
 
MIT License. Feel free to modify and reuse. 
 
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


  
    ISSSUE RESOLVED IN THIS SCRIPT: WIM FILE > 4G and UEFI Boot. 
 
 1. Though some UEFI implementation may support NTFS, most of the them only support FAT/FAT32, 
 2. FAT32 max file size is 4G, 
 3. Windows WIM file is too big, even the install.wim on retail DVD may > 4G
 4. Most captured WIM file are super large, sometime > 10G
 
    
    SOLUTION
 
 1. Create two partitions in the USB disk, first FAT32 for UEFI Boot, boot.wim to load the WINPE or Setup Environment
 2. Windows Setup will look for the install.wim on all volumes, so put the install.wim on the second NTFS partition
     In this script, set the FAT32 active, so no matter UEFI or Legacy will boot from the same code/file FAT32 path to make it easier to maintain the code.
