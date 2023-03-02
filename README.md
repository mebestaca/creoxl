   	 _____ _______ ______ ______ __ ___ __		     ___      ______
   	|     |  __   |	  ___|	    |   |  |  |       __ __ |_  |    |___   |
   	|   --|	     -|   ___|	|   |-    -|  |__    |  |  | _| |_ __ 	 |  |
  	|_____|___|___|______|______|__|___|_____|    \___/ |_____|__|	 |__|



  ## BRIEF:
	This package can programmatically produce a formatted excel file and also
	import contents of excel file into a block or table.
  
  ## REQUIREMENTS:
	WPSOffice or Excel and set either as default program.

  ## HOW TO IMPORT:
 	Copy the following to your project:
	1. CREOXL (Package Spec)
	2. CREOXL (Package Body)
	3. IMPORT_CUSTOM_BLOCK (Implementation is found in comments)
	4. IMPORT_CUSTOM_TABLE (Implementation is found in comments)

  ## HOW TO USE:

  ## WHEN GENERATING EXCEL FILE:
  	1. Start with CREOXL.INITIALIZE;
  	2. Format your excel file using the available functions within this library
	     to fit your needs. 
  	3. Use CREOXL.DISPLAY_OUTPUT; or SAVEAS('YourPath')
	     NOTE: The former will open an untitled formatted excel file
	     and the latter saves the formatted file to the specified path 
  	4. End with CREOXL.TERMINATE; 
	     NOTE: This will release all the handles and prevent memory leaks

  ## WHEN IMPORTING:
  	1. You can use IMPORT_TO_BLOCK or IMPORT_TO_TABLE
	     NOTE: 
 	     When using IMPORT_TO_BLOCK use IMPORT_CUSTOM_BLOCK function
	     to assign values to your block. DO NOT edit the functions within
	     this library unless you know what you are doing.
	     When using IMPORT_TO_TABLE use IMPORT_CUSTOM_TABLE function
	     to map fields and values.
    2. You can write your own custom importing function by using the following
	     guide:
     
		   IMPORT_INITIALIZE
		   --
		   write your loop here and use GET_TEXT to fetch the data
		    --
	     IMPORT_TERMINATE

  ## CHANGE LOGS:
  	v1.0 Initial release
  	v1.1 Now supports importing
	       changed RELEASE_OBJECTS to TERMINATE
	       added INITIALIZE_IMPORT
	       added INITIALIZE_TERMINATE
	       added GET_TEXT
	  v1.2 added SAVEAS function
	  v1.3 added SET_ORIENTATION function
	  v1.4 added SET_IMAGE function
	  v1.5 added IMPORT_TO_BLOCK function
	  v1.6 added IMPORT_TO_TABLE function
	       changed import implementation
	  v1.7 exposed BORDER_WDITH
